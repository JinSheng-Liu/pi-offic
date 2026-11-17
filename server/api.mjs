import {structuralTargets,auditStructures} from './document/structural-audit.mjs'
import {EditorRequests} from './editor-requests.mjs'
import {prepareNativeToc,validateNativeToc} from './document/native-toc.mjs'
import {prepareThreeLineTable,applyThreeLineTableOOXML,validateThreeLineTable} from './document/native-table-borders.mjs'
import {hasNativeToc} from './document/toc-layout.mjs'
import {diagnosticFetch} from './model-transport.mjs'
import {attachPiSession} from './pi-session.mjs'
import {createChecklist,resumeChecklist,updateChecklist,runToTaskBoundary} from './task-checklist.mjs'
import {randomUUID} from 'node:crypto'
import {auditPlan} from './document/plan-audit.mjs'
import {executeFormatPlan} from './document/format-plan.mjs'
import {createInspectionReader} from './document/inspection-cache.mjs'
import {auditFrontMatter} from './document/front-matter-audit.mjs'
import {saveTaskResult,readTaskResult,taskResult} from './task-result.mjs'
import { createServer } from 'node:http'
import { readFile, writeFile, appendFile, mkdir } from 'node:fs/promises'
import { extname, join, resolve, sep } from 'node:path'
import { Type } from '@earendil-works/pi-ai'
import { Agent } from '@earendil-works/pi-agent-core'
import { streamSimple } from '@earendil-works/pi-ai/api/openai-completions'
import { SessionManager } from '../vendor/pi/packages/coding-agent/src/core/session-manager.ts'
import { loadSkillsFromDir } from '../vendor/pi/packages/coding-agent/src/core/skills.ts'
import { DocumentStore } from './store.mjs'
import { inspect, problem } from './engine.mjs'
import { validateOperations } from './document/rules.mjs'
import { projectInspection } from './document/skill-routing.mjs'
import { budgetDocumentContext, inspectModelPayload } from './document/read-context.mjs'

const api = '/api'
const root = resolve(process.env.PI_DOCX_DATA_DIR || 'data/documents')
const store = new DocumentStore(root)
const editorRequests = new EditorRequests()
const sessions = new Map()
const initializing = new Map()
const activeTasks = new Set()
const taskPlans = new Map()
const planListeners = new Map()
const taskChecklists = new Map()
const modelSettings = {
  apiKey: process.env.OPENAI_API_KEY || '',
  baseUrl: process.env.PI_API_BASE_URL || 'https://api.openai.com/v1',
  model: process.env.PI_MODEL || 'gpt-4o-mini',
}
const extensionByType = { '.js': 'text/javascript', '.css': 'text/css', '.html': 'text/html', '.svg': 'image/svg+xml', '.wasm': 'application/wasm' }
const json = (res, status, value) => { res.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' }); res.end(JSON.stringify(value)) }
const sendEvent = (res, event, data) => res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`)

async function readBody(req, limit) {
  const chunks = []
  let size = 0
  for await (const chunk of req) {
    size += chunk.length
    if (size > limit) throw problem('请求内容超过大小限制', 413)
    chunks.push(chunk)
  }
  return Buffer.concat(chunks)
}
async function readJson(req) {
  const body = await readBody(req, 64 * 1024)
  try { return JSON.parse(body.toString('utf8')) } catch { throw problem('请求 JSON 无效') }
}
const currentRevision = req => Number(req.headers['x-document-revision'])
const privateMetadata = m => ({ ...m, versions: m.versions.map(({ file, ...v }) => v) })
const urlFor = (id, suffix = '') => `${api}/documents/${id}${suffix}`

async function getAgent(id) {
  if (sessions.has(id)) return sessions.get(id)
  if (initializing.has(id)) return initializing.get(id)
  const pending = createAgent(id)
  initializing.set(id, pending)
  try { return await pending } finally { initializing.delete(id) }
}

async function createAgent(id) {
  const existing = sessions.get(id)
  if (existing) return existing
  const skillsDir = resolve('docs/skills')
  const discovered = loadSkillsFromDir({ dir: skillsDir, source: 'project' })
  if (discovered.diagnostics.length) throw problem('DOCX 技能文件无法载入', 500)
  const skills = discovered.skills
  const coreSkills = await Promise.all(skills.filter(s => ['docx-formatting', 'docx-validation'].includes(s.name))
    .map(async s => ({ name: s.name, content: await readFile(join(s.baseDir, 'SKILL.md'), 'utf8') })))
  const systemPrompt = [
    '你是 Pi DOCX Studio 的格式处理助手。使用中文。只围绕当前文档工作。',
    '仅通过已提供的文档工具读取和修改文档。绝不输出、生成或执行 XML、文件路径、Shell 命令。',
    '读取结果可能因上下文预算省略。遇到omittedForContextBudget必须定点重读当前版本，不得把省略当作空文档。长文档逐区域读取和执行，避免先累计全篇详细格式；写入后旧索引须重新确认。',
    '整段角色排版优先使用scope=namedStyle，明确indices、styleName和font/paragraph；字体间距写入命名样式而非只改外观。章节须明确outlineLevel，非章节设置10；段内标题使用scope=characterStyle，提供styleName、单段indices、准确text和font，保留正文段落样式；普通局部强调用scope=text。命名样式每步独立提交，可放入execute_format_plan。',
    '全文排版先把用户每项要求列为执行清单，再优先使用execute_format_plan按章节、表格等完整小步骤连续执行。每步保存后实时推送；不能只改几段样本就结束。失败或阻塞必须说明剩余项目，已提交步骤不要重复。布局/目录页码未验证不能声称完成。',
    '先读取用户明确涉及的文档范围，再用 update_task_checklist 创建本次计划，不预设论文八大分类；不添加用户未要求的修改。计划条目描述具体目标和检查方式，不把整个文档的修改与视觉验收合并成一个项目。仅因未检查不能标blocked：已执行待回读用awaitingVerification，只有真实工具失败才能用blocked，并提供工具结果中的evidenceId及具体target；局部失败不覆盖其他独立区域。需要用户决定时用needsDecision并提出具体question。修改任务先调用 update_task_checklist 读取或补充用户要求清单。清单必须覆盖全部明确要求；已有项目不可删除或改名。每个范围处理后使用 addressed 并提供当前revision和实际操作/检查依据；addressed只表示该范围已处理，不代表视觉通过。blocked必须引用toolFailures中的真实调用ID并解释；原文不含该范围用notApplicable并说明读取依据。局部受阻不阻断其他区域，所有pending都必须处理或说明阻塞后才能总结。',
    '自然语言格式指令明确时直接执行，不重复询问。真正无法判断范围时先询问。',
    '用户明确要求校正手工标题序号时，独立调用apply_formatting的renumberHeadings：仅支持Heading1–9或直接大纲层级且有阿拉伯数字前缀的简单标题，不能处理中文数字、复杂域或原生列表。独立显示OMML公式可用numberEquations添加SEQ编号，format支持arabic/upperRoman/lowerRoman/upperLetter/lowerLetter，prefix/suffix默认括号；不识别图片公式，不覆盖已有编号，不支持按章重启。结构不明确时报告限制。两项操作完成后重新读取文档，不能宣称页面视觉已验证。',
    '用户明确要求全文底端居中阿拉伯页码时使用scope=ensurePageNumbers独立提交：保留已有页脚内容与编号起点，空白页脚补PAGE域，现有页码居中，分节编号格式decimal；复杂页脚拒绝需报告。',
    '表格适应版心使用scope=fitTableWidth和tableIndex独立提交，按所在分节缩放列宽；全文表格文字统一使用scope=tableText，包含嵌套单元格；纯嵌入图片段落对齐使用scope=inlinePictures，混排文字和浮动图片另行检查。不要只处理分页读取的样本。',
    '修改前读取结构和相关格式。操作必须声明范围与 revision。不得改写正文，不得推断不确定的标题。',
    '用户要求已有题注位置与格式规范时，先调用normalize_caption_layout落实安全文本框转换和字体；受阻文本框逐项报告，不阻塞其他格式工作。',
    '用户要求生成或规范正式目录，完成标题层级和正文布局后必须调用update_native_toc；手工目录replace，已有TOC用refresh。不能以普通文字排版代替目录生成。normalize_toc_layout仅用于已有目录局部格式：分离混合标题并设置原生目录页码右制表位和点引导线；保留原页码，不声称已重算。普通多列Tab段落不能靠字号或缩进解决独立换行；用户要求整理全文且连续等列段落明确表达同一张表、未禁止结构变化时，用convert_tabular_text转表格。不要转换目录、普通首行Tab或语义不明内容；转换后重读索引。',
    '图表题注优先使用 add_captions 工具：从 captionTargets 选择唯一的正文顶层图/表目标，figure 插在图后、table 插在表前；仅当一级标题明确写出“第X章”时使用章节编号。图表已有题注不得重复添加；已有同章题注位于目标之后时停止并说明目前不能安全重排静态编号。已有题注格式使用段落工具；手工编号转SEQ使用convertCaptions，完整选同组连续编号，必须报告静态章号和未验证关联，不重复新增。',
    '设置首行或左缩进前检查 leadingWhitespace 和 indentWarning。原有制表符会与段落缩进共同影响实际文字起点，不能仅因 firstLineIndent 等于目标值就声称缩进已验收；不要一律清零标题、列表缩进或删除制表符。遇到冲突按具体段落检查，无法验证时明确报告。',
    '同级标题可能混用前导 Tab 与段落缩进，必须比较实际起点。保留单个前导 Tab 时，可按制表位验证悬挂首行（例如 leftIndent:21、firstLineIndent:-21），不能机械对所有文档套用或忽略自定义制表位。中英分别指定字体时，同时检查 font.asciiDirect、hAnsiDirect 和 eastAsiaDirect；font.name 不能代表所有文字的字体。',
    '左侧大纲由编辑器只读获取，不要为了显示导航主动修改文档。用户明确要求规范标题层级或标题样式时，才对确定的独立章节设置 Heading1–Heading9，并落实用户指定的字体字号等格式。不要仅凭颜色或字号判定标题，不要把与正文混排的整段设为标题。',
    '用户要求段内标题与正文不同字体时，使用 scope=text、一个段落索引及准确 text；重复文字必须指定从0开始的 occurrence。不得把整段正文误设为标题字体。',
    '排版应遵守用户指定字号、行距和边距，不以缩小字号填满页面。明确的独立标题可设置 keepNext=true、keepLines=true，正文通常不应全部 keepLines=true。根据已读取的属性处理不必要的 pageBreakBefore，保留用户明确要求的分页。不擅自增删文字或加入换行。',
    'body 范围仅排除已知特殊对象，不能可靠识别 Normal 样式的封面、签署区或目录。复杂文档必须使用已检查的明确 indices；不把 unclassified-text 当作已确认正文。',
    '结构或目标操作校验通过不等于需求全部完成，更不等于页面视觉验收。最终按用户要求逐项报告已完成、未完成和未验证；未渲染实际页面时禁止声称分页、断行或留白已经验收。',
    '仅当 apply_formatting 返回 INVALID_FORMAT_ARGUMENTS 且 retryable=true、committed=false 时，可按 correction 在相同 revision 修正参数并重试一次；分节操作中的段落格式必须按用户明确的段落范围另列操作，不能猜测范围。其他内容保护或保存校验失败必须停止。revision conflict 必须停止，不能换 revision 重试。已成功提交的操作绝不重复，根据工具真实结果报告成功、部分成功、未完成。',
    '结构读取的layoutRisks必须纳入需求覆盖说明：floatingObjects是只读浮动对象索引，不是段落indices。浮动文本框题注不在普通段落修改范围，嵌入图片居中不代表所有图表排版完成。风险信号不等于已观察到页面碰撞。',
    '先使用 inspect_document 获取结构及推荐技能；根据文档类型和用户任务调用 load_document_skill 读取相关技能，再执行操作。建议不等于确定分类；混合文档按区域使用技能。',
    'inspect_document 默认只返回40项精简结构；先读取coverage区域和blockedTextboxes，按用户范围列出待处理项；任务结束时逐项核对，未识别/未验证不能报完成。execute_format_plan 和 validate_document 返回的 verifiedTargets、failedTargets 是当前版本的程序核验证据：已通过目标不要逐段重复读取，只有失败目标、范围不明目标或结构变化影响的目标才用indices定点读取完整格式。lineSpacing提供规则及倍数/磅值，tabStops提供制表位信息。使用 pagination.nextOffset 继续读取；summary 不包含全部段落，禁止把样本当全文。figures/captionTables 的 index 仅给题注工具使用。',
    ...skills.map(s => `可按需加载技能 ${s.name}: ${s.description}`),
    ...coreSkills.map(s => `\n<skill name="${s.name}">\n${s.content}\n</skill>`),
  ].join('\n')
  const agentDir = join(root, id, 'agent')
  await mkdir(agentDir, { recursive: true })
  const modelId = modelSettings.model
  const model = {
    id: modelId, name: modelId, api: 'openai-completions', provider: 'openai',
    baseUrl: modelSettings.baseUrl,
    reasoning: false, input: ['text'],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow: 128000, maxTokens: 4096,
  }
  const documentTools = makeTools(id, skills)
  const sessionManager = SessionManager.continueRecent(agentDir, join(agentDir, 'sessions'))
  const history = sessionManager.buildSessionContext().messages
  const transport=diagnosticFetch(async event=>{const folder=join(agentDir,'diagnostics');await mkdir(folder,{recursive:true});await appendFile(join(folder,'transport.jsonl'),JSON.stringify(event)+'\n','utf8')})
  const agent = new Agent({ initialState: { model, messages: history },
    transformContext: budgetDocumentContext,
    streamFn: (m, context, options) => streamSimple(m, context, { ...options, apiKey: modelSettings.apiKey, maxTokens: 4096, maxRetries:2, maxRetryDelayMs:10000, fetch:transport, timeoutMs:120000,
      onPayload: async payload => {
        const stats=inspectModelPayload(payload)
        if(options?.sessionId===id&&stats.toolCount===0)throw new Error('文档会话工具声明缺失，已停止发送；已保存版本保留，请重新加载会话。')
        try { const folder=join(agentDir,'diagnostics');await mkdir(folder,{recursive:true})
          await writeFile(join(folder,'last-request.json'),JSON.stringify({...stats,time:new Date().toISOString()}),'utf8')
        } catch { /* Diagnostic disk failures must not interrupt model requests. */ }
      } }),
    toolExecution: 'sequential', sessionId: id })
  const runtime=await attachPiSession({agent,sessionManager,agentDir,systemPrompt,tools:documentTools,apiKey:modelSettings.apiKey})
  const entry = { agent, runtime, sessionManager, active: false }
  sessions.set(id, entry)
  return entry
}

function makeTools(id, skills) {
  const readInspection=createInspectionReader(inspect)
  const textResult = value => ({ content: [{ type: 'text', text: JSON.stringify(value) }], details: value })
  const documentTools = [
    {name:'audit_reference_front_matter',label:'检查前置页参考差异',description:'用户指定参考文档后，处理前及处理后检查封面、授权书、原创性声明的覆盖和格式差异。仅为结构属性检查，不代表视觉通过。必须使用真实已上传参考ID及版本。',parameters:Type.Object({referenceId:Type.String(),referenceRevision:Type.Integer({minimum:1})},{additionalProperties:false}),execute:async(_call,args)=>{const meta=await store.metadata(id);return textResult({revision:meta.revision,...auditFrontMatter(await store.bytes(id,meta.revision),await store.bytes(args.referenceId,args.referenceRevision))})}},
    {name:'restore_reference_declaration',label:'按参考文档恢复原创性声明',description:'仅用户明确指定参考文档时使用。referenceId与referenceRevision必须是已上传参考文档的真实ID和版本，不可猜测。仅恢复文字相同、样式和图片依赖兼容的以标题和分节定位的原创性声明，保留当前正文和分节；不支持不同姓名或题目的字段替换。执行后重新读取全部索引。',executionMode:'sequential',
      parameters:Type.Object({revision:Type.Integer({minimum:1}),referenceId:Type.String(),referenceRevision:Type.Integer({minimum:1})},{additionalProperties:false}),
      execute:async(_call,args,signal)=>{const result=await store.workflow.execute(id,args.revision,{kind:'referenceDeclaration',referenceId:args.referenceId,referenceRevision:args.referenceRevision},signal);return textResult({revision:result.document.revision,validation:result.validation})}},
    {name:'restore_reference_authorization',label:'按参考文档恢复授权书',description:'仅用户明确指定参考文档时使用。referenceId与referenceRevision必须是已上传参考文档的真实ID和版本，不可猜测。仅恢复文字相同、样式和图片依赖兼容的以标题和分节定位的授权书，保留当前正文和分节；不支持不同姓名或题目的字段替换。执行后重新读取全部索引。',executionMode:'sequential',
      parameters:Type.Object({revision:Type.Integer({minimum:1}),referenceId:Type.String(),referenceRevision:Type.Integer({minimum:1})},{additionalProperties:false}),
      execute:async(_call,args,signal)=>{const result=await store.workflow.execute(id,args.revision,{kind:'referenceAuthorization',referenceId:args.referenceId,referenceRevision:args.referenceRevision},signal);return textResult({revision:result.document.revision,validation:result.validation})}},
    {name:'restore_reference_cover',label:'按参考文档恢复封面',description:'仅用户明确指定参考文档时使用。referenceId与referenceRevision必须是已上传参考文档的真实ID和版本，不可猜测。仅恢复文字相同、样式和图片依赖兼容的首节封面，保留当前正文和分节；不支持不同姓名或题目的字段替换。执行后重新读取全部索引。',executionMode:'sequential',
      parameters:Type.Object({revision:Type.Integer({minimum:1}),referenceId:Type.String(),referenceRevision:Type.Integer({minimum:1})},{additionalProperties:false}),
      execute:async(_call,args,signal)=>{const result=await store.workflow.execute(id,args.revision,{kind:'referenceCover',referenceId:args.referenceId,referenceRevision:args.referenceRevision},signal);return textResult({revision:result.document.revision,validation:result.validation})}},
    { name: 'load_document_skill', label: '加载排版技能', description: '按当前文档类型和任务读取已注册的中文排版技能。只接受技能目录中的名称，不接受路径。',
      parameters: Type.Object({ name: Type.String() }, { additionalProperties: false }), execute: async (_call, args) => {
        const skill = skills.find(s => s.name === args.name)
        if (!skill) throw problem('未知技能，请使用系统提示中的技能名称')
        return textResult({ name: skill.name, content: await readFile(join(skill.baseDir, 'SKILL.md'), 'utf8') })
      } },
    { name: 'inspect_document', label: '读取文档结构', description: '分页读取结构；默认40段，最多80项。默认精简结构，不含格式；指定indices默认返回完整格式，也可detail=formatting。summary 返回统计和技能建议；paragraphs 返回可用于格式工具的真实索引；cells 返回单元格与 paragraphIndices，可用于分列对齐；floatingObjects 返回只读浮动对象，objectIndex不能用于段落操作；figures/captionTables 的索引仅用于题注工具。必须跟随 nextOffset 才能检查完整文档。',
      parameters: Type.Object({
        detail: Type.Optional(Type.Union([Type.Literal('structure'), Type.Literal('formatting')])),
        view: Type.Optional(Type.Union(['summary','paragraphs','tables','cells','floatingObjects','sections','captions','figures','captionTables'].map(v => Type.Literal(v)))),
        offset: Type.Optional(Type.Integer({ minimum: 0 })), limit: Type.Optional(Type.Integer({ minimum: 1, maximum: 80 })),
        indices: Type.Optional(Type.Array(Type.Integer({ minimum: 0 }), { minItems: 1, maxItems: 80 })),
      }, { additionalProperties: false }), execute: async (_call, args) => {
        const meta = await store.metadata(id), bytes = await store.bytes(id,meta.revision)
        return textResult({ revision: meta.revision, name: meta.name, ...projectInspection(await readInspection(bytes), args) })
      } },
    { name: 'analyze_formatting', label: '分析格式', description: '比较指定正文、标题或全文的格式一致性。',
      parameters: Type.Object({ scope: Type.Union([Type.Literal('body'), Type.Literal('headings'), Type.Literal('all')]),
        heading: Type.Optional(Type.Integer({ minimum: 1, maximum: 9 })) }), execute: async (_call, args) => {
        const data = await readInspection(await store.bytes(id))
        const ps = data.paragraphs.filter(p => args.scope === 'all' || (args.scope === 'body' ? p.role === 'unclassified-text' : p.heading > 0 && (!args.heading || p.heading === args.heading)))
        return textResult({ revision: (await store.metadata(id)).revision, count: ps.length,
          styles: summary(ps.map(p => p.style)), fonts: summary(ps.map(p => JSON.stringify(p.font))),
          alignments: summary(ps.map(p => p.paragraph.alignment)), example: ps.slice(0, 8), limitations: data.limitations })
      } },
    { name: 'add_captions', label: '添加图表题注', description: '为 inspect_document 返回的正文顶层嵌入式图片或表格添加 Word 原生 Caption 段落与 SEQ 编号域。图题放图后，表题放表前。targetIndex 必须使用 captionTargets.figures/tables 对应条目的 index 属性，不是数组序号。只支持章节号可从前置 Heading1 的“第X章”可靠识别，拒绝浮动图、重复题注及插入到已有题注中间。插入后重新读取结构再设置字体等格式，旧段落索引已失效。', executionMode: 'sequential',
      parameters: Type.Object({ revision: Type.Integer({minimum:1}), captions: Type.Array(Type.Object({
        kind: Type.Union([Type.Literal('figure'),Type.Literal('table')]),
        targetIndex: Type.Integer({minimum:0}), text: Type.String({minLength:1,maxLength:300}),
      }, {additionalProperties:false}), {minItems:1,maxItems:20}) }), execute: async (_call,args,signal) => {
        const result=await store.addCaptions(id,args.revision,args.captions,signal)
        return textResult({revision:result.document.revision,affected:result.affected,captions:result.captions,validation:result.validation})
      } },
    { name: 'apply_formatting', label: '应用格式', description: '对明确匹配的段落、段内文字或分节执行经过校验的格式操作。namedStyle独立提交indices、styleName与font/paragraph，定义并应用真实命名样式，其他属性和非目标段落保持；同名不同定义显示短后缀区分。convertCaptions必须单独提交indices，完整选择从1连续递增的同一图/表编号组，将普通手工编号原位转成SEQ域；保留原文，不自动验证对象关联，章节前缀仍为静态文字。每个 scope=sections 操作只含 sectionIndices 与 page、margins 中至少一种，允许同时设置 page 和 margins；不得包含 font、paragraph、style。段前距、段后距、字体、行距等必须放在独立的段落 scope 操作中。scope=text 需要单个 indices、准确 text 和 font，支持带制表符和手动换行的普通段落，不支持域或其他复杂内联结构，并创建版本。分节页面尺寸和边距以 twips 为单位。font.color 使用 #RRGGBB 六位十六进制字符串，例如红色 #FF0000、黑色 #000000；字号以 pt 为单位。', executionMode: 'sequential',
      parameters: Type.Object({ revision: Type.Integer({ minimum: 1 }), operations: Type.Array(Type.Object({
        scope: Type.Union([Type.Literal('all'), Type.Literal('body'), Type.Literal('tableText'), Type.Literal('inlinePictures'), Type.Literal('headings'), Type.Literal('indices'), Type.Literal('sections'), Type.Literal('text'),Type.Literal('tableRows'),Type.Literal('tableCells'),Type.Literal('floatingObjects'),Type.Literal('tocTitle'),Type.Literal('footerPageNumber'),Type.Literal('convertCaptions'),Type.Literal('coverLayout'),Type.Literal('namedStyle'),Type.Literal('characterStyle'),Type.Literal('fitTableWidth'),Type.Literal('ensurePageNumbers'),Type.Literal('renumberHeadings'),Type.Literal('numberEquations')]),
        format: Type.Optional(Type.Union(['arabic','upperRoman','lowerRoman','upperLetter','lowerLetter'].map(value => Type.Literal(value)), {description:'仅numberEquations：公式编号格式'})),
        prefix: Type.Optional(Type.String({maxLength:20,description:'仅numberEquations：编号前缀，默认左括号'})),
        suffix: Type.Optional(Type.String({maxLength:20,description:'仅numberEquations：编号后缀，默认右括号'})),
        cover: Type.Optional(Type.Object({widthPt:Type.Number({minimum:200,maximum:500}),labelWidthPt:Type.Number({minimum:40})},{additionalProperties:false,description:'仅coverLayout：明确授权封面结构重排；indices按顺序指定连续普通题名/信息段落，合并题名续段、清理定位空白，保留其他文字，生成无边框两列；完成后重读所有索引'})),
        alignment: Type.Optional(Type.Union(['Left','Centered','Right'].map(v=>Type.Literal(v)),{description:'仅footerPageNumber：独立提交，sectionIndices明确指定已有独立PAGE页码的分节；font只支持name和size，不能更改编号或目录页码'})),
        objectIndices: Type.Optional(Type.Array(Type.Integer({minimum:0}),{minItems:1,maxItems:150})),
        objectLayout: Type.Optional(Type.Union([Type.Literal('inlinePicture'),Type.Literal('captionInFlow')],{description:'对象进入文档流：只能独立提交一个操作，从floatingObjects读取可转换对象；完成后所有段落和对象索引必须重读'})),
        tableIndex: Type.Optional(Type.Integer({minimum:0})),
        rowIndices: Type.Optional(Type.Array(Type.Integer({minimum:0}),{minItems:1,maxItems:100})),
        cellTargets: Type.Optional(Type.Array(Type.Object({rowIndex:Type.Integer({minimum:0}),cellIndex:Type.Integer({minimum:0})},{additionalProperties:false}),{minItems:1,maxItems:100})),
        table: Type.Optional(Type.Object({repeatHeader:Type.Optional(Type.Boolean()),cantSplit:Type.Optional(Type.Boolean()),autoHeight:Type.Optional(Type.Boolean()),verticalAlignment:Type.Optional(Type.Union([Type.Literal('Top'),Type.Literal('Center'),Type.Literal('Bottom')]))},{additionalProperties:false})),
        indices: Type.Optional(Type.Array(Type.Integer({ minimum: 0 }))),
        text: Type.Optional(Type.String({minLength:1,maxLength:1000,description:'scope=text 时准确原文，仅修改该文字的字体'})),
        occurrence: Type.Optional(Type.Integer({minimum:0,description:'同段重复匹配的序号，从0开始；唯一匹配时可省略'})),
        heading: Type.Optional(Type.Integer({ minimum: 1, maximum: 9 })),
        sectionIndices: Type.Optional(Type.Array(Type.Integer({ minimum: 0 }))),
        page: Type.Optional(Type.Object({ widthTwips: Type.Optional(Type.Integer({ minimum: 1000, maximum: 31680 })),
          heightTwips: Type.Optional(Type.Integer({ minimum: 1000, maximum: 31680 })), orientation: Type.Optional(Type.Union([Type.Literal('portrait'), Type.Literal('landscape')])) })),
        margins: Type.Optional(Type.Object({ top: Type.Optional(Type.Integer({ minimum: 0, maximum: 31680 })),
          right: Type.Optional(Type.Integer({ minimum: 0, maximum: 31680 })), bottom: Type.Optional(Type.Integer({ minimum: 0, maximum: 31680 })),
          left: Type.Optional(Type.Integer({ minimum: 0, maximum: 31680 })), gutter: Type.Optional(Type.Integer({ minimum: 0, maximum: 31680 })) })),
        styleName: Type.Optional(Type.String({description:'scope=namedStyle或characterStyle时必填，例如公文·标题、论文·正文。单独提交，使用indices与font/paragraph定义可复用段落样式；只清理本次规定的格式，其他强调保留。'})),
        style: Type.Optional(Type.String()), font: Type.Optional(Type.Object({
          name: Type.Optional(Type.String({description:'统一字体；同时设置latinName时，西文使用latinName'})), latinName: Type.Optional(Type.String({description:'仅西文字体，保留中文字体；英文要求Time New Roman时使用实际名称Times New Roman'})), size: Type.Optional(Type.Number({ minimum: 1, maximum: 200, description: '字号单位 pt：二号22，三号16' })),
          color: Type.Optional(Type.String({ pattern: '^#?[0-9a-fA-F]{6}$', description: '红色 #FF0000，黑色 #000000' })),
          bold: Type.Optional(Type.Boolean()), italic: Type.Optional(Type.Boolean()),
          underline: Type.Optional(Type.Union(['None','Single','Double','Dotted','Dash','Wave'].map(value => Type.Literal(value)))),
          strikeThrough: Type.Optional(Type.Boolean()), doubleStrikeThrough: Type.Optional(Type.Boolean()),
          subscript: Type.Optional(Type.Boolean()), superscript: Type.Optional(Type.Boolean()),
        }, { additionalProperties: false })),
        paragraph: Type.Optional(Type.Object({
          keepNext: Type.Optional(Type.Boolean({description:'与下一段同页，适用于独立标题'})),
          keepLines: Type.Optional(Type.Boolean({description:'段中不分页；避免对全部长正文开启'})),
          widowControl: Type.Optional(Type.Boolean({description:'孤行控制'})),
          pageBreakBefore: Type.Optional(Type.Boolean({description:'段前分页；仅在确定需要时设置或清除'})),
          outlineLevel: Type.Optional(Type.Integer({ minimum: 1, maximum: 10, description: '导航标题层级：1–9；10为正文。仅修改语义层级，不更改字体字号。' })),
          alignment: Type.Optional(Type.Union(['Left','Centered','Right','Justified'].map(value => Type.Literal(value)), { description: '居中 Centered；两端对齐 Justified' })),
          tocPageTabPosition: Type.Optional(Type.Number({exclusiveMinimum:0,maximum:720,description:'目录条目右对齐点引导制表位，pt；明确indices且原段无制表位、单Tab后为页码；不重算页码'})),
          leadingTabPosition: Type.Optional(Type.Number({exclusiveMinimum:0,maximum:720,description:'仅明确indices、单前导Tab、无已有制表位/缩进/列表的普通段落可用；设Tab后文字起点pt，保留Tab，不能与缩进或style同操作，不等于firstLineIndent'})),
          leftIndent: Type.Optional(Type.Number()), rightIndent: Type.Optional(Type.Number()), firstLineIndent: Type.Optional(Type.Number()),
          spaceBefore: Type.Optional(Type.Number()), spaceAfter: Type.Optional(Type.Number()), lineSpacing: Type.Optional(Type.Number({exclusiveMinimum:0,maximum:720,description:'固定行距，单位pt，必须大于0；不可用0表示单倍'})),
          lineSpacingMultiple: Type.Optional(Type.Number({minimum:1,maximum:10,description:'单倍填1，1.5倍填1.5，双倍填2；不能与固定行距同时设置'})),
        }, { additionalProperties: false, description: '缩进和间距单位均为 pt' })),
      }), { minItems: 1, maxItems: 30 }) }), execute: async (_call, args, signal) => {
        const meta = await store.metadata(id)
        if (meta.revision !== args.revision) throw problem('文档版本已变化，请重新读取；不能自动换版本重试', 409)
        try { validateOperations(args.operations) } catch (error) {
          return { ...textResult({ code: 'INVALID_FORMAT_ARGUMENTS', committed: false, retryable: true,
            revision: meta.revision, correction: error.message }), isError: true }
        }
        const result = await store.apply(id, args.revision, args.operations, signal)
        return textResult({ revision: result.document.revision, affected: result.affected, validation: result.validation })
      } },
    { name: 'validate_document', label: '验证文档', description: '验证最近一次格式操作和当前 DOCX 包。',
      parameters: Type.Object({}), execute: async () => {
        const meta = await store.metadata(id), bytes = await store.bytes(id)
        const record = meta.versions.at(-1)
        return textResult({ revision: meta.revision, package: projectInspection(await readInspection(bytes), { view: 'summary' }), validation: record.validation,
          // Formatting plans are non-structural. A later formatting revision does
          // not invalidate earlier target evidence; re-audit the current bytes
          // and let the audit report the actual regression instead of forcing a
          // model to reread every old target.
          plans: await Promise.all((taskPlans.get(id)||[]).map(async plan=>({steps:plan.steps,audit:await auditPlan(bytes,plan.operations||[])}))),
          coverage: '仅代表包检查与最近提交记录；请按用户要求回读目标段落，页面视觉效果未验证。' })
      } },
  ]
  const formatting=documentTools.find(tool=>tool.name==='apply_formatting')
  const generalOperation=formatting.parameters.properties.operations.items
  generalOperation.properties.scope=Type.Union(generalOperation.properties.scope.anyOf.filter(value=>!['ensurePageNumbers','fitTableWidth'].includes(value.const)))
  formatting.parameters.properties.operations.items=Type.Union([generalOperation,Type.Object({scope:Type.Literal('ensurePageNumbers')},{additionalProperties:false}),Type.Object({scope:Type.Literal('fitTableWidth'),tableIndex:Type.Integer({minimum:0}),columnRatios:Type.Optional(Type.Array(Type.Number({exclusiveMinimum:0,maximum:100}),{minItems:2,maxItems:20}))},{additionalProperties:false})])
  documentTools.push({name:'execute_format_plan',label:'分批执行排版计划',description:'将已确认范围的格式要求组织为完整步骤清单并顺序执行，每步保存一版并推送预览。每步使用apply_formatting相同操作格式；结构转换独立处理后重读。失败即停止，返回全部待处理步骤；执行成功仍不代表全文或页面验收通过。',
    parameters:Type.Object({revision:Type.Integer({minimum:1}),steps:Type.Array(Type.Object({label:Type.String({minLength:1,maxLength:160}),operations:formatting.parameters.properties.operations},{additionalProperties:false}),{minItems:1,maxItems:100})},{additionalProperties:false}),
    execute:async(_call,args,signal)=>{
      const reports=taskPlans.get(id)||[];taskPlans.set(id,reports)
      let slot=reports.length
      const report=await executeFormatPlan(store,id,args.revision,args.steps,signal,progress=>{reports[slot]={...structuredClone(progress),operations:args.steps};planListeners.get(id)?.({planIndex:slot,...progress})})
      return {...textResult(report),...(report.state==='partial'?{isError:true}:{})}
    }})
  documentTools.push({name:'normalize_caption_layout',label:'规范已有题注布局',description:'用户要求已有图表题注规范排版时使用。只将结构与关联明确的浮动题注移入正文流，保留受阻对象；重读后按指定字体居中格式化独立题注候选。不新增题名、不改编号，不保证自动编号域。返回未处理对象，不能称全部题注完成。',
    parameters:Type.Object({revision:Type.Integer({minimum:1}),numbering:Type.Optional(Type.Union([Type.Literal('preserve'),Type.Literal('sequence')],{description:'仅用户要求真实自动编号时使用sequence；保留当前编号，断档或冲突拒绝。默认保留现有编号。'})),font:Type.Object({name:Type.String({minLength:1,maxLength:100}),latinName:Type.Optional(Type.String({minLength:1,maxLength:100})),size:Type.Number({minimum:1,maximum:200})},{additionalProperties:false})},{additionalProperties:false}),
    execute:async(_call,args,signal)=>{
      let revision=args.revision
      if((await store.metadata(id)).revision!==revision)throw problem('文档版本已变化',409)
      let data=await readInspection(await store.bytes(id,revision))
      const safe=data.floatingObjects.filter(o=>o.textBoxCount&&o.flowCapability.supported&&o.flowCapability.mode==='captionInFlow')
      if(safe.length){const result=await store.apply(id,revision,[{scope:'floatingObjects',objectIndices:safe.map(o=>o.objectIndex),objectLayout:'captionInFlow'}],signal);revision=result.document.revision;data=await readInspection(await store.bytes(id,revision))}
      const eligible=data.captions.filter(c=>c.kind!=='unknown'&&(c.associationVerified||c.style==='Caption'||c.numbering==='sequence-field'))
      const indices=eligible.map(c=>c.index)
      if(indices.length){
        const operations=['figure','table'].map(kind=>({kind,indices:eligible.filter(c=>c.kind===kind).map(c=>c.index)})).filter(group=>group.indices.length).map(group=>({scope:'indices',indices:group.indices,font:args.font,paragraph:{alignment:'Centered',firstLineIndent:0,leftIndent:0,rightIndent:0,keepLines:true,keepNext:group.kind==='table'}}))
        if(operations.length){const result=await store.apply(id,revision,operations,signal);revision=result.document.revision}
      }
      let numberingResult='preserved'
      if(args.numbering==='sequence'){
        const manual=eligible.filter(c=>c.numbering!=='sequence-field').map(c=>c.index)
        if(manual.length){
          try{const result=await store.apply(id,revision,[{scope:'convertCaptions',indices:manual}],signal);revision=result.document.revision;numberingResult='converted'}
          catch(error){if(error.status!==422)throw error;numberingResult=error.message}
        }
      }
      return textResult({revision,numberingResult,unconfirmedCandidates:data.captions.filter(c=>!indices.includes(c.index)).map(c=>({index:c.index,text:c.text})),converted:safe.length,formatted:indices.length,blocked:data.floatingObjects.filter(o=>o.textBoxCount&&!o.flowCapability.supported).map(o=>({objectIndex:o.objectIndex,reason:o.flowCapability.reason})),requirementsVerified:false,visualVerified:false})
    }})
  documentTools.push({name:'apply_three_line_table',label:'设置三线表',description:'使用当前 DOCX Editor 原生边框命令设置明确表格为三线表。tableIndex 必须来自当前 inspect_document.tables 的顶层表格索引；仅保留顶线、表头下线和底线，清除竖线及内部横线。执行后必须回读 DOCX 边框，验证失败不提交。',
    parameters:Type.Object({revision:Type.Integer({minimum:1}),tableIndex:Type.Integer({minimum:0})},{additionalProperties:false}),
    execute:async(_call,args,signal)=>store.exclusive(id,args.revision,async meta=>{
      signal?.throwIfAborted()
      const source=await store.bytes(id,args.revision)
      const prepared=prepareThreeLineTable(source,args.tableIndex)
      let output
      let method='DOCX Editor 原生边框命令'
      try {
        output=await editorRequests.request(id,{revision:args.revision,mode:'threeLineTable',paraId:prepared.paraId,tableIndex:args.tableIndex},prepared.bytes)
      } catch (error) {
        if (!/无法定位|无法选择目标表格|目标表格选择失败|编辑器尚未就绪/.test(error.message || '')) throw error
        output=applyThreeLineTableOOXML(prepared.bytes,args.tableIndex)
        method='OOXML 三线表结构化回退'
      }
      signal?.throwIfAborted()
      const validation=validateThreeLineTable(prepared.bytes,output,args.tableIndex)
      validation.checks=[...validation.checks,`处理方式：${method}`]
      const document=await store.commit(meta,output,'设置三线表',validation,undefined,store.taskContext.getStore())
      return textResult({revision:document.revision,tableIndex:args.tableIndex,method,validation,visualVerified:false})
    })}),
  documentTools.push({name:'update_native_toc',label:'生成或更新原生目录',description:'通过当前编辑器原生命令生成真正TOC目录。先完成正文标题和布局。已有原生目录用refresh；手工目录用replace并提供连续条目indices（不含目录标题），条目必须对应正文标题；没有目录用insert和paragraphIndex，在该段之前插入。保存后重新读取段落索引。不得用普通条目样式代替正式目录。',
    parameters:Type.Object({revision:Type.Integer({minimum:1}),mode:Type.Union([Type.Literal('insert'),Type.Literal('replace'),Type.Literal('refresh')]),indices:Type.Optional(Type.Array(Type.Integer({minimum:0}),{minItems:1,maxItems:100})),paragraphIndex:Type.Optional(Type.Integer({minimum:0}))},{additionalProperties:false}),
    execute:async(_call,args,signal)=>store.exclusive(id,args.revision,async meta=>{
      signal?.throwIfAborted()
      const source=await store.bytes(id,args.revision)
      const prepared=await prepareNativeToc(source,args)
      const output=await editorRequests.request(id,{revision:args.revision,mode:prepared.mode,paraId:prepared.paraId},prepared.bytes)
      signal?.throwIfAborted()
      const validation=validateNativeToc(prepared.bytes,output,prepared.headings)
      await inspect(output)
      const document=await store.commit(meta,output,'原生目录生成或更新',validation,undefined,store.taskContext.getStore())
      return textResult({revision:document.revision,nativeToc:true,pageNumbersUpdatedByEditor:true,validation,visualVerified:false})
    })})
  documentTools.push({name:'normalize_toc_layout',label:'规范目录布局',description:'分离混合目录标题并按指定字体居中；按所在分节版心设置原生目录条目右对齐页码和点引导线。保留目录域和文字，不重算页码；非标准条目需检查。',
    parameters:Type.Object({revision:Type.Integer({minimum:1}),font:Type.Object({name:Type.String({minLength:1,maxLength:100}),size:Type.Number({minimum:1,maximum:200}),bold:Type.Optional(Type.Boolean())},{additionalProperties:false})},{additionalProperties:false}),
    execute:async(_call,args,signal)=>{
      let revision=args.revision
      if((await store.metadata(id)).revision!==revision)throw problem('文档版本已变化',409)
      const tocBytes=await store.bytes(id,revision)
      if(!hasNativeToc(tocBytes))throw problem('未找到真实 TOC 目录域；普通文字目录不能按原生目录更新。需要通过编辑器原生目录能力生成，请调用update_native_toc完成生成或转换。',422)
      let data=await readInspection(tocBytes)
      const mixed=data.layoutRisks.mixedTocTitleIndices||[]
      if(mixed.length){
        const result=await store.apply(id,revision,[{scope:'tocTitle',indices:mixed}],signal)
        revision=result.document.revision
        data=await readInspection(await store.bytes(id,revision))
      }
      const indices=data.paragraphs.filter(p=>p.tableNestingLevel===0&&/^(目录|目\s+录|contents|table of contents)$/i.test(p.text.trim())).map(p=>p.index)
      if(indices.length){
        const result=await store.apply(id,revision,[{scope:'indices',indices,font:args.font,paragraph:{alignment:'Centered',firstLineIndent:0,leftIndent:0,rightIndent:0,keepNext:true,keepLines:true}}],signal)
        revision=result.document.revision
      }
      const entries=await store.apply(id,revision,[{scope:'tocEntries'}],signal)
      revision=entries.document.revision
      return textResult({revision,separated:mixed.length,formatted:indices.length,formattedEntries:entries.affected,validation:entries.validation,pageNumbersRecalculated:false,visualVerified:false,requirementsVerified:false})
    }})
  documentTools.push({name:'convert_tabular_text',label:'将分列文字转为表格',description:'用户要求整理全文/排版且已确认连续Tab段落表达同一张表时使用；明确禁止结构变化或含义不明时不得转换。2–100个连续顶层纯文字段落、每行2–8列且列数相同；保留文字和字符格式，转为真实表格，各列独立换行；识别出的分析表、指标表会在同一次提交中完成分类排版。拒绝目录、域、书签、图片等复杂内容。完成后重新读取段落索引。',
    parameters:Type.Object({revision:Type.Integer({minimum:1}),indices:Type.Array(Type.Integer({minimum:0}),{minItems:2,maxItems:100})},{additionalProperties:false}),
    execute:async(_call,args,signal)=>{
      const result=await store.apply(id,args.revision,[{scope:'tabularText',indices:args.indices}],signal)
      return textResult({revision:result.document.revision,affected:result.affected,validation:result.validation,requiresReread:true,visualVerified:false})
    }})
  documentTools.push({name:'update_task_checklist',label:'更新任务清单',description:'读取当前任务清单；传items增补或更新任务项。不得缩小用户范围；已处理需具体操作或检查依据，不可凭空声称验证通过。',
    parameters:Type.Object({items:Type.Optional(Type.Array(Type.Object({id:Type.String({minLength:1,maxLength:80}),label:Type.String({minLength:1,maxLength:160}),state:Type.Union(['pending','awaitingVerification','addressed','blocked','needsDecision','notApplicable'].map(v=>Type.Literal(v))),revision:Type.Optional(Type.Integer({minimum:1})),reason:Type.Optional(Type.String({maxLength:2000})),evidenceId:Type.Optional(Type.String({maxLength:160})),target:Type.Optional(Type.String({maxLength:300})),question:Type.Optional(Type.String({maxLength:500}))},{additionalProperties:false}),{minItems:1,maxItems:60}))},{additionalProperties:false}),
    execute:async(_call,args)=>{
      const active=taskChecklists.get(id)
      if(!active)throw problem('没有运行中的任务',409)
      const revision=(await store.metadata(id)).revision
      if(args.items) {
        const closing=args.items.filter(item=>['addressed','notApplicable'].includes(item.state)&&active.targets.some(target=>target.id===item.id))
        if(closing.length){
          const audit=await auditStructures(await store.bytes(id,revision),active.targets)
          if(closing.some(item=>!audit.find(check=>check.id===item.id)?.passed))throw problem('目录或表格结构验收未通过，不能用文字格式检查关闭该任务项',422)
        }
        updateChecklist(active.checklist,args.items,revision,active.evidence)
      }
      await active.persist()
      return textResult({revision,...active.checklist,toolFailures:active.evidence,requirementsVerified:false,visualVerified:false})
    }})
  return documentTools

}
const summary = items => Object.entries(items.reduce((out, item) => { out[item] = (out[item] || 0) + 1; return out }, {})).map(([value, count]) => ({ value, count }))

async function chat(req, res, id) {
  if (activeTasks.has(id) || store.locks.has(id)) throw problem('该文档已有任务运行中', 409)
  activeTasks.add(id)
  try { return await runChat(req, res, id) } finally { activeTasks.delete(id) }
}

async function runChat(req, res, id) {
  const input = await readJson(req)
  if (typeof input.message !== 'string' || !input.message.trim() || input.message.length > 4000) throw problem('消息需为 1–4000 个字符')
  if (!Number.isInteger(input.revision) || input.revision !== (await store.metadata(id)).revision) throw problem('文档版本已变化，请重新加载后重试', 409)
  if (!modelSettings.apiKey) throw problem('请先配置模型 API 密钥。', 503)
  const session = await getAgent(id)
  if (session.active) throw problem('该文档已有 AI 任务运行中', 409)
  session.active = true
  const abort = new AbortController()
  abort.signal.addEventListener('abort', () => { void session.runtime.abort() }, { once: true })
  const disconnect = () => abort.abort()
  req.on('aborted', disconnect)
  res.on('close', disconnect)
  session.cancel = disconnect
  res.writeHead(200, { 'content-type': 'text/event-stream; charset=utf-8', 'cache-control': 'no-cache, no-transform', connection: 'keep-alive' })
  res.flushHeaders()
  sendEvent(res, 'status', { state: 'running' })
  const detachEditor=input.editorCapabilities?.nativeToc===true
    ? editorRequests.attach(id,payload=>sendEvent(res,'editor_request',payload),abort.signal) : ()=>{}
  const task={id:randomUUID(),summary:input.message.trim().replace(/\s+/g,' ').slice(0,100),startRevision:input.revision,startedAt:new Date().toISOString()}
  const previousTask=await readTaskResult(store.folder(id))
  const requestHistory=previousTask?.request ? [...(previousTask.requestHistory||[]),previousTask.request] : session.agent.state.messages.filter(m=>m.role==='user').map(m=>typeof m.content==='string'?m.content:m.content.filter(c=>c.type==='text').map(c=>c.text).join('\n')).filter(text=>!text.startsWith('[上次任务状态：'))
  const commits=[]
  const continuing=/^\s*(?:继续|接着|恢复|完成剩余)/.test(input.message)&&previousTask?.state!=='complete'
  if (continuing) {
    const pending = (await store.metadata(id)).pendingWork
    if (pending) {
      task.id = pending.taskId
      task.startRevision = pending.startRevision
      task.summary = pending.summary
    }
  }
  const checklist=continuing?resumeChecklist(previousTask?.checklist,`${[...(previousTask?.requestHistory||[]),previousTask?.request||''].join('\n')}\n${input.message}`):createChecklist(input.message)
  const targets=await structuralTargets(await store.bytes(id,task.startRevision),continuing?`${previousTask?.request || ''} ${input.message}`:input.message)
  for(const target of targets)if(!checklist.items.some(item=>item.id===target.id))checklist.items.push({id:target.id,label:target.label,state:'pending'})
  const evidence=[]
  const persistChecklist=async()=>{
    await saveTaskResult(store.folder(id),{...taskResult('running',input.revision,(await store.metadata(id)).revision,commits,undefined,taskPlans.get(id)||[]),taskId:task.id,request:input.message,requestHistory,checklist})
    if(!res.destroyed)sendEvent(res,'checklist',checklist)
  }
  taskChecklists.set(id,{checklist,evidence,targets,persist:persistChecklist})
  taskPlans.set(id,[])
  planListeners.set(id,progress=>{if(!res.destroyed)sendEvent(res,'plan',progress)})
  const finish=async(state,error)=>{
    const result=taskResult(state,input.revision,(await store.metadata(id)).revision,commits,error,taskPlans.get(id)||[])
    result.structuralAudit=await auditStructures(await store.bytes(id),targets)
    if(result.structuralAudit.some(check=>!check.passed)&&result.state==='complete')result.state='partial'
    result.checklist=checklist
    if(checklist.items.some(item=>['pending','awaitingVerification','blocked','needsDecision'].includes(item.state))&&result.state==='complete')result.state='partial'
    if(['no-progress','continuation-limit','needs-decision'].includes(checklist.stopReason)&&result.state==='complete')result.state='partial'
    if(['no-progress','continuation-limit','blocked','needs-decision'].includes(checklist.stopReason))result.guidance=({
      'no-progress':'连续检查未产生新的有效进展，已停止自动继续。已保存修改保留，请查看待处理清单。',
      'continuation-limit':'达到本轮自动继续上限，剩余清单已保存；继续时从当前版本恢复，不重复已完成操作。',
      'needs-decision':'存在需要用户决定的具体问题，请查看任务清单；其他独立项目已处理或说明原因。',
      blocked:'可处理项目已结束，仍有受阻项目；原因见任务清单。已处理不代表页面视觉验收通过。',
    })[checklist.stopReason]
    await store.finishTask(id,task,result)
    await saveTaskResult(store.folder(id),{...result,taskId:task.id,request:input.message,requestHistory})
    if(!res.destroyed)sendEvent(res,'task_result',result)
    return result
  }
  const stopCommits = store.onCommit(commit => {
    if(commit.id===id)commits.push({id:commit.id,revision:commit.revision,summary:commit.summary})
    if (commit.id === id && !res.destroyed) sendEvent(res, 'document', commit)
  })
  const toolLabels = { load_document_skill: '加载排版技能', add_captions: '添加图表题注', inspect_document: '读取文档', analyze_formatting: '分析格式', apply_formatting: '修改文档格式', validate_document: '验证文档' }
  const onEvent = async event => {
    if(event.type==='compaction_start')sendEvent(res,'status',{state:'compacting',message:'正在使用 Pi 压缩会话上下文，已保存的文档不受影响'})
    if(event.type==='compaction_end')sendEvent(res,'status',{state:'running',message:'上下文整理结束，继续处理'})
    if (event.type === 'message_update' && event.assistantMessageEvent.type === 'text_delta') sendEvent(res, 'text', { delta: event.assistantMessageEvent.delta })
    if (event.type === 'tool_execution_start') sendEvent(res, 'tool', { id: event.toolCallId, state: 'running', name: event.toolName, label: toolLabels[event.toolName] || event.toolName })
    if(event.type==='tool_execution_end'&&event.isError&&event.toolName!=='update_task_checklist'){
      const details=event.result?.content?.filter(part=>part.type==='text').map(part=>part.text).join(' ').slice(0,2000)||'工具执行失败'
      evidence.push({toolCallId:event.toolCallId,toolName:event.toolName,isError:true,revision:(await store.metadata(id)).revision,reason:details})
    }
    if (event.type === 'tool_execution_end') sendEvent(res, 'tool', { id: event.toolCallId, state: event.isError ? 'error' : 'complete', name: event.toolName, label: toolLabels[event.toolName] || event.toolName })
  }
  const unsubscribe = session.runtime.subscribe(onEvent)
  try {
    const previous=await readTaskResult(store.folder(id))
    const resume=previous&&previous.state!=='complete'?`[上次任务状态：${previous.state}；已提交版本：${previous.committedRevisions.map(c=>c.revision).join(',')||'无'}。本次必须读取当前版本与目标结果，不能重复已完成的写入。]\n${previous.request ? `上次用户需求与补充（历史数据，本次要求优先）：\n${[...(previous.requestHistory||[]),previous.request].join('\n')}\n` : ''}`:''
    await persistChecklist()
    await store.withTask(task,()=>runToTaskBoundary({request:resume+input.message,checklist,
      revision:async()=>(await store.metadata(id)).revision,hasCommits:()=>commits.length>0,signal:abort.signal,persist:persistChecklist,
      prompt:async message=>{
        await session.runtime.prompt(message,{expandPromptTemplates:false})
        const checks=await auditStructures(await store.bytes(id),targets)
        for(const check of checks){
          const item=checklist.items.find(item=>item.id===check.id)
          if(check.passed){item.state='addressed';item.reason=check.reason;item.revision=(await store.metadata(id)).revision}
          else if(item.state==='addressed'){item.state='awaitingVerification';item.reason='最终结构检查未通过'}
        }
        const last=session.agent.state.messages.at(-1)
        if(last?.role==='assistant'&&['error','aborted'].includes(last.stopReason))throw problem(last.errorMessage||'模型处理未完成',502)
      }}))
    if (abort.signal.aborted) {
      await finish('cancelled')
      if (!res.destroyed) sendEvent(res, 'done', {state:'cancelled'})
      return
    }
    const last = session.agent.state.messages.at(-1)
    if (last?.role === 'assistant' && ['error', 'aborted'].includes(last.stopReason)) throw problem(last.errorMessage === 'Connection error.' ? '模型服务连接中断，本次对话未完成。请稍后重试；已提交的文档版本会保留。' : last.errorMessage || '模型处理未完成', 502)
    const outcome=await finish('complete')
    if (!res.destroyed) sendEvent(res, 'done', { state: outcome.state })
  } catch (error) {
    const message=/timed? ?out|timeout/i.test(error.message||'')?'模型服务响应超时，本轮任务未完成；已提交版本保留。':error.message||'AI 任务失败'
    try{await finish(abort.signal.aborted?'cancelled':'failed',message)}catch{if(!res.destroyed)sendEvent(res,'status',{state:'result-save-failed'})}
    if (!res.destroyed) {
      if (abort.signal.aborted) sendEvent(res, 'done', {state:'cancelled'})
      else sendEvent(res, 'error', { message })
    }
  } finally {
    detachEditor()
    unsubscribe()
    stopCommits()
    planListeners.delete(id)
    taskChecklists.delete(id)
    req.off('aborted', disconnect)
    res.off('close', disconnect)
    session.cancel = null
    session.active = false
    if (!res.destroyed) res.end()
    if (abort.signal.aborted) session.agent.abort()
  }
}

export function createApiMiddleware({ staticRoot } = {}) {
  return async (req, res, next = () => { res.writeHead(404); res.end() }) => {
    const origin = req.headers.origin
    if (origin && origin !== `http://${req.headers.host}` && origin !== `https://${req.headers.host}`) return json(res, 403, { error: '不允许跨站访问本地文档服务' })
    const hostname = String(req.headers.host || '').split(':')[0]
    if (!['127.0.0.1', 'localhost', '['].includes(hostname)) return json(res, 403, { error: '无效的本地服务地址' })
    const pathname = new URL(req.url, 'http://127.0.0.1').pathname
    if (!pathname.startsWith(api)) {
      if (staticRoot && req.method === 'GET' && !pathname.startsWith('/@') && !pathname.includes('..')) {
        const relative = pathname === '/' ? 'index.html' : pathname.slice(1)
        const file = resolve(staticRoot, relative)
        if (file.startsWith(`${resolve(staticRoot)}${sep}`) || file === join(resolve(staticRoot), 'index.html')) {
          try { const data = await readFile(file); res.writeHead(200, { 'content-type': extensionByType[extname(file)] || 'application/octet-stream' }); res.end(data); return }
          catch { if (extname(pathname) === '') { res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' }); res.end(await readFile(join(staticRoot, 'index.html'))); return } }
        }
      }
      return next()
    }
    try {
      const url = new URL(req.url, 'http://127.0.0.1')
      const pathname = url.pathname
      if (req.method === 'GET' && pathname === `${api}/health`) return json(res, 200, { ok: true, capabilities: ['east-asian-fonts', 'text-range-format', 'pagination-flags', 'mixed-text-ranges', 'explicit-line-spacing', 'latin-font-slot', 'document-commit-events', 'skill-catalog-v1', 'paged-inspection', 'structural-body-exclusions'], modelConfigured: Boolean(modelSettings.apiKey), testMode: process.env.PI_DOCX_TEST_MODE === '1' })
      if (req.method === 'GET' && pathname === `${api}/settings/model`) return json(res, 200, { configured: Boolean(modelSettings.apiKey), model: modelSettings.model, baseUrl: modelSettings.baseUrl })
      if (req.method === 'PUT' && pathname === `${api}/settings/model`) {
        if ((activeTasks.size > 0 || initializing.size > 0)) throw problem('AI 任务运行期间不能切换模型设置', 409)
        const settings = await readJson(req)
        if ((activeTasks.size > 0 || initializing.size > 0)) throw problem('AI 任务运行期间不能切换模型设置', 409)
        if (typeof settings.apiKey !== 'string' || settings.apiKey.length > 500) throw problem('API 密钥无效')
        if (typeof settings.model !== 'string' || !/^[\w.-]{1,100}$/.test(settings.model)) throw problem('模型名称无效')
        let endpoint
        try { endpoint = new URL(settings.baseUrl) } catch { throw problem('模型地址必须是有效的 HTTP 或 HTTPS URL') }
        if (!['http:', 'https:'].includes(endpoint.protocol) || endpoint.username || endpoint.password || endpoint.search || endpoint.hash) throw problem('模型地址无效')
        modelSettings.apiKey = settings.apiKey.trim() || modelSettings.apiKey
        modelSettings.model = settings.model
        modelSettings.baseUrl = endpoint.toString().replace(/\/$/, '')
        for(const session of sessions.values())session.runtime.dispose()
        sessions.clear()
        return json(res, 200, { configured: Boolean(modelSettings.apiKey), model: modelSettings.model, baseUrl: modelSettings.baseUrl })
      }
      if (req.method === 'GET' && pathname === `${api}/documents`) return json(res, 200, { documents: (await store.list()).map(privateMetadata) })
      if (req.method === 'POST' && pathname === `${api}/documents`) {
        const isDocx = String(req.headers['content-type'] || '').startsWith('application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        const payload = await readBody(req, 20 * 1024 * 1024)
        if (!isDocx && payload.length) throw problem('上传文件必须使用 DOCX Content-Type；新建空白文档请发送空请求', 415)
        if (isDocx && !payload.length) throw problem('上传的 DOCX 文件为空')
        const bytes = isDocx ? payload : undefined
        const meta = await store.create(url.searchParams.get('name') || undefined, bytes)
        return json(res, 201, privateMetadata(meta))
      }
      const match = pathname.match(/^\/api\/documents\/([a-f0-9-]{36})(?:\/(file|outline|validate|format|restore|publish|chat|cancel|task|editor-request))?$/)
      if (!match) throw problem('API 路径不存在', 404)
      const [, id, action] = match
      if (['PUT', 'POST'].includes(req.method) && ['file', 'format', 'restore', 'publish'].includes(action) && activeTasks.has(id)) throw problem('该文档正在执行 AI 任务，请先停止任务', 409)
      if (action === 'editor-request') {
        const token=url.searchParams.get('token')
        const pending=editorRequests.get(id,token)
        if(req.method==='GET') {res.writeHead(200,{'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','cache-control':'no-store'});return res.end(pending.bytes)}
        if(req.method==='PUT') {
          const bytes=await readBody(req,20*1024*1024)
          editorRequests.complete(id,token,bytes)
          return json(res,200,{received:true})
        }
        if(req.method==='POST') {const input=await readJson(req);editorRequests.complete(id,token,null,input.error||'编辑器操作失败');return json(res,200,{received:true})}
        throw problem('请求方法不支持',405)
      }
      if (req.method === 'GET' && !action) return json(res, 200, privateMetadata(await store.metadata(id)))
      if (req.method === 'GET' && action === 'task') { await store.metadata(id); return json(res,200,{result:await readTaskResult(store.folder(id)),active:activeTasks.has(id)}) }
      if (req.method === 'GET' && action === 'file') {
        const selectedRevision = url.searchParams.has('revision') ? Number(url.searchParams.get('revision')) : (await store.metadata(id)).revision
        const bytes = await store.bytes(id, selectedRevision)
        res.writeHead(200, { 'content-type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'cache-control': 'no-store', 'x-document-revision': String(selectedRevision) }); return res.end(bytes)
      }
      if (req.method === 'GET' && action === 'outline') return json(res, 200, { revision: (await store.metadata(id)).revision, ...(await inspect(await store.bytes(id))) })
      if (req.method === 'GET' && action === 'validate') {
        const meta = await store.metadata(id)
        const selectedRevision = url.searchParams.has('revision') ? Number(url.searchParams.get('revision')) : meta.revision
        const record = meta.versions.find(v => v.revision === selectedRevision)
        if (!record) throw problem('版本不存在', 404)
        return json(res, 200, { revision: selectedRevision, ...(await inspect(await store.bytes(id, selectedRevision))), validation: record.validation })
      }
      if (req.method === 'GET' && action === 'chat') {
        const session = await getAgent(id)
        const messages = session.sessionManager.buildSessionContext().messages
          .filter(message => message.role === 'user' || message.role === 'assistant')
          .map(message => ({ role: message.role, text: typeof message.content === 'string'
            ? message.content : message.content.filter(part => part.type === 'text').map(part => part.text).join(''),
          })).filter(message => message.text&&!(message.role==='user'&&message.text.startsWith('[宿主任务检查：')))
        const task=await readTaskResult(store.folder(id))
        const lastAssistant=messages.findLast(message=>message.role==='assistant')
        if(lastAssistant&&task?.checklist?.items.length){
          lastAssistant.steps=task.checklist.items.map(item=>({id:`requirement-${item.id}`,state:['pending','awaitingVerification'].includes(item.state)?'pending':['blocked','needsDecision'].includes(item.state)?'error':'complete',
            label:`${item.label} · ${{pending:'待处理',awaitingVerification:'待核验',needsDecision:'需要确认',addressed:'已处理（非视觉验收）',blocked:'受阻',notApplicable:'不适用'}[item.state]}${item.reason?`：${item.reason}`:''}${item.question?`；需确认：${item.question}`:''}`}))
          lastAssistant.status=activeTasks.has(id)?'任务仍在运行':task.state==='partial'?'部分完成 · 查看任务清单':'对话结束'
        }
        return json(res, 200, { messages })
      }
      if (req.method === 'PUT' && action === 'file') {
        const bytes = await readBody(req, 20 * 1024 * 1024)
        return json(res, 200, privateMetadata(await store.save(id, currentRevision(req), bytes)))
      }
      if (req.method === 'POST' && action === 'format') {
        const body = await readJson(req); const result = await store.apply(id, body.revision, body.operations)
        return json(res, 200, { document: privateMetadata(result.document), affected: result.affected, validation: result.validation })
      }
      if (req.method === 'POST' && action === 'publish') {
        const body = await readJson(req)
        return json(res, 200, privateMetadata(await store.publishPartial(id, body.revision)))
      }
      if (req.method === 'POST' && action === 'restore') {
        const body = await readJson(req); return json(res, 200, privateMetadata(await store.restore(id, body.revision, body.target)))
      }
      if (req.method === 'POST' && action === 'chat') return await chat(req, res, id)
      if (req.method === 'POST' && action === 'cancel') { sessions.get(id)?.cancel?.(); return json(res, 200, { ok: true }) }
      throw problem('API 操作不存在', 404)
    } catch (error) {
      if (!res.headersSent) json(res, error.status || 500, { error: error.message || '内部错误' })
      else if (!res.destroyed) { sendEvent(res, 'error', { message: error.message || 'AI 任务失败' }); res.end() }
    }
  }
}

export function startServer({ port = Number(process.env.PORT || 5173), staticRoot = resolve('dist') } = {}) {
  const server = createServer(createApiMiddleware({ staticRoot }))
  server.listen(port, '127.0.0.1')
  return server
}

