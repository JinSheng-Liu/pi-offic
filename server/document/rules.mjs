import { problem } from './package.mjs'
const fontKeys = new Set(['name', 'latinName', 'size', 'bold', 'italic', 'color', 'underline', 'strikeThrough', 'doubleStrikeThrough', 'subscript', 'superscript'])
export const paginationKeys = new Set(['keepNext', 'keepLines', 'widowControl', 'pageBreakBefore'])
const paragraphKeys = new Set([...paginationKeys, 'alignment', 'leftIndent', 'rightIndent', 'firstLineIndent', 'spaceBefore', 'spaceAfter', 'lineSpacing', 'lineSpacingMultiple', 'leadingTabPosition', 'tocPageTabPosition', 'outlineLevel'])
const marginKeys = new Set(['top', 'right', 'bottom', 'left', 'gutter'])

export function validatePage(page) {
  if (!page || typeof page !== 'object' || Object.keys(page).some(key => !['widthTwips', 'heightTwips', 'orientation'].includes(key))) throw problem('页面设置字段无效')
  for (const key of ['widthTwips', 'heightTwips']) if (page[key] !== undefined && (!Number.isInteger(page[key]) || page[key] < 1000 || page[key] > 31680)) throw problem('页面宽高必须为 1000–31680 twips 的整数')
  if (page.orientation !== undefined && !['portrait', 'landscape'].includes(page.orientation)) throw problem('页面方向必须为 portrait 或 landscape')
  if (!Object.keys(page).length) throw problem('没有页面设置属性')
}

export function validateMargins(margins) {
  if (!margins || typeof margins !== 'object' || !Object.keys(margins).length || Object.keys(margins).some(key => !marginKeys.has(key))) throw problem('页边距字段无效')
  for (const [key, value] of Object.entries(margins)) if (!Number.isInteger(value) || value < 0 || value > 31680) throw problem(`页边距 ${key} 必须为 0–31680 twips 的整数`)
}

export function validateOperations(operations) {
  if (!Array.isArray(operations) || !operations.length || operations.length > 30) throw problem('需要 1–30 项格式操作')
  for (const op of operations) {
    if (!op || typeof op !== 'object' || Object.keys(op).some(k => !['scope','indices','heading','font','paragraph','style','styleName','sectionIndices','page','margins','text','occurrence','tableIndex','rowIndices','cellTargets','table','objectIndices','objectLayout','alignment','cover','columnRatios','format','prefix','suffix'].includes(k))) throw problem('未知格式操作字段')
    if (!['tocEntries','tabularText','characterStyle','namedStyle','all','body','tableText','inlinePictures','headings','indices','sections','text','tableRows','tableCells','floatingObjects','tocTitle','footerPageNumber','fitTableWidth','ensurePageNumbers','convertCaptions','coverLayout','renumberHeadings','numberEquations'].includes(op.scope)) throw problem('无效操作范围')
    if(op.columnRatios!==undefined&&op.scope!=='fitTableWidth')throw problem('列宽比例仅用于表格版心适配')
    if (op.scope !== 'numberEquations' && ['format', 'prefix', 'suffix'].some(key => Object.hasOwn(op, key))) throw problem('编号格式参数仅用于公式编号')
    if(op.scope==='tocEntries') {
      if(operations.length!==1||Object.keys(op).length!==1)throw problem('目录条目整理需独立提交')
      continue
    }
    if(op.scope==='tabularText') {
      if(operations.length!==1||Object.keys(op).some(k=>!['scope','indices'].includes(k))||!Array.isArray(op.indices)||op.indices.length<2||op.indices.length>100||op.indices.some((v,i)=>!Number.isInteger(v)||v<0||i>0&&v!==op.indices[i-1]+1))throw problem('Tab 转表格需独立提交 2–100 个连续段落索引')
      continue
    }
    if(op.scope==='characterStyle') {
      if(Object.keys(op).some(k=>!['scope','indices','text','occurrence','styleName','font'].includes(k))||typeof op.styleName!=='string'||!op.styleName.trim()||op.styleName.length>80||/[\x00-\x1f]/.test(op.styleName))throw problem('字符样式需提供名称、单段索引、准确文字和字体')
      const {styleName,...range}=op
      validateOperations([{...range,scope:'text'}]);continue
    }
    if(op.scope==='namedStyle') {
      if(operations.length!==1||Object.keys(op).some(k=>!['scope','indices','styleName','font','paragraph'].includes(k))||typeof op.styleName!=='string'||!op.styleName.trim()||op.styleName.length>80||/[\x00-\x1f]/.test(op.styleName))throw problem('命名样式需独立提交段落索引、名称及格式')
      if(op.paragraph?.leadingTabPosition!==undefined||op.paragraph?.tocPageTabPosition!==undefined)throw problem('命名样式不支持专项制表位')
      validateOperations([{scope:'indices',indices:op.indices,font:op.font,paragraph:op.paragraph}])
      continue
    }
    if(op.styleName!==undefined)throw problem('styleName仅用于namedStyle')
    if(op.scope==='ensurePageNumbers') {
      if(operations.length!==1||Object.keys(op).length!==1)throw problem('全文页码规范需独立提交且仅指定scope')
      continue
    }
    if(op.scope==='fitTableWidth') {
      if(operations.length!==1||Object.keys(op).some(k=>!['scope','tableIndex','columnRatios'].includes(k))||!Number.isInteger(op.tableIndex)||op.tableIndex<0)throw problem('表格版心适配需独立提交tableIndex')
      if(op.columnRatios!==undefined&&(!Array.isArray(op.columnRatios)||op.columnRatios.length<2||op.columnRatios.length>20||op.columnRatios.some(v=>!Number.isFinite(v)||v<=0||v>100)))throw problem('列宽比例必须是2至20个正数且不超过100')
      continue
    }
    if(op.scope==='coverLayout') {
      if(operations.length!==1||Object.keys(op).some(k=>!['scope','indices','cover'].includes(k))||!Array.isArray(op.indices)||op.indices.length<2||op.indices.length>20||op.indices.some(i=>!Number.isInteger(i)||i<0)||new Set(op.indices).size!==op.indices.length||!op.cover||Object.keys(op.cover).some(k=>!['widthPt','labelWidthPt'].includes(k))||!Number.isFinite(op.cover.widthPt)||!Number.isFinite(op.cover.labelWidthPt)||op.cover.widthPt<200||op.cover.widthPt>500||op.cover.labelWidthPt<40||op.cover.labelWidthPt>op.cover.widthPt-100)throw problem('封面重排需独立提交连续段落索引及有效总宽和标签列宽（pt）')
      continue
    }
    if(op.cover!==undefined)throw problem('cover仅用于coverLayout')
    if(op.scope==='convertCaptions') {
      if(operations.length!==1||Object.keys(op).some(k=>!['scope','indices'].includes(k))||!Array.isArray(op.indices)||!op.indices.length||op.indices.length>200||op.indices.some(i=>!Number.isInteger(i)||i<0)||new Set(op.indices).size!==op.indices.length)throw problem('题注转换必须独立提交不重复的明确段落索引')
      continue
    }
    if (op.scope === 'renumberHeadings') {
      if (operations.length !== 1 || Object.keys(op).length !== 1) throw problem('标题编号校正需独立提交且仅指定scope')
      continue
    }
    if (op.scope === 'numberEquations') {
      if (operations.length !== 1 || Object.keys(op).some(k => !['scope', 'format', 'prefix', 'suffix'].includes(k))) throw problem('公式编号需独立提交')
      if (op.format !== undefined && !['arabic', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter'].includes(op.format)) throw problem('公式编号格式无效')
      for (const key of ['prefix', 'suffix']) if (op[key] !== undefined && (typeof op[key] !== 'string' || op[key].length > 20 || /[\r\n\u0000-\u001f]/u.test(op[key]))) throw problem('公式编号前后缀无效')
      continue
    }
    if(op.scope==='footerPageNumber') {
      if(operations.length!==1||Object.keys(op).some(k=>!['scope','sectionIndices','font','alignment'].includes(k))||!Array.isArray(op.sectionIndices)||!op.sectionIndices.length||op.sectionIndices.some(i=>!Number.isInteger(i)||i<0)||(!op.font&&!op.alignment)) throw problem('页脚页码格式需单独提交明确分节与字体或对齐')
      if(op.alignment!==undefined&&!['Left','Centered','Right'].includes(op.alignment))throw problem('页码对齐无效')
      if(op.font&&(!Object.keys(op.font).length||Object.keys(op.font).some(k=>!['name','size'].includes(k))||op.font.name!==undefined&&(typeof op.font.name!=='string'||!op.font.name.trim()||op.font.name.length>100)||op.font.size!==undefined&&(!Number.isFinite(op.font.size)||op.font.size<1||op.font.size>200||op.font.size*2!==Math.round(op.font.size*2))))throw problem('页码字体或字号无效，字号须为半磅倍数')
      continue
    }
    if(op.alignment!==undefined)throw problem('alignment仅用于页脚页码，段落请使用paragraph.alignment')
    if (op.scope==='tocTitle') {
      if(operations.length!==1||Object.keys(op).some(k=>!['scope','indices'].includes(k))||!Array.isArray(op.indices)||!op.indices.length||op.indices.some(i=>!Number.isInteger(i)||i<0)||new Set(op.indices).size!==op.indices.length) throw problem('目录标题分离必须单独提交有效且不重复的段落索引')
      continue
    }
    if (op.scope==='floatingObjects') {
      if(operations.length!==1||Object.keys(op).some(k=>!['scope','objectIndices','objectLayout','alignment','cover','columnRatios'].includes(k))||!Array.isArray(op.objectIndices)||!op.objectIndices.length||op.objectIndices.length>150||op.objectIndices.some(i=>!Number.isInteger(i)||i<0)||new Set(op.objectIndices).size!==op.objectIndices.length||!['inlinePicture','captionInFlow'].includes(op.objectLayout)) throw problem('对象转换必须单独提交，提供不重复的对象索引与明确布局方式')
      continue
    }
    if (op.objectIndices!==undefined||op.objectLayout!==undefined) throw problem('对象字段只能用于 floatingObjects')
    if (['tableRows','tableCells'].includes(op.scope)) {
      const allowed = new Set(['scope','tableIndex',op.scope==='tableRows'?'rowIndices':'cellTargets','table'])
      if (Object.keys(op).some(key=>!allowed.has(key))) throw problem('表格操作包含不适用的字段')
      if (!Number.isInteger(op.tableIndex)||op.tableIndex<0) throw problem('表格索引无效')
      if (op.scope==='tableRows') {
        if (!Array.isArray(op.rowIndices)||!op.rowIndices.length||op.rowIndices.length>100||op.rowIndices.some(i=>!Number.isInteger(i)||i<0)||new Set(op.rowIndices).size!==op.rowIndices.length||!op.table||Object.keys(op.table).some(k=>!['repeatHeader','cantSplit','autoHeight'].includes(k))||!Object.keys(op.table).length||Object.values(op.table).some(v=>typeof v!=='boolean')) throw problem('表格行操作需提供不重复的有效行索引及布尔型重复表头、禁止跨页或自动行高属性')
      } else if (!Array.isArray(op.cellTargets)||!op.cellTargets.length||op.cellTargets.length>100||op.cellTargets.some(c=>!c||typeof c!=='object'||Object.keys(c).some(k=>!['rowIndex','cellIndex'].includes(k))||!Number.isInteger(c.rowIndex)||c.rowIndex<0||!Number.isInteger(c.cellIndex)||c.cellIndex<0)||!op.table||Object.keys(op.table).some(k=>k!=='verticalAlignment')||!['Top','Center','Bottom'].includes(op.table.verticalAlignment)) throw problem('单元格操作需提供有效物理行列索引及垂直对齐方式')
      continue
    }
    if (['tableIndex','rowIndices','cellTargets','table'].some(key=>Object.hasOwn(op,key))) throw problem('表格字段只能用于 tableRows 或 tableCells 范围')
    if (op.indices !== undefined && !['indices','text'].includes(op.scope)) throw problem('indices 只能用于 indices 或 text 范围，不能与全文或其他范围混用')
    if (['indices','text'].includes(op.scope) && (!Array.isArray(op.indices) || !op.indices.length || op.indices.some(i => !Number.isInteger(i) || i < 0))) throw problem('段落索引无效')
    if (op.scope === 'sections' && (!Array.isArray(op.sectionIndices) || !op.sectionIndices.length || (!op.page && !op.margins))) throw problem('分节操作必须提供 sectionIndices，并设置 page 或 margins')
    if (op.scope === 'sections' && (op.font || op.paragraph || op.style)) throw problem('分节操作只接受 page 和 margins；请将 font、paragraph、style 格式拆成独立的段落操作')
    if (op.scope === 'text' && (op.indices.length !== 1 || typeof op.text !== 'string' || !op.text.length || op.text.length > 1000 || op.paragraph || op.style || !Object.keys(op.font || {}).length)) throw problem('文字范围需要一个段落索引、准确原文和字体属性')
    if (op.scope !== 'text' && (op.text !== undefined || op.occurrence !== undefined)) throw problem('文字定位字段只能用于 text 范围')
    if (op.occurrence !== undefined && (!Number.isInteger(op.occurrence) || op.occurrence < 0)) throw problem('出现序号必须为非负整数')
    if (op.heading !== undefined && (!Number.isInteger(op.heading) || op.heading < 1 || op.heading > 9 || op.scope !== 'headings')) throw problem('标题层级必须为 1–9')
    if (op.style !== undefined && !/^(Normal|Title|Subtitle|Heading[1-9])$/.test(op.style)) throw problem('仅支持原生正文与标题样式')
    if (op.scope !== 'sections' && (!op.style && !Object.keys(op.font || {}).length && !Object.keys(op.paragraph || {}).length || op.sectionIndices || op.page || op.margins)) throw problem('段落操作格式属性无效')
    if (op.scope === 'sections') {
      if (op.page) validatePage(op.page)
      if (op.margins) validateMargins(op.margins)
    }
    if (op.paragraph?.lineSpacing !== undefined && op.paragraph?.lineSpacingMultiple !== undefined) throw problem('固定行距与多倍行距不能同时设置')
    if (op.paragraph?.tocPageTabPosition !== undefined && (op.scope !== 'indices' || op.paragraph.leadingTabPosition !== undefined || op.style)) throw problem('目录页码制表位只能用于明确索引，不能与样式或前导Tab混用')
    if (op.paragraph?.leadingTabPosition !== undefined && (op.scope !== 'indices' || op.paragraph.firstLineIndent !== undefined || op.paragraph.leftIndent !== undefined || op.style)) throw problem('前导Tab定位只能使用明确索引，且不能同时修改缩进或样式')
    if (op.font?.subscript && op.font?.superscript) throw problem('上标和下标不能同时启用')
    for (const [key, value] of Object.entries(op.font || {})) {
      if (!fontKeys.has(key)) throw problem(`不支持字体属性 ${key}`)
      if (['name','latinName'].includes(key) && (typeof value !== 'string' || !value.trim() || value.length > 100)) throw problem('字体名称无效')
      if (key === 'size' && (typeof value !== 'number' || !Number.isFinite(value) || value < 1 || value > 200)) throw problem('字号必须为 1–200 pt')
      if (['bold','italic','strikeThrough','doubleStrikeThrough','subscript','superscript'].includes(key) && typeof value !== 'boolean') throw problem('粗体和斜体必须为布尔值')
      if (key === 'underline' && !['None', 'Single', 'Double', 'Dotted', 'Dash', 'Wave'].includes(value)) throw problem('不支持的下划线类型')
      if (key === 'color' && (typeof value !== 'string' || !/^#?[0-9a-fA-F]{6}$/.test(value))) throw problem('颜色必须为 #RRGGBB')
    }
    for (const [key, value] of Object.entries(op.paragraph || {})) {
      if (!paragraphKeys.has(key)) throw problem(`不支持段落属性 ${key}`)
      if (key === 'lineSpacing' && (typeof value !== 'number' || !Number.isFinite(value) || value <= 0 || value > 720)) throw problem('固定行距必须大于0；单倍行距请使用 lineSpacingMultiple: 1')
      if (key === 'lineSpacingMultiple') { if (typeof value !== 'number' || !Number.isFinite(value) || value < 1 || value > 10) throw problem('多倍行距必须为1至10倍') }
      else if (paginationKeys.has(key)) { if (typeof value !== 'boolean') throw problem('分页属性必须为布尔值') }
      else if (key === 'outlineLevel') { if (!Number.isInteger(value) || value < 1 || value > 10) throw problem('大纲层级必须为 1–9，或 10（正文）') }
      else if (key === 'alignment') { if (!['Left','Centered','Right','Justified','Center','Justify'].includes(value)) throw problem('无效对齐方式') }
      else if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 720 || (key !== 'firstLineIndent' && value < 0)) throw problem('段落数值必须为有效 pt 值')
    }
  }
}

