import { W } from './package.mjs'
import {flowTarget} from './object-flow.mjs'
import {findMixedTocTitles} from './toc-layout.mjs'
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

// Structural risk signals, not proof of a rendered collision. XML object indices
// deliberately do not pretend to be SDK paragraph indices or editable anchors.
export function inspectLayoutRisks(document) {
  const anchors = Array.from(document.getElementsByTagNameNS(WP, 'anchor'))
  const floatingObjects = anchors.map((anchor, index) => {
    const boxes = Array.from(anchor.getElementsByTagNameNS(W, 'txbxContent'))
    const texts = boxes.map(box => Array.from(box.getElementsByTagNameNS(W, 't')).map(t => t.textContent || '').join(''))
    const wrap = Array.from(anchor.childNodes).find(n => n.namespaceURI === WP && n.localName.startsWith('wrap'))
    const position = axis => {
      const node = Array.from(anchor.childNodes).find(n => n.namespaceURI === WP && n.localName === axis)
      const offset = node?.getElementsByTagNameNS(WP, 'posOffset')[0]
      return node ? {relativeFrom:node.getAttribute('relativeFrom'),offsetEmu:offset ? Number(offset.textContent) : null} : undefined
    }
    const mode=boxes.length?'captionInFlow':'inlinePicture'
    let flowCapability
    try {flowTarget(anchor,mode);flowCapability={mode,supported:true}} catch(error){flowCapability={mode,supported:false,reason:error.message}}
    return { objectIndex:index, readOnly:true, flowCapability, wrap:wrap?.localName || 'unspecified', textBoxCount:boxes.length,
      text:texts.join('\n').slice(0,300), textTruncated:texts.join('\n').length>300,
      positionH:position('positionH'),positionV:position('positionV') }
  })
  const tocInstructions = Array.from(document.getElementsByTagNameNS(W,'instrText')).filter(n=> /\bTOC\b/.test(n.textContent || '')).length
  const noWrap = floatingObjects.filter(o=>o.wrap==='wrapNone').length
  const textBoxes = floatingObjects.filter(o=>o.textBoxCount>0).length
  const warnings = []
  const mixedTocTitleIndices=findMixedTocTitles(document)
  if(mixedTocTitleIndices.length)warnings.push({code:'TOC_MIXED_TITLE',count:mixedTocTitleIndices.length,message:'目录标题与目录域共用段落，存在标题拆分和分页异常风险；规范目录前先分离标题并重新读取索引。'})
  const body=document.getElementsByTagNameNS(W,'body')[0]
  const front=Array.from(body?.childNodes||[]).filter(n=>n.namespaceURI===W&&n.localName==='p').slice(0,20)
  const plain=p=>Array.from(p.getElementsByTagNameNS(W,'t')).map(t=>t.textContent||'').join('')
  const titleIndex=front.findIndex(p=>/本科毕业论文|学位论文|毕业设计/.test(plain(p)))
  if(titleIndex>=0){
    const cover=front.slice(titleIndex+1)
    const boundary=cover.findIndex(p=>/授权书|原创性声明|^摘\s*要/.test(plain(p)))
    const fields=(boundary<0?cover:cover.slice(0,boundary)).filter(p=>/^(学生姓名|学\s*号|学\s*院|专业班级|指导教师|起止时间)/.test(plain(p)))
    const spaced=fields.filter(p=>/[：:][ \u3000]{2,}\S/.test(plain(p)))
    if(spaced.length)warnings.push({code:'COVER_SPACE_ALIGNMENT',count:spaced.length,message:'封面信息栏仍使用连续空格定位，统一字号或左缩进不能保证标签和值对齐；需按模板改为统一制表位或无边框布局表格，封面未验收。'})
    const heading=cover.findIndex(p=>/^题\s*目[：:]/.test(plain(p)))
    if(heading>=0&&cover[heading+1]&&plain(cover[heading+1]).trim()&&!fields.includes(cover[heading+1])&&cover[heading+1].getElementsByTagNameNS(W,'jc')[0]?.getAttributeNS(W,'val')==='center')warnings.push({code:'COVER_SPLIT_TITLE',count:1,message:'封面题目后紧接独立居中段落，可能把续行当作独立标题；需核对题名连续性、标签与题名起点，不能只验收字号。'})
  }
  if (noWrap) warnings.push({code:'FLOATING_NO_WRAP',count:noWrap,message:'无环绕浮动对象可能覆盖正文；普通段落/嵌入图片格式检查不覆盖其视觉位置。'})
  if (textBoxes) warnings.push({code:'FLOATING_TEXTBOX',count:textBoxes,message:'浮动文本框文字未作为普通段落处理；包括图表题注时，不能宣称全部题注已完成。'})
  if (tocInstructions) warnings.push({code:'TOC_LAYOUT_UNVERIFIED',count:tocInstructions,message:'存在原生目录域；域保留不等于制表位、分页或目录页码正确。'})
  return { floatingObjects, layoutRisks: { floatingCount:anchors.length, floatingTextBoxCount:textBoxes,
    noWrapCount:noWrap, tocInstructions, mixedTocTitleIndices, warnings, visualVerified:false } }
}
