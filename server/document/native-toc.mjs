import { randomBytes } from 'node:crypto'
import { zipSync } from 'fflate'
import { XMLSerializer } from '@xmldom/xmldom'
import { unpack, xml, W, canonical, compareProtected, problem } from './package.mjs'
import { inspect } from './structure.mjs'
import { hasNativeToc } from './toc-layout.mjs'
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
const elements = node => Array.from(node.childNodes).filter(n => n.nodeType === 1)
const encode = doc => new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
const tocBlocks = doc => [...doc.getElementsByTagNameNS(W,'sdt')].filter(n => [...n.getElementsByTagNameNS(W,'instrText')].some(i => /^\s*TOC(?:\s|$)/i.test(i.textContent)))
const title = value => value.trim().replace(/^(?:第[一二三四五六七八九十百\d]+[章节]|[一二三四五六七八九十百]+[、．.]|\d+(?:\.\d+)*[.、．]?)[\s　]*/, '').replace(/\s/g,'')

export async function prepareNativeToc(bytes, request) {
  const parts = unpack(bytes), doc = xml(parts['word/document.xml'])
  const body = doc.getElementsByTagNameNS(W,'body')[0]
  const data = await inspect(bytes)
  const paragraphs = []
  const walk = node => { if(node.namespaceURI===W && node.localName==='p'){paragraphs.push(node);return} for(const child of elements(node))walk(child) }
  walk(body)
  let anchor
  if (request.mode === 'refresh') {
    const blocks = tocBlocks(doc)
    if (blocks.length !== 1 || blocks[0].parentNode !== body) throw problem('更新要求唯一的正文原生目录；复杂目录需单独定位',422)
    anchor = blocks[0].getElementsByTagNameNS(W,'p')[0]
  } else {
    if (hasNativeToc(bytes)) throw problem('已存在原生目录，请更新而非重复插入',422)
    if (!data.paragraphs.some(p=>p.heading>=1 && p.heading<=3)) throw problem('请先建立正文标题层级，再生成目录',422)
    if (request.mode === 'replace') {
      const indices = request.indices
      if(!Array.isArray(indices)||!indices.length||indices.length>100||indices.some((n,i)=>!Number.isInteger(n)||n<0||i>0&&n!==indices[i-1]+1))throw problem('手工目录转换需要连续条目索引',422)
      const selected = indices.map(index=>paragraphs[index])
      const headingNames = new Set(data.paragraphs.filter(p=>p.heading>=1&&p.heading<=3&&!indices.includes(p.index)).map(p=>title(p.text)))
      for (const [i,p] of selected.entries()) {
        if (!p || p.parentNode!==body || i>0 && elements(body).indexOf(p)!==elements(body).indexOf(selected[i-1])+1) throw problem('目录条目必须为连续顶层段落',422)
        if (elements(p).some(n=>n.namespaceURI!==W||!['pPr','r'].includes(n.localName)) || [...p.getElementsByTagNameNS(W,'r')].some(r=>elements(r).some(n=>!['rPr','t','tab'].includes(n.localName))) || p.getElementsByTagNameNS(W,'sectPr').length) throw problem('手工目录含复杂对象，不能安全替换',422)
        const label = data.paragraphs[indices[i]].text.replace(/\t\s*[0-9ivxlcdm]+\s*$/i,'')
        if (!headingNames.has(title(label))) throw problem('目录条目与正文标题不对应，拒绝删除',422)
      }
      anchor = selected.at(-1).nextSibling
      while(anchor && anchor.nodeType!==1)anchor=anchor.nextSibling
      selected.forEach(p=>body.removeChild(p))
    } else anchor = paragraphs[request.paragraphIndex]
    if (!anchor || anchor.namespaceURI!==W || anchor.localName!=='p' || anchor.parentNode!==body) throw problem('目录插入位置必须为明确顶层段落之前',422)
  }
  if (!anchor) throw problem('无法定位目录',422)
  let paraId = anchor.getAttributeNS(W14,'paraId')
  if(!paraId){paraId=randomBytes(4).toString('hex').toUpperCase();anchor.setAttributeNS(W14,'w14:paraId',paraId)}
  parts['word/document.xml']=encode(doc)
  return { bytes:zipSync(parts), paraId, headings:data.paragraphs.filter(p=>p.heading>=1&&p.heading<=3).map(p=>p.text.trim()), mode:request.mode==='refresh'?'refresh':'insert' }
}

export function validateNativeToc(prepared, output, headings) {
  const before=unpack(prepared),after=unpack(output)
  const source=xml(before['word/document.xml']),result=xml(after['word/document.xml'])
  const blocks=tocBlocks(result)
  if(blocks.length!==1||blocks[0].parentNode!==result.getElementsByTagNameNS(W,'body')[0])throw problem('原生目录生成结果无效',422)
  const links=[...blocks[0].getElementsByTagNameNS(W,'hyperlink')]
  if(!links.length)throw problem('原生目录没有可跳转的章节条目',422)
  if (headings) {
    const labels = [...blocks[0].getElementsByTagNameNS(W,'p')].map(p => {
      let value = ''
      const read = node => {
        if(node.namespaceURI===W && node.localName==='t')value+=node.textContent
        else if(node.namespaceURI===W && node.localName==='tab')value+='\t'
        else for(const child of elements(node))read(child)
      }
      read(p)
      return value.split('\t')[0].trim()
    }).filter(Boolean)
    const matches = (value, index) => value === headings[index] ||
      value.startsWith(headings[index]) && /^[0-9ivxlcdm]+$/i.test(value.slice(headings[index].length).trim())
    if (labels.length !== headings.length || labels.some((value,index)=>!matches(value,index))) throw problem(`目录条目与正文标题不一致（实际${labels.length}项，预期${headings.length}项），拒绝保存`,422)
  }
  const anchors=new Set(links.map(n=>n.getAttributeNS(W,'anchor')))
  const bookmarks=[...result.getElementsByTagNameNS(W,'bookmarkStart')]
  if([...anchors].some(name=>!name||!bookmarks.some(n=>n.getAttributeNS(W,'name')===name)))throw problem('目录跳转目标缺失',422)
  const clean=doc=>{
    const old=tocBlocks(doc)
    const names=new Set(old.flatMap(b=>[...b.getElementsByTagNameNS(W,'hyperlink')].map(n=>n.getAttributeNS(W,'anchor'))))
    old.forEach(b=>b.parentNode.removeChild(b))
    const ids=new Set()
    for(const n of [...doc.getElementsByTagNameNS(W,'bookmarkStart')])if(names.has(n.getAttributeNS(W,'name'))&&/^_Toc/.test(n.getAttributeNS(W,'name'))){ids.add(n.getAttributeNS(W,'id'));n.parentNode.removeChild(n)}
    for(const n of [...doc.getElementsByTagNameNS(W,'bookmarkEnd')])if(ids.has(n.getAttributeNS(W,'id')))n.parentNode.removeChild(n)
    for(const n of [...doc.getElementsByTagName('*')]){n.removeAttributeNS(W14,'paraId');n.removeAttributeNS(W14,'textId')}
  }
  clean(source);clean(result)
  if(canonical(source.documentElement)!==canonical(result.documentElement))throw problem('原生目录操作改变了非目标正文或格式，拒绝保存',422)
  before['word/document.xml']=encode(source);after['word/document.xml']=encode(result)
  const validation=compareProtected(zipSync(before),zipSync(after))
  validation.checks.push('真实TOC域及章节跳转验证','非目录正文与格式保持')
  return {...validation,limitations:[...validation.limitations,'目录分页以当前编辑器结果为准，跨软件视觉尚未验收']}
}
