import test from 'node:test'
import assert from 'node:assert/strict'
import {format} from '../server/engine.mjs'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
import {unpack,xml,W} from '../server/document/package.mjs'
import {paragraphDocument} from './helpers/docx-fixture.mjs'
import {zipSync} from 'fflate'
import {auditPlan} from '../server/document/plan-audit.mjs'
const op={scope:'namedStyle',indices:[0],styleName:'公文·标题',font:{name:'黑体',latinName:'Times New Roman',size:22},paragraph:{alignment:'Centered',firstLineIndent:0,spaceAfter:6,outlineLevel:10}}

test('explicit named fonts override inherited theme slots without changing original theme or unselected paragraph',async()=>{
 const parts=unpack(paragraphDocument(['标题 English','原文保留']))
 parts['word/styles.xml']=new TextEncoder().encode(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`)
 const result=await format(zipSync(parts),[op]),read=readEditorDocument(result.bytes)
 const fonts=read.paragraphs[0].runs[0].effectiveFonts
 assert.equal(fonts.eastAsia,'黑体');assert.equal(fonts.ascii,'Times New Roman')
 assert.equal(fonts.eastAsiaTheme,undefined);assert.equal(fonts.asciiTheme,undefined)
 assert.equal(read.paragraphs[1].runs[0].effectiveFonts.eastAsiaTheme,'minorEastAsia')
})

test('audit rejects a theme competing with explicit fonts at the same level',async()=>{
 const formatted=await format(paragraphDocument(['正文 English']),[op]),parts=unpack(formatted.bytes)
 const source=new TextDecoder().decode(parts['word/styles.xml'])
 parts['word/styles.xml']=new TextEncoder().encode(source.replace(/<w:rFonts(?=\s)/g,'<w:rFonts w:eastAsiaTheme="minorEastAsia"'))
 const report=await auditPlan(zipSync(parts),[{label:'字体',operations:[op]}])
 assert.ok(report.failures.some(f=>f.target.endsWith('.eastAsia')))
})
test('named style stores formatting in style definition and reuses it without changing other paragraphs',async()=>{
 const input=paragraphDocument(['公文标题 Title','正文保留'])
 const first=await format(input,[op]),read=readEditorDocument(first.bytes)
 assert.match(read.paragraphs[0].styleId,/^PiStyle/)
 assert.equal(read.paragraphs[0].font.size,22)
 assert.equal(read.paragraphs[0].alignment,'Centered')
 assert.equal(read.paragraphs[0].runs[0].value.rPr?.sz,undefined)
 const styles=xml(unpack(first.bytes)['word/styles.xml'])
 const style=Array.from(styles.getElementsByTagNameNS(W,'style')).find(n=>n.getAttributeNS(W,'styleId')===read.paragraphs[0].styleId)
 assert.equal(style.getElementsByTagNameNS(W,'name')[0].getAttributeNS(W,'val'),'公文·标题')
 assert.equal(style.getElementsByTagNameNS(W,'sz')[0].getAttributeNS(W,'val'),'44')
 assert.equal(style.getElementsByTagNameNS(W,'qFormat').length,1)
 const second=await format(first.bytes,[op])
 assert.equal(readEditorDocument(second.bytes).paragraphs[0].styleId,read.paragraphs[0].styleId)
 assert.equal(xml(unpack(second.bytes)['word/styles.xml']).getElementsByTagNameNS(W,'style').length,styles.getElementsByTagNameNS(W,'style').length)
})
test('named style keeps local emphasis and refuses invalid scopes',async()=>{
 const input=await format(paragraphDocument(['正文强调','未选择']),[{scope:'text',indices:[0],text:'强调',font:{bold:true}}])
 const result=await format(input.bytes,[{...op,styleName:'公文·正文',font:{name:'宋体',size:12}}])
 assert.equal(readEditorDocument(result.bytes).paragraphs[0].runs.at(-1).font.bold,true)
 await assert.rejects(format(input.bytes,[{...op,indices:[999]}]))
 await assert.rejects(format(input.bytes,[op,op]))
})

test('named style is reusable on another paragraph and changed definitions do not restyle previous targets',async()=>{
 const first=await format(paragraphDocument(['甲','乙','丙']),[op])
 const second=await format(first.bytes,[{...op,indices:[1]}])
 const before=readEditorDocument(second.bytes)
 assert.equal(before.paragraphs[0].styleId,before.paragraphs[1].styleId)
 const third=await format(second.bytes,[{...op,indices:[1],font:{...op.font,size:16}}])
 const after=readEditorDocument(third.bytes)
 assert.equal(after.paragraphs[0].font.size,22)
 assert.equal(after.paragraphs[1].font.size,16)
 assert.notEqual(after.paragraphs[0].styleId,after.paragraphs[1].styleId)
 assert.notEqual(after.paragraphs[0].styleName,after.paragraphs[1].styleName)
 const fourth=await format(third.bytes,[{...op,indices:[1],font:{...op.font,size:16}}])
 assert.equal(readEditorDocument(fourth.bytes).paragraphs[1].styleId,after.paragraphs[1].styleId)
})
test('named heading styles retain outline and inherited font slots in final plan audit',async()=>{
 const {auditPlan}=await import('../server/document/plan-audit.mjs')
 const heading={...op,styleName:'论文·一级标题',paragraph:{outlineLevel:1,lineSpacingMultiple:1.5,keepNext:true,firstLineIndent:0}}
 const result=await format(paragraphDocument(['第一章 绪论','正文']),[heading])
 assert.equal(readEditorDocument(result.bytes).paragraphs[0].outlineLevel,1)
 const audit=await auditPlan(result.bytes,[{label:'标题样式',operations:[heading]}])
 assert.deepEqual(audit.failures,[])
 assert.deepEqual(audit.unsupported,[])
 assert.equal(audit.planTargetsVerified,true)
})
test('named style replaces conflicting explicit font and indentation but preserves untouched text',async()=>{
 const old=await format(paragraphDocument(['测试 English 123','不改']),[{scope:'indices',indices:[0],font:{name:'ＭＳ 明朝',size:20,color:'#FF0000'},paragraph:{firstLineIndent:-24,lineSpacing:30}}])
 const operation={...op,font:{...op.font,color:'#000000'},paragraph:{firstLineIndent:24,lineSpacingMultiple:1.5}}
 const result=await format(old.bytes,[operation]),p=readEditorDocument(result.bytes).paragraphs[0]
 assert.equal(p.firstLineIndent,24)
 assert.equal(p.font.size,22)
 assert.equal(p.runs[0].effectiveFonts.eastAsia,'黑体')
 assert.equal(p.runs[0].effectiveFonts.ascii,'Times New Roman')
 assert.equal(p.runs[0].value.rPr?.rFonts,undefined)
 assert.equal(p.p.pPr?.ind,undefined)
})

test('inherited theme font attributes with different key order do not reject named styles',async()=>{
 const {zipSync}=await import('fflate')
 const parts=unpack(paragraphDocument(['通知标题']))
 const doc=xml(parts['word/styles.xml'])
 const defaults=doc.getElementsByTagNameNS(W,'rPrDefault')[0]
 const pr=defaults.getElementsByTagNameNS(W,'rPr')[0]
 const fonts=doc.createElementNS(W,'w:rFonts')
 for(const [key,val] of Object.entries({asciiTheme:'minorHAnsi',hAnsiTheme:'minorHAnsi',eastAsiaTheme:'minorEastAsia',cstheme:'minorBidi'}))fonts.setAttributeNS(W,'w:'+key,val)
 pr.appendChild(fonts)
 const {XMLSerializer}=await import('@xmldom/xmldom')
 parts['word/styles.xml']=new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
 const result=await format(zipSync(parts),[{...op,font:{name:'宋体',size:26,color:'#FF0000'}}])
 assert.equal(readEditorDocument(result.bytes).paragraphs[0].font.size,26)
})
