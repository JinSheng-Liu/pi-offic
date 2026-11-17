import {readFileSync} from 'node:fs'
import {zipSync} from 'fflate'
import {createBlankDocument} from '../../server/document/editor-automation.mjs'
import {unpack,W} from '../../server/document/package.mjs'
export const escapeXml=value=>String(value).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;')
export const paragraph=text=>`<w:p><w:r><w:t${/^\s|\s$/.test(text)?' xml:space="preserve"':''}>${escapeXml(text)}</w:t></w:r></w:p>`
export function table(rows) {
 const columns=Math.max(...rows.map(r=>r.length))
 return `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid>${Array.from({length:columns},()=>`<w:gridCol w:w="${Math.floor(9000/columns)}"/>`).join('')}</w:tblGrid>${rows.map(row=>`<w:tr>${row.map(value=>`<w:tc><w:tcPr><w:tcW w:w="${Math.floor(9000/columns)}" w:type="dxa"/></w:tcPr>${paragraph(value)}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`
}
export function documentFixture(blocks) {
 const parts=unpack(createBlankDocument())
 parts['word/document.xml']=new TextEncoder().encode(`<w:document xmlns:w="${W}"><w:body>${blocks}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`)
 return zipSync(parts)
}
export const paragraphDocument=texts=>documentFixture(texts.map(paragraph).join(''))

export function richParagraph(text,pPr='',rPr='') {
 const content=String(text).split(/([\t\n\r\f\u000e\u000b])/).map(t=>t==='\t'?'<w:tab/>':t==='\r'?'<w:cr/>':t==='\f'?'<w:br w:type="page"/>':t==='\u000e'?'<w:br w:type="column"/>':['\n','\u000b'].includes(t)?'<w:br/>':`<w:t${/^\s|\s$/.test(t)?' xml:space="preserve"':''}>${escapeXml(t)}</w:t>`).join('')
 return `<w:p>${pPr?`<w:pPr>${pPr}</w:pPr>`:''}<w:r>${rPr?`<w:rPr>${rPr}</w:rPr>`:''}${content}</w:r></w:p>`
}
export function withInlinePicture(before='',after='') {
 const parts=unpack(readFileSync(new URL('../fixtures/inline-picture.docx',import.meta.url)))
 let source=new TextDecoder().decode(parts['word/document.xml'])
 source=source.replace('<w:body>',`<w:body>${before}`).replace('<w:sectPr',`${after}<w:sectPr`)
 parts['word/document.xml']=new TextEncoder().encode(source)
 return zipSync(parts)
}
