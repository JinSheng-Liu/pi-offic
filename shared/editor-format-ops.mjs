import {bodyStoryRoot,storyParagraphs,directParagraphProperties,mergedProperties,
  mergedParagraphMarkProperties,runPropertyEdits,paragraphTextOf} from '@docx-editor.dev/core/store'

const property=(localName,attributes)=>({localName,attributes:Object.fromEntries(Object.entries(attributes).map(([k,v])=>[k,String(v)]))})
export function paragraphProperties(format={}) {
  const properties=[],spacing={},ind={}
  if(format.alignment)properties.push(property('jc',{val:({Left:'left',Centered:'center',Center:'center',Right:'right',Justified:'both',Justify:'both'})[format.alignment]}))
  for(const [key,name] of [['spaceBefore','before'],['spaceAfter','after']])if(format[key]!==undefined)spacing[name]=Math.round(format[key]*20)
  if(format.lineSpacing!==undefined)Object.assign(spacing,{line:Math.round(format.lineSpacing*20),lineRule:'exact'})
  if(format.lineSpacingMultiple!==undefined)Object.assign(spacing,{line:Math.round(format.lineSpacingMultiple*240),lineRule:'auto'})
  for(const [key,name] of [['leftIndent','left'],['rightIndent','right']])if(format[key]!==undefined)ind[name]=Math.round(format[key]*20)
  if(format.firstLineIndent!==undefined)Object.assign(ind,format.firstLineIndent<0?{hanging:Math.round(-format.firstLineIndent*20)}:{firstLine:Math.round(format.firstLineIndent*20)})
  if(Object.keys(spacing).length)properties.push(property('spacing',spacing))
  if(Object.keys(ind).length)properties.push(property('ind',ind))
  for(const key of ['keepNext','keepLines','widowControl','pageBreakBefore'])if(format[key]!==undefined)properties.push(property(key,{val:format[key]?'1':'0'}))
  return properties
}
export function fontProperties(font={}) {
  const properties=[]
  if(font.name||font.latinName)properties.push(property('rFonts',{
    ...(font.name?{ascii:font.name,hAnsi:font.name,eastAsia:font.name,cs:font.name}:{}),
    ...(font.latinName?{ascii:font.latinName,hAnsi:font.latinName}:{}),
  }))
  if(font.size!==undefined)for(const name of ['sz','szCs'])properties.push(property(name,{val:Math.round(font.size*2)}))
  for(const [key,names] of [['bold',['b','bCs']],['italic',['i','iCs']],['strikeThrough',['strike']],['doubleStrikeThrough',['dstrike']]])
    if(font[key]!==undefined)for(const name of names)properties.push(property(name,{val:font[key]?'1':'0'}))
  if(font.color)properties.push(property('color',{val:font.color.replace(/^#/,'')}))
  if(font.underline)properties.push(property('u',{val:({None:'none',Single:'single',Double:'double',Dotted:'dotted',Dash:'dash',Wave:'wave'})[font.underline]}))
  if(font.subscript!==undefined||font.superscript!==undefined)properties.push(property('vertAlign',{val:font.subscript?'subscript':font.superscript?'superscript':'baseline'}))
  return properties
}
/**
 * @param {object} part
 * @param {number[]} indices
 * @param {object} op
 * @param {string[] | undefined} [paragraphIds]
 */
export function compileFormatOps(part,indices,op,paragraphIds) {
  const ids=paragraphIds ?? storyParagraphs(bodyStoryRoot(part)).map(node=>node.id),ops=[]
  const font=fontProperties(op.font),paragraph=paragraphProperties(op.paragraph)
  for(const index of indices) {
        const paragraphId=ids[index]
        if(!paragraphId)throw new Error('编辑器段落索引超出范围')
        if(paragraph.length) {
          let props=directParagraphProperties(part,paragraphId)
          for(const incoming of paragraph) {
            const attributes={...(props.find(p=>p.localName===incoming.localName)?.attributes||{}),...incoming.attributes}
            if(incoming.localName==='ind') {
              if('firstLine' in incoming.attributes)delete attributes.hanging
              if('hanging' in incoming.attributes)delete attributes.firstLine
              for(const [point,char] of [['firstLine','firstLineChars'],['hanging','hangingChars'],['left','leftChars'],['right','rightChars']])
                if(point in incoming.attributes)delete attributes[char]
              if('firstLine' in incoming.attributes)delete attributes.hangingChars
              if('hanging' in incoming.attributes)delete attributes.firstLineChars
            }
            props=mergedProperties(props,{...incoming,attributes})
          }
          ops.push({op:'setParagraphProperties',paragraphId,properties:props})
        }
        if(font.length) {
          const start=op.range?.start??0,end=op.range?.end??(paragraphTextOf(part,paragraphId)?.length||0)
          for(const edit of runPropertyEdits(part,paragraphId,start,end,font))
            ops.push({op:'setRunProperties',paragraphId,...edit})
          if(!op.range)ops.push({op:'setParagraphMarkProperties',paragraphId,properties:mergedParagraphMarkProperties(part,paragraphId,font)})
        }
  }
  return ops
}
