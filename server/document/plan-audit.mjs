import {inspect} from './structure.mjs'

// Checks only explicit plan targets. A model-authored plan is not evidence that
// every requirement in natural language was captured, nor that pages look right.
export async function auditPlan(bytes,steps){
  const data=await inspect(bytes),expected=new Map(),sectionExpected=new Map(),footerExpected=new Map(),checks=[],unsupported=[]
  for(const step of steps)for(const op of step.operations){
    if(op.scope==='ensurePageNumbers'){
      for(const section of data.pageNumbers)checks.push({label:step.label,target:`section:${section.index}.pageNumber`,passed:section.format==='decimal'&&section.footers.some(f=>f.type==='default')&&section.footers.every(f=>f.valid&&f.pageFields.length===1&&f.pageFields[0].alignment==='center')})
      continue
    }
    if(op.scope==='fitTableWidth'){
      const table=data.tables[op.tableIndex]
      checks.push({label:step.label,target:`table:${op.tableIndex}.width`,passed:Boolean(table&&table.availableWidthTwips>0&&table.declaredWidth.type==='dxa'&&table.declaredWidth.twips>0&&table.declaredWidth.twips<=table.availableWidthTwips)})
      if (op.columnRatios) {
        const widths = table?.gridWidthsTwips || []
        const total = widths.reduce((sum, width) => sum + width, 0)
        const weight = op.columnRatios.reduce((sum, ratio) => sum + ratio, 0)
        const expected = op.columnRatios.map(ratio => Math.floor(total * ratio / weight))
        expected[expected.length - 1] += total - expected.reduce((sum, width) => sum + width, 0)
        checks.push({ label: step.label, target: `table:${op.tableIndex}.columnRatios`,
          passed: total > 0 && widths.length === expected.length && widths.every((width, index) => width === expected[index]) })
      }
      continue
    }
    if(op.scope==='footerPageNumber'){
      for(const index of op.sectionIndices){
        const e=footerExpected.get(index)||{font:{}}
        if(op.alignment)e.alignment=op.alignment==='Centered'?'center':op.alignment.toLowerCase()
        Object.assign(e.font,op.font);e.label=step.label;footerExpected.set(index,e)
      }
      continue
    }
    if(op.scope==='sections'){
      for(const index of op.sectionIndices)for(const kind of ['page','margins'])for(const [key,value]of Object.entries(op[kind]||{}))
        sectionExpected.set(`${index}.${kind}.${key}`,{index,kind,key,value,label:step.label})
      continue
    }
    const supported=['namedStyle','all','tableText','inlinePictures','headings','indices']
    if(!supported.includes(op.scope)){unsupported.push({label:step.label,scope:op.scope,reason:'此类型尚未纳入跨批次最终属性验收；单批校验不等于全文完成'});continue}
    const ps=data.paragraphs.filter(p=>op.scope==='all'||op.scope==='tableText'&&p.tableNestingLevel>0||op.scope==='inlinePictures'&&p.inlinePictureCount>0&&!p.text.trim()||op.scope==='headings'&&p.heading>0&&(!op.heading||p.heading===op.heading)||['indices','namedStyle'].includes(op.scope)&&op.indices.includes(p.index))
    if(!ps.length)checks.push({label:step.label,target:op.scope,passed:false,reason:'没有匹配目标'})
    for(const p of ps){
      const e=expected.get(p.index)||{font:{},paragraph:{}}
      Object.assign(e.font,op.font);Object.assign(e.paragraph,op.paragraph)
      if(op.paragraph?.lineSpacingMultiple!==undefined)delete e.paragraph.lineSpacing
      if(op.paragraph?.lineSpacing!==undefined)delete e.paragraph.lineSpacingMultiple
      if(op.font?.name&&!op.font.latinName)e.font.latinName=op.font.name
      if(op.style){e.style=op.style;delete e.styleName}
      if(op.scope==='namedStyle'){e.styleName=op.styleName.trim();delete e.style}
      e.label=step.label;expected.set(p.index,e)
    }
  }
  for(const [index,e]of footerExpected){
    const footers=data.pageNumbers[index]?.footers||[],fields=footers.flatMap(f=>f.pageFields)
    checks.push({label:e.label,target:`section:${index}.footerPageNumber`,passed:footers.every(f=>f.valid)&&fields.length>0&&fields.every(f=>(!e.alignment||f.alignment===e.alignment)&&(!Object.keys(e.font).length||f.fonts.length>0&&f.fonts.every(font=>(!e.font.name||['ascii','hAnsi','eastAsia'].every(key=>font[key]===e.font.name))&&(e.font.size===undefined||font.size===Math.round(e.font.size*2)/2))))})
  }
  for(const [target,e]of sectionExpected)checks.push({label:e.label,target:`section:${target}`,passed:data.sections[e.index]?.[e.kind]?.[e.key]===e.value})
  for(const [index,e]of expected){
    const p=data.paragraphs[index]
    const check=(key,passed)=>checks.push({label:e.label,target:`paragraph:${index}.${key}`,passed})
    if(e.style)check('style',p.style===e.style)
    if(e.styleName)check('styleName',[e.styleName,`${e.styleName} · ${p.style.slice(-6)}`].includes(p.styleName))
    for(const [key,value]of Object.entries(e.font)){
      if(key==='name')check('eastAsia',p.font.effectiveFonts.length>0&&p.font.effectiveFonts.every(f=>f.eastAsia===value&&!f.eastAsiaTheme))
      else if(key==='latinName')check('latin',p.font.effectiveFonts.length>0&&p.font.effectiveFonts.every(f=>f.ascii===value&&f.hAnsi===value&&!f.asciiTheme&&!f.hAnsiTheme))
      else if(Object.hasOwn(p.font,key))check('font.'+key,key==='color'?String(p.font[key]).replace(/^#/,'').toLowerCase()===String(value).replace(/^#/,'').toLowerCase():p.font[key]===value)
      else unsupported.push({label:e.label,target:index,property:key,reason:'尚无对应最终属性读数'})
    }
    for(const [key,value]of Object.entries(e.paragraph)){
      if(key==='leadingTabPosition'||key==='tocPageTabPosition'){
        const tabs=p.tabStops.direct
        check(key,tabs.length===1&&tabs[0].positionTwips===Math.round(value*20)&&tabs[0].alignment===(key==='leadingTabPosition'?'left':'right'))
      }
      else if(key==='lineSpacingMultiple')check(key,p.lineSpacing.rule==='auto'&&p.lineSpacing.multiple===value)
      else if(Object.hasOwn(p.pagination,key))check(key,p.pagination[key]===value)
      else if(Object.hasOwn(p.paragraph,key))check(key,p.paragraph[key]===(({Center:'Centered',Justify:'Justified'})[value]||value))
      else unsupported.push({label:e.label,target:index,property:key,reason:'尚无对应最终属性读数'})
    }
  }
  const failures=checks.filter(c=>!c.passed)
  return {checked:checks.length,passed:checks.length-failures.length,failures,unsupported,
    verifiedTargets:checks.filter(c=>c.passed).map(c=>c.target),
    failedTargets:failures.map(c=>c.target),
    planTargetsVerified:checks.length>0&&checks.every(c=>c.passed)&&unsupported.length===0,requirementsVerified:false,visualVerified:false,
    note:'只验证已声明的计划目标，不代表自然语言需求无遗漏，也不代表页面效果通过。'}
}
