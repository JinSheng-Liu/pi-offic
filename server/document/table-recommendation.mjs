import {W} from './package.mjs'
const children=(n,name)=>Array.from(n?.childNodes||[]).filter(c=>c.namespaceURI===W&&c.localName===name)
const text=n=>Array.from(n.getElementsByTagNameNS(W,'t')).map(t=>t.textContent||'').join('')

// Deliberately narrow: never apply a generic center-all rule to arbitrary tables.
export function databaseTableRecommendation(table,tableIndex,paragraphIndices){
  const rows=children(table,'tr'),cells=rows.map(row=>children(row,'tc'))
  if(rows.length<2||cells.some(row=>row.length!==4)||table.getElementsByTagNameNS(W,'tbl').length)return null
  if(cells.flat().some(cell=>['gridSpan','vMerge','hMerge'].some(name=>cell.getElementsByTagNameNS(W,name).length)))return null
  const headers=cells[0].map(cell=>text(cell).replace(/\s/g,''))
  if(!/^(字段名|字段名称)$/.test(headers[0])||!/^(类型|数据类型|字段类型)$/.test(headers[1])||!/^(约束|约束条件)$/.test(headers[2])||!/^(注释|说明|备注)$/.test(headers[3]))return null
  const sqlType=/^(?:bigint|int|integer|smallint|tinyint|varchar|nvarchar|char|text|longtext|decimal|numeric|float|double|boolean|bool|bit|date|datetime|timestamp|time|blob|json|uuid)(?:\s*\(\s*\d+(?:\s*,\s*\d+)?\s*\))?(?:\s+unsigned)?$/i
  if(cells.slice(1).some(row=>!sqlType.test(text(row[1]).trim())||!/^[A-Za-z_][\w$]*$/.test(text(row[0]).trim())))return null
  const centered=[],left=[],targets=[]
  for(const [rowIndex,row]of cells.entries())for(const [cellIndex,cell]of row.entries()){
    const ps=children(cell,'p')
    if(!ps.length||ps.length!==cell.getElementsByTagNameNS(W,'p').length||ps.some(p=>!paragraphIndices.has(p)))return null
    // This is a content-length heuristic, not a measured visual line count.
    const long=rowIndex>0&&(Array.from(text(cell).trim()).length>24||ps.length>1||cell.getElementsByTagNameNS(W,'br').length>0)
    ;(long?left:centered).push(...ps.map(p=>paragraphIndices.get(p)))
    targets.push({rowIndex,cellIndex})
  }
  const operations=[]
  for(const [indices,alignment]of [[centered,'Centered'],[left,'Left']])if(indices.length)operations.push({scope:'indices',indices,paragraph:{alignment,firstLineIndent:0,leftIndent:0,rightIndent:0,spaceBefore:0,spaceAfter:0}})
  for(let i=0;i<targets.length;i+=100)operations.push({scope:'tableCells',tableIndex,cellTargets:targets.slice(i,i+100),table:{verticalAlignment:'Center'}})
  return {kind:'database-schema',reason:'四列表头及字段标识符、SQL类型匹配',policy:'短内容水平居中、长内容左对齐；单元格垂直居中；清除段落缩进与段间距',longTextThreshold:24,requiresUserIntent:true,operations}
}

// Narrative matrices need reading space, not the database-table center-all policy.
export function analysisTableRecommendation(table, tableIndex, paragraphIndices) {
  const rows = children(table, 'tr'), cells = rows.map(row => children(row, 'tc'))
  if (rows.length < 2 || cells.some(row => row.length !== 3)) return null
  if (table.getElementsByTagNameNS(W, 'tbl').length ||
      cells.flat().some(cell => ['gridSpan', 'vMerge', 'hMerge'].some(name => cell.getElementsByTagNameNS(W, name).length))) return null
  const headers = cells[0].map(cell => text(cell).replace(/\s/g, ''))
  const analysis = /^(问题|存在问题|共性问题)$/.test(headers[0]) &&
    /^(原因|形成原因|原因分析)$/.test(headers[1]) && /^(影响|主要影响|影响分析)$/.test(headers[2])
  const indicators = /^(指标|指标名称)$/.test(headers[0]) &&
    /^(公开数据|数据|统计数据)$/.test(headers[1]) && /^(分析含义|含义|分析说明)$/.test(headers[2])
  if (!analysis && !indicators) return null
  const header = [], body = [], targets = []
  for (const [rowIndex, row] of cells.entries()) for (const [cellIndex, cell] of row.entries()) {
    const paragraphs = children(cell, 'p')
    if (!paragraphs.length || paragraphs.length !== cell.getElementsByTagNameNS(W, 'p').length ||
        paragraphs.some(p => !paragraphIndices.has(p))) return null
    ;(rowIndex === 0 ? header : body).push(...paragraphs.map(p => paragraphIndices.get(p)))
    targets.push({ rowIndex, cellIndex })
  }
  const operations = []
  for (const [indices, alignment] of [[header, 'Centered'], [body, 'Left']]) {
    for (let i = 0; i < indices.length; i += 100) operations.push({ scope: 'indices', indices: indices.slice(i, i + 100),
      paragraph: { alignment, firstLineIndent: 0, leftIndent: 0, rightIndent: 0,
        spaceBefore: 0, spaceAfter: 0, lineSpacingMultiple: 1.15 } })
  }
  for (let i = 0; i < targets.length; i += 100) operations.push({ scope: 'tableCells', tableIndex,
    cellTargets: targets.slice(i, i + 100), table: { verticalAlignment: 'Center' } })
  for (let i = 0; i < rows.length; i += 100) operations.push({ scope: 'tableRows', tableIndex,
    rowIndices: rows.slice(i, i + 100).map((_, offset) => i + offset), table: { autoHeight: true } })
  return { kind: analysis ? 'analysis-matrix' : 'indicator-matrix', reason: analysis ? '问题、原因、影响三列表头匹配' : '指标、公开数据、分析含义三列表头匹配', requiresUserIntent: true,
    policy: '表头居中、内容左对齐；表内1.15倍行距、无缩进段距；自适应行高；用户明确格式优先',
    operations, widthOperation: { scope: 'fitTableWidth', tableIndex, columnRatios: analysis ? [24, 38, 38] : [24, 30, 46] } }
}
