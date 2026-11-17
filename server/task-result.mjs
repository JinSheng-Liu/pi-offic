import {readFile,writeFile} from 'node:fs/promises'
import {atomicReplace} from './atomic-replace.mjs'
import {join} from 'node:path'
import {randomUUID} from 'node:crypto'

export async function saveTaskResult(folder,result) {
  const temporary=join(folder,`task-${randomUUID()}.tmp`)
  await writeFile(temporary,JSON.stringify(result,null,2),{encoding:'utf8',flag:'wx'})
  await atomicReplace(temporary,join(folder,'last-task.json'))
}
export async function readTaskResult(folder) {
  try{return JSON.parse(await readFile(join(folder,'last-task.json'),'utf8'))}
  catch(error){if(error.code==='ENOENT')return null;throw error}
}
export function taskResult(state,startRevision,currentRevision,commits,error,plans=[]) {
  const conversationState=state
  if(state==='complete' && commits.length)state='partial'
  return {state,conversationState,plans,startRevision,currentRevision,committedRevisions:commits.map(c=>({revision:c.revision,summary:c.summary})),
    requirementsVerified:false,finishedAt:new Date().toISOString(),...(error?{error}:{}),
    guidance:conversationState==='complete'?'本轮对话已结束；已提交修改不代表整份需求验收通过。请核对执行清单、未完成项目和页面效果。':
      commits.length?'本轮已提交的版本保留。继续前读取当前版本和目标结果，仅处理未完成项，不重复已提交操作。':'本轮没有提交修改。继续前仍需核对当前版本。'}
}
