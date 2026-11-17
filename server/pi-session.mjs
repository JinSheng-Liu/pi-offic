import {AgentSession} from '../vendor/pi/packages/coding-agent/src/core/agent-session.ts'
import {SettingsManager} from '../vendor/pi/packages/coding-agent/src/core/settings-manager.ts'
import {DefaultResourceLoader} from '../vendor/pi/packages/coding-agent/src/core/resource-loader.ts'
import {ModelRuntime} from '../vendor/pi/packages/coding-agent/src/core/model-runtime.ts'
import {AuthStorage} from '../vendor/pi/packages/coding-agent/src/core/auth-storage.ts'
import {convertToLlm} from '../vendor/pi/packages/coding-agent/src/core/messages.ts'

export async function attachPiSession({agent,sessionManager,agentDir,systemPrompt,tools,apiKey,compaction}){
  const settingsManager=SettingsManager.inMemory({compaction:{enabled:true,...compaction}})
  const resourceLoader=new DefaultResourceLoader({cwd:agentDir,agentDir,settingsManager,
    noExtensions:true,noSkills:true,noPromptTemplates:true,noThemes:true,
    systemPromptOverride:()=>systemPrompt,agentsFilesOverride:()=>({agentsFiles:[]})})
  await resourceLoader.reload()
  const modelRuntime=await ModelRuntime.create({credentials:AuthStorage.inMemory({openai:{type:'api_key',key:apiKey}}),modelsPath:null,refreshOnCreate:false})
  agent.convertToLlm=convertToLlm
  return new AgentSession({agent,sessionManager,settingsManager,cwd:agentDir,resourceLoader,modelRuntime,
    baseToolsOverride:Object.fromEntries(tools.map(tool=>[tool.name,tool])),initialActiveToolNames:tools.map(tool=>tool.name),allowedToolNames:tools.map(tool=>tool.name)})
}
