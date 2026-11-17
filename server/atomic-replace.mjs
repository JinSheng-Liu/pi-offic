import {rename} from 'node:fs/promises'
import {setTimeout} from 'node:timers/promises'

// Windows readers or antivirus can briefly hold the destination open.
// Never unlink it: the previous metadata must remain readable on failure.
export async function atomicReplace(source,target,{move=rename,wait=setTimeout,platform=process.platform}={}){
  for(let attempt=0;;attempt++){
    try{return await move(source,target)}catch(error){
      if(platform!=='win32'||!['EPERM','EACCES','EBUSY'].includes(error.code)||attempt>=4)throw error
      await wait(25*2**attempt)
    }
  }
}
