import {digest} from './package.mjs'
// One snapshot per tool session, keyed by actual file bytes, not a stale revision.
export function createInspectionReader(inspect) {
  let key,pending
  return async bytes=>{
    const hash=digest(bytes)
    if(key===hash&&pending)return pending
    key=hash
    const request=Promise.resolve().then(()=>inspect(bytes))
    pending=request
    try{return await request}catch(error){if(pending===request){key=undefined;pending=undefined}throw error}
  }
}
