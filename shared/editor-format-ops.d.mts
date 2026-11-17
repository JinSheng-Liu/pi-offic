import type {OoxmlPart,TreeDocOp,OoxmlProperty} from '@docx-editor.dev/core/store'
export interface FormatTarget {indices:number[];op:{font?:Record<string,string|number|boolean>;paragraph?:Record<string,string|number|boolean>;range?:{start:number;end:number}}}
export function compileFormatOps(part:OoxmlPart,indices:number[],op:FormatTarget['op']):TreeDocOp[]
export function paragraphProperties(format?:FormatTarget['op']['paragraph']):OoxmlProperty[]
export function fontProperties(font?:FormatTarget['op']['font']):OoxmlProperty[]
