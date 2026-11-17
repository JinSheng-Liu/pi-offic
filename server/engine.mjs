// Compatibility facade. Document writes are coordinated by DocumentWorkflow.
export { MAX_BYTES, digest, problem, unpack, compareProtected } from './document/package.mjs'
export { inspect } from './document/structure.mjs'
export { validateOperations } from './document/rules.mjs'
export { format } from './document/editor-adapter.mjs'
export { insertCaptions } from './document/ooxml-captions.mjs'
export { numberEquations, renumberHeadings } from './document/numbering-operations.mjs'
