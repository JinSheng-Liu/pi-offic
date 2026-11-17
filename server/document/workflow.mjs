import { runDocumentJob } from './document-jobs.mjs'

// One tool invocation is one atomic batch. Earlier committed batches survive cancellation.
export class DocumentWorkflow {
  constructor(store) { this.store = store }
  async execute(id, revision, batch, signal) {
    return this.store.exclusive(id, revision, async meta => {
      signal?.throwIfAborted()
      const source = await this.store.bytes(id, revision)
      const reference = ['referenceCover', 'referenceAuthorization', 'referenceDeclaration'].includes(batch.kind)
        ? await this.store.bytes(batch.referenceId, batch.referenceRevision) : undefined
      const {result, summary, validation} = await runDocumentJob('batch', {source, reference, batch, revision}, signal)
      signal?.throwIfAborted()
      const document = await this.store.commit(meta, result.bytes, summary, validation, result.replay, this.store.taskContext.getStore())
      return {document, affected:result.affected, validation,
        ...(result.captions ? {captions:result.captions} : {})}
    })
  }
}
