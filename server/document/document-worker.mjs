import { parentPort, workerData } from 'node:worker_threads'
import { computeBatch } from './batch-compute.mjs'
import { auditPlan } from './plan-audit.mjs'
import { withPackageCache } from './package.mjs'

try {
  const { kind, data } = workerData
  const result = await withPackageCache(async () => kind === 'batch'
    ? await computeBatch(data)
    : kind === 'audit' ? await auditPlan(data.bytes, data.steps)
      : (() => { throw Error('未知后台文档任务') })())
  parentPort.postMessage({ result })
} catch (error) {
  parentPort.postMessage({ error: { message: error.message, name: error.name, status: error.status } })
}
