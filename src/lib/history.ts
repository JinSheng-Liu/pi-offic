import type { HistoryGroup, HistoryTask, Version } from '../types'

export function groupHistory(versions: Version[] = [], tasks: Record<string, HistoryTask> = {}): HistoryGroup[] {
  const groups: HistoryGroup[] = []
  const reversed = [...versions].reverse()

  for (const version of reversed) {
    const taskId = version.task?.id
    if (!taskId) {
      groups.push({
        id: `rev-${version.revision}`,
        versions: [version],
      })
      continue
    }

    const resolvedTask: HistoryTask = {
      ...version.task!,
      ...(tasks[taskId] || {}),
    }
    const previous = groups.at(-1)
    if (previous && previous.task?.id === taskId) {
      previous.versions.push(version)
    } else {
      groups.push({
        id: `task-${taskId}-${version.revision}`,
        task: resolvedTask,
        versions: [version],
      })
    }
  }

  return groups
}
