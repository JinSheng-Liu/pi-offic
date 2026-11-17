// All editor mutations share one lane. UI-only events stay responsive, and a
// failed update cannot prevent the server's next native-editor request replying.
export function createEditorEventHandler(
  handle: (event: string, payload: any) => void | Promise<void>,
) {
  let tail = Promise.resolve()
  return (event: string, payload: any): void | Promise<void> => {
    if (event !== 'document' && event !== 'editor_request') return handle(event, payload)
    const task = tail.then(() => handle(event, payload))
    tail = task.catch(() => {})
    return task
  }
}
