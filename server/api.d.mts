import type { IncomingMessage, ServerResponse } from 'node:http'

export function createApiMiddleware(options?: { staticRoot?: string }):
  (req: IncomingMessage, res: ServerResponse, next?: () => void) => Promise<void>
export function startServer(options?: { port?: number; staticRoot?: string }): ReturnType<typeof import('node:http').createServer>
