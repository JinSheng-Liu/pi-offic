import { startServer } from './api.mjs'

const server = startServer()
server.on('listening', () => console.log(`Pi DOCX Studio listening on http://127.0.0.1:${server.address().port}`))
for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, () => server.close(() => process.exit(0)))
