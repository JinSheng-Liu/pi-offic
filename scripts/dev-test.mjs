import { resolve } from 'node:path'

// Override any .env data directory before importing the backend through Vite.
process.env.PI_DOCX_DATA_DIR = resolve('data/testing/documents')
process.env.PI_DOCX_TEST_MODE = '1'
process.argv = [process.argv[0], process.argv[1], '--configLoader', 'native', '--host', '127.0.0.1', '--port', '5175', '--strictPort']
console.log('测试工作区：http://127.0.0.1:5175；数据目录：data/testing/documents')
await import('../node_modules/vite/bin/vite.js')
