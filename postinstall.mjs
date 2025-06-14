import fs from 'node:fs'

await fs.promises.mkdir('public', { recursive: true })
await fs.promises.copyFile('node_modules/xlsx-populate/browser/xlsx-populate.min.js', 'public/xlsx-populate.min.js')
