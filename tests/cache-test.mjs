/**
 * 解析缓存单测：命中不重读、mtime 失效、LRU 淘汰、并发去重。
 * 用法：node tests/cache-test.mjs
 */
import { readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { performance } from 'node:perf_hooks'
import { OfficeCache } from '../lib/cache.js'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const fixture = (name) => join(root, 'tests', 'fixtures', name)
let failures = 0
const check = (name, condition, detail = '') => {
  if (condition) console.log(`  ok - ${name}`)
  else { failures++; console.log(`  FAIL - ${name} ${detail}`) }
}

const caps = { maxFileBytes: 200 * 1024 * 1024, maxDownloadBytes: 100 * 1024 * 1024, downloadTimeoutMs: 30000, allowedHosts: [] }
const exec = { signal: new AbortController().signal }

// ① 命中：第二次 open 明显更快（跳过读盘+解析）
const cache = new OfficeCache({ maxEntries: 6, maxChars: 30_000_000 })
const t0 = performance.now()
await cache.open(fixture('big.xlsx'), exec, caps)
const t1 = performance.now()
const doc1 = await cache.open(fixture('big.xlsx'), exec, caps)
const t2 = performance.now()
check(`首次解析 ${(t1 - t0).toFixed(1)}ms > 命中 ${(t2 - t1).toFixed(2)}ms`, t2 - t1 < (t1 - t0) / 2)

// ② 命中返回同一对象（零拷贝）
const doc2 = await cache.open(fixture('big.xlsx'), exec, caps)
check('命中返回同一实例', doc1 === doc2)

// ③ mtime/size 失效：同路径换内容后重新解析（旧条目被清）
import { zipSync, strToU8 } from 'fflate'
const makeDocx = (marker) => zipSync({
  'word/document.xml': strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标记 ${marker}</w:t></w:r></w:p></w:body></w:document>`),
})
const tmp = join(root, 'tests', 'fixtures', 'tmp-mtime.docx')
await writeFile(tmp, makeDocx('甲'))
const first = await cache.open(tmp, exec, caps)
check('修改前内容', first.units[0].text.includes('标记 甲'), first.units[0].text)
await writeFile(tmp, makeDocx('乙'))
const second = await cache.open(tmp, exec, caps)
check('mtime/size 失效后重新解析', second !== first)
check('失效后内容正确', second.units[0].text.includes('标记 乙'), second.units[0].text)

// ④ LRU 淘汰：maxEntries=1 时新条目挤掉旧条目
const tiny = new OfficeCache({ maxEntries: 1, maxChars: 30_000_000 })
const a = await tiny.open(tmp, exec, caps)
const b = await tiny.open(fixture('sample.xlsx'), exec, caps)
check('LRU 淘汰旧条目', a !== b)
const aAgain = await tiny.open(tmp, exec, caps)
check('被淘汰后重新解析为新实例', aAgain !== a)

// ⑤ 并发去重：同 key 并发共享一次解析
const shared = new OfficeCache({ maxEntries: 6, maxChars: 30_000_000 })
const [x, y] = await Promise.all([
  shared.open(fixture('big.xlsx'), exec, caps),
  shared.open(fixture('big.xlsx'), exec, caps),
])
check('并发共享同一实例', x === y)

console.log(failures === 0 ? '\nALL PASS' : `\n${failures} FAILURES`)
process.exit(failures === 0 ? 0 : 1)
