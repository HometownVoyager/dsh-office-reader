/**
 * 测试小工具：用最小 ctx 桩把 dsh-office-reader 的工具注册出来，直接调用 execute。
 *
 * 工具渲染（render）与返回值是分开的，所以这里断言的是 execute 的原始返回值；要验证渲染文案时
 * 用 `render(name, value)` 单独跑一遍。
 * 用法：`const { tools, exec, render } = await loadReaderTools()`
 */
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const libArg = process.argv[2] ?? 'dsh-office-reader'
const readerLib = join(here, '..', '..', libArg, 'lib', 'index.js')

/** Register the reader plugin against a stub ctx and return its tools. */
export async function loadReaderTools(pluginPath = readerLib) {
  const module = await import(new URL(`file://${pluginPath.replace(/\\/gu, '/')}`).href)
  const registered = []
  const ctx = {
    effect: (fn) => fn(),
    tools: { register: (tool) => { registered.push(tool); return () => {} } },
  }
  // Fill the plugin's Config defaults the way the loader would: schemastery schemas are callable and
  // resolve every `default()` for us, so the tests run against the deployed defaults.
  let config = {}
  try {
    config = module.Config({})
  } catch {
    config = {}
  }
  module.apply(ctx, config)
  const tools = new Map(registered.map((tool) => [tool.name, tool]))
  const exec = {
    // A real AbortSignal: the tools forward it into node:fs, which validates the instance.
    signal: new AbortController().signal,
    agent: { session: { header: { cwd: process.cwd() } } },
  }
  return {
    tools,
    exec,
    config,
    /** Run one tool and return its raw result. */
    async call(name, args) {
      const tool = tools.get(name)
      if (tool === undefined) throw new Error(`工具未注册：${name}（已注册：${[...tools.keys()].join('、')}）`)
      return tool.execute(args, exec)
    },
    /** Render one tool's result the way the model would see it (tool render takes `(args, value)`). */
    render(name, value, args = {}) {
      const renderFn = tools.get(name)?.output?.render
      return renderFn === undefined ? String(value) : renderFn(args, value)
    },
  }
}

/** Tiny assertion helper shared by the P2 reader tests. */
export function makeChecker(label) {
  let failures = 0
  let checks = 0
  const check = (name, condition, detail = '') => {
    checks++
    if (condition) console.log(`  ok   - ${name}`)
    else { failures++; console.log(`  FAIL - ${name} ${detail}`) }
  }
  return {
    check,
    done() {
      console.log(`\n${failures === 0 ? `ALL PASS（${checks} 项）` : `${failures} FAILURES（${checks} 项）`} — ${label}`)
      return failures === 0 ? 0 : 1
    },
  }
}
