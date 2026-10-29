import { describe, test, expect } from 'bun:test'
import { mkdtempSync, writeFileSync, readFileSync, existsSync, rmSync } from 'fs'
import { join } from 'path'
import { tmpdir } from 'os'
import { moveFile } from '../src/teams/attachments.ts'

function withDir(fn: (dir: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'move-test-'))
  try { fn(dir) } finally { rmSync(dir, { recursive: true, force: true }) }
}
const codeErr = (msg: string, code: string) => Object.assign(new Error(msg), { code })

describe('moveFile', () => {
  test('a plain rename moves the file', () => withDir(dir => {
    const src = join(dir, 'src'), dst = join(dir, 'dst')
    writeFileSync(src, 'abc')
    moveFile(src, dst)
    expect(readFileSync(dst, 'utf8')).toBe('abc')
    expect(existsSync(src)).toBe(false)
  }))

  test('EXDEV falls back to copy and unlink', () => withDir(dir => {
    const src = join(dir, 'src'), dst = join(dir, 'dst')
    writeFileSync(src, 'payload')
    const calls: [string, string][] = []
    moveFile(src, dst, (a, b) => {
      calls.push([a, b])
      throw codeErr('EXDEV: cross-device link not permitted', 'EXDEV')
    })
    expect(calls).toEqual([[src, dst]])
    expect(readFileSync(dst, 'utf8')).toBe('payload')
    expect(existsSync(src)).toBe(false)
  }))

  test('another rename error is rethrown and nothing is copied', () => withDir(dir => {
    const src = join(dir, 'src'), dst = join(dir, 'dst')
    writeFileSync(src, 'x')
    const err = codeErr('EACCES: denied', 'EACCES')
    let caught: unknown
    try { moveFile(src, dst, () => { throw err }) } catch (e) { caught = e }
    expect(caught).toBe(err)
    expect(existsSync(src)).toBe(true)
    expect(existsSync(dst)).toBe(false)
  }))

  test('EXDEV onto an existing destination keeps both files', () => withDir(dir => {
    const src = join(dir, 'src'), dst = join(dir, 'dst')
    writeFileSync(src, 'new')
    writeFileSync(dst, 'old')
    let caught: any
    try {
      moveFile(src, dst, () => { throw codeErr('EXDEV', 'EXDEV') })
    } catch (e) { caught = e }
    expect(caught?.code).toBe('EEXIST')
    expect(readFileSync(dst, 'utf8')).toBe('old')
    expect(readFileSync(src, 'utf8')).toBe('new')
  }))

  test('a copy that fails on the source leaves an existing destination alone', () => withDir(dir => {
    const src = join(dir, 'missing'), dst = join(dir, 'dst')
    writeFileSync(dst, 'old')
    let caught: any
    try {
      moveFile(src, dst, () => { throw codeErr('EXDEV', 'EXDEV') })
    } catch (e) { caught = e }
    expect(caught?.code).toBe('ENOENT')
    expect(readFileSync(dst, 'utf8')).toBe('old')
  }))

  test('processOne moves the temp file with moveFile', () => {
    const text = readFileSync(join(import.meta.dir, '..', 'src', 'teams', 'attachments.ts'), 'utf8')
    expect(text).not.toContain('renameSync(tempPath')
    expect(text.split('moveFile(tempPath, finalPath)').length - 1).toBe(2)
  })
})
