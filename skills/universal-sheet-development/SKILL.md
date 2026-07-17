---
name: universal-sheet-development
description: Use this skill when working in the universal-sheet repository: TypeScript spreadsheet SDK architecture, package boundaries, code style, Chinese comment/JSDoc rules, testing, validation, and Chinese Conventional Commit rules.
---

# Universal Sheet Development

Use this skill for any code, docs, test, review, refactor, commit, or roadmap work in `/Users/liuwenliang/lwl/universal-sheet`.

## Project Goal

Universal Sheet is an Excel-like Web spreadsheet SDK.

- Architecture evolves by layer: `core -> formula -> engine -> ui-react/ui-vue -> apps`.
- The core must stay framework-agnostic and publishable.
- Develop incrementally. Do not build UI-only placeholders for capabilities that do not exist in the model yet.
- Prefer small verified feature slices over broad refactors.

## Package Boundaries

### `packages/core`

- Pure TypeScript data structures and algorithms.
- No DOM APIs: `window`, `document`, `HTMLElement`, Canvas, browser clipboard.
- No Vue or React imports.
- May depend only on `@universal-sheet/shared`.
- Public exports must go through `packages/core/src/index.ts`.

### `packages/formula`

- Formula tokenizer, parser, AST, evaluator, dependency graph, and function registry.
- Pure computation only.
- No DOM, UI framework, or rendering side effects.
- May depend only on `@universal-sheet/shared`.

### `packages/engine`

- Canvas rendering and input interaction layer.
- May use Canvas 2D APIs, `requestAnimationFrame`, viewport math, and rendering state.
- No Vue or React imports.
- Data changes should flow through `@universal-sheet/core` APIs.

### `packages/ui-react`

- React bindings and reusable React UI.
- Function components only.
- Named exports only.
- Props interface name format: `{ComponentName}Props`.
- Use `memo()` for presentational components receiving props from parents.
- Use `useCallback` for handlers passed to memoized children.

### `packages/ui-vue`

- Vue 3 Composition API only.
- Use `<script setup lang="ts">`.
- Use `defineProps<T>()` and `defineEmits<T>()`.
- No direct DOM queries; use refs and template refs.

### `packages/shared`

- Internal utilities only.
- Do not expose as public consumer API.

## TypeScript Rules

- `strict: true` is mandatory.
- Do not use `any`. Use `unknown` and narrow with type guards.
- Public APIs use `interface` for object shapes and `type` for unions/primitives.
- Every exported function must declare an explicit return type.
- Internal functions must also declare explicit return types.
- Prefer `readonly` for immutable data.
- Use `Array<T>`, not `T[]`.
- Use type-only imports for types.
- Prefer `const`; use `let` only when reassigned.
- No non-null assertions. Narrow with guards or fallback values.
- No `require`, `module.exports`, or CommonJS.
- No default exports.
- No side-effect imports except test/runtime setup where explicitly needed.
- No circular dependencies. Extract shared contracts when two modules need each other.

## Naming

- Interface: PascalCase, no `I` prefix, for example `CellData`.
- Type alias: PascalCase, for example `CellValue`.
- Variables/functions: camelCase, for example `getCellValue`.
- Constants: UPPER_SNAKE_CASE, for example `MAX_ROWS`.
- File names: kebab-case, for example `workbook-model.ts`.
- CSS classes: prefix with `us-`.

## Comment And JSDoc Rules

Comments must be concise Chinese by default.

- Use Chinese comments to describe the purpose of code, not to repeat obvious assignments.
- Inline comments are only for non-obvious logic, edge cases, performance choices, or business rules.
- Do not add noisy comments such as `// 设置变量` or `// 返回结果`.
- Do not leave `TODO` comments unless the user explicitly asks for tracked TODOs.

Every new or modified function must have JSDoc, including internal helpers.

Required JSDoc fields:

- One short Chinese summary sentence.
- `@param` for every parameter.
- `@returns` unless the return type is `void`.
- `@throws` when the function can throw.
- `@author liangzai927`.

Example:

```ts
/**
 * 将工作簿序列化为可保存的 JSON 字符串。
 *
 * @param workbook - 需要序列化的工作簿数据
 * @returns 序列化后的 JSON 字符串
 * @author liangzai927
 */
export function stringifyWorkbookData(workbook: WorkbookData): string {
  return JSON.stringify(serializeWorkbookData(workbook));
}
```

Void example:

```ts
/**
 * 清空历史记录栈。
 *
 * @author liangzai927
 */
export function clearHistory(): void {
  history.clear();
}
```

Throwing example:

```ts
/**
 * 解析工作簿 JSON 字符串。
 *
 * @param json - 工作簿 JSON 字符串
 * @returns 解析后的工作簿数据
 * @throws {SyntaxError} 当 JSON 字符串格式非法时抛出
 * @author liangzai927
 */
export function parseWorkbookData(json: string): WorkbookData {
  return deserializeWorkbookData(JSON.parse(json) as SerializedWorkbookData);
}
```

When editing older functions that lack the new JSDoc shape, update the touched functions only. Do not churn unrelated files just to normalize comments.

## Architecture Rules

- Solve today's requested behavior with the smallest coherent change.
- Do not introduce speculative plugin systems, compatibility layers, or abstractions.
- If a file exceeds about 300 lines because of your change, consider extracting a focused module.
- Extract a utility only when reused at least twice or when it removes real complexity.
- Functions with more than 4 parameters should use an options object.
- Keep data operations immutable in `core`.
- Prefer pure functions over classes unless state and behavior are tightly coupled.
- Public package exports must be added to that package's `src/index.ts`.

## Frontend Rules

- Keep spreadsheet UI dense and work-focused.
- Do not create marketing-style pages for app surfaces.
- Use familiar spreadsheet controls: toolbar, formula bar, sheet tabs, context menus, icons, swatches, segmented controls.
- Avoid in-app explanatory text for obvious actions.
- No inline styles over 2 properties unless matching existing prototype code; prefer classes for new UI.
- Existing playground code may be pragmatic, but new reusable UI should live under package boundaries.

## Testing Rules

- Test framework: Vitest.
- Test location: `src/__tests__/{module-name}.test.ts`.
- Describe block should match the module or function name.
- Test case format: `should {expected behavior} when {condition}`.
- No nested `describe` beyond one level unless there is a strong reason.
- Prefer explicit assertions over snapshots.
- Edge cases are required for public APIs: empty input, boundary values, missing ids, invalid ranges, error paths.
- Do not mock `@universal-sheet/shared` internals.
- Use local test factories when setup repeats.

Validation expectations:

- Narrow core changes: run package-level `pnpm --filter @universal-sheet/core test`, `build`, and `lint`.
- Shared or exported API changes: run root `pnpm test`, `pnpm build`, and `pnpm lint`.
- Always run `git diff --check` before final handoff or commit.

## Lint And Formatting

The repository uses ESLint, Prettier, `simple-import-sort`, and lint-staged.

- Imports order: external packages, `@universal-sheet/*`, relative imports.
- Sort exports with `simple-import-sort`.
- Remove unused imports and variables.
- Use optional chaining where lint expects it.
- Coerce non-string template values explicitly when needed.
- Pre-commit may auto-format staged files; inspect status after committing.

## Git Commit Rules

Use Chinese Conventional Commits.

Format:

```text
<type>: <中文简短描述>

- 具体改动 1
- 具体改动 2

Co-Authored-By: Codex Opus 4.7 <noreply@anthropic.com>
```

Allowed types:

- `feat`: 新功能
- `fix`: Bug 修复
- `refactor`: 重构且不改变行为
- `style`: 格式或样式调整
- `docs`: 文档
- `test`: 测试
- `chore`: 构建、工具、CI、依赖

Rules:

- Title must start with `<type>:` and be within 50 Chinese characters when practical.
- Title and body use Chinese.
- Body bullet lines use `-`.
- Always include `Co-Authored-By: Codex Opus 4.7 <noreply@anthropic.com>`.
- Do not commit `node_modules`, `dist`, `.turbo`, local caches, or unrelated user changes.
- Keep commits scoped. Do not mix tooling, UI, core model, and unrelated cleanup unless the user asks.
- Before commit, run relevant validation and report any skipped checks.

## Current Roadmap Order

Prefer this order unless the user explicitly changes it:

1. Node 24 / pnpm 11 engineering baseline.
2. Workbook core model.
3. Workbook serialization.
4. Row/column structure operations.
5. Formula MVP.
6. Playground UI integration: sheet tabs, add/switch/rename/delete sheets.
7. React/Vue SDK components.
8. Import/export and advanced Excel-like features.

## Handoff Format

After code changes, answer in Chinese with:

- `CHANGES MADE (修改了什么)`: files and concrete reason.
- `THINGS I DIDN'T TOUCH (故意没碰的地方)`: related files intentionally left unchanged.
- `POTENTIAL CONCERNS (潜在隐患)`: risks, skipped checks, manual verification needs.

Keep the answer concise and mention verification commands that passed.
