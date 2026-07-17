# AGENTS.md

## Project: Universal Sheet

Excel-like frontend spreadsheet SDK. Framework-agnostic core, publishable to npm, supports Vue 3 and React. Features developed incrementally layer by layer: core → engine → UI.

---

## TypeScript Conventions

### Strict typing — zero tolerance for `any`

```ts
// BAD
function getCell(row: any, col: any): any { ... }

// GOOD
function getCell(row: number, col: number): CellData | null { ... }
```

- `strict: true` enforced globally (see tsconfig.base.json)
- Prefer `interface` for public APIs, `type` for unions/primitives
- Every function must declare explicit return type (no inference for public exports)
- Use `readonly` on data that should not mutate
- Use `const` assertions and `as const` for literal types
- `unknown` over `any` — narrow with type guards before use
- Barrel exports only: each package exports via `src/index.ts`, no deep imports between packages

### Naming

| Category          | Convention                | Example                                 |
| ----------------- | ------------------------- | --------------------------------------- |
| Interface         | PascalCase, no `I` prefix | `CellData`, `SheetConfig`               |
| Type alias        | PascalCase                | `CellValue = string \| number`          |
| Enum member       | PascalCase                | `CellType.Text`                         |
| Variable/function | camelCase                 | `getCellValue`, `rowIndex`              |
| Constant          | UPPER_SNAKE_CASE          | `MAX_ROWS`, `DEFAULT_COL_WIDTH`         |
| Private members   | camelCase, no `_` prefix  | `internalState` (use `private` keyword) |
| File name         | kebab-case                | `cell-model.ts`, `sheet-engine.ts`      |

---

## Package Architecture Constraints

### `packages/core` — Pure logic, zero DOM

- **Forbidden**: `document`, `window`, `HTMLElement`, Vue/React imports
- **Allowed**: Pure TS data structures, algorithms, event emitters
- **Dependencies**: Only `@universal-sheet/shared`

### `packages/engine` — Canvas rendering, no framework code

- **Forbidden**: Vue/React imports, direct DOM manipulation for UI
- **Allowed**: Canvas 2D API, `OffscreenCanvas`, requestAnimationFrame
- **Dependencies**: `@universal-sheet/core`, `@universal-sheet/shared`

### `packages/formula` — Formula parsing, pure computation

- **Forbidden**: DOM, framework imports, side effects
- **Allowed**: Parser combinators, AST manipulation, decimal arithmetic
- **Dependencies**: `@universal-sheet/shared` only

### `packages/shared` — Internal utilities only

- **Never imported by external consumers** (not in npm exports)
- Small utils: debounce, throttle, event emitter, type guards

### `packages/ui-vue` — Vue 3 bindings

- Vue 3 Composition API only (no Options API)
- `<script setup lang="ts">` required
- `defineProps<T>()` with interface, `defineEmits<T>()` with type
- No direct DOM queries — use `ref`/`template ref`

### `packages/ui-react` — React bindings

- Function components only, no class components
- Hooks for all state/logic
- `React.memo` on pure presentational components
- Props interfaces prefixed with component name

---

## Vue 2 Component Template (if `packages/ui-vue2` is added)

```vue
<script lang="ts">
import Vue from 'vue';

interface Props {
  modelValue: string;
  disabled?: boolean;
}

export default Vue.extend({
  name: 'UsComponent',
  props: {
    modelValue: { type: String, required: true },
    disabled: { type: Boolean, default: false },
  },
  model: {
    prop: 'modelValue',
    event: 'input',
  },
  methods: {
    handleChange(value: string): void {
      this.$emit('input', value);
    },
  },
});
</script>

<template>
  <div class="us-component" :class="{ 'us-disabled': disabled }">
    <!-- Template content -->
  </div>
</template>
```

### Vue 2 constraints

- Use `Vue.extend()` for type-safe component definitions
- `v-model` uses `modelValue` prop + `input` event (standard Vue 2 model)
- Props: declare types via constructor (`String`, `Number`, `Boolean`, `Object`, `Array`)
- No mixins — prefer scoped slots or utility functions
- CSS class prefix: `us-`
- One component per file

---

## Vue 3 Component Template

```vue
<script setup lang="ts">
interface Props {
  /** Description of the prop */
  readonly modelValue: string;
  /** Optional flag with default */
  readonly disabled?: boolean;
}

const props = withDefaults(defineProps<Props>(), {
  disabled: false,
});

interface Emits {
  (e: 'update:modelValue', value: string): void;
}

const emit = defineEmits<Emits>();
</script>

<template>
  <div class="us-component" :class="{ 'us-disabled': props.disabled }">
    <!-- Template content -->
  </div>
</template>
```

### Vue constraints

- CSS class prefix: `us-` (universal-sheet)
- Scoped styles via `<style scoped>` or CSS modules
- No inline styles over 2 properties — extract to class
- Component name in kebab-case for the file, PascalCase for import
- One component per file

---

## React Component Template

```tsx
import { memo, useCallback, useState } from 'react';

interface SheetViewProps {
  /** The data model to render */
  readonly data: SheetData;
  /** Called when a cell is clicked */
  readonly onCellClick?: (row: number, col: number) => void;
}

export const SheetView = memo(function SheetView({ data, onCellClick }: SheetViewProps) {
  const [selectedCell, setSelectedCell] = useState<[number, number] | null>(null);

  const handleCellClick = useCallback(
    (row: number, col: number) => {
      setSelectedCell([row, col]);
      onCellClick?.(row, col);
    },
    [onCellClick],
  );

  return <div className="us-sheet-view">{/* JSX content */}</div>;
});
```

### React constraints

- CSS class prefix: `us-`
- Named exports only (no default exports)
- `memo()` on components that receive props from parent
- `useCallback` on event handlers passed to memo'd children
- No inline styles over 2 properties — extract to CSS
- Props interface: `{ComponentName}Props`

---

## Comment Standards

### JSDoc for every function (exported or internal)

Every function must have a JSDoc block describing **what it does**, its **parameters**, and its **return value**.

```ts
/**
 * Parses a formula string into an AST.
 * Uses recursive descent rather than regex to handle nested function calls.
 *
 * @param input - Raw formula string (e.g. "=SUM(A1, B1)")
 * @returns The root node of the parsed AST
 * @throws {FormulaError} If input is empty or malformed
 */
export function parseFormula(input: string): FormulaNode { ... }
```

- `@param` — every parameter documented with type and purpose
- `@returns` — return value described (omit only for `void`)
- `@throws` — documented if the function can throw
- Internal helper functions follow the same rule — no exceptions

### When to write inline comments

- **Non-obvious logic**: Algorithm choice, performance consideration, edge case handling
- **Why, not what**: Explain the reason behind a decision, not what the code does

```ts
// Recalculate only dirty cells to avoid O(n²) on large sheets
for (const cell of dirtyCells) { ... }
```

### When NOT to write comments

- Self-documenting code: `getCellValue(row, col)` needs no comment
- Variable declarations with clear names
- Simple getters/setters

---

## Test Rules

### Framework: Vitest

### File location

```
packages/core/src/cell-model.ts
packages/core/src/__tests__/cell-model.test.ts
```

### Naming

- Test file: `{module-name}.test.ts`
- Describe block: the module/function name
- Test case: `should {expected behavior} when {condition}`

### Template

```ts
import { describe, it, expect } from 'vitest';
import { parseFormula } from '../formula-parser';

describe('parseFormula', () => {
  it('should parse a simple addition expression', () => {
    const result = parseFormula('=A1+B1');
    expect(result).toEqual({
      type: 'binary',
      operator: '+',
      left: { type: 'cell-ref', ref: 'A1' },
      right: { type: 'cell-ref', ref: 'B1' },
    });
  });

  it('should throw FormulaError when given empty input', () => {
    expect(() => parseFormula('')).toThrow('FormulaError');
  });

  it('should handle nested function calls', () => {
    const result = parseFormula('=SUM(A1, AVERAGE(B1:B10))');
    expect(result.type).toBe('function-call');
  });
});
```

### Test constraints

- **Coverage target**: 80%+ on core/engine/formula, 60%+ on UI packages
- **No test nesting** beyond `describe` → `it` (no nested describes)
- **One assertion context per test**: test one behavior per `it` block
- **No shared mutable state** between tests — reset in `beforeEach`
- **Do NOT mock** `@universal-sheet/shared` internals
- **Data factories**: Extract repeated test data setup into `test-utils.ts` files
- **No snapshot tests** — prefer explicit `toEqual`/`toMatchObject` assertions
- **Edge cases required** per public function: empty input, boundary values, error paths

---

## Modularization

- **Strict ES module**: `import`/`export` only, no `require()`, no `module.exports`
- **One responsibility per file**: if a file exceeds 300 lines, split it
- **No barrel re-export chains**: `index.ts` only re-exports from siblings, never re-exports another barrel
- **Import order**: external packages → `@universal-sheet/` internal → relative imports
- **No side-effect imports**: every import must bind a value or type (exception: `import 'vitest'` in test files)
- **No circular dependencies**: if two modules need each other, extract a shared interface to `shared/`

### Encapsulation

- **Extract reusable logic**: any utility used ≥ 2 times becomes a named function in the appropriate module
- **Pure functions over methods**: prefer `function doThing(data: Data): Result` over class methods when no instance state is required
- **No magic numbers/strings**: extract to named constants at module top
- **Parameter limit**: functions with > 4 params should accept an options object

```ts
// BAD: inline magic, non-reusable
const total = price * 1.13 + 5;
// BAD: too many positional params
function drawCell(x: number, y: number, w: number, h: number, color: string, text: string): void { ... }

// GOOD: extracted constant + function
const TAX_RATE = 0.13;
const SHIPPING_FLAT = 5;
function calculateTotal(price: number): number { return price * (1 + TAX_RATE) + SHIPPING_FLAT; }
// GOOD: options object
function drawCell(pos: CellPos, size: CellSize, style: CellStyle): void { ... }
```

---

## Code Style

- **Concision over ceremony**: 3 similar lines is better than a premature abstraction
- No classes unless state + behavior are tightly coupled — prefer functions + interfaces
- No half-finished implementations, no `// TODO` comments, no dead code
- No backwards-compatibility shims or feature flags
- No `export default` — always named exports
- Imports: external packages first, then internal `@universal-sheet/` packages, then relative

---

## ESLint / Pre-commit Rules

The pre-commit hook runs `eslint --fix` + `prettier --write`. Code that doesn't pass will block the commit. Follow these rules proactively.

### Import sorting (`simple-import-sort`)

```ts
// BAD — unsorted imports
import { SheetRenderer } from '@universal-sheet/engine';
import { useState } from 'react';
import type { CellPosition } from '@universal-sheet/core';

// GOOD — externals first (react), then @universal-sheet/*, then relative
import { useState } from 'react';
import type { CellPosition } from '@universal-sheet/core';
import { SheetRenderer } from '@universal-sheet/engine';
```

- External npm packages first
- `@universal-sheet/*` second
- Relative imports (`./`, `../`) last
- Type-only imports (`import type`) sorted within each group

### Prefer `const` over `let`

```ts
// BAD
let result = computeValue();

// GOOD
const result = computeValue();
```

Only use `let` when the variable is **actually reassigned**.

### Type-only imports (`@typescript-eslint/consistent-type-imports`)

```ts
// BAD
import { CellPosition, SheetData, createSheetData } from '@universal-sheet/core';

// GOOD
import type { CellPosition, SheetData } from '@universal-sheet/core';
import { createSheetData } from '@universal-sheet/core';
```

Use `import type` for types, separate from runtime imports.

### Unused variables (`@typescript-eslint/no-unused-vars`)

No unused imports or variables. After refactoring, remove imports that are no longer needed. Prefix intentional unused args with `_`.

### Prefer optional chain (`@typescript-eslint/prefer-optional-chain`)

```ts
// BAD
if (!obj || obj.value === null) continue;

// GOOD
if (obj?.value == null) continue;
```

### Array types (`@typescript-eslint/array-type`)

```ts
// BAD
function foo(): string[] {}

// GOOD
function foo(): Array<string> {}
```

Use `Array<T>` not `T[]`.

### Restrict template expressions (`@typescript-eslint/restrict-template-expressions`)

```ts
// BAD — boolean/object in template
`cell: ${cell}`
// GOOD — number is allowed; anything else must be coerce
`cell: ${String(cell)}`;
```

### No non-null assertions (`@typescript-eslint/no-non-null-assertion`)

```ts
// BAD (warning)
const v = config.field!;

// GOOD — narrow with type guard or use ?? fallback
const v = config.field ?? defaultValue;
```

### File ignores

Config files (`vite.config.ts`, `.vitepress/config.ts`) are ignored by ESLint. They don't need to follow project rules.

---

## Git Commit Conventions

Follow [Conventional Commits](https://www.conventionalcommits.org/).

### Format

```
<type>: <简短描述>

- 具体改动 1
- 具体改动 2
```

中文描述，`<type>` 使用英文。

### Types

| Type       | 用途               |
| ---------- | ------------------ |
| `feat`     | 新功能             |
| `fix`      | Bug 修复           |
| `refactor` | 重构（不改变功能） |
| `style`    | 格式/样式调整      |
| `docs`     | 文档               |
| `test`     | 测试               |
| `chore`    | 构建/工具/CI       |

### 示例

```
feat: 框选/多选/批量调整 + 回车换行 + 角格全选

- 鼠标拖拽框选单元格范围
- 点击列头/行头选中整列/整行
- 拖拽批量调整列宽行高
- Enter 提交并下移，Shift+Enter 换行
```

### 规则

- 标题以 `type:` 开头，50 字符以内
- 标题用中文描述主要改动
- 正文用 `-` 列出具体改动点
- 结尾必须有 `Co-Authored-By: Codex Opus 4.7 <noreply@anthropic.com>`
- 提交前确保 `pnpm build` 和 `pnpm lint` 通过
- 不提交 `node_modules`、`dist`、`.turbo`
