# Getting Started

## Installation

```bash
pnpm add @universal-sheet/core @universal-sheet/engine @universal-sheet/formula
```

## Quick Start

```ts
import { createSheet } from '@universal-sheet/core';

const sheet = createSheet({
  rows: 100,
  cols: 26,
});

sheet.setCell(0, 0, 'Hello, Universal Sheet!');
```
