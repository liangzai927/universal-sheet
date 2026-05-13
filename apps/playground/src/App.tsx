import { VERSION as coreVersion } from '@universal-sheet/core';
import { VERSION as engineVersion } from '@universal-sheet/engine';
import { VERSION as formulaVersion } from '@universal-sheet/formula';

export default function App() {
  return (
    <div style={{ padding: '2rem', fontFamily: 'system-ui, sans-serif' }}>
      <h1>Universal Sheet - Playground</h1>
      <ul>
        <li>@universal-sheet/core: v{coreVersion}</li>
        <li>@universal-sheet/engine: v{engineVersion}</li>
        <li>@universal-sheet/formula: v{formulaVersion}</li>
      </ul>
    </div>
  );
}
