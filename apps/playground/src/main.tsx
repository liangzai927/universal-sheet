import './index.css';

import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';

import App from './App';

const root = document.getElementById('app');
if (!root) throw new Error('Mount element #app not found in DOM.');

createRoot(root).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
