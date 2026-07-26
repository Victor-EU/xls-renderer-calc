import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.js';
// The grid's own stylesheet, imported by its published specifier exactly as a
// consumer would. Without it the ⚠ refusal marker is unstyled text.
import '@xlscalc/xlsx-preview/view/style.css';
import './styles.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
