import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './ui/App';
import './styles.css';

const container = document.getElementById('root');
if (!container) throw new Error('Root element missing');

// The desktop shell marks itself so assets resolve next to the executable.
(window as { __SHEETCRAFT_PACKAGED__?: boolean }).__SHEETCRAFT_PACKAGED__ =
  typeof window !== 'undefined' && window.location.protocol === 'file:';

createRoot(container).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
