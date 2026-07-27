import './polyfills/buffer';
import { lazy, Suspense } from 'react';
import ReactDOM from 'react-dom/client';
import './main.css';

document.getElementById('_defaultStyles')?.remove();

const Excel = lazy(() => import('./view/excel/Excel.tsx'));

export default function App() {
  return (
    <Suspense fallback={<div className='app-loading' aria-label='Loading spreadsheet' />}>
      <Excel />
    </Suspense>
  );
}

ReactDOM.createRoot(document.getElementById('root')!).render(<App />);
