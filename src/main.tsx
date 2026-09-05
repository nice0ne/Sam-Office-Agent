import React from 'react';
import ReactDOM from 'react-dom/client';
import { App } from './App';
import './index.css';
import { HostType } from './types';

export function renderApp(host: HostType = 'BrowserDev', targetElement?: any) {
  if (typeof document === 'undefined' && !targetElement) return;
  const root = targetElement ?? (typeof document !== 'undefined' ? document.getElementById('root') : null);
  if (root) {
    if (typeof root.render === 'function') {
      root.render(
        <React.StrictMode>
          <App initialHost={host} />
        </React.StrictMode>
      );
      return;
    }
    ReactDOM.createRoot(root).render(
      <React.StrictMode>
        <App initialHost={host} />
      </React.StrictMode>
    );
  }
}

export function detectHostAndRender(customRenderApp: (host: HostType) => void = renderApp) {
  // Check if running inside Microsoft Office
  const office = typeof window !== 'undefined' ? (window as any).Office : undefined;
  if (office && typeof office.onReady === 'function') {
    office.onReady((info: any) => {
      let detectedHost: HostType = 'BrowserDev';
      const officeHostType = office.HostType;
      if (info?.host) {
        if (info.host === officeHostType?.Excel || info.host === 'Excel') detectedHost = 'Excel';
        else if (info.host === officeHostType?.Word || info.host === 'Word') detectedHost = 'Word';
        else if (info.host === officeHostType?.PowerPoint || info.host === 'PowerPoint') detectedHost = 'PowerPoint';
      }
      customRenderApp(detectedHost);
    });
  } else {
    // Standalone browser mode for local development
    customRenderApp('BrowserDev');
  }
}

detectHostAndRender();
