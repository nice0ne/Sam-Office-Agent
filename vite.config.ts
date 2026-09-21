import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';
import path from 'path';
import fs from 'fs';
import os from 'os';

const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
const certPath = path.join(certDir, 'localhost.crt');
const keyPath = path.join(certDir, 'localhost.key');
const hasOfficeCerts = fs.existsSync(certPath) && fs.existsSync(keyPath);

function localProxyPlugin() {
  return {
    name: 'local-proxy',
    configureServer(server: any) {
      server.middlewares.use('/api/proxy', async (req: any, res: any) => {
        if (req.method === 'OPTIONS') {
          res.setHeader('Access-Control-Allow-Origin', '*');
          res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
          res.setHeader('Access-Control-Allow-Headers', '*');
          res.statusCode = 200;
          res.end();
          return;
        }

        const urlObj = new URL(req.url, 'http://localhost');
        const targetUrl = req.headers['x-target-url'] || urlObj.searchParams.get('target');
        if (!targetUrl) {
          res.statusCode = 400;
          res.setHeader('Content-Type', 'application/json');
          res.setHeader('Access-Control-Allow-Origin', '*');
          res.end(JSON.stringify({ error: 'Missing target URL' }));
          return;
        }

        try {
          const chunks: any[] = [];
          for await (const chunk of req) {
            chunks.push(chunk);
          }
          const bodyBuffer = Buffer.concat(chunks);

          const forwardHeaders: Record<string, string> = {};
          if (req.headers['content-type']) forwardHeaders['content-type'] = req.headers['content-type'];
          if (req.headers['authorization']) forwardHeaders['authorization'] = req.headers['authorization'];

          const response = await fetch(targetUrl, {
            method: req.method,
            headers: forwardHeaders,
            body: bodyBuffer.length > 0 ? bodyBuffer : undefined,
          });

          res.setHeader('Access-Control-Allow-Origin', '*');
          res.setHeader('Content-Type', response.headers.get('content-type') || 'application/json');
          res.statusCode = response.status;

          if (response.body) {
            const reader = response.body.getReader();
            while (true) {
              const { done, value } = await reader.read();
              if (done) break;
              res.write(value);
            }
          }
          res.end();
        } catch (err: any) {
          res.statusCode = 502;
          res.setHeader('Access-Control-Allow-Origin', '*');
          res.setHeader('Content-Type', 'application/json');
          res.end(JSON.stringify({ error: `Proxy Error: ${err.message}` }));
        }
      });
    },
  };
}

export default defineConfig({
  plugins: [
    react(),
    localProxyPlugin(),
    ...(hasOfficeCerts ? [] : [basicSsl()]),
  ],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
  server: {
    port: 5173,
    cors: true,
    https: hasOfficeCerts
      ? {
          key: fs.readFileSync(keyPath),
          cert: fs.readFileSync(certPath),
        }
      : true,
  },
  test: {
    globals: true,
    environment: 'node',
    setupFiles: './tests/setup.ts',
    passWithNoTests: true,
  },
});
