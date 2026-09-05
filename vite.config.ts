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

export default defineConfig({
  plugins: [
    react(),
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
