import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('node_modules/react') || id.includes('node_modules/react-dom') || id.includes('node_modules/react-router')) return 'react';
          if (id.includes('node_modules/chart.js') || id.includes('node_modules/react-chartjs-2')) return 'charts';
          if (id.includes('node_modules/xlsx')) return 'spreadsheet';
          if (id.includes('node_modules/lucide-react')) return 'icons';
        },
      },
    },
  }
})
