import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  // Porta fixa: o localStorage (notas, comentários, pastas) é separado por
  // origem, incluindo a porta. Se o Vite caísse para 5174 por a 5173 estar
  // ocupada, os dados pareceriam ter sumido. Com strictPort ele falha em vez
  // de trocar de porta.
  server: { port: 5173, strictPort: true },
  preview: { port: 5173, strictPort: true },
})
