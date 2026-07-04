// @ts-check
import { defineConfig } from 'astro/config';

import tailwindcss from '@tailwindcss/vite';

export default defineConfig({
  // Update after Netlify assigns your site URL
  site: 'https://praveen-karthika-wedding.netlify.app',
  output: 'static',
  vite: {
    plugins: [tailwindcss()]
  }
});