const { defineConfig } = require('vite');

module.exports = defineConfig({
  base: './',
  clearScreen: false,
  build: {
    outDir: 'dist',
    emptyOutDir: true,
  },
});
