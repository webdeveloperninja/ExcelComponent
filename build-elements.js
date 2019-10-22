const fs = require('fs-extra');
const concat = require('concat');

(async function build() {
  const files = [
    './dist/graph-tutorial/runtime-es5.js',
    './dist/graph-tutorial/polyfills-es5.js',
    './dist/graph-tutorial/scripts.js',
    './dist/graph-tutorial/main-es5.js'
  ];
  await fs.ensureDir('elements');
  await concat(files, 'elements/elements.js');
  await fs.copyFile('./dist/graph-tutorial/styles.css', 'elements/styles.css');
})();
