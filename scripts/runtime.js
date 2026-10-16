'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const {createRequire} = require('module');

// Prefer the desktop runtime; retain existing skill dependencies for TeX parsing.
function loadDependency(name) {
  const roots = [process.env.CODEX_NODE_MODULES,
    path.join(os.homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules'),
    path.join(__dirname, '../node_modules'),
    path.join(os.homedir(), '.agents/skills/docx-editor-cn/node_modules')].filter(Boolean);
  for (const root of roots) {
    const candidate = path.join(root, name);
    if (fs.existsSync(candidate)) return createRequire(path.join(path.dirname(root),'_skill_resolver.cjs'))(name);
  }
  return require(name);
}
module.exports = {loadDependency};
