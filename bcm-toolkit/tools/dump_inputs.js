// Print the website's input cells (sheet + address) as JSON for the Python tools.
global.window = {};
require('../js/template-data.js');
const SCHEMA = require('../js/schema.js');
const keep = {};
SCHEMA.sheets.forEach(d => { if (d.keepDefaults) keep[d.sheet] = true; });
// keep: sheets whose template values are defaults (methodology lists), not data.
process.stdout.write(JSON.stringify(SCHEMA.inputCells(window.BCM_TEMPLATE_DATA).map(c => Object.assign(c, { keep: !!keep[c.sheet] }))));
