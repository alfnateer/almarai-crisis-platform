// Print the website's input cells (sheet, address, keep) as JSON for the Python tools.
// keep = the template's value is a default (support lists, risk assumption map), not data.
const SCHEMA = require('../js/schema.js');
process.stdout.write(JSON.stringify(SCHEMA.inputCells()));
