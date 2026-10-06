'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
for (const name of fs.readdirSync(root).filter(name => name.endsWith('.gs'))) {
  new vm.Script(fs.readFileSync(path.join(root, name), 'utf8'), { filename: name });
}
const client = fs.readFileSync(path.join(root, 'Client.html'), 'utf8');
const match = client.match(/^<script>\s*([\s\S]*?)\s*<\/script>\s*$/);
if (!match) throw new Error('Client.html must contain exactly one script block.');
new vm.Script(match[1], { filename: 'Client.html' });
const leaflet = fs.readFileSync(path.join(root, 'Leaflet.html'), 'utf8');
const leafletMatch = leaflet.match(/^<script>\s*([\s\S]*?)\s*<\/script>\s*$/);
if (!leafletMatch) throw new Error('Leaflet.html must contain exactly one script block.');
new vm.Script(leafletMatch[1], { filename: 'Leaflet.html' });
const spatialOrder = fs.readFileSync(path.join(root, 'SpatialOrder.html'), 'utf8');
const spatialOrderMatch = spatialOrder.match(/^<script>\s*([\s\S]*?)\s*<\/script>\s*$/);
if (!spatialOrderMatch) throw new Error('SpatialOrder.html must contain exactly one script block.');
new vm.Script(spatialOrderMatch[1], { filename: 'SpatialOrder.html' });
JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
process.stdout.write('Syntax check passed.\n');
