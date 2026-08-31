// Why does Chrome kern Liberation Mono when our parser finds nothing?
// Dump the GPOS structure rather than guess.
import { readFileSync } from 'node:fs';

const path = process.argv[2];
const buf = readFileSync(path);
const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
const u16 = (o) => dv.getUint16(o);
const u32 = (o) => dv.getUint32(o);
const tag = (o) => String.fromCharCode(buf[o], buf[o + 1], buf[o + 2], buf[o + 3]);

const numTables = u16(4);
let gpos = null;
for (let i = 0; i < numTables; i++) {
    const rec = 12 + i * 16;
    if (tag(rec) === 'GPOS') gpos = u32(rec + 8);
}
if (gpos === null) { console.log('no GPOS'); process.exit(0); }

const scriptList = gpos + u16(gpos + 4);
const featureList = gpos + u16(gpos + 6);
const lookupList = gpos + u16(gpos + 8);

console.log('scripts:');
const scriptCount = u16(scriptList);
for (let i = 0; i < scriptCount; i++) {
    console.log('  ', tag(scriptList + 2 + i * 6));
}

console.log('features:');
const featureCount = u16(featureList);
const kernLookups = new Set();
for (let i = 0; i < featureCount; i++) {
    const rec = featureList + 2 + i * 6;
    const t = tag(rec);
    const feature = featureList + u16(rec + 4);
    const n = u16(feature + 2);
    const idx = [];
    for (let j = 0; j < n; j++) idx.push(u16(feature + 4 + j * 2));
    console.log('  ', t, '-> lookups', idx.join(','));
    if (t === 'kern') idx.forEach(x => kernLookups.add(x));
}

console.log('lookups referenced by kern:', [...kernLookups].join(',') || '(none)');
const lookupCount = u16(lookupList);
for (const i of kernLookups) {
    if (i >= lookupCount) continue;
    const lk = lookupList + u16(lookupList + 2 + i * 2);
    const type = u16(lk);
    const subCount = u16(lk + 4);
    console.log(`  lookup ${i}: type=${type} subtables=${subCount}`);
    for (let s = 0; s < subCount; s++) {
        const sub = lk + u16(lk + 6 + s * 2);
        console.log(`    subtable ${s}: posFormat=${u16(sub)} valueFormat1=0x${u16(sub + 4).toString(16)} valueFormat2=0x${u16(sub + 6).toString(16)}`);
    }
}
