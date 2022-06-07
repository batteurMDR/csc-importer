const XLSX = require('xlsx');
const Papa = require('papaparse');

const wb = XLSX.readFile('./prevot.xlsx');
const prevotJsonContent = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw: true});
const prevotRawContent = wb.Sheets[wb.SheetNames[0]];


const fireworks = [];
for (let product of prevotJsonContent) { 
    if (product.length !== 16) {
        continue;
    } 
    const firework = {};             
    firework.ref = `${product[1]}`;
    firework.caliber = product[2];
    firework.name = product[3];
    firework.numberOfShots = product[3].match(/([0-9]{1,3}) tirs?/) ?? 1;   
    firework.certification = product[4] ? product[4] : '';
    firework.category = product[5];
    firework.activeWeight = product[6] ? product[6] : 0;
    firework.duration = product[7] ? product[7] : 2;
    firework.safetyDistances = product[8] ? product[8] : 0;
    firework.transportCategory = product[11];
    firework.price = product[12];
    firework.discountable = product[13] === 'oui';
    // firework.providerStock = product[14] !== 'non';
    // firework.nextArrival = product[15] ? product[15] : '';
    firework.provider = 'prevot';
    firework.ascentTime = 0;
    fireworks.push(firework);
}


for (const product in prevotRawContent) {
    const ref = prevotRawContent[product];

    if (ref.l) {
        const firework = fireworks.find((firework) => firework.name === ref.l.display);
        if (firework) {
            firework.video = ref.l.Target;
        }
    }
}

const fs = require('fs');

fs.writeFile('./fireworkImport.csv', fireworks.slice(0, 10).map((f) => {
    return [
        f.name,
        `Cat: ${f.category} / Cal: ${f.caliber} / Cert: ${f.certification} / Dist: ${f.safetyDistances} / Div: ${f.transportCategory} / Disc: ${f.discountable}`,
        f.duration,
        f.ascentTime,
        f.numberOfShots,
        2022,
        f.price,
        f.ref,
        'JPA',
        'Other',
        f.video,
        '',
        0
    ].join(', ');
}).join('\n'), {encoding: 'utf-8', flag: 'a'}, (err) => {
    console.log(err);
});