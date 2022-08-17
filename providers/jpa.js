const XLSX = require('xlsx');
const fs = require('fs');

// JPA
module.exports.transform = (inputFilePath, outputFilePath) => {
    const jpaWb = XLSX.readFile(inputFilePath);
    const jpaJsonContent = XLSX.utils.sheet_to_json(jpaWb.Sheets[jpaWb.SheetNames[0]], {header:1, raw: true});
    const jpaRawContent = jpaWb.Sheets[jpaWb.SheetNames[0]];

    const jpaFireworks = [];
    for (let product of jpaJsonContent) { 
        if (product.length !== 16) {
            continue;
        }
        const firework = {};             
        firework.ref = `${product[1]}`;
        firework.caliber = product[2];
        firework.name = product[3];
        firework.numberOfShots = product[3].match(/([0-9]{1,3}) tirs?/) ? parseInt(product[3].match(/([0-9]{1,3}) tirs?/)[1]) : 1;   
        firework.certification = product[4] ? product[4] : null;
        firework.category = product[5] ? product[5].toUpperCase() : null;
        firework.activeWeight = product[6] ? (product[6] * 1000) : 0;
        firework.duration = product[7] ? product[7] : 2;
        firework.safetyDistances = product[8] ? product[8] : 0;
        firework.transportCategory = product[11];
        firework.conditioning = {
            unit: 1,
            bte: null,
            ca: null
        };
        firework.priceWithoutTaxes = {
            unit: product[12],
            bte: 0,
            ca: 0
        };
        firework.priceWithTaxes = {
            unit: product[13],
            bte: 0,
            ca: 0
        };
        firework.discountable = product[14] === 'oui';
        firework.isAvailable = product[15] !== 'non';
        firework.nextArrival = product[16] ? product[16] : null;
        firework.provider = 'jpa';
        firework.ascentTime = 0;
        firework.video = null;
        jpaFireworks.push(firework);
    }


    for (const product in jpaRawContent) {
        const ref = jpaRawContent[product];

        if (ref.l) {
            const firework = jpaFireworks.find((firework) => firework.name === ref.l.display);
            if (firework) {
                firework.video = ref.l.Target;
            }
        }
    }

    // Write JPA
    fs.writeFileSync(outputFilePath, JSON.stringify(jpaFireworks, null, 4));

    return jpaFireworks.length;
}
