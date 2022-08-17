const XLSX = require('xlsx');
const fs = require('fs');

// BREZAC
module.exports.transform = (inputFilePath, outputFilePath) => {
    const brezacWb = XLSX.readFile(inputFilePath);
    const brezacArticlesJsonContent = XLSX.utils.sheet_to_json(brezacWb.Sheets['Articles'], {header:1, raw: true});
    const brezacUnitPricesJsonContent = XLSX.utils.sheet_to_json(brezacWb.Sheets['Prix PRO PCS'], {header:1, raw: true});
    const brezacBtePricesJsonContent = XLSX.utils.sheet_to_json(brezacWb.Sheets['Prix PRO BTE'], {header:1, raw: true});
    const brezacCaPricesJsonContent = XLSX.utils.sheet_to_json(brezacWb.Sheets['Prix PRO CA'], {header:1, raw: true});
    const brezacBteConditioningJsonContent = XLSX.utils.sheet_to_json(brezacWb.Sheets['Unités BTE'], {header:1, raw: true});
    const brezacCaConditioningJsonContent = XLSX.utils.sheet_to_json(brezacWb.Sheets['Unités CA'], {header:1, raw: true});

    const brezacFireworks = [];
    for (let product of brezacArticlesJsonContent) {
        if (product.length !== 24 || !product[7] || ['STD', 'TRANSFERT', 'COLOR STD', 'COLOR', 'ACCPYRO'].includes(product[7]) || !['1', 'OPPORTUNITE'].includes(product[5])) {
            continue;
        }
        const firework = {};         
        firework.ref = `${product[3]}`;
        firework.caliber = product[9];
        firework.name = product[4];
        firework.numberOfShots = 1;   
        firework.certification = product[15] ? product[15] : null;
        firework.category = product[13] ? product[13] : null;
        firework.activeWeight = product[16] ? (product[16] * 1000) : 0;
        firework.duration = product[17] ? parseFloat(product[17].replace(",", ".")) : 2;
        firework.safetyDistances = product[18] ? product[18] : 0;
        firework.transportCategory = product[14];
        firework.conditioning = {
            unit: null,
            bte: null,
            ca: null
        };
        firework.priceWithoutTaxes = {
            unit: 0,
            bte: 0,
            ca: 0
        };
        firework.priceWithTaxes = {
            unit: 0,
            bte: 0,
            ca: 0
        };
        firework.discountable = false;
        firework.isAvailable = product[20] >= 0;
        firework.nextArrival = null;
        firework.provider = 'brezac';
        firework.ascentTime = 0;
        firework.video = product[23] ? product[23] : null;
        brezacFireworks.push(firework);
    }

    for (let brezacFirework of brezacFireworks) {
        const unitHtPrice = brezacUnitPricesJsonContent.find((row) => row[3] === brezacFirework.ref);
        if (unitHtPrice) {
            brezacFirework.conditioning.unit = 1;
            brezacFirework.priceWithoutTaxes.unit = unitHtPrice[9];
            brezacFirework.priceWithTaxes.unit = unitHtPrice[9] * 1.2;
        }
        const bteHtPrice = brezacBtePricesJsonContent.find((row) => row[3] === brezacFirework.ref);
        if (bteHtPrice) {
            brezacFirework.conditioning.bte = 0;
            brezacFirework.priceWithoutTaxes.bte = bteHtPrice[9];
            brezacFirework.priceWithTaxes.bte = bteHtPrice[9] * 1.2;

            const bteConditioning = brezacBteConditioningJsonContent.find((row) => row[3] === brezacFirework.ref);
            if (bteConditioning) {
                brezacFirework.conditioning.bte = parseInt(bteConditioning[4].slice(0, -4));
            }
        }
        const caHtPrice = brezacCaPricesJsonContent.find((row) => row[3] === brezacFirework.ref);
        if (caHtPrice) {
            brezacFirework.conditioning.ca = 0;
            brezacFirework.priceWithoutTaxes.ca = caHtPrice[9];
            brezacFirework.priceWithTaxes.ca = caHtPrice[9] * 1.2;

            const caConditioning = brezacCaConditioningJsonContent.find((row) => row[3] === brezacFirework.ref);
            if (caConditioning) {
                brezacFirework.conditioning.ca = parseInt(caConditioning[4].slice(0, -3));
            }
        }
    }

    // Write Brezac
    fs.writeFileSync(outputFilePath, JSON.stringify(brezacFireworks, null, 4));

    return brezacFireworks.length;
}
