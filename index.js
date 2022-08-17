const providersToProcess = ['jpa', 'brezac'];

for (const provider of providersToProcess) {
    console.log(`Start ${provider}`);
    let providerFunction = require(`./providers/${provider}`);

    const inputFilePath = `./input/${provider}.xlsx`;
    const outputFilePath = `./output/${provider}.json`;

    const totalRefFromProvider = providerFunction.transform(inputFilePath, outputFilePath);
    console.log(`End ${provider} with ${totalRefFromProvider} references`);
    console.log('');
}
