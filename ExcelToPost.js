const XLSX = require('xlsx');
const axios = require('axios');
const https = require('https');

const excelFilePath = "C:/Users/kbrka/Downloads/Products.xlsx"; 
const workbook = XLSX.readFile(excelFilePath);
const sheetName = workbook.SheetNames[0];
const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]); 

const agent = new https.Agent({  
    rejectUnauthorized: false 
});

async function hitApiForProducts(products) {
    const url = 'https://localhost:7194/Products';

    for (let i = 0; i < products.length; i++) {
        const product = products[i];
        const requestBody = {
            name: product.name,
            price: product.price, 
            description: product.description
        };

        try {
            const response = await axios.post(url, requestBody, {
                httpsAgent: agent,
                headers: {
                    'Content-Type': 'application/json'
                }
            });
            console.log(`Request ${i + 1}: Success with status ${response.status}`);
        } catch (error) {
            console.log(`Request ${i + 1}: Failed with error ${error.message}`);
        }
    }
}

hitApiForProducts(sheetData);
