const keys = require('./assets/keys');
const googleMapsClient = require('@google/maps').createClient({
    key: keys.googleMaps,
    Promise: Promise
});

const XLSX = require('js-xlsx');
const workbook = XLSX.readFile('./assets/distances.xlsx')
const first_sheet_name = workbook.SheetNames[0];
const worksheet = workbook.Sheets[first_sheet_name];
const counter = {index: 2};

const asyncGrab = function (value1, value2){
return new Promise((resolve, reject)=>{ 
    googleMapsClient.distanceMatrix(
        {
            origins: [value1],
            destinations: [value2],
            units:'imperial'
        }).asPromise().then((response) => {
            if(response.json.rows[0].elements[0].distance){
            resolve(response.json.rows[0].elements[0].distance.text)
            } else {resolve('not available');}
            })
});
}

fillTheSheet(worksheet, counter).then(() => {
    console.log(counter.index);
    worksheet['!ref'] = `A1:E${counter.index}`;
XLSX.writeFile(workbook, 'result.xlsx');
}
)  

async function fillTheSheet(worksheet, counter){
for (let i = 1; i<2; i++){
    let result;
    const firstAddress = worksheet[`B${counter.index}`].v;
    const secondAddress = worksheet[`D${counter.index}`].v;
    if(!secondAddress.includes('p.o.') && 
    !secondAddress.includes('P.O.') && 
    !secondAddress.includes('P O box') && 
    !secondAddress.includes('PO box') &&
    !secondAddress.includes('P O BOX') &&
    !secondAddress.includes('PO BOX')
){
        result = await asyncGrab(firstAddress, secondAddress);
    }
    else { result="not available"}
    worksheet[`E${counter.index}`] = {
        t:'s',
        v:result,
            };
    counter.index++;            
    }

}
