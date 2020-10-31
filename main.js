const { PythonShell } = require('python-shell');
const fs = require('fs');

/* VARIABLES */
const ccaPrice = 450;
const pythonPath = `${process.cwd()}/python/`;

let date = new Date();
let dateString = `${date.getMonth() + 1}/${date.getFullYear()}`;

/* READ OLD JSON FILES */
/* PRICELIST */
let customerPricesJson = JSON.parse(fs.readFileSync('./dataOld/customerPrices.json', 'utf-8'));
let customerPricesid = customerPricesJson._id;
delete customerPricesJson['_id'];
let customerPricesKeys = Object.keys(customerPricesJson);

/* BACKUP */
let customerBackUpJson = JSON.parse(fs.readFileSync('./dataOld/customerBackUp.json', 'utf-8'));
let customerBackUpid = customerBackUpJson._id;
delete customerBackUpJson['_id'];
// let customerBackUpKeys = Object.keys(customerBackUpJson);

/* CUSTOMERNUMBER - NAME */
let customerNumberNameJson = JSON.parse(
  fs.readFileSync('./dataOld/customerNumberName.json', 'utf-8')
);

/* PRICELIST - NUMBER */
let customerPricelistNumberJson = JSON.parse(
  fs.readFileSync('./dataOld/customerPricelistNumber.json', 'utf-8')
);

/* UPDATE THE BACKUP */
customerPricesKeys.forEach((key) => {
  if (customerBackUpJson[key]) {
    customerBackUpJson[key][dateString] = customerPricesJson[key];
  } else {
    customerBackUpJson[key] = {};
    customerBackUpJson[key][dateString] = customerPricesJson[key];
  }
});
/* WRITE JSON FILE */
customerBackUpJson['_id'] = customerBackUpid;
fs.writeFileSync('./dataNew/customerBackUp.json', JSON.stringify(customerBackUpJson));

/* UPDATE PRICES */
customerPricesKeys.forEach((key) => {
  /* GET INDEX KEYS */
  customerPricesJson[key].CCA = ccaPrice;
  let idx = Object.keys(customerPricesJson[key]);
  idx = idx.slice(0, idx.length - 4);

  idx.forEach((indx) => {
    customerPricesJson[key][indx][4] = customerPricesJson[key][indx][3] + ccaPrice;
  });

  let customerPrices = { ...customerPricesJson };
  let customerPricelist = {};
  customerPricelist[key] = customerPrices[key];
  customerPricelist.PRICELIST = customerPricelistNumberJson[key];
  customerPricelist.HEADER = customerNumberNameJson[key];
  customerPricelistString = JSON.stringify(customerPricelist);

  /* PYSHELL OPTIONS */
  let options = {
    mode: 'text',
    pythonOptions: ['-u'],
    scriptPath: pythonPath,
    args: [customerPricelistString, 'none'],
  };

  /* CREATE PYSHELL  */
  let pyshell = new PythonShell('conversion.py', options);

  pyshell.end(function (err, code, signal) {
    if (err) {
      console.log(err);
    }
  });
});

/* WRITE JSON FILE */
customerPricesJson['_id'] = customerPricesid;
fs.writeFileSync('./dataNew/customerPrices.json', JSON.stringify(customerPricesJson));
