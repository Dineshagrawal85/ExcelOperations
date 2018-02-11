let express = require('express');
let XLSX = require('xlsx')
let router = express.Router();

/**
 * Function to parse json from received excel inside "report" key
 * @return JSON parsed JSON
 */
router.post('/', function(req, res, next) {
  if(req.files && req.files.report){
  	let workbook = XLSX.read(req.files.report.data,{'type':'buffer'});
  	let sheet_name_list = workbook.SheetNames;
  	let xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  	return res.jsonp(xlData)
  }else{
  	return res.status(400).send("file not found")
  }
});

module.exports = router;