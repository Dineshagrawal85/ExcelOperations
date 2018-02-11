var express = require('express');
var router = express.Router();
var xl = require('excel4node');

/**
 * Funtion to response request for download xlsx
 * @return xlsx file with name Template.xlsx
 */
router.get('/', function(req, res, next) {
    //DUMMY EXCEL TEMPLATE
    /*
    	GIVE TYPE FOR DIFFERENT TYPE OF VALIDATIONS
    	EX:- list, whole
    */
    var dummy_template = [{
            'prime_head': 'Name'
        },
        {
            'prime_head': 'description'
        },
        {
            'prime_head': 'device'
        },
        {
            'prime_head': 'label',
            'sec_head': [{
                    'name': 'label-initial',
                    'type': 'list',
                    'values': ['one,two']
                },
                {
                    'name': 'label-mid',
                    'type': 'list',
                    'values': ['three,four']
                },
                {
                    'name': 'label-final',
                    'type': 'whole',
                    'values': [1, 100]
                }
            ]
        }
    ]
    generateExcel(dummy_template, function(err, wb, ws) {
        if (err) {
            return next(err)
        }
        //pipe workbook directly to response
        wb.write('Template.xlsx', res);
    })


});

module.exports = router;


/**
 * Function to accept a custom JSON and create a excel in-momory
 * @return Object workbook, worksheet
 */
function generateExcel(excel_template, cb) {
    try {
        var wb = new xl.Workbook();
        var ws = wb.addWorksheet('options');

        var style = wb.createStyle({
            font: {
                color: '#000000',
                size: 16,
                bold: true
            },
            numberFormat: '$#,##0.00; ($#,##0.00); -',
            alignment: {
                horizontal: 'center'
            }
        });


        var styleForSubHeading = wb.createStyle({
            font: {
                color: '#000000',
                size: 14,
                bold: true
            },
            numberFormat: '$#,##0.00; ($#,##0.00); -',
            alignment: {
                horizontal: 'center'
            }
        });

        ws.row(2).freeze();
        var MAX_ROW_LIMIT = 1002

        var main_row_number = 1
        var current_column_number = 1

        excel_template.map(function(_val) {
            if (_val) {
                var sec_head = _val.sec_head
                var sec_head_length = 0
                if (sec_head && sec_head.length) {
                    sec_head_length = sec_head.length
                }
                if (sec_head_length == 0) {
                    ws.cell(main_row_number + 1, current_column_number).string(_val.prime_head).style(style);
                    current_column_number += 1;
                } else {
                    var cells_to_merge = []
                    ws.cell(main_row_number, current_column_number).string(_val.prime_head).style(style);
                    _val.sec_head.map(function(__sec_head) {
                        ws.cell(main_row_number + 1, current_column_number).string(__sec_head.name).style(styleForSubHeading);
                        cells_to_merge.push(main_row_number)
                        cells_to_merge.push(current_column_number)
                        var columnAlpha = xl.getExcelAlpha(current_column_number);
                        if (columnAlpha) {
                            var sqref = columnAlpha + '3:' + columnAlpha + MAX_ROW_LIMIT
                            if (__sec_head.type == 'list') {
                                ws.addDataValidation({
                                    type: 'list',
                                    allowBlank: true,
                                    prompt: 'Choose from dropdown',
                                    promptTitle: '',
                                    error: 'Invalid choice was chosen',
                                    showDropDown: true,
                                    sqref: sqref,
                                    formulas: __sec_head.values
                                });
                            } else if (__sec_head.type == 'whole') {
                                ws.addDataValidation({
                                    type: 'whole',
                                    operator: 'between',
                                    prompt: 'Please Enter value b/w 0 to 100',
                                    error: 'Invalid Value',
                                    allowBlank: 1,
                                    sqref: sqref,
                                    formulas: __sec_head.values
                                });
                            }
                        }
                        current_column_number += 1;
                    })
                    ws.cell(...cells_to_merge.slice(0, 2), ...cells_to_merge.slice(-2), true)
                }
            }
        })
        return cb(null, wb, ws)

    } catch (e) {
        return cb(e)
    }

}