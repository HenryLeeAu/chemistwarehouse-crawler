"use strict";

var _request = require("request");

var _request2 = _interopRequireDefault(_request);

var _cheerio = require("cheerio");

var _cheerio2 = _interopRequireDefault(_cheerio);

var _assert = require("assert");

var _excel4node = require("excel4node");

var _excel4node2 = _interopRequireDefault(_excel4node);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var inc = 1;
var url = "https://www.chemistwarehouse.com.au/Shop-OnLine/81/Vitamins?size=120&page=" + inc;
var weathers = [];
var totlPage = 1;
var run = function run() {
  if (inc <= totlPage) {
    (0, _request2.default)(url, function (err, res, body) {
      var $ = _cheerio2.default.load(body);
      if (inc == 1) {
        console.log(Math.ceil(parseInt($('.Pager').eq(0).find('.pager-count>b').text().replace('Results')) / 120));
        totlPage = Math.ceil(parseInt($('.Pager').eq(0).find('.pager-count>b').text().replace('Results')) / 120);
      }

      $('.product-list-container td').each(function (i, elem) {

        if ($(this).find('.product_image_overlay').attr('class') == "product_image_overlay") {
          console.log('true', i);
          weathers.push({
            'price': parseInt($(this).find('.Price').text().replace(/(\r\n\t|\n|\r\t)/gm, "").replace('$', '')),
            'name': $(this).find('.product-container').attr('title'),
            'save': parseInt($(this).find('.Save').text().replace(/(\r\n\t|\n|\r\t)/gm, " ").replace('SAVE', '').replace('$', '')),
            'url': $(this).find('.product-container').attr('href'),
            'img': $(this).find('img').eq(0).attr('src')

          });
        } else {}
      });

      inc++;
      run();
    });
  } else {
    console.log(weathers);
    console.log(weathers.length);

    var wb = new _excel4node2.default.Workbook();

    // Add Worksheets to the workbook
    var ws = wb.addWorksheet('Sheet 1');
    var ws2 = wb.addWorksheet('Sheet 2');
    // Create a reusable style
    var style = wb.createStyle({
      font: {
        color: '#FF0800',
        size: 12
      },

      numberFormat: '$#,##0.00; ($#,##0.00); -'
    });
    weathers.map(function (item, index) {
      ws.cell(index + 2, 1).string(item.name);
      ws.cell(index + 2, 2).number(item.save).style(style);
      ws.cell(index + 2, 3).number(item.price).style(style);
      ws.cell(index + 2, 4).string("https://www.chemistwarehouse.com.au" + item.url);
      ws.cell(index + 2, 5).string(item.img);
    });
    ws.cell(1, 1).string('Name');
    ws.cell(1, 2).string('Save');
    ws.cell(1, 3).string('half price').style(style);
    ws.cell(1, 4).string('url');
    ws.cell(1, 5).string('img');

    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    //ws.cell(1,1).number(100).style(style);

    // Set value of cell B1 to 300 as a number type styled with paramaters of style
    //ws.cell(1,2).number(200).style(style);

    // Set value of cell C1 to a formula styled with paramaters of style
    //ws.cell(1,3).formula('A1 + B1').style(style);

    // Set value of cell A2 to 'string' styled with paramaters of style
    //ws.cell(2,1).string('string').style(style);

    // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.//
    //ws.cell(3,1).bool(true).style(style).style({font: {size: 14}});

    wb.write('Excel.xlsx');
  }
};
run();