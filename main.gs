const CONFIG_SHEET = 'configs'
const MAIL_SUBJECT = 'VN_STOCK ALERT'
const PRICE_UNIT = 1000
const PRICE_CURRENCY = 'VND'
var Config = {}

function run() {
  loadConfig()
  let msg = setDataGoogleSheet()
  let currentHour = calcTime('+7.0').getHours()
  if (9 < currentHour && currentHour < 15 && Config.enable_send_mail[0] && msg != '') {
    let htmlBody = "<h2>Hi, our stock's prices has been changed!</h2>" + msg
    sendMail(htmlBody)
  }
}

function calcTime(offset) {
    // create Date object for current location
    var d = new Date();

    // convert to msec
    // subtract local time zone offset
    // get UTC time in msec
    var utc = d.getTime() + (d.getTimezoneOffset() * 60000);

    // create new Date object for different city
    // using supplied offset
    var localDate = new Date(utc + (3600000*offset));
    return localDate
}

function main() {
  run();
  SpreadsheetApp.getUi()
  .createMenu('Utils')
  .addItem('Refresh', 'setDataGoogleSheet')
  .addItem('Enable send mail','enableSendMail')
  .addItem('Disable send mail','disableSendMail')
  .addToUi();
}

function loadConfig() {
  let sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET);
  let data = sheet.getDataRange().getValues();
  for (var i in data) {
    let row = data[i];
    if (Config[row[0]]) {
      Config[row[0]].push(row[1])
    } else {
      Config[row[0]] = [row[1]]
    }
  }
  Logger.log(Config)
}

function getStockdata(symbols) {
  const response = UrlFetchApp.fetch('https://priceservice.vndirect.com.vn/priceservice/secinfo/snapshot/q=codes:' + symbols)
  const json = response.getContentText()
  const data = JSON.parse(json)

  const stockData = data.reduce(
    (accumulator, d) => {
      const arrStock = d.split('|')
      accumulator[arrStock[3]] = {
        price: arrStock[19],
        r_price: arrStock[8],
        h_price: arrStock[13],
        l_price: arrStock[14],
        volume: arrStock[36],
        f_buy_volume: arrStock[37],
        f_sell_volume: arrStock[38]
      }
    return accumulator
  }, {})

  return stockData
}


function setDataGoogleSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('stocks')
  var symbols = sheet.getRange('A3:A10').getValues()
  symbols = symbols.filter(function(s) {
    return s[0].length > 0
  }).map(function(s) {
    return s[0]
  })

  Logger.log(symbols)

  const stocks = getStockdata(symbols)
  let startRow = 3
  let message = ''
  for (let idx = 0; idx < symbols.length; idx++) {
    let code = symbols[idx]
    if (!stocks[code]) continue
    
    const curStock = stocks[code]
    const row = startRow + idx
    
    const currentStock = sheet.getRange(row, 3, 1, 1).getValue()
    const boughtPrice = sheet.getRange(row, 4, 1, 1).getValue()
    const refPriceRange = sheet.getRange(row, 5, 1, 1)
    const priceRange = sheet.getRange(row, 6, 1, 1)
    const highestPriceRange = sheet.getRange(row, 7, 1, 1)
    const lowestPriceRange = sheet.getRange(row, 8, 1, 1)
    const totalVolumeRange = sheet.getRange(row, 9, 1, 1)
    const foreignBuyRange = sheet.getRange(row, 10, 1, 1)
    const foreignSellRange = sheet.getRange(row, 11, 1, 1)
    
    const oldPrice = priceRange.getValue()
    const refPrice = stocks[code]['r_price']
    const curPrice = stocks[code]['price']
    
    if (curPrice > refPrice) {
      priceRange.setFontColor('green')
    } else if (curPrice < refPrice) {
      priceRange.setFontColor('red')
    }
    if (oldPrice - curPrice != 0) {
      if (curPrice >= boughtPrice) {
        let profit = curPrice - boughtPrice
        let profitPercent = parseFloat(profit*100/boughtPrice).toFixed(2)
        if (profitPercent < Config.alert_threshold_profit_percent[0]) {
          continue
        }
        let totalProfit = parseFloat(profit * currentStock*PRICE_UNIT).toFixed(0)
        message += `<p>Code: ${makeProfitFontTag(code)} | Profit(%): ${makeProfitFontTag(profitPercent)} | Hold Stocks: ${currentStock} | Total Profit: ${makeProfitFontTag(totalProfit)} ${PRICE_CURRENCY}</p>`
      } else {
        let takeLoss = boughtPrice - curPrice
        let lossPercent = parseFloat(takeLoss*100/boughtPrice).toFixed(2)
        if (lossPercent < Config.alert_threshold_loss_percent[0]) {
          continue
        }
        let totalLoss = parseFloat(takeLoss * currentStock*PRICE_UNIT).toFixed(0)
        message += `<p>Code: ${makeLossFontTag(code)} | Loss(%): ${makeLossFontTag(lossPercent)} | Hold Stocks: ${currentStock} | Total Loss: ${makeLossFontTag(totalLoss)} ${PRICE_CURRENCY}</p>`
      }
    } 
    
    refPriceRange.setValue(curStock['r_price'])
    priceRange.setValue(curStock['price'])
    highestPriceRange.setValue(curStock['h_price'])
    lowestPriceRange.setValue(curStock['l_price'])
    totalVolumeRange.setValue(curStock['volume'])
    foreignBuyRange.setValue(curStock['f_buy_volume'])
    foreignSellRange.setValue(curStock['f_sell_volume'])
  }
  sheet.getRange(1, 1, 1, 1).setValue('Last update: ' + new Date().toLocaleString('vn-VI', { timeZone: 'Asia/Ho_Chi_Minh' }))
  return message;
}
  
function makeLossFontTag(con) {
   return `<strong style='color:red'>${con}</strong>`
}

function makeProfitFontTag(con) {
   return `<strong style='color:green'>${con}</strong>`
}
/**
 * Google trigger function. When the sheet is opened, a custom menu is produced.
 * 
 */

function onOpen() {
  main()
}

function enableSendMail() {
  Config.enableSendMail = true
}

function disableSendMail() {
  Config.enableSendMail = false
}

/**
 * Send mail
 */
function sendMail(htmlBody) {
  for (var i in Config.email_receiver) {
    MailApp.sendEmail(Config.email_receiver[i], MAIL_SUBJECT, '', {htmlBody: htmlBody});  
  }  
}


/********************************************************
 * Mail HTML body
 */
function createHeader(type, content, body) {
  return body + `<${type}>${content}</${type}>`
}

function createTblHeader(body) {
  let TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
  let header = '<table ' + TABLEFORMAT +' ">';
  return body + header
}

function createTblEnd(body) {
  return body + "</table>";
}

function createRow(isHeader, colVals, body, colProps) {
  let colTagStart = '<td>';
  let colTagEnd = '</td>';
  if (isHeader) {
    colTagStart = '<th';
    colTagEnd = '</th>';
  }
  let row = '';
  for (var i in ColNames) {
    colStr = colTagStart
    if (colProps.length > 0) {
      colStr += " " + colProps[i] + ">"
    }
    row += `${colStr}${colVals[i]}${colTagEnd}`;
  }
  row = `<tr>${row}</tr>`
  return body + row
}