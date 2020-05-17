function getStocksData(symbols) {
  const response = UrlFetchApp.fetch("https://priceservice.vndirect.com.vn/priceservice/secinfo/snapshot/q=codes:" + symbols)
  const json = response.getContentText()
  const data = JSON.parse(json)

  const stockData = data.reduce(
    (accumulator, d) => {
      const arrStock = d.split("|")
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
  const sheet = SpreadsheetApp.getActive().getSheetByName("stocks")
  let symbols = sheet.getRange("A3:A20").getValues()
  symbols = symbols.filter(function(s) {
    return s[0].length > 0
  }).map(function(s) {
    return s[0]
  })

  const stocks = getStocksData(symbols)
  let startRow = 3
  for (let idx = 0; idx < symbols.length; idx++) {
    if (!stocks[symbols[idx]]) continue

    const curStock = stocks[symbols[idx]]
    const row = startRow + idx

    const refPriceRange = sheet.getRange(row, 5, 1, 1)
    const priceRange = sheet.getRange(row, 6, 1, 1)
    const highestPriceRange = sheet.getRange(row, 7, 1, 1)
    const lowestPriceRange = sheet.getRange(row, 8, 1, 1)
    const totalVolumeRange = sheet.getRange(row, 9, 1, 1)
    const foreignBuyRange = sheet.getRange(row, 10, 1, 1)
    const foreignSellRange = sheet.getRange(row, 11, 1, 1)

    const refPrice = stocks[symbols[idx]]['r_price']
    const curPrice = stocks[symbols[idx]]['price']

    if (curPrice > refPrice) {
      priceRange.setFontColor('green')
    } else if (curPrice < refPrice) {
      priceRange.setFontColor('red')
    }

    refPriceRange.setValue(curStock['r_price'])
    priceRange.setValue(curStock['price'])
    highestPriceRange.setValue(curStock['h_price'])
    lowestPriceRange.setValue(curStock['l_price'])
    totalVolumeRange.setValue(curStock['volume'])
    foreignBuyRange.setValue(curStock['f_buy_volume'])
    foreignSellRange.setValue(curStock['f_sell_volume'])
  }

  sheet.getRange(1, 1, 1, 1).setValue("Last update: " + new Date().toLocaleString('vn-VI', { timeZone: 'Asia/Ho_Chi_Minh' }))
}

function main() {
  setDataGoogleSheet()
  SpreadsheetApp.getUi()
  .createMenu('Stock Utils')
  .addItem('setDataGoogleSheet')
  .addToUi()
}
