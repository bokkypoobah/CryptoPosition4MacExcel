# CryptoPosition4MacExcel
Crypto Position For Excel On The Mac

This repository contains an Excel for Mac spreadsheet that will allow you to retrieve rates from https://min-api.cryptocompare.com.

Unlike Excel for Windows where you can easily retrieve data from web pages, Excel for Mac requires your to create Queries to retrieve data from web pages.

## Creating Your Own Queries Files

There are a few sample Queries files in [Queries](https://github.com/bokkypoobah/CryptoPosition4MacExcel/tree/master/Queries) directory.

### [CryptoCompareETH](https://github.com/bokkypoobah/CryptoPosition4MacExcel/blob/master/Queries/CryptoCompareETH)
    WEB
    1
    https://min-api.cryptocompare.com/data/price?fsym=ETH&tsyms=BTC,ETH,AUD,USD
    
    Selection=Cell
    Formatting=None
    PreFormattedTextToColumns=True
    ConsecutiveDelimitersAsOne=True
    SingleBlockTextImport=False

### [CryptoCompareGNT](https://github.com/bokkypoobah/CryptoPosition4MacExcel/blob/master/Queries/CryptoCompareGNT)
    WEB
    1
    https://min-api.cryptocompare.com/data/price?fsym=GNT&tsyms=BTC,ETH,AUD,USD
    
    Selection=Cell
    Formatting=None
    PreFormattedTextToColumns=True
    ConsecutiveDelimitersAsOne=True
    SingleBlockTextImport=False

### [CryptoCompareBTC](https://github.com/bokkypoobah/CryptoPosition4MacExcel/blob/master/Queries/CryptoCompareBTC)
    WEB
    1
    https://min-api.cryptocompare.com/data/price?fsym=BTC&tsyms=BTC,ETH,AUD,USD
    
    Selection=Cell
    Formatting=None
    PreFormattedTextToColumns=True
    ConsecutiveDelimitersAsOne=True
    SingleBlockTextImport=False



Enjoy. (c) BokkyPooBah 2016. The MIT licence.
