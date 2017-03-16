# CryptoPosition4MacExcel
Crypto Position For Excel On The Mac

This repository contains an Excel for Mac spreadsheet that will allow you to retrieve rates from https://min-api.cryptocompare.com.

Unlike Excel for Windows where you can easily retrieve data from web pages, Excel for Mac requires your to create Queries to retrieve data from web pages.

## Creating Your Own Queries Files

There are a few sample Queries files in [Queries](https://github.com/bokkypoobah/CryptoPosition4MacExcel/tree/master/Queries) directory. Using the template below, you can customise your own files with the required `fsym` (from symbol) and list of `tsyms` (to symbols). Save this file to your Mac in the `$HOME/Library/
Group Containers/UBF8T346G9.Office/User Content.localized/Queries/` subdirectory. Note that when you are viewing this directory structure in Finder, the subdirectory `User Content.localized` is displayed as `User Content`.

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

## Excel Macros

The spreadsheet contains the following Excel macros:

    ' Option Explicit
    ' Crypto Position For Excel On The Mac
    '
    ' If you find this spreadsheet useful, please send any ETH or token donations to
    ' 0x000001f568875f378bf6d170b790967fe429c81a
    '
    ' Enjoy. (c) BokkyPooBah 2016. The MIT licence.

    Sub RefreshRates()
        ActiveWorkbook.RefreshAll
    End Sub

    ' Note that this is not a proper JSON string parser, but
    ' just a simple function to search for some text and return
    ' the following number
    Public Function getNumber(ccy As String, json As String) As Variant
        Dim startPos As Integer
        Dim endPos As Integer
        Dim temp As String

        startPos = InStr(json, ccy) + Len(ccy) + 2
        temp = Mid(json, startPos)
        endPos = InStr(temp, ",")
        If (endPos = 0) Then
            endPos = InStr(temp, "}")
        End If

        getNumber = Val(Mid(temp, 1, endPos - 1))
    End Function



Enjoy. (c) BokkyPooBah 2016. The MIT licence.
