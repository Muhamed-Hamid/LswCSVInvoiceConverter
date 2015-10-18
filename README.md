# LswCSVInvoiceConverter
Leaseweb CSV Invoice Converter is a python tool to convert your CSV invoice in leaseweb to a pretty cool and clean xlsx output.

## Features
- Convert CSV invoice to XLSX format compatible with Microsoft Office, LibreOffice, Google Sheets
- Formatting your services info into separated sheets including ( Servers Pricing - VPS - Windows Licenses - Plesk Licenses - cPanel Licenses - Cabling )
- Formatting servers components info including (Uplink - KVM - Server - RackSpace - Traffic - IP - APC - Support - Switches)
- Writing out the total server price.
- Formatting your licenses info including license name and the price.
- Additional sheet for Extras on your invoice including the items info and the price for each.
- Reports sheet generating statistics of your services including (Number of cPanel licenses, Number of Plesk licenses, Number of Windows licenses, Number of Servers, Number of VPS, Total cPanel Licenses Cost, Total Plesk Licenses Cost, Total Windows Licenses Cost, Total Servers Cost, Total VPS Cost, Total Cabling, Total Extra, Total cPanel Licenses Cost, Total Plesk Licenses Cost, Total Windows Licenses Cost, Total Servers Cost, Total VPS Cost, Total Cabling, Total Extra )
- Generating reports indicates Total Invoice amount from XLSX and from CSV And the number of imported items.

## Requirements
- Python3
- openpyxl and xlsxwriter modules

## Modules Installation
```
pip install xlsxwriter
pip install openpyxl
```
## How to use
- Make sure you have installed the required modules
- Clone or download the repo to your working directory
- Download your CSV invoice from Leaseweb Customer Portal into your working directory
- Run dimofinf.py and type the CSV filename including the extension. Ex. 2015-1111.csv
- After the execution you should find an exported file with the name invoice.xlsx into the same working directory

## Supported Services
- Dedicated Servers
- Virtual Servers
- Private Racks
- Switches and Cabling
- Plesk, cPanel and Windows license

## DISCLAIMER
Despite we are doing our best to improve this tool, Dimofinf does not provide any kind of warranty for the results and you should manually check the exported data if you find any missing or corrupted items.
Your feedback is highly appreciated to continue the development any notes, suggestions or any kind of help to improve the tool please do not hesitate to share.
