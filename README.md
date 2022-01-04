# QR Code Generator

### Desktop application for generating QR code.
___
The application was written specifically for the e-commerce website
[onshop.am](https://www.onshop.am).

![mygif](./QRCodeGenerator.gif)

The user simply copies the URL of the desired product to the clipboard, then in the application itself clicks on the "fill in automatically" button. Fields such as link, title, price and color of the QR code are filled in with data automatically.

The data can also be filled in manually.

The output is a file in docx format, which displays the name of the web site, the name of the product and the QR code.
___
- ### The project is completely written in `python`.
- ### Web scraping and parsing of data provided by `Requests` and `Beautiful Soup` libraries.
- ### A `qrcode` library is used to generate the QR code.
- ### The `python-docx-template’s` was used to create the template of docx file.
- ### For creating the GUI was used `Tkinter`. 