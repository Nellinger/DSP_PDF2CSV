# Order PDF 2 Order CSV for DSP

## Description
This project is designed to extract data customer and order data from PDF files and convert it into a CSV format. 

## Installation
1. Portable App: Simply download the app and run the executable.

## Usage
- Put the PDF files in the `input` folder and then run the app. The CSV files will be generated in the `output` folder.

## License
This project is licensed under the MIT License.

## Dependencies
This project uses the following libraries:
### Tabula
	- License under the MIT License:
	Copyright (C) 2012-2020 Manuel Aristarán jazzido@jazzido.com

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

### UglyToad.PdfPig
	- License under the MIT License:
	The MIT License (MIT)

Copyright (c) 2014 Eliot Jones

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## To-Do
- [ ] 

## Known Issues
- The app has been tested on a couple of PDF files and the quality of result may vary on different PDF files.
- Since line breaks are not always detected correctly, the app may not be able to extract the data correctly. Some manual intervention may be required. Currently some key words are used to detect the data, but this may not work in all cases.
- Therefore it is highly recommended to check the CSV files generated for any errors.

## Contact
For support, please contact `software@nellinger.com`.
