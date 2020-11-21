## A module for split the PDF then send to outlook
This is a module for split the PDF file then send to your outlook.
Because it need to call the Windows API, it only can be used on Windows

## Modules to be used
pywin32: `pip install pywin32`

PyPDF2: `pip install PyPDF2`

## How to used
1. Open the Outlook
2. Copy the PDF file into C:\temp\pdf_split
3. Input the start page number
4. Input the end page number
Then you will find that the split pdf file is sent to your mailbox separately.And you can find the file in C:\temp\pdf_split\result
