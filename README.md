OfficeConverter
===============

## Prerequisites

1. .NET FW 3.5.1
2. Office 15.00.00 with Interop Lib
3. IIS 7
4. Windows 2008 R 2
5. Visual Studio .NET 2012
6. Log4Net

## Introduction

Provides Webservice for automatic conversion of Office documents into PDF. 

Feature uses propietary software packages for conversion being able to 
interprete the fonts (the look&feel) of the documents stored in Office formats. 

Written in .NET ASP, the software runs on IIS 7.0 using .NET Framework 3.5. (Windows(R) Server 2008 R2). No running installation of DNSCore are needed to test and run the service. The project provides a testpage for conversion checks.  

Supported and tested formats so far are: 
Word DOCX, DOC

Output: 
Actually PDF 1.5 based on standard office capabilities (the output recieved is converted to PDF/A in DNSCore and validated with Apache pdfbox)

## Setup on IIS 

You have to setup MS Office installation running on host for word automation. Crucial to setup is setting the COM rights correctly. 
* http://social.msdn.microsoft.com/Forums/en-US/0f5448a7-72ed-4f16-8b87-922b71892e07/word-2007-documentsopen-returns-null-in-aspnet
* http://www.codeproject.com/Articles/19993/MS-Word-Automation-from-ASP-NET-and-Publishing

## Conversion check / "Smoke test"

Tests on conversion could be carried out via URL: http://Server/Converter.aspx 

## Integration in DNSCore 

The equivalent conversion routine has to be set up like this:
<pre>ID | pdf           | de.uzk.hki.da.convert.DocxConversionStrategy         | LZA
_DOCX                 | http://Server/Handler.ashx 
</pre>

Please add the conversion policies based on pronom identifiers for Office documents at your needs. Please refer to ContenBroker documentation for DNSCore setup. 

## Debugging 

The application generates Logfiles using Log4NET 
