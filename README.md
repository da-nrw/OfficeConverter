OfficeConverter
===============

Provides Webservice for automatic conversion of Office documents into PDF. 

The main feature is, the feature uses propietary software packages for conversion being able to 
interprete the fonts (the whole look&feel) of the documents stored in Office documents. 

Written in .NET ASP, the software runs on IIS 7.0 using .NET Framework 3.5. (Windows(R) Server 2008 R2) No running installation of DNSCore are needed to test and run the service. The project provides a testpage for conversion checks.  

Supported and tested formats so far are: 
Word DOCX, DOC

Output: 
Actually PDF 1.5 (the output is converted to PDF/A in DNSCore and validated with pdfbox)

## Setup on IIS 

You have to setup your Office installation for word automation, esp. the COM rights like it is explained in these forums:
* http://social.msdn.microsoft.com/Forums/en-US/0f5448a7-72ed-4f16-8b87-922b71892e07/word-2007-documentsopen-returns-null-in-aspnet
* http://www.codeproject.com/Articles/19993/MS-Word-Automation-from-ASP-NET-and-Publishing

## Conversion check / "Smoke test"

Tests on conversion could be carried out via URL: http://Server/Converter.aspx 

## Integration in DNSCore 

The equivalent conversion routine has to be set up like this:
<pre>ID | pdf           | de.uzk.hki.da.convert.DocxConversionStrategy         | LZA
_DOCX                 | http://Server/Handler.ashx 
</pre>


Please add the conversion policies for Office documents on your needs. 

## Debugging 

The Logfile generates Logfiles using Log4NET 
