/*
DA-NRW Software Suite | ContentBroker
Copyright (C) 2013 Historisch-Kulturwissenschaftliche Informationsverarbeitung
Universität zu Köln

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <http://www.gnu.org/licenses/>.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using System.IO;
using log4net;
/// <summary> 
/**
 * @Author Jens Peters DA-NRW 2013 
 * 
 * DOCX -> PDF Conversion 
 * Assuming .NET FW 3.5.1
 * Quite easy to code, but hard to set up on IIS: 
 * http://social.msdn.microsoft.com/Forums/en-US/0f5448a7-72ed-4f16-8b87-922b71892e07/word-2007-documentsopen-returns-null-in-aspnet
 * http://www.codeproject.com/Articles/19993/MS-Word-Automation-from-ASP-NET-and-Publishing
 * Much more has to be done! 
 */
/// </summary>
public class WordToPdf
{

    string file;
    object outputFileName;
    private static readonly ILog log = LogManager.GetLogger("WordToPdf");


	public WordToPdf(string filename)
	{
        log4net.Config.XmlConfigurator.Configure();
        this.file = filename;
        log.Debug("Converting request: " + filename);
	}
    public void word2PDF()
    {   //Creating the instance of Word Application
        // C# doesn't have optional arguments so we'll need a dummy value
        // in .NET FW 4, this might be more effective and even thread safe!!
        log.Debug("trying to kill old word.exe Procs");

        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("winword.exe"))
        {
            p.Kill();
        }
        
        object oMissing = System.Reflection.Missing.Value;
        log.Debug("trying to load word.exe");
        Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
        log.Debug("trying to load Document Object");
        
        Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
        
        FileInfo fi = new FileInfo(file);

        Object filename = file;
        log.Debug("about to open document now!");
        doc = word.Documents.Open(ref filename, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        doc.Activate();

        object fileFormat = (object)Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;

        file = file.Replace(fi.Extension, ".pdf");
        this.outputFileName = (Object) file.Replace(".docx", ".pdf");
        log.Debug("trying to save as " + this.outputFileName);
        
        doc.SaveAs(ref outputFileName, ref fileFormat, ref oMissing, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing);

        // word has to be cast to type _Application so that it will find
        // the correct Quit method.

        ((Microsoft.Office.Interop.Word._Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
        word = null;
        }

    public string getOutPutFileName() {
        return (string)outputFileName;
    }

    public void cleanup() {

        try
        {

            File.Delete(file);
            File.Delete((string)outputFileName);

            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("winword.exe"))
            {
                p.Kill();
            }
        }
        catch (Exception ex)
        {
            log.Error(ex.Message);
        }
    }
}