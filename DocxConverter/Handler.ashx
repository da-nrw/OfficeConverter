<%@ WebHandler Language="C#" Class="Handler" %>
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
using System.Web;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using log4net;


/**
 * Author Jens Peters DA-NRW 2013 
 * 
 * DOCX -> PDF Conversion 
 * Quite easy to code, but hard to set up on IIS: 
 * http://social.msdn.microsoft.com/Forums/en-US/0f5448a7-72ed-4f16-8b87-922b71892e07/word-2007-documentsopen-returns-null-in-aspnet
 * http://www.codeproject.com/Articles/19993/MS-Word-Automation-from-ASP-NET-and-Publishing
 * Much more could be done !!
 */
public class Handler : IHttpHandler {

    private static readonly ILog log = LogManager.GetLogger("IHttpHandler");


    public void ProcessRequest (HttpContext context) {
        log4net.Config.XmlConfigurator.Configure();


        string filePath = System.Web.HttpContext.Current.Server.MapPath("~/");
        if (context.Request.QueryString["cleanup"]!=null)
        {
            context.Response.Write("cleaned<br>");
            log.Debug("Cleanup called!");
             System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(@filePath);
            IEnumerable < System.IO.FileInfo > fileInfo = dir.GetFiles("*.*", SearchOption.AllDirectories);
           
            IEnumerable<System.IO.FileInfo> fileQuery =
            from file in fileInfo
            where file.Extension == ".pdf" 
                || file.Extension == ".docx"
                && file.CreationTime < DateTime.Now.AddDays(-1)
            orderby file.Name
            select file;

            foreach (FileInfo file in fileQuery)
            {
                log.Debug("File " + file + " deleted");
                file.Delete();
            }
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
            {
                p.Kill();
            }
        }


        if (context.Request.Files.Count <= 0 && context.Request.QueryString["cleanup"]==null) 
        {
            
            context.Response.Write("No file uploaded");
        }
        else
        {
            for (int i = 0; i < context.Request.Files.Count; ++i)
            {
                HttpPostedFile file = context.Request.Files[i];

                if (file.ContentType == "application/msword" ||
                 file.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                {
                    file.SaveAs(filePath + file.FileName);
                    log.Debug("File " + filePath + file.FileName + " saved");

                    try
                    {
                        WordToPdf w2pdf = new WordToPdf(filePath + file.FileName);
                        w2pdf.word2PDF();

                        FileInfo fi = new FileInfo(w2pdf.getOutPutFileName());
                        context.Response.AddHeader("Content-Disposition", "attachment; filename=" + fi.Name);
                        context.Response.AddHeader("Content-Length", fi.Length.ToString());
                        context.Response.ContentType = "application/octet-stream";
                        context.Response.AppendHeader("Content-Disposition", "attachment; filename=fi.Name");

                        context.Response.TransmitFile(w2pdf.getOutPutFileName());
                        context.Response.End();
                    }
                    catch (Exception ex)
                    {
                        log.Error("Converison failed! " + ex.StackTrace);
                        context.Response.Write("Error " + ex.Message); }
                } else context.Response.Write("Wrong Content Type");
            }
        }
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}