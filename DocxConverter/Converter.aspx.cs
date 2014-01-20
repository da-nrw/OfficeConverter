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
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using System.IO;

/**
 * Author Jens Peters DA-NRW 2013 
 * 
 * DOCX -> PDF Conversion 
 * Quite easy to code, but hard to set up on IIS: 
 * http://social.msdn.microsoft.com/Forums/en-US/0f5448a7-72ed-4f16-8b87-922b71892e07/word-2007-documentsopen-returns-null-in-aspnet
 * http://www.codeproject.com/Articles/19993/MS-Word-Automation-from-ASP-NET-and-Publishing
 * 
 */
public partial class Converter : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
    }

    void Application_Start(object sender, EventArgs e)
    {
        // Code that runs on application startup
        log4net.Config.XmlConfigurator.Configure(new FileInfo(Server.MapPath("~/Web.config")));
    }

    protected void UploadButton_Click(object sender, EventArgs e)
    {
        if (FileUploadControl.HasFile)
        {
            try
            {
               string filename = Path.GetFileName(FileUploadControl.FileName);
               if (FileUploadControl.PostedFile.ContentType == "application/msword" ||
                   FileUploadControl.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                           
                   ) {
               FileUploadControl.SaveAs(Server.MapPath("~/") + filename);
                
               WordToPdf w2pdf = new WordToPdf(Server.MapPath("~/") +filename);
                   w2pdf.word2PDF();
                StatusLabel.Text = "Upload status: File uploaded and converted !";
                FileInfo fo = new FileInfo (w2pdf.getOutPutFileName());
                   OutgoingLabel.Text  = "<a href=\"" +fo.Name +"\">" + fo.Name + "</a>";               
                   }
               else StatusLabel.Text = FileUploadControl.PostedFile.ContentType +" is not allowed";
            } 
            catch (Exception ex)
           {
              StatusLabel.Text = "Upload status: The file could not be uploaded. The following error occured: " + ex.Message + " " + ex.StackTrace;
           }
        }
    }

}