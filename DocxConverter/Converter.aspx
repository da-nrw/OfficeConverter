<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Converter.aspx.cs" Inherits="Converter" %>
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

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload id="FileUploadControl" runat="server" />
    <asp:Button runat="server" id="UploadButton" text="Upload" onclick="UploadButton_Click" />
    <br /><br />
    <asp:Label runat="server" id="StatusLabel" text="Upload status: " />
    <asp:Label runat="server" id="OutgoingLabel" text="" />
    
    </div>
    </form>
    <p>
        HKI 2013 jp -DA-NRW</p>
</body>
</html>
