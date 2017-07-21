<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TrendsTable.aspx.vb" Inherits="TrendCalcDemo.TrendsTable" %>
<%@ Register Src="~/TrendSummaryTable.ascx" TagName="TrendSummaryTable" TagPrefix="xp" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Trends Table</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <xp:TrendSummaryTable ID="TrendSummaryTable1" runat="server"></xp:TrendSummaryTable>
        </div>
    </form>
</body>
</html>
