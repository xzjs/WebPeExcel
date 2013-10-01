<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <br />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="上传文件" />
        <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
    
        <br />
        时间<asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <br />
        日期<asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
        <br />
        专业<asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="LinqDataSource1" DataTextField="College" DataValueField="College">
        </asp:DropDownList>
        <asp:LinqDataSource ID="LinqDataSource1" runat="server" ContextTypeName="WPEDataContext" EntityTypeName="" GroupBy="College" OrderGroupsBy="key" Select="new (key as College, it as Total)" TableName="Total">
        </asp:LinqDataSource>
        <br />
        年级<asp:DropDownList ID="DropDownList2" runat="server">
            <asp:ListItem>2011</asp:ListItem>
            <asp:ListItem>2010</asp:ListItem>
        </asp:DropDownList>
        <br />
        测试教师人数<asp:TextBox ID="TextBox3" runat="server"></asp:TextBox>
        <br />
        <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="分配" />
    
    </div>
    </form>
</body>
</html>
