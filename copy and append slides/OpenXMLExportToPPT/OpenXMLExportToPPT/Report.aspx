<%@ Page Title="" Language="C#" MasterPageFile="~/Report.Master" AutoEventWireup="true" CodeBehind="Report.aspx.cs" Inherits="OpenXMLExportToPPT.Report1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
      <div class="bs-example">
        <form runat="server" class="form-horizontal">
                <h1>Create Presentation</h1>
            <div class="form-group">
                <div class="col-xs-offset-2 col-xs-10">
                    <asp:Button id="btn_submit" CssClass="btn btn-primary" OnClick="btn_submit_Click" OnClientClick="return confirm('Your ppt has been created !');"  runat="server" Text="Create PPT"></asp:Button>
                </div>
            </div>
            
        </form>
    </div>
</asp:Content>

