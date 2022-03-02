<%@ Page Title="Início" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="searchInfoCNPJS._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <h3>Clique no botão para iniciar a busca dos CNPJs</h3>
    <br />
    <asp:Button ID="btnConsultar" runat="server" Text="Consultar" OnClick="btnConsultar_Click"/>

</asp:Content>


