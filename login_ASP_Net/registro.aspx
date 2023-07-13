<%@ Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OLEDB" %>

<html><head>
<meta name="GENERATOR" Content="ASP Express 2.1">
<title>Registro</title>

<script language="VB" runat="server">


'Declaração de variáveis
Dim conec as OLEDBConnection
Dim data_rd as OLEDBDatareader
Dim comando as OLEDBCommand

'Abaixo a declaração do data souce
Dim DS as string ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("usuarios.mdb")

'Abaixo a rotina que insere os valores no banco de dados e checa excessões
Sub registrar(Source as Object, E as EventArgs)


Dim SQL_SQL as string = "INSERT INTO cadastro (nome, login, senha)  VALUES ('"& txtnome.text & "','"& txtlogin2.text & "', '"& txtsenha2.text & "')"

Dim conec as New OleDbConnection(DS)
Dim comando as New OleDbCommand(SQL_SQL, conec)


if txtnome.text = "" then
label1.text = "Digite seu nome!"

else if txtlogin2.text = "" then
label1.text = "Digite seu login!"

else if txtsenha2.text = "" then
label1.text = "Digite sua senha!"

else if txtsenha2.text <> txtsenha3.text then
label1.text = "Senhas diferentes"

else


'Se todos os campos forem preenchiidos corretamente o comando é executado
conec.Open()
data_rd=comando.ExecuteReader(system.data.CommandBehavior.CloseConnection)

response.Redirect("index.aspx")

end if



End Sub
</script>

<style type="text/css">
<!--
.style1 {background-color: #FFFFFF}
.style2 {background-color: #FFFFFF; color: #FF0000; }
.style3 {background-color: #FFFFFF; color: #006600; }
-->
</style></head>
<body bgcolor="white">
<form Name="frmSignup" runat="server">

<asp:placeholder ID="ph2" visible="true" runat="server">
<p align="left" class="style3">
  <font face="Verdana" size="5"><span style="font-weight: 700">Fa&ccedil;a sua inscri&ccedil;&atilde;o </span></font><a href="index.aspx">Voltar ao login</a></p>
<p align="left" class="style2">
  <asp:Label ID="label1" runat="server" />  </p>
<table width="450" border="1">
  <tr>
    <td width="144"><b>Nome</b></td>
    <td width="290"><asp:TextBox ID="txtnome" runat="server" Columns="42" /></td>
  </tr>
  <tr>
    <td><b>Login</b></td>
    <td><span class="style1">
      <asp:TextBox ID="txtlogin2" runat="server"  />      
</span></td>
  </tr>
  <tr>
    <td><b>Senha</b></td>
    <td><span class="style1">
      <asp:TextBox ID="txtsenha2" runat="server" TextMode="Password" />      
</span>
      <div align="right"></div></td>
  </tr>
  <tr>
    <td><b>Repetir Senha</b></td>
    <td><span class="style1">
      <asp:TextBox ID="txtsenha3" runat="server" TextMode="Password" />            
      <asp:Button ID="buttonlg" Text="OK" OnClick="registrar" runat="server" />      
    </span></td>
  </tr>
</table>
<p>&nbsp;</p>
</asp:placeholder>


  <table border=0 align="center">
   <tr><td>
</form>
<tr></td>

</body>
</html>
