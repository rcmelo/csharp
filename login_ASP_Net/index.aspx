<%@ Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OLEDB" %>

<html>
<head>

<title>Login</title>

<script language="VB" runat="server">


Dim conec as OLEDBConnection
Dim data_rd as OLEDBDatareader
Dim comando as OLEDBCommand

Sub logar(Source as Object, E as EventArgs)

'Abaixo declaramos o nosso data source
Dim DS as string ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("usuarios.mdb")


'Seleciona os dados do usuário.
Dim SQL_SQL as string = "Select id , nome , login , senha  from cadastro Where login = '" & txtlogin.text & "' and senha = '" & txtsenha.text & "'"

Dim conec as New OleDbConnection(DS)
Dim comando as New OleDbCommand(SQL_SQL, conec)

conec.Open()
data_rd = comando.ExecuteReader(system.data.CommandBehavior.CloseConnection)

if Not data_rd.Read() then
   label1.text = "Login ou Senha incorretos!"
   data_rd.Close
   conec.Close
else
   
   Session("id")=data_rd("id")
   Session("nome") = data_rd("nome")

   response.Redirect("area_restrita.aspx")


End If

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
  <font face="Verdana" size="5"><span style="font-weight: 700">Digite seu login e Senha </span></font> </p>
<p align="left" class="style2">
  <asp:Label ID="label1" runat="server" />  </p>
<div align="left">
  <table border="0" align="left">
    <tr>
      <td align="right" valign="Top"><b>Login</b></td>
        <td align="Left" valign="Top"><asp:TextBox id="txtlogin" runat="server" /></td>
      </tr>
    <tr>
      <td align="Right" valign="Top"><b>Senha </b></td>
        <td align="Left" valign="Top"><span class="style1">
          <asp:TextBox ID="txtsenha" runat="server" TextMode="Password" />          
          <asp:Button ID="buttonlg" Text="Entrar" OnClick="logar" runat="server" />          
        </span></td>
      </tr>
    <tr>
      <td height="29" Colspan="2" align="Center" valign="Top">Se n&atilde;o &eacute; cadastrado <a href="registro.aspx">Registre-se Aqui </a></td>
      </tr>
    <tr>
      <td align="Center" valign="Top" Colspan="2">&nbsp;</td>
    </tr>
  </table>
</div>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
</asp:placeholder>


  <table border=0 align="center">
   <tr>
     <td><div align="left"></div>
</form>
<tr></td>

</body>
</html>
