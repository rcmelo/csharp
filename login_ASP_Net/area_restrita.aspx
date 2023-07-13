<%@ Page Language="VB" Debug="true" %>

<script language="VB" runat="server">


'A rotina abaixo redireciona para o login caso o a sessao ID contenha um valor nulo
Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) 

dim sessao as string
sessao = Session("id")

Select case sessao

case ""

response.Redirect("index.aspx")


case else

lblnome.text = Session("nome") 
lblid.text = Session("id") 


end select

End Sub

</script>



<html>
<head>
<title>Area Restrita</title>
<style type="text/css">
<!--
.style1 {
	color: #009900;
	font-weight: bold;
	font-style: italic;
}
-->
</style>
</head>
<body>
 <p>Seja bem vindo <span class="style1"><%= Session("nome") %></span>
   <a href="logout.aspx">Sair</a> </p>
<form runat="server">
  <table width="369" border="1">
    <tr>
      <td width="88"><div align="center"><strong>ID</strong></div></td>
      <td width="265"><div align="center"><strong>NOME</strong></div></td>
    </tr>
    <tr>
      <td><div align="center">
        <asp:Label ID="lblid" runat="server" />
      </div>      </td>
      <td><div align="center">
        <asp:Label ID="lblnome" runat="server" />
      </div></td>
    </tr>
  </table>
  <br>
</form>
</body>
</html>
