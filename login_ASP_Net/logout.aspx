<%@ Page Language="VB" Debug="true" %>

<script language="VB" runat="server">
'Ao carregar esta p�gina o script abaixo abandona a sessao e redireciona o usu�rio para o login.

Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) 

session.Abandon()
Response.Redirect("index.aspx")

End Sub

</script>
