<%

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
<!-- #include file="topo.asp" -->
<%

' Dados do e-mail
strEmail    = request.form("email")
strIdProd   = request.form("idProd")
strNomeProd = request.form("nome")
strFab      = request.form("Fabricante")
strAssunto  = "Solicitação de Produtos"

'Busca o nome do cliente
set rs = abredb.execute("select nome from clientes where email='" & Cripto(Session("usuario"),true) & "'")
if not rs.eof then
strNome = Cripto(rs("nome"),false)
end if
rs.close
set rs = nothing

If strEmail = "" Or instr(strEmail, "@") = 0 Or instr(strEmail, ".") = 0 then
   response.redirect "produtos.asp?produto="&strIdProd&"&erro=Preencha o E-Mail Corretamente!"
end if

'Monta a Mensagem
StrMsg = "Id: " & strIdProd & vbNewLine
strMsg = strMsg & "Nome do Produto: " & strNomeProd & vbNewLine
strMsg = strMsg & "Fabricante: " & strFab & vbNewLine
strMsg = strMsg & "Quem Solicitou: " & strEmail

%>

<!--#include file="email.asp"-->
<%EnviaEmail Application("HostLoja"), Application("ComponenteLoja"), strEmail, "", emailloja, strAssunto, strMsg%>

   			   	 <table><tr><td align="left" valign="top">
				 				<table border="0" cellspacing="4" cellpadding="4" width=590><tr><td><font face="tahoma,arial,helvetica" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> » <%=strLg14%><br><br><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><br><br><p align=center><font style=font-size:17px; color=#000000><b><%=strLg124%></b></font></p><p align=center><%=strLg125%> <b><%if strNome="" then Response.Write "Visitante" else Response.Write strNome end if %></b>, <%=strLg252%><br><%=strLg127%><br></p><p align=center></b><a href=default.asp><img src="<%=dirlingua%>/imagens/continuar.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg68%>';return true;" border=0></a></p> <br><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
								</table></td></tr>
				</table>
<!-- #include file="baixo.asp" -->