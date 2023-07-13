<% 

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
	  <tr><td valign=top>
<%

'Função para chamar os produtos aleatoreamente na vitime inicial
set rs = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok';")
if rs.eof or rs.bof then
rs.close
set rs = nothing
set atualizar = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok';")
if rs.eof or rs.bof then
rs.close
set rs = nothing
set atualizar = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok';")
end if
else
set atualizar = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs("idprod")&";")
end if
set rs2 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&";")
if rs2.eof or rs2.bof then
rs2.close
set rs2 = nothing
set atualizar2 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs2 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&";")
if rs2.eof or rs2.bof then
rs2.close
set rs2 = nothing
set atualizar2 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs2 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&";")
end if
else
set atualizar2 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs2("idprod")&";")
end if
set rs3 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&";")
if rs3.eof or rs3.bof then
rs3.close
set rs3 = nothing
set atualizar3 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs3 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&";")
if rs3.eof or rs3.bof then
rs3.close
set rs3 = nothing
set atualizar3 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs3 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&";")
end if
else
set atualizar3 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs3("idprod")&";")
end if
set rs4 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&";")
if rs4.eof or rs4.bof then
rs4.close
set rs4 = nothing
set atualizar4 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs4 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&";")
if rs4.eof or rs4.bof then
rs4.close
set rs4 = nothing
set atualizar4 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs4 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> '"&rs("idprod")&"' and idprod <> '"&rs2("idprod")&"' and idprod <> '"&rs3("idprod")&"';")
end if
else
set atualizar4 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs4("idprod")&";")
end if
set rs5 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&";")
if rs5.eof or rs5.bof then
rs5.close
set rs5 = nothing
set atualizar5 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs5 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&";")
if rs5.eof or rs5.bof then
rs5.close
set rs5 = nothing
set atualizar5 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs5 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&";")
end if
else
set atualizar5 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs5("idprod")&";")
end if
set rs6 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&";")
if rs6.eof or rs6.bof then
rs6.close
set rs6 = nothing
set atualizar6 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs6 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs2("idprod")&" and idprod <> "&rs("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&";")
if rs6.eof or rs6.bof then
rs6.close
set rs6 = nothing
set atualizar6 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok';")
set rs6 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs2("idprod")&" and idprod <> "&rs("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&";")
end if
else
set atualizar6 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs6("idprod")&";")
end if

intProdID1 = rs("idprod")
intProdID2 = rs2("idprod")
intProdID3 = rs3("idprod")
intProdID4 = rs4("idprod")
intProdID5 = rs5("idprod")
intProdID6 = rs6("idprod")

'Formatação dos preços dos produtos
precito1 = formatNumber(rs("preco"), 2)
precito2 = formatNumber(rs2("preco"), 2)
precito3 = formatNumber(rs3("preco"), 2)
precito4 = formatNumber(rs4("preco"), 2)
precito5 = formatNumber(rs5("preco"), 2)
precito6 = formatNumber(rs6("preco"), 2)%>

	 <table BORDER="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#bfbfbf>
             <form action="comprar.asp" method="post" name="comprar1">
		 	  <table BORDER="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>		 
				  <table align=center width=280 cellspacing="1" cellpadding="1"><tr>
				  <td width=80 height=100><img src="produtos/<%=rs.fields("impeq")%>" width="75" border="0"></td>
				  <td valign=center width=200><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs.fields("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%= strlgmoeda & " " & precito1%><br><b><%=strLg28%></b><%if rs.fields("estoque") = "s" then response.write " " & strLg26 else response.write " " & strLg27 end if%>&nbsp;<%=strLg25%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID1 %>"><input type="hidden" name="intQuant" value=1><%if rs.fields("estoque") = "s" then response.write "<a href=""JavaScript: document.comprar1.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>&nbsp;&nbsp;" end if%>&nbsp;&nbsp;<a href="produtos.asp?produto=<%=rs.fields("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr>
		        </table>
		  </table>
	  </table>
	   
	  <table BORDER="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#bfbfbf>
	         <form action="comprar.asp" method="post" name="comprar2">
	   		 <table BORDER="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>
			  		<table align=center width=280 cellspacing="1" cellpadding="1"><tr>
					<td width=80 height=100><img src="produtos/<%=rs2.fields("impeq")%>" width="75" border="0"></td>
				  <td valign=center width=200><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs2.fields("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%= strlgmoeda & " " & precito2%><br><b><%=strLg28%></b><%if rs2.fields("estoque") = "s" then response.write " " & strLg26 else response.write " " & strLg27 end if%>&nbsp;<%=strLg25%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID2 %>"><input type="hidden" name="intQuant" value=1><%if rs2.fields("estoque") = "s" then response.write "<a href=""JavaScript: document.comprar2.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>&nbsp;&nbsp;" end if%>&nbsp;&nbsp;<a href="produtos.asp?produto=<%=rs2.fields("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr>					
				    </table>
		     </table>
	  </table>
	   
	  <table BORDER="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#bfbfbf>
	         <form action="comprar.asp" method="post" name="comprar3">
	   		 <table BORDER="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>
			  		<table align=center width=280 cellspacing="1" cellpadding="1"><tr>
					<td width=80 height=100><img src="produtos/<%=rs3.fields("impeq")%>" width="75" border="0"></td>
				  <td valign=center width=200><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs3.fields("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%= strlgmoeda & " " & precito3%><br><b><%=strLg28%></b><%if rs3.fields("estoque") = "s" then response.write " " & strLg26 else response.write " " & strLg27 end if%>&nbsp;<%=strLg25%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID3 %>"><input type="hidden" name="intQuant" value=1><%if rs3.fields("estoque") = "s" then response.write "<a href=""JavaScript: document.comprar3.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>&nbsp;&nbsp;" end if%>&nbsp;&nbsp;<a href="produtos.asp?produto=<%=rs3.fields("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr>
					</table>
			 </table>
	 </table>
	  
</td><td valign=top>

	 <table BORDER="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#bfbfbf>
           <form action="comprar.asp" method="post" name="comprar4">
		 	<table BORDER="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>
					  <table align=center width=280 cellspacing="1" cellpadding="1">
					  <tr><td width=80 height=100><img src="produtos/<%=rs4.fields("impeq")%>" width="75" border="0"></td>
				      <td valign=center width=200><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs4.fields("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%= strlgmoeda & " " & precito4%><br><b><%=strLg28%></b><%if rs4.fields("estoque") = "s" then response.write " " & strLg26 else response.write " " & strLg27 end if%>&nbsp;<%=strLg25%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID4 %>"><input type="hidden" name="intQuant" value=1><%if rs4.fields("estoque") = "s" then response.write "<a href=""JavaScript: document.comprar4.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>&nbsp;&nbsp;" end if%>&nbsp;&nbsp;<a href="produtos.asp?produto=<%=rs4.fields("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr>					  
					  </table>
			</table>
	</table>
		
	<table BORDER="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#bfbfbf>
	<form action="comprar.asp" method="post" name="comprar5">
			  <table BORDER="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>
			   		 <table align=center width=280 cellspacing="1" cellpadding="1">
					 <tr><td width=80 height=100><img src="produtos/<%=rs5.fields("impeq")%>" width="75" border="0"></td>
				     <td valign=center width=200><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs5.fields("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%= strlgmoeda & " " & precito5%><br><b><%=strLg28%></b><%if rs5.fields("estoque") = "s" then response.write " " & strLg26 else response.write " " & strLg27 end if%>&nbsp;<%=strLg25%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID5 %>"><input type="hidden" name="intQuant" value=1><%if rs5.fields("estoque") = "s" then response.write "<a href=""JavaScript: document.comprar5.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>&nbsp;&nbsp;" end if%>&nbsp;&nbsp;<a href="produtos.asp?produto=<%=rs5.fields("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr>					 
					 </table>
			  </table>
	</table>
		
	<table BORDER="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#bfbfbf>
	<form action="comprar.asp" method="post" name="comprar6">
			  <table BORDER="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>
			   		 <table align=center width=280 cellspacing="1" cellpadding="1">
					 <tr><td width=80 height=100><img src="produtos/<%=rs6.fields("impeq")%>" width="75" border="0"></td>
				     <td valign=center width=200><font style=font-size:11px;font-family:<%=fonte%>><b><font color=000000><%=rs6.fields("nome")%></b><br><br><b><%=strLg29%></b>&nbsp;<%= strlgmoeda & " " & precito6%><br><b><%=strLg28%></b><%if rs6.fields("estoque") = "s" then response.write " " & strLg26 else response.write " " & strLg27 end if%>&nbsp;<%=strLg25%><div align=right><br><input type="hidden" name="intProdID" value="<%= intProdID6 %>"><input type="hidden" name="intQuant" value=1><%if rs6.fields("estoque") = "s" then response.write "<a href=""JavaScript: document.comprar6.submit();""><img src="&dirlingua&"/imagens/comprar_2.gif border=0 alt=""Comprar Produto"" onMouseOut=""window.status='';return true;"" onMouseOver=""window.status='Comprar Produto';return true;""></a>&nbsp;&nbsp;" end if%>&nbsp;&nbsp;<a href="produtos.asp?produto=<%=rs6.fields("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="color:#FFFFFF;text-decoration:none;"><img src=<%=dirlingua%>/imagens/detalhes.gif border="0"></a></div></td></form></tr>
				     </table>
			  </table>
	 </table> 
</td></tr><tr><td colspan=2><img src=<%=dirlingua%>/imagens/ofertasbaixo.gif border=0></td></tr>

<%
'Fecha as tabelas
rs.close
set rs=nothing
rs2.close
set rs2=nothing
rs3.close
set rs3=nothing
rs4.close
set rs4=nothing
rs5.close
set rs5=nothing
rs6.close
set rs6=nothing
%>
