

<!-- #include file="topo.asp" -->
	 <table><tr><td align="left" valign="top">
	 				<table border="0" cellspacing="4" cellpadding="4" width=590><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> » <%=strLg1%><br><br><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table>
<%
'Chama os produtos comprados

set pedidos = abredb.Execute("SELECT idprod, quantidade FROM pedidos WHERE idcompra='" & intOrderID & "'")
if pedidos.eof then
Session("orderID") = "" %>

				<br><br><div align=center><p><font face=<%=fonte%> style='font-size:17px'><b><%=strLg49%><br> <%=strLg50%><br><br></b><a href=default.asp><img src="<%=dirlingua%>/imagens/continuar.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg68%>';return true;" border=0></a><b></font></p></div><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
			</table></td></tr>
	</table>
	<!-- #include file="baixo.asp" -->
<%
response.end
else
intOrderID = cstr(Session("orderID"))
%>  
		   <font face="<%=fonte%>" style=font-size:13px><strong><%=strLg65%></strong></font></p><p><font face="<%=fonte%>" style=font-size:10px;color:000000><%=strLg56%><br><%=strLg57%></font></p><font face="<%=fonte%>" style=font-size:10px;color:red><%=request.querystring("erro")%></font><form action="atualizapedido.asp" method="post" name="registro1">
		   		 <table border="0" cellpadding="2" cellspacing="1" width="570" align=center>
				 	<tr bgcolor="<%=cor3%>"><td width="15%" align="left" valign="center" nowrap><font color="#000000" face="<%=fonte%>" style=font-size:11px;><strong><%=strLg58%><strong></font></strong></strong></td><td width="39%" align="left" valign="center" nowrap><font color="#000000" face="<%=fonte%>" style=font-size:11px;><strong><%=strLg59%><strong></font></strong></strong></td><td width="16%" align="left" valign="center"><font color="#000000" face="<%=fonte%>" style=font-size:11px;><strong><%=strLg60%><strong></font></strong></strong></td><td width="20%" align="right" valign="center" nowrap><font color="#000000" face="<%=fonte%>" style=font-size:11px;><strong><%=strLg61%><strong></font></strong></strong></td><td width="10%" align="right" valign="center" nowrap><font color="#000000" face="<%=fonte%>" style=font-size:11px;><strong><%=strLg62%><strong></font></strong></strong></td></tr>
<%
while not pedidos.EOF
idprod = pedidos("idprod")
quantidade = pedidos("quantidade")
set produtos = abredb.Execute("SELECT preco, nome, peso FROM produtos WHERE idprod="&idprod&";")
if produtos.eof then
Session("orderID") = ""
response.redirect "carrinhodecompras.asp"
else
end if
preco = produtos("preco")
peso = produtos("peso")
nome = produtos("nome")

intProdID = idprod
strProdNome = nome
pesoz = peso
intProdPrice = preco
intQuant = quantidade
if session("estado") = "" then
intFrete = 0
else
intFrete = valorfrete
end if

'Calcula o total do frete
intTotalFrete = intTotalFrete + (intQuant * intProdPrice)	
intTotal = intTotalFrete + intFrete	
subpreco = formatNumber(intProdPrice,2)
totpreco = formatNumber((intQuant * intProdPrice),2)%>
		 <script LANGUAGE="JScript">
		 	function AbortEntry(sMsg, eSrc){window.alert(sMsg);eSrc.focus();}
			function HandleError<%= intProdID %>(eSrc){var val = parseInt(eSrc.value);
				 if (isNaN(val)){
				 	document.registro1.quant<%= intProdID %>.value = '<%=intQuant%>';
				}
				if (val <= 0) {
				   document.registro1.quant<%= intProdID %>.value = '<%=intQuant%>';
				}
			}
		</script>
				 <input type="hidden" name="peso<%= intProdID %>" value="<%=pesoz%>">
				 <tr><td width="15%" valign="center" align="left"><font face="<%=fonte%>" size="2"> <input name="quant<%= intProdID %>" size="2" value="<%=intQuant%>" onChange="HandleError<%= intProdID %>(this)" style=font-size:11px;font-family:<%=fonte%>; maxlength=2></font></td><td width="39%" align="left"><font face="<%=fonte%>" style=font-size:11px;color:000000><b><%= strProdNome %></b></font></td><td width="16%" align="right"><font face="<%=fonte%>" style=font-size:11px;color:000000><%= strLgMoeda & " " & subpreco %></font></td><td width="20%" align="right"><font face="<%=fonte%>" style=font-size:11px;color:000000><%= strLgMoeda & " " & totpreco %></font></td><td width="10%" align="right"><center><a href="atualizapedido.asp?acao=remover&produto=<%= intProdID %>&frete=<%=estado2%>" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg62%>';return true;"><img src=<%=dirlingua%>/imagens/del.gif border=0></a></td></tr>
<%pedidos.MoveNext
wend
pedidos.Close
set pedidos = Nothing
produtos.Close
set produtos = Nothing

'Valor total da compra
totaldasuacompra = formatNumber(intTotalFrete,2)
suacompra = formatNumber(intTotal,2)
%>    
	  	  <tr bgcolor=<%=cor3%>><td colspan=5 heigth=5></tr>
		  <tr><td valign="center" align="left" colspan=3><font face="<%=fonte%>" style=font-size:11px;color:000000><strong><%=strLg63%><strong></font></strong></strong></td><td colspan="3" align="right" valign="center"><font face="<%=fonte%>" style=font-size:11px;color:000000><strong><b><%= strLgMoeda & " " & totaldasuacompra %></b><strong></font></strong></strong></td></tr>
		  <!--#include file="frete.asp" -->
		  <tr bgcolor=<%=cor3%>><td colspan=5 heigth=5></td></tr>
		  <tr><td valign="center" align="left" colspan=3><font face="<%=fonte%>" style=font-size:11px;color:000000><strong><%=strLg65%><strong></font></strong></strong></td><td colspan=2 align="right" valign="center"><font face="<%=fonte%>" style=font-size:11px;color:000000><strong><b><%= strLgMoeda & " " & suacompra %></b><strong></font></strong></strong></td><% session("totalcompra1") = intTotal %></tr>
		  <tr><td colspan=5><br></td></tr><tr><td colspan=5 align=center><table><tr><td valign=top><center><input type="image" name="control" name="Atualizar" src=<%=dirlingua%>/imagens/atualiza.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg66%>';return true;">&nbsp;&nbsp;&nbsp;</td></form><td valign=top><form action="<%=link%>" method="post" id="form1a" name="form1a"><input type="image" name="control" value="fpedido" src=<%=dirlingua%>/imagens/pross2.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg67%>';return true;"></td></tr>
		 </table></td></tr>
	  </table><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><br><center></b><a href=default.asp><img src="<%=dirlingua%>/imagens/continuar.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg68%>';return true;" border=0></a></center>
<% end if %>
   </td></tr>
  </table></td></tr>
 </table>
 <!-- #include file="baixo.asp" -->