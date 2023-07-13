<%

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
<!-- #include file="topo.asp" -->

<script>
	function cadmail(){
		window.open('cadmail.asp?email='+document.emailx.email.value,'email','resizable=no,width=270,height=180,scrollbars=no');
		document.emailx.reset()
	}
	function limpa() {
	document.emailx.email.value=''
	}
	
</script>

		 <table width=100% cellspacing="3" cellpadding="1" align=right>
		 		<tr><td colspan=2><img src=<%=dirlingua%>/imagens/ofertasdodia.gif border=0></td></tr>
				<!-- #include file="produto.asp" -->
				<tr><td>
				<table BORDER="0" CELLSPACING="1" CELLPADDING="0"><tr><td bgcolor=#bfbfbf>
					   <table BORDER="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>
					   		  <table border=0 align=center width=280  height=125 cellspacing="0" cellpadding="0">
                              <tr><td align=right valign=bottom height=21 colspan=2><img src=<%=dirlingua%>/imagens/trofeu.gif width=15 height=21 border=0>&nbsp;&nbsp;<font face=<%=fonte%> style=font-size:11px;color:000000;><b>CAMPEÕES DE VENDA</b></td></tr>
							  <tr><td height=1 colspan=2><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr>
				<%
				if VersaoDb = 1 then
				set rs = abredb.execute("SELECT idprod, nome, preco FROM produtos ORDER BY contador DESC LIMIT 0,3")
				else
				set rs = abredb.execute("SELECT TOP 3 idprod, nome, preco FROM produtos ORDER BY contador DESC")
				end if
				cont = 1
				if not rs.eof then
				  for x = 1 to 3%>
							  <tr><td valign=top width=15><font face=<%=fonte%> style=font-size:11px;color:000000;><%=cont & ") "%></font></td><td valign=top width=260><font face=<%=fonte%> style=font-size:11px;color:000000;><a href="produtos.asp?produto=<%=rs("idprod")%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='+ Detalhes';return true;" style="text-decoration: none;"><%=rs("nome")%><br><b><%= strlgMoeda & " " & formatnumber(rs("preco"),2)%></b></font></A></td></tr>
							  <%cont = cont + 1%>
				              <%rs.movenext%>
				  <%next%>
				<%end if%>
					          </table>
					   </table>
				</table>
			    </td><td>
				<%rs.close%>
				<%set rs = nothing%>
				
		<form method=post action="javascript: cadmail()" name=emailx>

		<table BORDER="0" CELLSPACING="1" CELLPADDING="0">
			   <tr><td bgcolor=#bfbfbf>
			   <table BORDER="0" CELLSPACING="1" CELLPADDING="3"><tr><td bgcolor=#ffffff>
			   		  <table align=center width=280 height=125 cellspacing="1" cellpadding="1"><tr><td><font color=000000><font style=font-size:11px;font-family:<%=fonte%>><b><%=strLg158%></b><br><%=strLg159%><br><br><center><font color=red size=1><%=erro%></font><Br><center><input type=text size=30 value="<%=strLg160%>" style=font-size:11px;font-family:<%=fonte%>;font-weight:bold name=email onfocus="limpa();"><br><input type=image src=<%=dirlingua%>/imagens/cadastro.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=srtLg161%>';return true;"><br><br></center> </td></tr>
					  </table>
		       </table>
	    </table>
		</td></tr>

		<tr><td colspan=2></td></tr>
		<tr><td colspan=2></td></tr>
 </table>
    </form>
	<!-- #include file="baixo.asp" -->