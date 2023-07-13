<% 

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
  	  </td></tr></table>
	</table>
	
	<div id="layer1" style="position:absolute; z-index:8;background-color:ffffff; width: 775px;">
	<div id="layer1" style="position:absolute; left:0px; top:-2px; width:775px; z-index:1"><table border=0 cellspacing=0 width=775 cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor="bcbcbc"></td></tr><tr><td height=5></td></tr></table></div>
	<div id="layer1" style="position:absolute; left:0px; top:-7px; width:775px; z-index:1"><table border=0 cellspacing=0 width=775 cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor="bcbcbc"></td></tr><tr><td height=5></td></tr></table></div>

	<table border=0 bgcolor=#ffffff cellpadding=3 cellspacing=3 width=770 align=center>
		   <tr><td height=10></td><td></td></tr><tr align=center><td width=40></td><td  width=660><font face="<%=fonte%>" style=font-size:11px>&nbsp;|&nbsp;<a class=baixo  href="default.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><%=strLg4%></a>&nbsp;|&nbsp;
<%'Inicia os links inferiores
Set menui = abredb.Execute("SELECT * FROM sessoes WHERE ver='s' ORDER by nome;")
While Not menui.EOF%>
	  	  			<font face="<%=fonte%>" style=font-size:11px><a class=baixo  href="sessoes.asp?item=<%= menui("id") %>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%= menui("nome") %>';return true;"><%= menui("nome") %></a>&nbsp;|&nbsp;
<%menui.MoveNext
Wend
menui.Close
Set menui = Nothing%> 
<a class=baixo  href="quemsomos"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg20%>';return true;"><%=strLg20%></a>&nbsp;|&nbsp;<a class=baixo  href="seguranca.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg19%>';return true;"><%=strLg19%></a>&nbsp;|&nbsp;<a class=baixo  href="como.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg16%>';return true;"><%=strLg16%></a>&nbsp;|&nbsp;<a class=baixo  href="carrinhodecompras.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg1%>';return true;"><%=strLg1%></a>&nbsp;|&nbsp;<a class=baixo  href="<%=link%>"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg8%>';return true;"><%=strLg8%></a>&nbsp;|&nbsp;<a class=baixo  href="registro.asp"  onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg5%>';return true;"><%=strLg5%></a></div>
   			</td></tr>
	</tbody></table>
	<table border=0 bgcolor=#ffffff cellpadding=0 cellspacing=0 width="775">
		   <tr><td>&nbsp;</td></tr><tr><td colspan="2" height="20"  bgcolor="<%=cor2%>" width="100%"><font face="<%=fonte%>" style=font-size:11px;color:<%=fontebranca%>><b>&nbsp;&nbsp;&nbsp;<%response.write "© Copyright 2001/"&year(now)&" "&nomeloja&"&nbsp;"%></b>|<b>&nbsp;<%call BaixoC()%>
		   </td></tr>
	</table>
	</div>
	</div>
</form>
<%
'Fecha o banco de bados
abredb.close
set abredb = nothing

if err.number <> 0 then
Select Case Err.number
Case 0
Case 424
Case 3021
Case -2147352567
Case Else
Response.clear
Response.Write "<title>Erro!</title><center><br><br><font face=tahoma style=font-size:11px><h1>Erro no Sistema!</h1><br><br>N°. do erro:" & Err.number & "<br>Descrição do erro:" & Err.Description & "<br><hr size=1><br><font color=red><b>Informe ao administrador do sistema os erros acima!</b></font><br><br>Pedimos desculpas pela inconveniência.<br><br><hr size=1><br>Se você estiver inserindo dados em sua loja NÃO use: aspas simples ('), barra (/) e parêntes (())<br><Br><a href=""javascript: history.go(-1);"">Voltar para página anterior</a>"
End Select
else
end if
%>
