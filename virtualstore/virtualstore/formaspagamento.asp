<%

'IN�CIO DO C�DIGO
'----------------------------------------------------------------------------------------------------------------
'Este c�digo est� otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
<!-- #include file="topo.asp" -->
<%
'Ve se o usuario est� logado
if session("usuario") = "" then
response.redirect "fechapedido.asp?compra=login"
end if
intOrderID = cstr(Session("orderID"))%>
		   <table><tr><td align="left" valign="top"><table border="0" cellspacing="4" cellpadding="4" width=590><tr><td>
		   				  <script language="javascript">
						  function limpar(){
						  		   document.registro1.reset()
						  }
						  </script>
  		   <form action=pagamento.asp method=post name=registro1>
				 <input type="hidden" name="frete1" value="<%=session("frete1")%>">
				 <input type="hidden" name="email1" value="<%=session("usuario")%>">
				 <input type="hidden" name="pedido1" value="<%=session("pedido1")%>">
				 <input type="hidden" name="nome1" value="<%=session("nome1")%>">
				 <input type="hidden" name="ende1" value="<%=session("ende1")%>">
				 <input type="hidden" name="bairro1" value="<%=session("bairro1")%>">
				 <input type="hidden" name="cidade1" value="<%=session("cidade1")%>">
				 <input type="hidden" name="est1" value="<%=session("est1")%>">
				 <input type="hidden" name="cep1" value="<%=session("cep1")%>">
				 <input type="hidden" name="pais1" value="<%=session("pais1")%>">
				 <input type="hidden" name="fone1" value="<%=session("fone1")%>">
				 <input type="hidden" name="compra1" value="<%=session("compra1")%>">
				 <input type="hidden" name="totalae" value="<%=session("totalcompra1")%>">
				 <input type="hidden" name="idcompra" value="<%=intOrderID%>">
				 		<font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg4%>';return true;"><b><%=strLg4%></b></a> � <%=strLg8%><br><br>
							  <div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></div>
<%
'Verifica se h� itens na compra
if cstr(Session("orderID")) = "" then%>
   							<br><br><div align=center><p><font face=<%=fonte%> style=font-size:17px><b><%=strLg49%><br><%=strLg50%><br><br></b><a href=default.asp><img src="<%=dirlingua%>/imagens/continuar.gif" onMouseOut="window.status='';return true;" onMouseOver="window.status='Continuar Compras';return true;" border=0></a><b></font></p></div><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table>
				</td></tr>
			</table></td></tr>
		</table>
		<!-- #include file="baixo.asp" -->
<%
response.end
else
%>


  			 		  <font face="<%=fonte%>" style=font-size:13px><strong><%=strLg91%></strong></font></p>
					  <font face="<%=fonte%>" style=font-size:11px><%=strLg92%><br><%=strLg93%> <a href=fechapedido.asp?compra=ok onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg71%>';return true;"><b><%=strLg71%></b></a>
					  		<div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></div>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111"><tr><td colspan=2><font style=font-size:15px><b>�</b>  <font face="<%=fonte%>" style=font-size:11px;color:000000><strong><%=strLg96%></strong></font> <font style=font-size:11px color=red><%=request.querystring("erro")%></font><br></td></tr><tr><td valign=botton><font style=font-size:11px>&nbsp;&nbsp;&nbsp;<%=strLg97%><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <div align="center">
      <center>
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber2">
        <tr>
          <td align="center"><font style=font-size:11px face="<%=fonte%>"><img src=<%=dirlingua%>/imagens/visa.gif border=0></td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>"> <img src=<%=dirlingua%>/imagens/mastercard.gif border=0></td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>"> <img src=<%=dirlingua%>/imagens/amex.gif border=0></td>
          <td align="center"><font style=font-size:11px face="<%=fonte%>"> <img src=<%=dirlingua%>/imagens/diners.jpg border=0></td>
        </tr>
        <tr>
          <td align="center"><input type=radio name=cartao value=V></td>
          <td align="center"><input type=radio name=cartao value=M></td>
          <td align="center"><input type=radio name=cartao value=A></td>
          <td align="center"><input type=radio name=cartao value=D></td>
        </tr>
      </table>
      </center>
    </div>
    </td><td rowspan=3><font style=font-size:11px><%=strLg98%> </font><br><input type="text" name="numerocartao" size="30" style=font-family:<%=fonte%>;font-size:11px;color:#000000;><br><font style=font-size:11px><%=strLg99%> </font><br><select style="font-family:<%=fonte%>;font-size:11px;color:#000000;" name="cartaomes"><option value=""> <%=strLg148%> <option value="01">Jan <option value="02">Fev<option value="03">Mar<option value="04">Abr<option value="05">Mai <option value="06">Jun <option value="07">Jul<option value="08">Ago<option value="09">Set<option value="10">Out<option value="11">Nov<option value="12">Dez</option></select><b>/</b> <select name='cartaoano' style=font-size:11px;font-family:<%=fonte%>;> <Option value=""><%=strLg101%></Option><% ano = year(now)
'Calcula o ano da data atual + 10
 For data = ano to ano + 10
response.write "<Option value="&data&">"&data&"</Option>"
next
%>
  			   								 </select><br><font style=font-size:11px><%=strLg102%> </font><br>
    <input type="text" name="titularcartao" size="43" style=font-family:<%=fonte%>;font-size:11px;color:#000000;></td></tr>
											<tr><td valign=top>&nbsp;</td></tr>
	   </font>
					  <font face="Tahoma" style=font-size:11px color="#FF0000">
	   </font>
					  <font face="<%=fonte%>" style=font-size:11px><tr><td colspan=2><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr><tr><td colspan=2><font style=font-size:15px><b>�</b> <font face="<%=fonte%>" style=font-size:11px;color:000000><strong><%=strLg95%></strong></font> <br></td></tr>
											<tr><td valign=botton widht=100% colspan=2>
                                              <p align="center"><font style=font-size:11px>
                                              <br>
                                              <br>
                                              &nbsp;<%=strLg94%></font></p>
                                              <div align="center">
                                                <center>
                                                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1">
                                                  <tr>
                                                    <td width="25%" align="center">
                                                    <font style=font-size:11px face="<%=fonte%>"><img src="<%=dirlingua%>/imagens/pg_deposito.gif" border=0></font></td>
                                                    <td align="center">
                                                    <font style=font-size:11px face="<%=fonte%>"><img src="<%=dirlingua%>/imagens/pg_transf.gif" border=0></font></td>
                                                    <td width="25%" align="center">
                                                    <font style=font-size:11px face="<%=fonte%>">
                                                    <img src="<%=dirlingua%>/imagens/pg_sedexacobrar.gif" border=0></font></td>
                                                    <td width="25%" align="center">
                                              <font style=font-size:11px face="<%=fonte%>">
                                              <img src="<%=dirlingua%>/imagens/pg_boleto.gif" border=0></td>
                                                  </tr>
                                                  <tr>
                                                    <td width="25%" align="center">
					  <font face="<%=fonte%>" style=font-size:11px>
                                              <input type=radio name=cartao value=di checked></font></td>
                                                    <td width="25%" align="center">
					  <font face="<%=fonte%>" style=font-size:11px><input type=radio name=cartao value=db></font></td>
                                                    <td width="25%" align="center">
					  <font face="<%=fonte%>" style=font-size:11px><input type=radio name=cartao value=sc></font></td>
                                                    <td width="25%" align="center">
					  <font face="<%=fonte%>" style=font-size:11px>
                                                    <input type=radio name=cartao value=bl></font></td>
                                                  </tr>
                                                </table>
                                                </center>
                                              </div>
                                              </td></tr>
											<tr><td valign=top colspan=2>&nbsp;</td></tr>
											<tr><td colspan=2><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5>
                                              <p align="center">
					  <font face="<%=fonte%>" style=font-size:11px><br>
                                              <input type=image src=<%=dirlingua%>/imagens/prosseguir.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg67%>';return true;">&nbsp; <input type=image src=<%=dirlingua%>/imagens/limpar.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='<%=strLg89%>';return true;" OnClick="javascript: limpar()"></font></td></tr></table></td></tr><tr align=center><td valign=top>&nbsp;</td><td valign=top></form></td></tr>
							</table>
						</center>
</div>
<p><br>
&nbsp;</p>
						</form></a></td></tr>
				</table></td></tr>
		</table>
	   </font>
	   <!-- #include file="baixo.asp" -->
<%end if%>