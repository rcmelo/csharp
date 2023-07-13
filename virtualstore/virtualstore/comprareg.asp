<%

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
<!-- #include file="df.asp" -->
<%
'Loga o usuario
if request.querystring ("login") = "sim" then
username = request("email")
password = request("senha") 
set passSet = abredb.Execute("select email,senha from clientes where email='"&Cripto(username,true)&"';")
if passSet.EOF then 

'Fecha o banco de dados
abredb.close
set abredb = nothing
response.redirect"fechapedido.asp?compra=entrar&erro=- " & strLg183 & "&user=x"
else

'Valida a senha
real_password = cripto(trim(passSet("senha")),false)
if password = real_password then
Application.lock
session("usuario") = username
Application.unlock
session.timeout = 999
response.cookies(""&nomeloja&"")("usuario")= username
response.cookies(""&nomeloja&"").expires = "01/01/"&year(now) + 1
else 
abredb.close
set abredb = nothing
response.redirect "fechapedido.asp?compra=entrar&erro2=- " & strLg184 & "&user="&username
end if
end if 
abredb.close
set abredb = nothing
response.redirect "default.asp"
passSet.close
set passSet = nothing

'Fecha o banco de dados
abredb.close
set abredb = nothing
else
username = request("email")
password = request("senha") 
checkPass = "select email,senha from clientes where email='"&Cripto(username,true)&"';"
set passSet = abredb.Execute(checkPass)
if passSet.EOF then 
abredb.close
set abredb = nothing
response.redirect"fechapedido.asp?compra=login&erro=- " & strLg183 & "&user=x"
else
real_password = Cripto(trim(passSet("senha")),false)
if password = real_password then
Application.lock
session("usuario") = username
Application.unlock
session.timeout = 999
response.cookies(""&nomeloja&"")("usuario")= username
response.cookies(""&nomeloja&"").expires = "01/01/"&year(now) + 1
else
abredb.close
set abredb = nothing
response.redirect "fechapedido.asp?compra=login&erro2=- " & strLg184 & "&user="&username
end if
end if 
passSet.close
set passSet = nothing
abredb.close
set abredb = nothing
response.redirect "fechapedido.asp?compra=ok"
end if
%>