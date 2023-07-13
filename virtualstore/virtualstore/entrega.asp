<%

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
<!--#include file="topo.asp"-->
<%
'Chama as variaveis
strUser = request.form("usuario")
strIdcompra = Request.Form("idcompra")
strPedido = Request.Form("pedido")
strNome = Request.Form("nome")
strCartao = Request.Form("cartao")
strPresente = Request.Form("presentesim")
msgpres = Request.Form("cartao")
strEndereco = request.form("endereco")
strBairro = request.form("bairro")
strCidade = request.form("cidade")
strEstado = request.form("estadoz")
strOutropais = request.form("outroz")
strCep = request.form("cepzz")
strPais = request.form("pais")
strFone = request.form("fone")	
intOrderID = Request.form("idcompra")
fretezz = request.form("frete")
totocompraz = request.form("totalcompra")

'Valida os dados
if strPais = "" then
strPais = "Brasil"
end if
If strEndereco = "" then
erro = "- " & strLg174 & "<br>" 
abredb.close
set abredb = nothing
response.redirect "fechapedido.asp?compra=ok&erro2="&erro&"&endereco="&strEndereco&"&bairro="&strBairro&"&cidade="&strCidade&"&cep="&strCep&"&fone="&strFone&"&estado="&strEstado&"&pais="&strPais&"&outropais="&strOutropais&"&#pagamento"
end if
If strBairro = "" then
erro = "- " & strLg175 & "<br>" 
abredb.close
set abredb = nothing
response.redirect "fechapedido.asp?compra=ok&erro2="&erro&"&endereco="&strEndereco&"&bairro="&strBairro&"&cidade="&strCidade&"&cep="&strCep&"&fone="&strFone&"&estado="&strEstado&"&pais="&strPais&"&outropais="&strOutropais&"&#pagamento"
end if
If strCidade = "" then
erro = "- " & strLg176 & "<br>" 
abredb.close
set abredb = nothing
response.redirect "fechapedido.asp?compra=ok&erro2="&erro&"&endereco="&strEndereco&"&bairro="&strBairro&"&cidade="&strCidade&"&cep="&strCep&"&fone="&strFone&"&estado="&strEstado&"&pais="&strPais&"&outropais="&strOutropais&"&#pagamento"
end if
If strCep = "" then
erro = "- " & strLg177 & "<br>" 
abredb.close
set abredb = nothing
response.redirect "fechapedido.asp?compra=ok&erro2="&erro&"&endereco="&strEndereco&"&bairro="&strBairro&"&cidade="&strCidade&"&cep="&strCep&"&fone="&strFone&"&estado="&strEstado&"&pais="&strPais&"&outropais="&strOutropais&"&#pagamento"
end if
If strFone = "" then
erro = "- " & strLg178 & "<br>" 
abredb.close
set abredb = nothing
response.redirect "fechapedido.asp?compra=ok&erro2="&erro&"&endereco="&strEndereco&"&bairro="&strBairro&"&cidade="&strCidade&"&cep="&strCep&"&fone="&strFone&"&estado="&strEstado&"&pais="&strPais&"&outropais="&strOutropais&"&#pagamento"
end if

'Ve se foi escohida a opção presente
if strPresente = "on" then
strPresente = "Sim"
else
strPresente = "Não"
end if

'Formata a data
if CStr(len(Day(now))) = CStr("1") then
diazz =  "0"&Day(now)
else
diazz = Day(now)
end if
if CStr(len(Month(now))) = CStr("1") then
meszz =  "0"&Month(now)
else
meszz = Month(now)
end if
datazz = diazz&"/"&meszz&"/"&Year(now)

'SQL para gravar a compra
textosql = "UPDATE compras SET clienteid = '"&Cripto(request.form("usuario"),true)&"', presente = '"&strPresente&"', " _ 
& "msgpresente = '"&msgpres&"', status = 'Compra em Aberto', frete = '"&Cripto(request.form("frete"),true)&"', " _
& "totalcompra = '"&Cripto(request.form("totalcompra"),true)&"', endentrega = '"&Cripto(request.form("endereco"),true)&"', " _
& "bairroentrega = '"&Cripto(request.form("bairro"),true)&"', cidadeentrega = '"&Cripto(request.form("cidade"),true)&"', " _
& "estadoentrega = '"&request.form("estadoz")&"', " _
& "cepentrega = '"&Cripto(request.form("cepzz"),true)&"', paisentrega = 'Brasil', telentrega = '"&Cripto(request.form("fone"),true)&"', " _
& " datacompra = '"&datazz&"' WHERE idcompra = " & strIdcompra & ";"

'Grava a compra
set cadnovo = abredb.Execute(textosql)

'Cria as sessões para prosseguir a compra
session("frete1") = fretezz
session("pedido1") = strPedido
session("nome1") = strNome
session("ende1") = strEndereco
session("bairro1") = strBairro 
session("cidade1") = strCidade
session("est1") = strEstado
session("cep1") = strCep
session("pais1") = strPais
session("fone1") = strFone
session("compra1") = totocompraz

'Fecha o banco de dados
abredb.close
set abredb = nothing

response.redirect "formaspagamento.asp"
%>
