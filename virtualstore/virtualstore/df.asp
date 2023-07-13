<% 

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP

'Declaração das variaveis comuns
Dim razaoloja
Dim bancopag
Dim contapag
Dim pagpara
Dim varimg
Dim pesquisa
Dim strTextoHtml 
Dim conexao 
DIm abredb
Dim dados 
Dim senhamaster 
Dim nomeloja 
Dim slogan 
Dim emailloja 
Dim urlloja 
Dim tituloloja 
Dim endereco11 
Dim bairro11 
Dim cidade11 
Dim estado11 
Dim pais11 
Dim fone11 
Dim razao 
Dim Mes 
Dim meszz 
Dim diazz 
Dim dataz 
Dim i 
Dim dia 
Dim mez 
Dim strLink 
Dim strAcao 
Dim contacompra 
Dim contacli 
Dim estadoz 
Dim rs 
Dim r2 
Dim finalera 
Dim pag 
Dim pesss 
Dim pagdxx 
Dim pesqsa 
Dim catege 
Dim fDia 
Dim mezito 
Dim anito 
Dim data 
Dim Ano 
Dim j 
Dim ndiasmes 
Dim anozinho 
Dim palavra 
Dim inicial 
Dim final 
Dim restinho 
Dim totalreg 
Dim pagina2 
Dim pagina3 
Dim rs2 
Dim nSem 
Dim aDiasMes 
Dim strString 
Dim UploadRequest 
Dim objFSO 
Dim ObjFile 
Dim ObjStream 
Dim arquivodat 
Dim separador 
Dim senhaok 
Dim VersaoDb
Dim StringdeConexao

'Seta as datas para: dd/mm/aaaa e moedas para: R$ 0.00*
Session.LCID = 1046

'Comando para ativar o uso do response.redirect 
response.buffer = true

'Comando faz pular a linha em caso de erro
on error resume next

'Chama o config.asp
%><!--#include file="config.asp"--><%

'Chama o criptografia.asp
%><!--#include file="criptografia.asp"--><%

'Chama os dados da loja
senhamaster = application("senhamaster")
nomeloja = application("nomeloja")
emailloja = application("emailloja")
slogan = application("slogan")
urlloja = application("urlloja")
endereco11 = application("endereco11")
bairro11 = application("bairro11")
cidade11 = application("cidade11")
estado11 = application("estado11")
pais11 = application("pais11")
fone11 = application("fone11")
bancopag = application("bancopag")
agpag = application("agpag")
contapag = application("contapag")
pagpara = application("pagpara")
tituloloja = nomeloja&" - "&slogan
fonte = application("fonte")
maopropria = application("maopropria")
diasentrega = application("diasentrega")
corfundo = application("corfundo")
cor1 = application("cor1")
cor2 = application("cor2")
cor3 = application("cor3")
cor4= application("cor4")
cor5= application("cor5")
fontebranca = application("fontebranca")
fontepreta = application("fontepreta")
corborda = application("corborda")

'Variaveis para calculo do frete
if session("frete") = "" then
valorfrete = 0.00
else
valorfrete = session("frete") + maopropria
session.timeout = 999
end if
estatus = left(ucase(session("estado")),2)
session.timeout = 999
estaos = session("estado")
session.timeout = 999
if session("peso") = "" then
pesoz = 0
else
pesoz = session("peso")
end if
if session("estado2") = "" then
estado2 = ""
else
estado2 = session("estado2") 
end if

if session("df") = "sim" then 
session("df") = "nao"
response.end
end if

StrTopo = "<!--" & vbNewline
StrTopo = StrTopo & "O que você está procurando?" & vbNewline
StrTopo = StrTopo & "What are you lookin' for?" & vbNewline
StrTopo = StrTopo & "" & vbNewline
StrTopo = StrTopo & "Copyright(c) VirtuaStore Soluções em comércio eletrônico" & vbNewline
StrTopo = StrTopo & "All Rights Reserved." & vbNewline
StrTopo = StrTopo & "------------------------------" & vbNewline
StrTopo = StrTopo & "http://www.virtuastore.com.br" & vbNewline
StrTopo = StrTopo & "-->" & vbNewline

response.write StrTopo

Sub BaixoC()
response.write "<a class=menu  href=""http://www.virtuastore.com.br"" target=_new>Powered by VirtuaStore 1.2</a>"
End Sub
%>
