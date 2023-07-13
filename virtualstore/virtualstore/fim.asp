<%

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP


'Desloga o usuario
response.buffer = true
session("usuario") = request.form("usuarioz")
response.redirect "default.asp"
%>