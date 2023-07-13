
'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
<!--#include file="df.asp"-->
<html>
	  <head>
	  		<title><%=nomeloja%> - <%=slogan%></title>
			 <style type="text/css">
			 		<!--
					a:link       { color: <%=cor4%> }
					a:visited    { color: <%=cor4%> }
					a:hover      { color: <%=cor3%> }
					-->
			</style>
			   		</head>
							<body>
<%
'Valida o e-mail
strEmail = Request.querystring("email")
If strEmail = "" Or instr(strEmail, "@") = 0 Or instr(strEmail, ".") = 0 then%>
								 <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><br><br><center><b><font style=font-size:13px;font-family:<%=fonte%>><%=strLg152%><br><br><br><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><font style=font-size:11px><a href="javascript: window.close()" style="text-decoration:none;">:: <%=strLg153%> ::</a>
<%
response.end
end if
set abredb = Server.CreateObject("ADODB.Connection")
	abredb.Open(StringdeConexao)

'Grava e-mail no banco de dados
abredb.Execute("INSERT INTO newsletter (datacad, email) VALUES ('"&date&"', '"&strEmail&"');")
abredb.Close
set abredb = Nothing%> 
	<basefont face="<%=fonte%>"><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><br><center><font style=font-size:13px; color=#000000><b><%=strLg154%></b></font></p><p align=center><font style=font-size:11px; color=#000000><%=strLg155%> <%=nomeloja%>!<br><br><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table><center><font style=font-size:11px;><a href="javascript: window.close()" style="text-decoration:none;"><b>:: <%=strLg153%> ::</b></a></p>