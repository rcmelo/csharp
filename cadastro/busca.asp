<%
Response.Cookies("dado") = Request.Form("dado")
response.cookies("dado").expires = "6/25/03"


%><html>
<head>
<title>Busca</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>



<body bgcolor="#FFFFFF" text="#000000">
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="100%" border="1" cellspacing="0" bgcolor="#6699FF" bordercolor="#6699FF">
  <tr> 
    <td width="45%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><font color="#000000">Busca 
      por:</font> <br>
      <font color="#990000">Nome End. Cidade E-mail ou Tel.</font></b></font></td>
    <td width="55%" bgcolor="#FFFFFF" valign="baseline"> 
      <form name="form1" method="post" action="busca_executa.asp">
        <div align="left">
          <input type="text" name="dado">
          <input type="submit" name="Submit" value="( Buscar )">
        </div>
      </form></td>
  </tr>
</table>
<p>&nbsp;</p>

</body>
</html>
