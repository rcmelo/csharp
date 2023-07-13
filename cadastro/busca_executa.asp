<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/cadastro.asp" -->
<%
'   ******************** WWW.COMERCIALWEB.COM.BR *********************
'   *   Este sistema foi desenvolvido por: www.comercialweb.com.br   *
'   *   desenvolvimentos loja virtual dominio e hospedagem.          *
'   *   E-mail comercial@comercialweb.com.br                         *
'   *   caso tenha alguma dúvida leia o arquivo configuração.txt     *
' 	*   ou envie um E-mail: yy200@ig.com.br  a/c Luciano             *
'   ******************************************************************

'Declara variaveis
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' Executa string 
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_cadastro_STRING
  MM_editTable = "cadastro_cliente"
  MM_editColumn = "cod"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "confirma_exclusao.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cadastro_STRING
Recordset1.Source = "SELECT * FROM cadastro_cliente"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
%>



<%
response.buffer = "true"

'Abre a conexão com o banco de dados
	Set Conexao = Server.CreateObject("ADODB.Connection")
dbPath = "DBQ=" & Server.Mappath("dados/cadastro.mdb")
Conexao.Open "DRIVER={Microsoft Access Driver (*.mdb)};" & dbPath
		
	Set Recordset1 = Server.CreateObject("ADODB.RecordSet")
	Recordset1.Open "SELECT * FROM cadastro_cliente WHERE nome LIKE '%"& request.form("dado") &"%' OR cidade LIKE'%"& request.form("dado") &"%' OR cel LIKE'%"& request.form("dado") &"%' OR tel2 LIKE'%"& request.form("dado") &"%' OR tel1 LIKE'%"& request.form("dado") &"%' OR email LIKE'%"& request.form("dado") &"%' OR endereco LIKE'%"& request.form("dado") &"%'  " , Conexao, 1, 3
%>
<%

Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows

'Senão encontrar nenhum registro '

	if Recordset1.eof then
		response.write "<p>&nbsp;<p><center><font face='Verdana, Arial, Helvetica, sans-serif' size='3'><b>Nenhum cadastro encontrado<br> com estas características. <p><a href='javascript:history.back(1)'>voltar</a></b></font></center>"


		
	else
		%>
<html>
<head>
<title>Busca www.comercialweb.com.br</title>
<meta name="Author" content="Luciano www.comercialweb.com.br">
</head>
<body bgcolor="#FFFFFF">

<% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
<!--Inicio da tabela -->
<table width="100%" border="1" cellspacing="1" cellpadding="0" bordercolor="#6699FF">
  <tr> 
    <td colspan="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Nome:<font color="#0000CC"> 
      <%=Recordset1("nome")%></font></font></b></td>
    <td colspan="4"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Cidade:<font color="#0000CC"> 
      <%=Recordset1("cidade")%></font></font></b></td>
    <td width="15%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Estado:<font color="#0000CC"> 
      <%=Recordset1("estado")%></font></font></b></td>
    <td width="11%" bgcolor="#FF9900"> 
      <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="atualizar_form.asp?cadastro_cliente=<%=Recordset1("Cod")%>">Atualizar</a></font></b></div>
    </td>
  </tr>
  <tr> 
    <td colspan="5"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">End:<font color="#0000CC"> 
      <%=Recordset1("endereco")%></font></font></b></td>
    <td colspan="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Cep:<font color="#0000CC"> 
      <%=Recordset1("cep")%></font></font></b></td>
    <td rowspan="3" width="11%" bgcolor="#CC0000"> 
      <form name="form1" method="POST" action="<%=MM_editAction%>">
        <div align="center">
          <input type="submit" name="Submit" value="&gt;Excluir&lt;">
          <input type="hidden" name="MM_delete" value="true">
          <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("cod").Value %>">
        </div>
      </form>
    </td>
  </tr>
  <tr> 
    <td width="26%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Tel.res.: 
      <%=Recordset1("tel1")%></font></b></td>
    <td colspan="4"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Tel.com.:<font color="#0000CC"> 
      <%=Recordset1("tel2")%></font></font></b></td>
    <td colspan="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Tel.cel.:<font color="#0000CC"> 
      <%=Recordset1("cel")%></font></font></b></td>
  </tr>
  <tr> 
    <td colspan="4"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">E-mail:<font color="#990000"> 
      <a href="mailto:<%=Recordset1("email")%>"><%=Recordset1("email")%></a></font></font></b></td>
    <td colspan="3"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Data 
      cadastro:<font color="#0000CC"> <%=Recordset1("data")%></font></font></b></td>
  </tr>
  <tr bgcolor="#6699FF"> 
    <td colspan="8"> <font size="-2">&nbsp;</font></td>
  </tr>
</table>
<!--Fim da tabela -->
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
<!--Fecha a conexão -->
  <%end if %>
  <%
Recordset1.Close()
%>
</body>
</html>