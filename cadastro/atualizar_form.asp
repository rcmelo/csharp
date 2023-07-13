<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/cadastro.asp" -->

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_cadastro_STRING
  MM_editTable = "cadastro_cliente"
  MM_editColumn = "cod"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "confirma_atualizacao.asp"
  MM_fieldsStr  = "nome|value|endereco|value|cidade|value|estado|value|cep|value|email|value|tel1|value|tel2|value|cel|value|data|value"
  MM_columnsStr = "nome|',none,''|endereco|',none,''|cidade|',none,''|estado|',none,''|cep|',none,''|email|',none,''|tel1|',none,''|tel2|',none,''|cel|',none,''|data|',none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
Dim Recordset1__Cod
Recordset1__Cod = "0"
if (request.querystring("cadastro_cliente")   <> "") then Recordset1__Cod = request.querystring("cadastro_cliente")  
%>
<%
set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cadastro_STRING
Recordset1.Source = "SELECT *  FROM cadastro_cliente  WHERE cod = " + Replace(Recordset1__Cod, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Nome:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="nome" value="<%=(Recordset1.Fields.Item("nome").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="endereco" value="<%=(Recordset1.Fields.Item("endereco").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Cidade:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="cidade" value="<%=(Recordset1.Fields.Item("cidade").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Estado:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="estado" value="<%=(Recordset1.Fields.Item("estado").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Cep:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="cep" value="<%=(Recordset1.Fields.Item("cep").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">E-mail:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="email" value="<%=(Recordset1.Fields.Item("email").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tel1:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="tel1" value="<%=(Recordset1.Fields.Item("tel1").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tel2:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="tel2" value="<%=(Recordset1.Fields.Item("tel2").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Cel.:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="cel" value="<%=(Recordset1.Fields.Item("cel").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Data:</font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="data" value="<%=(Recordset1.Fields.Item("data").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline" bgcolor="#6699FF"> 
      <td nowrap align="right" colspan="2" bgcolor="#6699FF"> 
        <div align="center"> 
          <input type="submit" value="( Atualizar )">
        </div>
      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="true">
  <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("cod").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
%>

