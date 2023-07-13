<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/cadastro.asp" -->


<%
response.buffer = "true"
' *** Edit Operations: declare variables

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
Recordset1.Source = "SELECT *  FROM cadastro_cliente  ORDER BY nome"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 10
Dim Repeat1__index
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
Recordset1_total = Recordset1.RecordCount

' set the number of rows displayed on this page
If (Recordset1_numRows < 0) Then
  Recordset1_numRows = Recordset1_total
Elseif (Recordset1_numRows = 0) Then
  Recordset1_numRows = 1
End If

' set the first and last displayed record
Recordset1_first = 1
Recordset1_last  = Recordset1_first + Recordset1_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset1_total <> -1) Then
  If (Recordset1_first > Recordset1_total) Then Recordset1_first = Recordset1_total
  If (Recordset1_last > Recordset1_total) Then Recordset1_last = Recordset1_total
  If (Recordset1_numRows > Recordset1_total) Then Recordset1_numRows = Recordset1_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (Recordset1_total = -1) Then

  ' count the total records by iterating through the recordset
  Recordset1_total=0
  While (Not Recordset1.EOF)
    Recordset1_total = Recordset1_total + 1
    Recordset1.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (Recordset1.CursorType > 0) Then
    Recordset1.MoveFirst
  Else
    Recordset1.Requery
  End If

  ' set the number of rows displayed on this page
  If (Recordset1_numRows < 0 Or Recordset1_numRows > Recordset1_total) Then
    Recordset1_numRows = Recordset1_total
  End If

  ' set the first and last displayed record
  Recordset1_first = 1
  Recordset1_last = Recordset1_first + Recordset1_numRows - 1
  If (Recordset1_first > Recordset1_total) Then Recordset1_first = Recordset1_total
  If (Recordset1_last > Recordset1_total) Then Recordset1_last = Recordset1_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = Recordset1
MM_rsCount   = Recordset1_total
MM_size      = Recordset1_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Recordset1_first = MM_offset + 1
Recordset1_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (Recordset1_first > MM_rsCount) Then Recordset1_first = MM_rsCount
  If (Recordset1_last > MM_rsCount) Then Recordset1_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<html>
<head>
<title>Desenvovido por: www.comercialweb.com.br</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#990000" vlink="#990000" alink="#990000">
<b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> </font></b> 
<table width="100%" border="0" bgcolor="#6699FF">
  <tr>
    <td width="58%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Listando:(<font color="#990000"> 
      <font color="#FFFFFF"><%=(Recordset1_first)%></font> </font>) De: (<font color="#990000"> <font color="#FFFFFF"><%=(Recordset1_last)%></font> </font>)</font></b></td>
    <td width="1%"><b></b></td>
    <td width="41%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Total 
      de registros: (<font color="#990000"> <font color="#FFFFFF"><%=(Recordset1_total)%></font></font> </font></b><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
      )</font></b> </td>
  </tr>
</table>
<% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
<table width="100%" border="1" cellspacing="1" cellpadding="0" bordercolor="#6699FF">
  <tr> 
    <td colspan="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Nome:<font color="#0000CC"> 
      <%=Recordset1("nome")%></font></font></b></td>
    <td colspan="4"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Cidade:<font color="#0000CC"> 
      <%=Recordset1("cidade")%></font></font></b></td>
    <td width="15%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Estado:<font color="#0000CC"> 
      <%=Recordset1("estado")%></font></font></b></td>
    <td width="11%" bgcolor="#FF9900"> 
      <div align="center"><b><font size="1"><b><font face="Verdana" size="2" color="#000000"><a href="atualizar_form.asp?cadastro_cliente=<%=Recordset1("Cod")%>"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Atualizar</font></b></a></font></b></font></b></div>
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
          <input type="submit" name="Submit" value="Excluir" >
          <br>
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
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
<table border="0" width="50%" align="center">
  <tr bgcolor="#6699FF"> 
    <td width="23%" align="center"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"> 
        <% If MM_offset <> 0 Then %>
        <a href="<%=MM_moveFirst%>">Primeira</a> 
        <% End If ' end MM_offset <> 0 %>
        </font></b></div>
    </td>
    <td width="31%" align="center"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"> 
        <% If MM_offset <> 0 Then %>
        <a href="<%=MM_movePrev%>">Anterior</a> 
        <% End If ' end MM_offset <> 0 %>
        </font></b></div>
    </td>
    <td width="23%" align="center"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"> 
        <% If Not MM_atTotal Then %>
        <a href="<%=MM_moveNext%>">Pr&oacute;xima</a> 
        <% End If ' end Not MM_atTotal %>
        </font></b></div>
    </td>
    <td width="23%" align="center"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"> 
        <% If Not MM_atTotal Then %>
        <a href="<%=MM_moveLast%>">&Uacute;ltima</a> 
        <% End If ' end Not MM_atTotal %>
        </font></b></div>
    </td>
  </tr>
</table>
<div align="center"><br>
</div>
<hr align="center">
<div align="center"><font color="#FFFF66" size="1" face="Arial, Helvetica, sans-serif"><b><font size="2"><font color="#000000" size="1">Desenvolvido 
  por: <a href="http://www.comercialweb.com.br" target="_blank" class=menu3><font color="#000099">COMERCIALWEB</font></a></font></font></b></font> 
  <font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><b><br>
  E-mail:<font color="#000066"> <a href="mailto:comercial@comercialweb.com.br">comercial@comercialweb.com.br 
  </a></font></b></font> </div>
</body>
</html>
<%
Recordset1.Close()
%>
