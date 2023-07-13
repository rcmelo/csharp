<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/cadastro.asp" -->

<%
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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_cadastro_STRING
  MM_editTable = "cadastro_cliente"
  MM_editRedirectUrl = "confirmacao_cadastro.asp"
  MM_fieldsStr  = "nome|value|endereco|value|cidade|value|estado|value|cep|value|email|value|tel1|value|tel2|value|cel|value"
  MM_columnsStr = "nome|',none,''|endereco|',none,''|cidade|',none,''|estado|',none,''|cep|',none,''|email|',none,''|tel1|',none,''|tel2|',none,''|cel|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
<html>
<head>
<title>inserir</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center" width="290">
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF" colspan="2">
        <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Todos 
          os campos s&atilde;o obrigat&oacute;rios</font></b></div>
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Nome: 
        </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="nome" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Endereco: 
        </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="endereco" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        Cidade: </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="cidade" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Estado: 
        </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="estado" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Cep: 
        </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="cep" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">E-mail: 
        </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="email" value="" size="32" onClick="MM_validateForm('email','','RisEmail');return document.MM_returnValue">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Tel.1: 
        </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="tel1" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Tel.2: 
        </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="tel2" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Cel.: 
        </font></b></td>
      <td bgcolor="#D5E2FF"> 
        <input type="text" name="cel" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right" bgcolor="#6699FF" colspan="2"> 
        <div align="center"> 
          <input type="submit" value="( Inserir )">
        </div>
      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
Recordset1.Close()
%>
