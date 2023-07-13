<% 


'IN�CIO DO C�DIGO
'----------------------------------------------------------------------------------------------------------------
'Este c�digo est� otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP


Function Cripto(StrCripto, BolAcao) 'In�cio de da fun��o de criptografia. Aonde o par�metro String � o valor que ser� criptografado ou descriptografado. E o par�metro BolAcao � um valor booleano (True ou False) para indicar se deve ser criptografado (True) ou descriptografado (False).

Dim X , i, n, TamChave, Valor 'Declara��o das vari�veis.
Dim aChaves(9) 'Declara��o do Array que contem as chaves de criptografia

aChaves(0) = 77 'Atribui��o dos valores dos caracteres que ser�o utilizados
aChaves(1) = 84 'como chave de criptografia
aChaves(2) = 79
aChaves(3) = 65
aChaves(4) = 73
aChaves(5) = 78
aChaves(6) = 67
aChaves(7) = 70
aChaves(8) = 82

X = Empty 'Atribui o valor vazio para a vari�vel X

For i = 1 To Len(StrCripto) 'Inicia um Loop For que deve ser repetido a quantidade de vezes de acordo com o tamanho da String que deve ser criptografada.
    n = Asc(Mid(StrCripto, i, 1))
	
    If n > 31 Then 'mant�m controles intactos
       n = n - 32 'desconsidera controles (Caracteres de 0 a 31)
       If BolAcao Then
          Valor = 1
       Else
          Valor = -1
       End If
       n = n + (aChaves((i - 1) Mod 9)) * Valor
       n = n Mod 224 ' Caracteres 256 - 32 desconsiderados
          If n < 0 Then
             n = 224 + n
          End If
       n = n + 32 ' Reajusta para caracteres normais
    End If
	
    X = X & Chr(n) 'Armazena na vari�vel X os caracteres que foram modificados
Next

       If BolAcao Then
          Cripto = EHtmlEncode(X) 'Retorna o valor da vari�vel X para a fun��o
       Else
         Cripto = EHtmlDecode(X) 'Retorna o valor da vari�vel X para a fun��o
       End If


End Function

Function EHtmlEncode(StrString)
StrString = Replace(StrString, ",", ",")
StrString = Replace(StrString, "'", "'")
StrString = Replace(StrString, """", """")
StrString = Replace(StrString, "=", "=")
StrString = Replace(StrString, ".", ".")
StrString = Replace(StrString, "(", "(")
StrString = Replace(StrString, "set", "set")
StrString = Replace(StrString, "SET", "SET")
StrString = Replace(StrString, " where ", "where")
StrString = Replace(StrString, " WHERE ", "WHERE")
StrString = Replace(StrString, ")", "(")
StrString = Replace(StrString, Chr(32), Chr(32))
StrString = Replace(StrString, Chr(13), Chr(13))
StrString = Replace(StrString, Chr(0), Chr(0))
StrString = Replace(StrString, Chr(10) & Chr(10), Chr(10) & Chr(10))
StrString = Replace(StrString, Chr(10), Chr(10))
StrString = Replace(StrString, Chr(9), Chr(9))
StrString = Replace(StrString, Chr(11), Chr(11))
StrString = Replace(StrString, Chr(12), Chr(12))
StrString = Replace(StrString, Chr(60), Chr(60))
StrString = Replace(StrString, "/", "/")
StrString = Replace(StrString, "\", "\")
EHtmlEncode = StrString
End Function

Function EHtmlDecode(StrString)
StrString = Replace(StrString, "_a", "_a")
StrString = Replace(StrString, "_b", "_b")
StrString = Replace(StrString, "_c", "_c")
StrString = Replace(StrString, "_d", "_d")
StrString = Replace(StrString, "_e", "_e")
StrString = Replace(StrString, "_f", "_f")
StrString = Replace(StrString, "_g", "_g")
StrString = Replace(StrString, "_h", "_h")
StrString = Replace(StrString, "_i", "_i")
StrString = Replace(StrString, "_j", "_j")
StrString = Replace(StrString, "_l", "_l")
StrString = Replace(StrString, "_m", "_m")
StrString = Replace(StrString, "_n", "_n")
StrString = Replace(StrString, "_o", "_o")
StrString = Replace(StrString, "_p", "_p")
StrString = Replace(StrString, "_q", "_q")
StrString = Replace(StrString, "_r", "_r")
StrString = Replace(StrString, "_s", "_s")
StrString = Replace(StrString, "_t", "_t")
StrString = Replace(StrString, "_u", "_u")
StrString = Replace(StrString, "_v", "_v")
StrString = Replace(StrString, "_x", "_x")
EHtmlDecode = StrString
End Function


%>