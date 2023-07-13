<%

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP


function trimtodigits(tstring)
'Remove letras da string
s="" 
ts=tstring
for x=1 to len(ts)
ch=mid(ts,x,1)
if asc(ch)>=48 and asc(ch)<=57 then
s=s & ch
end if
next
trimtodigits=s
end function

'Função q checa o numero do cartão
function checkcc(ccnumber,cctype)
ctype=ucase(cctype)
select case ctype
case "V"
cclength="13;16"
ccprefix="4"
case "M"
cclength="16"
ccprefix="51;52;53;54;55"
case "A"
cclength="15"
ccprefix="34;37"
case "D"
cclength="14"
ccprefix="300;301;302;303;304;305;36;38"
case else
cclength=""
ccprefix=""
end select
prefixes=split(ccprefix,";",-1)
lengths=split(cclength,";",-1)
number=trimtodigits(ccnumber)
prefixvalid=false
lengthvalid=false
for each prefix in prefixes
if instr(number,prefix)=1 then
prefixvalid=true
end if
next 
for each length in lengths
if cstr(len(number))=length then
lengthvalid=true
end if
next
result=0
if not prefixvalid then
result=result+1
end if
if not lengthvalid then
result=result+2
end if  
qsum=0
for x=1 to len(number)
ch=mid(number,len(number)-x+1,1)
'response.write ch
if x mod 2=0 then
sum=2*cint(ch)
qsum=qsum+(sum mod 10)
if sum>9 then 
qsum=qsum+1
end if
else
qsum=qsum+cint(ch)
end if
next
if qsum mod 10<>0 then
result=result+4
end if
if cclength="" then
result=result+8
end if
checkcc=result
end function%>
