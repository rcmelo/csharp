<%

'INÍCIO DO CÓDIGO
'----------------------------------------------------------------------------------------------------------------
'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP

Function EnviaEmail(Host,Componente,Email,NomeEmail,ParaEmail,Assunto,Mensagem)
Select Case Componente
Case "AspMail"
	 	on error resume next
	    Set eObjMail = Server.CreateObject("SMTPsvg.Mailer")
	    eObjMail.FromName = NomeEmail
	    eObjMail.FromAddress = Email
	    eObjMail.RemoteHost = Host
	    eObjMail.AddRecipient "", ParaEmail
	    eObjMail.Subject = Assunto
	    eObjMail.ContentType = "text/html"
	    eObjMail.BodyText = Mensagem	    
	    eObjMail.SendMail
	 	Set eObjMail = nothing

Case "AspEmail"
	    on error resume next
	    Set eObjMail = Server.CreateObject("Persits.MailSender")
	    eObjMail.Host = Host
	 	eObjMail.From = Email
	 	eObjMail.FromName = NomeEmail
	 	eObjMail.AddReplyTo Email
	 	eObjMail.AddAddress ParaEmail
	    eObjMail.Subject = Assunto
	    eObjMail.isHTML = true
	 	eObjMail.Body = Mensagem	 	
	 	eObjMail.Send
	 	Set eObjMail = nothing

Case "AspQmail"
	    on error resume next
		Set eObjMail = Server.CreateObject("SMTPsvg.Mailer")
		eObjMail.QMessage = 1
		eObjMail.FromName = NomeEmail
		eObjMail.FromAddress = Email
		eObjMail.RemoteHost = Host
		eObjMail.AddRecipient "", ParaEmail
		eObjMail.Subject = Assunto
		eObjMail.BodyText = Mensagem
		objNewMail.SendMail
		Set eObjMail = nothing
		
Case "CDONTS"
	    on error resume next
	 	Set eObjMail = Server.CreateObject("CDONTS.NewMail")
	 	eObjMail.to = ParaEmail
		eObjMail.from = NomeEmail & "<" & Email & ">"
		eObjMail.subject = Assunto
		eObjMail.Importance = 1
		eObjMail.BodyFormat = 0
		eObjMail.MailFormat = 0
		eObjMail.body = Mensagem		
		eObjMail.send
		Set eObjMail = nothing
		
End Select
End Function
%>