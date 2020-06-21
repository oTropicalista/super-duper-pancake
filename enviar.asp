<%
If Int(Request.Form.Count) > 0 Then


	sch = "http://schemas.microsoft.com/cdo/configuration/"
	Set cdoConfig = Server.CreateObject("CDO.Configuration")
		cdoConfig.Fields.Item(sch & "smtpserver") = "smtp.goesocial.com.br"
		cdoConfig.Fields.Item(sch & "sendusername") = "comercial@goesocial.com.br"
		cdoConfig.Fields.Item(sch & "sendpassword") = "abc#123$k"
		cdoConfig.Fields.Item(sch & "smtpauthenticate") = 1
		cdoConfig.Fields.Item(sch & "smtpserverport") = 587
		cdoConfig.Fields.Item(sch & "sendusing") = 2
		cdoConfig.Fields.Item(sch & "smtpconnectiontimeout") = 30
		cdoConfig.Fields.update
	Set cdoMessage = Server.CreateObject("CDO.Message")
	Set cdoMessage.Configuration = cdoConfig
	cdoMessage.From = "Go eSocial <comercial@goesocial.com.br>"
	cdoMessage.To = "suporte@technoplus.com.br"
	cdoMessage.Subject = "Formulário"
	cdoMessage.HTMLBody = "<p style='font:15px Calibri'>" & "NOME = " & Request.Form("form_nome") & "<br /><br />SOBRENOME = " & Request.Form("form_sobrenome") & "<br /><br />EMAIL = " & Request.Form("form_email") & "<br /><br />TELEFONE = " & Request.Form("form_celular") & "<br /><br />MENSAGEM = " & Request.Form("form_mensagem") & "</p>"
	On Error Resume Next
	cdoMessage.Send
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
	
	
	Response.Redirect "index.html?Enviado=Sim#contato"


Else
Response.Redirect "index.html"
End If
%>

















