<%
set objMessage = createobject("cdo.message")
set objConfig = createobject("cdo.configuration")
Set Flds = objConfig.Fields
 
Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="localhost"
 
' ' Passing SMTP authentication
Flds.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
Flds.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="info@vecinosdelossauces.com.ar"
Flds.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="bG6xj33?"
 
Flds.update
Set objMessage.Configuration = objConfig
objMessage.To = "facu.232@gmail.com"
objMessage.From = "info@vecinosdelossauces.com.ar"
objMessage.Subject = "ASUNTO"
objMessage.fields.update
objMessage.HTMLBody = "Mensaje"
objMessage.Send
%>