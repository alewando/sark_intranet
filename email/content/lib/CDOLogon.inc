<%

Sub LogonEx(bstrExchServer, bstrMailbox, bstrPassword)

	' This will create the CDO session object
	Set objOMSession = Server.CreateObject("MAPI.Session")

	bstrProfileInfo = bstrExchServer + vbLF + bstrMailbox

	objOMSession.Logon "", bstrPassword, False, True, 0, False, bstrProfileInfo

End Sub

%>

