<%
'---------------------------------------
'  Create instances of MAPI session		
'  and login, storing session in the	
'  users ASP Session  object			
'---------------------------------------
dim objOMSession
set objRenderApp = Application("RenderApplication")
if Session("MAPIsession") is nothing then
	set objOMSession = Server.CreateObject("MAPI.Session")
	on error resume next
	objOMSession.Logon "", "", False, True, 0, False, Application("ExchangeServer") & vbLF & Session("Username")
	if err = 0 then
		set Session("MAPIsession") = objOMSession
		Session("hImp") = objRenderApp.ImpID
		set Session("/Inbox") = objOMSession.Inbox
		set Session("/") = objOMSession.GetFolder(objOMSession.Inbox.FolderID)
		end if
	end if
	on error goto 0
%>
