<%
Public Function BAuthenticateUser
	
'Option Explicit
Response.Expires = 0

'	On Error Resume Next
	BAuthenticateUser = False

	strAT = Request.ServerVariables("AUTH_TYPE")

    If InStr(1, "_BasicNTLM", strAT, vbTextCompare) < 2 Then
	    Response.Buffer = True
        Response.Status = ("401 Unauthorized")
    	Response.AddHeader "WWW.Authenticate", "Basic"
        Response.End
	Else
        BAuthenticateUser = True
	End If
End Function


%>
