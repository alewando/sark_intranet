<%
	dim debug
	If Application("debug") Then
		debug = true
	Else
		debug = false
	End If
	Application("debug") = False
%>
<!-- #include file="../../script.asp" -->
<%
	sql = "UPDATE Tech_Links SET Num_Hits = (SELECT Num_Hits+1 FROM Tech_Links WHERE Tech_Links_ID = " & Request("LinkID") & ") " & _
			"WHERE Tech_Links_ID = " & Request("LinkID")
	'response.write sql
	DBQuery(sql)
	Application("debug") = debug
	Response.Redirect(Request("url"))
%>
