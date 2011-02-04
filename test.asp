

<%	'Get Weather Information
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	'path = "\\trickle\intranet\welcome\content\weather.txt"
	path = "c:\temp\weather.txt"
	set ts = fso.OpenTextFile(path)
	session("sW_HTML") = ts.ReadAll
	ts.close
	
	'Get CNN News	path = "c:\temp\news.txt"
	set ts = fso.OpenTextFile(path)
	session("sCNet_HTML") = ts.ReadAll
	ts.close
	
	
	Response.Write (session("sW_HTML"))
	Response.Write("<BR><BR>")
	Response.Write (session("sCNet_HTML"))
%>