<% Response.Buffer = TRUE %>

<!--
<% 
if (Len(Request.ServerVariables("LOGON_USER")) = 0 ) then
	Response.Status = "401 Unauthorized" 
%>
	<HTML><BODY><B>Error: Access is denied.</B><P></BODY></HTML>
<% else %>
-->
<html>

<head>
<title>RFC 1867 Upload Form</title>
</head>

<body>
<center>
	<hr>
	<h2>
		Upload your files here.
	</h2>
	<hr><br>
</center>
<form enctype="multipart/form-data" action="http://<%= Request.ServerVariables("SERVER_NAME") %>/scripts/cpshost.dll?PUBLISH?" method=post>
File to process: <input name="my_file" type="file"><br>
<!--
File to process: <input name="my_file" type="file"><br>
File to process: <input name="my_file" type="file"><br>
-->
Destination URL: <input name="TargetURL" value="http://<%= Request.ServerVariables("SERVER_NAME") %>/Uploads/"><br>
<input type="submit" value="Upload">
</form>
</body>

</html>
<!--
<% end if %>
-->
