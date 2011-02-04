<!--
Developers:   SSeissiger
Date:         8/24/00
Description:  Deletes a solution services member - Table: Tech_Specialists
				Performs the sql statement to delete
			    Info from form in update_ss
-->

<!-- #include file="../../script.asp" -->
<script language=javascript>

var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWindow(){
	window.close();
}

</script>

<html>

<head><title>Delete Solution Services Team Member</title></head>

<body bgcolor=silver onLoad='closeWindow()'>

<%
	'--------------------------------------------------------------------------------------
	' The above onload event calls closeWindow(), this will automatically close the current
	' pop up window. The main window is waiting for this to occur and will automatically 
	' refresh, when this terminates.
	'---------------------------------------------------------------------------------------
	If Request("DelFlag") = "True" And Request("Tech_Spec_ID")& "x" <> "x" Then
		' Ensures this page wasn't called mistakenly
		Tech_Spec_ID = Request("Tech_Spec_ID")
		Response.Write "Deleting Tech_Specialist_id=" & Tech_Spec_ID & "<br>"

		set rs = DBQuery("DELETE from Tech_Specialists WHERE Tech_Specialists_ID = " & Tech_Spec_ID)
	End If
%>

</body>
</html>