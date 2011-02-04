<!--
Developers:   David Martin
Date:         08/30/2000
Description:  Deletes an exam entry
-->

<!-- #include file="../../script.asp" -->
<script language=javascript>

var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWindow(){
	//if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close();
	}


</script>

<html>

<head><title>Delete employee class "deleteclass.asp"</title></head>


<body bgcolor=silver onLoad=closeWindow()>


<%
	ctid = Request.QueryString("ctid")
	Response.Write("<!-- ctid: " & ctid & "-->")

	set rs = DBQuery("DELETE from Classes WHERE Class_taken_id = " & ctid)
	'rs.close
       
%>

</body>
</html>