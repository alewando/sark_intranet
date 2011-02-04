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

<head><title>Delete employee exam "deleteexam.asp"</title></head>


<body bgcolor=silver onLoad=closeWindow()>


<%
	epid = Request.QueryString("epid")
	Response.Write("<!-- epid: " & epid & "-->")

	set rs = DBQuery("DELETE from Exams_passed WHERE exams_passed_id = " & epid)
	'rs.close
       
%>

</body>
</html>