<!--
Developers:   SSeissiger
Date:         07/19/2000
Description:  Deletes a vacation entry
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

<head><title>Delete employee vacation</title></head>


<body bgcolor=silver onLoad='closeWindow()'>


<%
	vacID = Request.QueryString("vacID")
	Response.Write("<!-- vacid: " & vacid & "-->")

	set rs = DBQuery("DELETE from vacation WHERE vacation_id = " & vacid)
	'rs.close
       
%>

</body>
</html>