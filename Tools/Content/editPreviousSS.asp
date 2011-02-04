<!--
Developers:   SSeissiger
Date:         08/24/00
Description:  Performs the sql statement to update a selected solution services member info
			  Info from form in update_ss
-->


<!-- #include file="../../script.asp" -->

<script language=javascript>

var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

</script>
	
<html>
<head>
<title>Update Solution Services Team Member Info</title>
</head>
 
<body bgcolor=silver onLoad='window.close();'>

<font size=1 face="ms sans serif, arial, geneva">Saving information...</font>

<script>
//GoHome();
</script>

<%
Tech_Spec_ID = Request.QueryString("tech_Spec_ID")

ls_area = Request("sArea")
ls_level = 	Request("sLevel")

sql = 	"UPDATE Tech_Specialists SET Tech_Area_ID = '" & ls_area & "', Tech_Specialist_Type_ID = '" & ls_level & "' WHERE Tech_Specialists_ID = " & Tech_Spec_ID

set rs = DBQuery(sql)
%>

</body>
</html>

<SCRIPT LANGUAGE=javascript>
<!--
function GoHome()
	{
		window.navigate("default.asp")
	}
//-->
</SCRIPT>
