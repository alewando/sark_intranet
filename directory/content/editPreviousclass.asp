<!--
Developer:   DMARTIN
Date:         09/12/2000
Description:  Displays listing of employee class information
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
<head>
<title>Update employee class info</title>
<body bgcolor=silver onLoad=closeWindow()>
</head>



<!-- <font size=1 face="ms sans serif, arial, geneva">Saving information...</font> -->


<%
ls_empid = Request.QueryString("empid")

ls_classid = Clean(trim(Request("txtclass_id")))
ls_ctid = Clean(trim(Request("txtct_id")))
ls_classname = Clean(trim(Request("txtclass_name")))
ls_classStartdate = Clean(trim(Request("txtclass_Start_date")))
ls_classEnddate = Clean(trim(Request("txtclass_End_date")))
ls_comments = Clean(trim(Request("txtComments")))
ls_locationid = Clean(trim(Request("txtlocation_ID")))	
	
set rs = DBQuery("UPDATE Classes SET  class_id = '" & ls_classid & "', Class_Start_Date = '" & ls_classStartdate & "', Class_End_Date = '" & ls_classEnddate & "', location_ID = '" & ls_locationid & "', comments = '" & ls_comments & "' WHERE class_taken_id = " & ls_ctid)
%>

</body>
</html>