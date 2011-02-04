<!--
Developers:   SSeissiger, KWahoff
Date:         07/19/2000
Description:  Displays listing of employee vacation information
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
<title>Update employee exam info</title>
</head>
<body bgcolor=silver onLoad=window.close()>


<!-- <font size=1 face="ms sans serif, arial, geneva">Saving information...</font> -->


<%
ls_empid = Request.QueryString("empid")

ls_examid = Clean(trim(Request("txtexam_id")))
ls_epid = Clean(trim(Request("txtep_id")))
ls_examnumber = Clean(trim(Request("txtexam_number")))
ls_examname = Clean(trim(Request("txtexam_name")))
ls_examdate = Clean(trim(Request("txtexam_date")))
ls_comments = Clean(trim(Request("txtComments")))
	
	
set rs = DBQuery("UPDATE exams_passed SET  exam_id = '" & ls_examid & "', exam_date = '" & ls_examdate & "', comments = '" & ls_comments & "' WHERE exams_passed_id = " & ls_epid)
%>

</body>
</html>