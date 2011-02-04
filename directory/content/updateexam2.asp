<!--
Developers:   David Martin
Date:         08/30/2000
Description:  called by addexam.asp : writes exam info to DB
-->


<!-- #include file="../../script.asp" -->

<html>
<head>
<title>Add employee exam info "updateexam2.asp"</title>
</head>

<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWindow(){
	//if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close();
	}
-->
</script>

<!-- 'closeWindow()' -->
<body bgcolor=silver onLoad=closeWindow()>

<!-- 
<font size=1 face="ms sans serif, arial, geneva">Saving information...</font>
-->

<%
ls_empid = Request("EmployeeID")

	ls_examid = Request("exam_id")
	ls_examname = Request("exam_name")
	ls_examdate = Request("exam_date")
	ls_comments = Request("Comments")
	
	set rs = DBQuery("INSERT into Exams_passed (exam_id, employeeid, exam_date, comments) Values('" & ls_examid & "', '" & ls_empid & "', '" & ls_examdate & "', '" & ls_comments & "')")
         
%>

</body>
</html>
-->