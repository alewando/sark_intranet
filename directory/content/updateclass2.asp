<!--
Developers:   David Martin
Date:         08/30/2000
Description:  called by addclass.asp : writes class info to DB
-->


<!-- #include file="../../script.asp" -->

<html>
<head>
<title>Add employee class info "updateclass2.asp"</title>
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
<body bgcolor=silver onLoad=window.close()>

<!-- 
<font size=1 face="ms sans serif, arial, geneva">Saving information...</font>
-->

<%
ls_empid = Request("EmployeeID")

	ls_classid = Request("class_id")
	ls_classname = Request("class_name")
	ls_classStartdate = Request("class_start_date")
	ls_classEnddate = Request("class_end_date")
	ls_locationid = Request("Location_ID")
	ls_comments = Request("Comments")
	
	DBQuery("INSERT into Classes (class_id, employeeid, class_start_date, Location_ID, comments, class_end_date) Values('" & ls_classid & "', '" & ls_empid & "', '" & ls_classStartdate & "', '" & ls_Locationid & "', '" & ls_comments & "', '" & ls_classEnddate & "')")
    
    set reqFrmSelected = Request("Selected")
    for each MyName in reqFrmSelected
		'Response.Write MyName & "<BR>"
		ls_empid = MyName
		newsql = "INSERT into Classes (class_id, employeeid, class_start_date, Location_ID, comments, class_end_date) Values('" & ls_classid & "', '" & ls_empid & "', '" & ls_classStartdate & "', '" & ls_Locationid & "', '" & ls_comments & "', '" & ls_classEnddate & "')"
		dbquery(newsql)
	next
%>

</body>
</html>
