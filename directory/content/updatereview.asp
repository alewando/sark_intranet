<!--
Developers:   Jason Moore and Mike Thierauf
Date:         07/17/2000
Description:  Displays listing of employee review information
-->


<!-- #include file="../../script.asp" -->

<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWindow()
{
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2)))
	{ 
		//self.opener.location.reload( );
		window.close();
	}
}
// -->
</script>

<html>
<head>
<title>Update employee review info</title>
</head>
<body bgcolor=silver onLoad='window.close();'>
<font size=1 face="ms sans serif, arial, geneva">Saving information...</font>

<%
ls_empid = Request("EmployeeID")

'-------------------------------------------
'	Update Review Info						
'-------------------------------------------
ls_last_promotion =	Clean(trim(Request("last_promotion")))
ls_last_review = 	Clean(trim(Request("last_review")))
ls_next_review = 	Clean(trim(Request("next_review")))
ls_sponsor_id =		CInt(Clean(trim(Request("sponsor_id"))))
ls_acct_mgr_id = 	CInt(Clean(trim(Request("account_mgr_id"))))
ls_comments =		Clean(trim(Request("comments")))
ls_insup =			Clean(trim(Request("InsertUpdate")))

if ls_sponsor_id = 0 then
	ls_sponsor_id = null
end if
if ls_acct_mgr_id = 0 then
	ls_acct_mgr_id = null
end if

if not isnull(ls_sponsor_id) and not isnull(ls_acct_mgr_id)and ls_insup = "U" then
	set rs = DBQuery("Update Review Set Last_promotion = '" & ls_last_promotion & _
	         "', last_review = '" & ls_last_review & "', next_review = '" & ls_next_review & _
	         "', sponsor_ID = " & ls_sponsor_id & ", account_mgr_id = " & ls_acct_mgr_id & _
	         ", comments = '" & ls_comments & "' where EmployeeID = " & ls_empid)
	elseif isnull(ls_sponsor_id) and isnull(ls_acct_mgr_id)and ls_insup = "U"  then
	set rs = DBQuery("Update Review Set Last_promotion = '" & ls_last_promotion & _
	         "', last_review = '" & ls_last_review & "', next_review = '" & ls_next_review & _
	         "', sponsor_ID = null " & ", account_mgr_id = null " & _
	         ", comments = '" & ls_comments & "' where EmployeeID = " & ls_empid)
	elseif not isnull(ls_sponsor_id) and isnull(ls_acct_mgr_id)and ls_insup = "U"  then
	set rs = DBQuery("Update Review Set Last_promotion = '" & ls_last_promotion & _
	         "', last_review = '" & ls_last_review & "', next_review = '" & ls_next_review & _
	         "', sponsor_ID = " & ls_sponsor_id & ", account_mgr_id = null" & _
	         ", comments = '" & ls_comments & "' where EmployeeID = " & ls_empid)
	elseif isnull(ls_sponsor_id) and not isnull(ls_acct_mgr_id)and ls_insup = "U"  then
	set rs = DBQuery("Update Review Set Last_promotion = '" & ls_last_promotion & _
	         "', last_review = '" & ls_last_review & "', next_review = '" & ls_next_review & _
	         "', sponsor_ID = null" & ", account_mgr_id = " & ls_acct_mgr_id & _
	         ", comments = '" & ls_comments & "' where EmployeeID = " & ls_empid)
end if
	
if not isnull(ls_sponsor_id) and not isnull(ls_acct_mgr_id)and ls_insup = "I" then
	set rsi = DBQuery("Insert into review (last_promotion, last_review, next_review, employeeID, sponsor_id, Account_Mgr_ID, comments) " & _
	         " Values('" & ls_last_promotion & "', '" & ls_last_review & "', '" & ls_next_review & _
	         "', " & ls_empid & ", " & ls_sponsor_id & ", " & ls_acct_mgr_id & ", '" & ls_comments & "')")
	elseif isnull(ls_sponsor_id) and isnull(ls_acct_mgr_id)and ls_insup = "I"  then
	set rsi = DBQuery("Insert into review (last_promotion, last_review, next_review, employeeID, comments) " & _
	         " Values('" & ls_last_promotion & "', '" & ls_last_review & "', '" & ls_next_review & _
	         "', " & ls_empid & ", '" & ls_comments & "')")
	elseif not isnull(ls_sponsor_id) and isnull(ls_acct_mgr_id)and ls_insup = "I"  then
	set rsi = DBQuery("Insert into review (last_promotion, last_review, next_review, employeeID, sponsor_id, comments) " & _
	         " Values('" & ls_last_promotion & "', '" & ls_last_review & "', '" & ls_next_review & _
	         "', " & ls_empid & ", " & ls_sponsor_id & ", '" & ls_comments & "')")
	elseif isnull(ls_sponsor_id) and not isnull(ls_acct_mgr_id)and ls_insup = "I"  then
	set rsi = DBQuery("Insert into review (last_promotion, last_review, next_review, employeeID, Account_Mgr_ID, comments) " & _
	         " Values('" & ls_last_promotion & "', '" & ls_last_review & "', '" & ls_next_review & _
	         "', " & ls_empid & ", " & ls_acct_mgr_id & ", '" & ls_comments & "')")
end if 
%>

<br>Done
</body>
</html>

