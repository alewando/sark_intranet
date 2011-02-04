<!--
Developer:    GINGRICE
Date:         08/18/1998
Description:  Displays detailed listing of consultants
-->


<!-- #include file="../../script.asp" -->

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
// -->
</script>


<html>
<head>
<title>Update employee info</title>
</head>
<body bgcolor=silver onLoad="closeWindow()">
<font size=1 face="ms sans serif, arial, geneva">Saving information...</font>

<%
ls_empid = Request("EmployeeID")
If Trim(Request("InsertUpdate"))= "U" Then 
	'-------------------------------------------
	'	Update Profile Info						
	'-------------------------------------------
	ls_college =		Clean(trim(Request("college")))
	ls_hobbies =		Clean(trim(Request("hobbies")))
	ls_vacation = 		Clean(trim(Request("vacation")))
	ls_favoriteshow = 	Clean(trim(Request("favoriteshow")))
	ls_inspired =		Clean(trim(Request("inspired")))
	ls_motto =			Clean(trim(Request("motto")))
	ls_moment =			Clean(trim(Request("moment")))
	ls_lotto =			Clean(trim(Request("lotto")))
	ls_petpeeve =		Clean(trim(Request("petpeeve")))
	ls_possession = 	Clean(trim(Request("possession")))
	ls_challenge = 		Clean(trim(Request("challenge")))
	set rs = DBQuery("Update Profile Set College = '" & ls_college & "', Hobbies_Sports_Interests = '" & ls_hobbies & "', Dream_Vacation = '" & ls_vacation & "', Favorite_Show_Ad = '" & ls_favoriteshow & "', Person_Inspired = '" & ls_inspired & "', Favorite_Motto = '" & ls_motto & "', Funniest_Embarrassing = '" & ls_moment & "', Lotto = '" & ls_lotto & "', Pet_Peeve = '" & ls_petpeeve & "', Prized_Possession = '" & ls_possession & "', Biggest_Challenge = '" & ls_challenge & "' where Employee_ID = " & ls_empid)
	
ElseIf Request("InsertUpdate") = "I" Then 
	'-------------------------------------------
	'	Insert Profile Information				
	'-------------------------------------------
	ls_college =		Clean(trim(Request("college")))
	ls_hobbies =		Clean(trim(Request("hobbies")))
	ls_vacation =		Clean(trim(Request("vacation")))
	ls_favoriteshow = 	Clean(trim(Request("favoriteshow")))
	ls_inspired =		Clean(trim(Request("inspired")))
	ls_motto =			Clean(trim(Request("motto")))
	ls_moment =			Clean(trim(Request("moment")))
	ls_lotto =			Clean(trim(Request("lotto")))
	ls_petpeeve =		Clean(trim(Request("petpeeve")))
	ls_possession = 	Clean(trim(Request("possession")))
	ls_challenge = 		Clean(trim(Request("challenge")))
	set rs = DBQuery("Insert into Profile (Employee_ID, College, Hobbies_Sports_Interests, Dream_Vacation, Favorite_Show_Ad, Person_Inspired, Favorite_Motto, Funniest_Embarrassing, Lotto, Pet_Peeve, Prized_Possession, Biggest_Challenge) " & " Values(" & ls_empid & ",'" & ls_college & "','" & ls_hobbies & "','" & ls_vacation & "','" & ls_favoriteshow & "','" & ls_inspired & "','" & ls_motto & "','" & ls_moment & "','" & ls_lotto & "','" & ls_petpeeve & "','" & ls_possession & "','" & ls_challenge & "')")
	
ElseIf Request("InsertUpdate") = "T" Then
	'---------------------------------------------------
	'	First delete all rows from employee_tech_xref	
	'   and then insert new rows						
	'---------------------------------------------------
	strSQL = "Delete from employee_tech_xref where employee_id = " & ls_empid & " AND" _
		& " employee_tech_xref.tech_id NOT IN (SELECT Tech.Tech_id FROM Tech" _
		& " WHERE Tech.Approved=0)"
	set rs = DBQuery(strSQL)
	For Each tech_id In Request("tech_expert")
		set rs = DBQuery("Insert into employee_tech_xref (employee_id, tech_id) values (" & ls_empid & "," & tech_id & ")")
		Next
		
ElseIf Request("InsertUpdate") = "IND" Then
	'---------------------------------------------------
	'	First delete all rows from employee_industry_xref	
	'   and then insert new rows (03/14/2000)						
	'---------------------------------------------------
	set rs = DBQuery("Delete from employee_industry_xref where employee_id = " & ls_empid)
	For Each industry_id In Request("ind_expert")
		set rs = DBQuery("Insert into employee_industry_xref (employee_id, industry_id) values (" & ls_empid & "," & industry_id & ")")
		Next

ElseIf Request("InsertUpdate") = "CERT" Then
	'---------------------------------------------------
	'	First delete all rows from employee_cert_xref	
	'   and then insert new rows (03/15/2000)						
	'---------------------------------------------------
	set rs = DBQuery("Delete from employee_cert_xref where employee_id = " & ls_empid)
	For Each cert_id In Request("certs")
		set rs = DBQuery("Insert into employee_cert_xref (employee_id, cert_id) values (" & ls_empid & "," & cert_id & ")")
		Next

'Added to allow updating of employee information		
ElseIf trim(Request("InsertUpdate")) = "E" Then
	ls_fname = Clean(trim(Request("FirstName")))
	ls_lname = Clean(trim(Request("LastName")))
	ls_address = Clean(trim(Request("Address")))
	ls_city = Clean(trim(Request("City")))
	ls_state = Clean(trim(Request("State")))
	ls_zip = Clean(trim(Request("Zip")))
	ls_hphone = Clean(trim(Request("HomePhone")))
	ls_spouse = Clean(trim(Request("SpouseName")))
	ls_pager = Clean(trim(Request("PagerNumber")))
	ls_pager_email = Clean(trim(Request("PagerEmail")))
	ls_birthday = Clean(trim(Request("Birthday")))
	ls_personal_website = Clean(trim(Request("PersonalWebsite")))
	ls_clientPhone = Clean(trim(Request("Client_Phone")))
	ls_cell = Clean(trim(Request("CellPhone")))
	ls_vmail = Clean(trim(Request("VoiceMail")))
	ls_ClientEmail = Clean(trim(Request("ClientEmail")))
	ls_ClientID = Clean(trim(Request("Client_ID")))
	if not isempty(ls_vmail) then
		set rs = DBQuery("Update Employee Set FirstName = '" & ls_fname & "', LastName = '" & ls_lname & "', Address = '" & ls_address & "', City = '" & ls_city & "', State = '" & ls_state & "', Zip = '" & ls_zip & "', HomePhone = '" & ls_hphone & "', SpouseName = '" & ls_spouse & "', PagerNumber = '" & ls_pager & "', PagerEmail = '" & ls_pager_email & "', CellPhone = '" & ls_Cell & "', Client_ID = " & ls_ClientID & ", ClientEmail = '" & ls_clientemail & "', Client_Phone = '" & ls_clientphone & "', Birthday = '" & ls_birthday & "', PersonalWebsite = '" & ls_personal_website & "', VoiceMail = '" & ls_vmail & "' where EmployeeID = " & ls_empid)
	else
		set rs = DBQuery("Update Employee Set FirstName = '" & ls_fname & "', LastName = '" & ls_lname & "', Address = '" & ls_address & "', City = '" & ls_city & "', State = '" & ls_state & "', Zip = '" & ls_zip & "', HomePhone = '" & ls_hphone & "', SpouseName = '" & ls_spouse & "', PagerNumber = '" & ls_pager & "', PagerEmail = '" & ls_pager_email & "', CellPhone = '" & ls_Cell & "', Client_ID = " & ls_ClientID & ", ClientEmail = '" & ls_clientemail & "', Client_Phone = '" & ls_clientphone & "', Birthday = '" & ls_birthday & "', PersonalWebsite = '" & ls_personal_website & "', VoiceMail = NULL where EmployeeID = " & ls_empid)
	end if
	'set rs = DBQuery("Update Employee Set FirstName = '" & ls_fname & "', LastName = '" & ls_lname & "', Address = '" & ls_address & "', City = '" & ls_city & "', State = '" & ls_state & "', Zip = '" & ls_zip & "', HomePhone = '" & ls_hphone & "', SpouseName = '" & ls_spouse & "', PagerNumber = '" & ls_pager & "', CellPhone = '" & ls_Cell & "', ClientEmail = '" & ls_clientemail & "', Client_Phone = '" & ls_clientphone & "', Birthday = '" & ls_birthday & "' where EmployeeID = " & ls_empid)
else
	'Response.redirect "www.microsoft.com"
End If

%>

<br>Done
</body>
</html>

