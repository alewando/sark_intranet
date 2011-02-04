<!--
Developer:		DMARSTON
Date:			03/28/2000
Description:	Allows an employee to be added to intranet database.
History:		Created - 03/28/2000-->

<!-- #include file="../section.asp" -->

<script language=javascript>
<!--
	
	function VerifyAddition(){
		document.frmInfo.Submitted.value = "True"
		document.frmInfo.submit()
		/*if (confirm("Are you sure you want to add this name?")){
			document.frmInfo.Submitted.value = "True"
			document.frmInfo.submit()
			}
		else{
			alert("Operation Canceled!")
			}*/
	}
	  
	function EmployeeSaved(){
		document.frmInfo.Submitted.value = "False"
		document.frmInfo.submit()
	}

// -->
</script>

<%
	if Request.Form("Submitted") = "True" then
		ls_EmptyFName = "False"
		ls_EmptyLName = "False"
		ls_EmptyTitle = "False"
		ls_EmptyUser = "False"
		ls_EmptyStart = "False"
		ls_dup_uname = "False"
		ls_complete = "False"
		ls_first_name = Request.form("FirstName")
		ls_last_name = Request.form("LastName")
		ls_title = Request.form("Title")
		ls_username = Request.form("Username")
		ls_start_date = Request.Form("StartDate")
		if ls_first_name <> "" then
			if ls_last_name <> "" then
				if ls_title <> "" then
					if ls_username <> "" then
						if ls_start_date <> "" then
							strSQL = "SELECT Employee.Username FROM Employee" _
								& " WHERE Employee.Username ='" & Request.Form("Username") & "'"
							set rs=DBQuery(strSQL)
							if not rs.eof then
								ls_dup_uname = "True"
							else
								strSQL = "INSERT INTO Employee (FirstName, LastName, EmployeeTitle_ID," _
									& " Username, Branch_ID, CanSponsor, UnlistedPersonal, OfficeLocation," _
									& " Client_ID, Client_Phone, StartDate) VALUES" _
									& " ('" & ls_first_name & "','" & ls_last_name & "','" _
									& ls_title & "','" & ls_username & "','2','0','0','Cincinnati'," _
									& "'2','721-3232','" & ls_start_date & "')"
								DBQuery(strSQL)
								ls_complete= "True"
							end if
						else
							ls_EmptyStart = "True"
						end if
					else
						ls_EmptyUser = "True"
					end if
				else
					ls_EmptyTitle = "True"
				end if
			else
				ls_EmptyLName = "True"
			end if
		else
			ls_EmptyFName = "True"
		end if
	end if
%>

<tr><td valign=top>
	<form name=frmInfo action='add_emp.asp' method=post>
	<input type=hidden name='Submitted' value='False'>

<%
if ls_complete = "True" then
    Response.write("<img src='../../common/images/checker.gif' height=32 width=32 border=0 align=left>")
	Response.Write("<table border=0 cellpadding=10>")
	Response.Write("<tr><td align='center'><font size=1 face='ms sans serif, arial, geneva' color=blue>")
	Response.Write("<b>New Employee has been saved.</b></font></td></tr>")
	Response.Write("<tr><td><input type=button class=button value='  OK  ' OnClick='EmployeeSaved();'></td></tr>")
else
	Response.Write("<table border=0 cellpadding=5>")
	Response.Write("<tr><td colspan=2><font size=2 color=red><b>")
	Response.write("ADD EMPLOYEE</b></td>")
	Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=blue><b>")
	Response.write("To add an employee, populate the following fields and click the 'Save' button when finished.</b></td>")
	Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>First Name:</b></td>")
	
	if ls_complete = "False" then
		Response.Write("<td><input type=text size=35 name=FirstName value=" & chr(34) & ls_first_name & chr(34) & "></td></tr>")
	else
		Response.Write("<td><input type=text size=35 name=FirstName value=" & chr(34) & chr(34) & "></td></tr>")
	end if
			
	if ls_EmptyFName = "True" then
		Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=red>")
		Response.Write("<b><i>*Please enter the new employees first name.</i></b></td>")
	end if
	
	Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Last Name:</b></td>")
	
	if ls_complete = "False" then
		Response.Write("<td><input type=text size=35 name=LastName value=" & chr(34) & ls_last_name & chr(34) & "></td></tr>")
	else
		Response.Write("<td><input type=text size=35 name=LastName value=" & chr(34) & chr(34) & "></td></tr>")
	end if
	
	if ls_EmptyLName = "True" then
		Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=red>")
		Response.Write("<b><i>*Please enter the new employees last name.</i></b></td>")
	end if
	
	Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Start Date:</b></td>")

	if ls_complete = "False" then
		Response.Write("<td><input type=text size=35 name=StartDate value=" & chr(34) & ls_start_date & chr(34) & "></td></tr>")
	else
		Response.Write("<td><input type=text size=35 name=StartDate value=" & Date & "></td></tr>")
	end if
	
	if ls_EmptyStart = "True" then
		Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=red>")
		Response.Write("<b><i>*Please enter the new employees start date.</i></b></td>")
	end if

		
	Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Title:</b></td>")
	Response.Write("<td><select name=Title>")

	NL = chr(13) & chr(10)
	if Request.Form("Submitted") = "True" then
		strSQL = "SELECT Employee_Title_ID, Title_Desc" _
			& " FROM Employee_Title" _
			& " WHERE Employee_Title_ID='" & ls_title & "'"
		set rs = DBQuery(strSQL)
		response.write("<option selected value=" & chr(34) & rs("Employee_Title_ID") & chr(34))
		response.write(">" & rs("Title_Desc") & "</option>" & NL)
		rs.close
	end if
	strSQL = "SELECT Employee_Title_ID, Title_Desc" _
		& " FROM Employee_Title" _
		& " ORDER BY Sort"
	set rs = DBQuery(strSQL)
	while not rs.eof
		response.write("<option value=" & chr(34) & rs("Employee_Title_ID") & chr(34))
		response.write(">" & rs("Title_Desc") & "</option>" & NL)
		rs.movenext
	wend
	rs.close
	
	Response.Write("</select></td></tr>")
	Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Username:</b></td>")

	if ls_complete = "False" then
		Response.Write("<td><input type=text size=35 name=Username value=" & chr(34) & ls_username & chr(34) & "></td></tr>")
	else
		Response.Write("<td><input type=text size=35 name=Username value=" & chr(34) & chr(34) & "></td></tr>")
	end if

	if ls_EmptyUser = "True" then
		Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=red>")
		Response.Write("<b><i>*Please enter the new employees Username.</i></b></td>")
	elseif ls_dup_uname = "True" then
			Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=red>")
			Response.Write("<b><i>*This username is already in use.</i></b></td>")
	end if
	
	Response.Write("<tr><td></td>")
	Response.Write("<td><input type=button class=button value='  Save  ' OnClick='VerifyAddition();'>")
	Response.Write("&nbsp")
	Response.Write("<input type=button class=button value=' Cancel ' OnClick='window.location.href=" & chr(34) & "default.asp" & chr(34) & "'>")
	Response.Write("</td>")
	
end if
%>	 

	</form></table></td></tr>
	
<!-- #include file="../../footer.asp" -->
		

