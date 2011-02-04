<!--
Developer:		DMARSTON
Date:			04/07/2000
Description:	Allows an employee to be edited in the intranet database.
History:		Created - 04/07/2000-->

<!-- #include file="../section.asp" -->

<STYLE>
#Title
{
    WIDTH: 200pt
}

#Text
{
	WIDTH: 200pt
}

</STYLE>

<script language=javascript>
<!--
	
	function SaveChanges(EmployeeID){
		document.frmInfo.EmpID.value = EmployeeID
		document.frmInfo.action.value = "EmployeeSave"
		document.frmInfo.submit()
	}
	  
	function EmployeeSelected(){
		document.frmInfo.EmpID.value = document.frmInfo.EmployeeID.value
		document.frmInfo.action.value = "EmployeeSelected"
		document.frmInfo.submit()
	}
	
	function NoOneSelected(){
		alert("Please select and employee!")
	}
	
	function BackToSelection(){
		document.frmInfo.action.value = ""
		document.frmInfo.submit()
	}
	
	function ReturnToEmployee(EmployeeID){
		document.frmInfo.EmpID.value = EmployeeID
		document.frmInfo.action.value = "EmployeeSelected"
		document.frmInfo.submit()
	}		

// -->
</script>

<tr><td valign=top>
	<form name=frmInfo action='chg_emp.asp' method=post>
	<input type=hidden name='action' value=''>
	<input type=hidden name='EmpID' value=''>
	

<%
NL = chr(13) & chr(10)

'If an employee has been selected from the combo box get the form information...
select case Request.Form("Action")
	case ""
		Response.Write("<table border=0 cellpadding=5>")
		Response.Write("<tr><td colspan=2><font size=2 color=red><b>")
		Response.write("EDIT EMPLOYEE</b></font></td>")
		Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=blue><b>")
		Response.write("To change employee information, select the employee you wish to edit, modify the appropriate fields and click the 'Save' button when finished.</b></td></tr>")
		
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Select Employee:</b></td>")
		Response.Write("<td><select name=EmployeeID id=Title onChange='EmployeeSelected();'>")
		strSQL = "SELECT EmployeeID, LastName, FirstName" _
			& " FROM Employee" _
			& " ORDER BY LastName"
		set rs_employees = DBQuery(strSQL)
		Response.Write("<option value=" & chr(34) & chr(34) & "></option>" & NL)
		while not rs_employees.eof
			response.write("<option value=" & chr(34) & rs_employees("EmployeeID") & chr(34))
			response.write(">" & rs_employees("LastName") & ", " & rs_employees("FirstName") & "</option>" & NL)
			rs_employees.movenext
		wend
		rs_employees.close
	
		Response.Write("</select></td></tr>")

		
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Last Name:</b></td>")
		Response.Write("<td><input type=text name=LastName id=Text value=" & chr(34) & chr(34) & "></td></tr>")
	
		'Creat the Title combo box.
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Title:</b></td>")
		Response.Write("<td><select name=Title id=Title>")
		Response.Write("</select></td></tr>")
	
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Username:</b></td>")
		Response.Write("<td><input type=text name=Username id=Text value=" & chr(34) & chr(34) & "></td></tr>")

		Response.Write("<tr><td></td>")
		Response.Write("<td>")
		Response.Write("<input type=button class=button value=' Cancel ' OnClick='window.location.href=" & chr(34) & "default.asp" & chr(34) & "' id=button1 name=button1>")
		Response.Write("</td>")

	case "EmployeeSelected"
		strSQL = "SELECT Employee.LastName, Employee.EmployeeTitle_ID," _
			& " Employee_Title.Title_Desc, Employee.Username, Employee.FirstName" _
			& " FROM Employee INNER JOIN" _
			& " Employee_Title ON" _
			& " Employee.EmployeeTitle_ID = Employee_Title.Employee_Title_ID" _
			& " WHERE (Employee.EmployeeID = " & Request.Form("EmpID") & ")"
		
		set rs_emp=DBQuery(strSQL)
	
		Response.Write("<table border=0 cellpadding=5>")
		Response.Write("<tr><td colspan=2><font size=2 color=red><b>")
		Response.write("EDIT EMPLOYEE</b></td>")
		Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=blue><b>")
		Response.write("To change employee information, select the employee you wish to edit, modify the appropriate fields and click the 'Save' button when finished.</b></td>")
		
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Employee:</b></td>")
		Response.Write("<td><font size=2 face='ms sans serif, arial, geneva' color=red><b>" & rs_emp("LastName") & ", " & rs_emp("FirstName") & "</b></font></td>")

		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Last Name:</b></td>")
		Response.Write("<td><input type=text name=LastName id=Text value=" & chr(34) & rs_emp("LastName") & chr(34) & "></td></tr>")
	
		'Creat the Title combo box with the above employee's title set as the initial value.
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Title:</b></td>")
		Response.Write("<td><select name=Title id=Title>")
		response.write("<option selected value=" & chr(34) & rs_emp("EmployeeTitle_ID") & chr(34))
		response.write(">" & rs_emp("Title_Desc") & "</option>" & NL)
		strSQL = "SELECT Employee_Title_ID, Title_Desc" _
				& " FROM Employee_Title"
		set rs_titles = DBQuery(strSQL)
		while not rs_titles.eof
			response.write("<option value=" & chr(34) & rs_titles("Employee_Title_ID") & chr(34))
			response.write(">" & rs_titles("Title_Desc") & "</option>" & NL)
			rs_titles.movenext
		wend
		rs_titles.close
		Response.Write("</select></td></tr>")
	
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Username:</b></td>")
		Response.Write("<td><input type=text name=Username id=Text value=" & chr(34) & rs_emp("Username") & chr(34) & "></td></tr>")
		Response.Write("<tr><td></td>")
		Response.Write("<td><input type=button class=button value='  Save  ' OnClick='SaveChanges(" & Request.Form("EmpID") & ");'>")
		Response.Write("&nbsp")
		Response.Write("<input type=button class=button value=' Cancel ' OnClick='BackToSelection();'>")
		Response.Write("</td>")
		rs_emp.close
	
	case "EmployeeSave"
		ls_Empty = ""
		if Request.Form("LastName") <> "" then
			if Request.Form("Username") <> "" then
				strSQL = "SELECT Employee.Username FROM Employee" _
						& " WHERE Employee.Username = '" & Request.Form("Username") & "'" _
						& " AND Employee.EmployeeID <> '" & Request.Form("EmpID") & "'"
				set rs=DBQuery(strSQL)
				if not rs.eof then
					Response.write("<img src='../../common/images/stop.gif' height=50 width=50 border=0 align=left>")
					Response.Write("<table border=0 cellpadding=10>")
					Response.Write("<tr><td align='center'><font size=2 face='ms sans serif, arial, geneva' color=blue>")
					Response.Write("<b>Username: <font size=2 face='ms sans serif, arial, geneva' color=red>'" & Request.Form("Username") & "'</font>, is already in use.</b></font></td></tr>")
					Response.Write("<tr><td><input type=button class=button value='  OK  ' OnClick='ReturnToEmployee(" & Request.Form("EmpID") & ");'></td></tr>")
				else
					strSQL = "UPDATE Employee" _
							& " SET LastName = '" & Request.Form("LastName") & "'," _
							& " EmployeeTitle_ID = '" & Request.Form("Title") & "'," _
							& " Username = '" & Request.Form("Username") & "'" _
							& " WHERE (EmployeeID = " & Request.Form("EmpID") & ")"
					DBQuery(strSQL)
					Response.write("<img src='../../common/images/checker.gif' height=32 width=32 border=0 align=left>")
					Response.Write("<table border=0 cellpadding=10>")
					Response.Write("<tr><td align='center'><font size=1 face='ms sans serif, arial, geneva' color=blue>")
					Response.Write("<b>Changes have been saved.</b></font></td></tr>")
					Response.Write("<tr><td><input type=button class=button value='  OK  ' OnClick='BackToSelection();'></td></tr>")
				end if
			else
				Response.write("<img src='../../common/images/stop.gif' height=50 width=50 border=0 align=left>")
				Response.Write("<table border=0 cellpadding=10>")
				Response.Write("<tr><td align='center'><font size=2 face='ms sans serif, arial, geneva' color=red>")
				Response.Write("<b>Employee's Username cannot be left blank.</b></font></td></tr>")
				Response.Write("<tr><td><input type=button class=button value='  OK  ' OnClick='ReturnToEmployee(" & Request.Form("EmpID") & ");'></td></tr>")
			end if
		else
			Response.write("<img src='../../common/images/stop.gif' height=50 width=50 border=0 align=left>")
			Response.Write("<table border=0 cellpadding=10>")
			Response.Write("<tr><td align='center'><font size=2 face='ms sans serif, arial, geneva' color=red>")
			Response.Write("<b>Employee's last name cannot be left blank.</b></font></td></tr>")
			Response.Write("<tr><td><input type=button class=button value='  OK  ' OnClick='ReturnToEmployee(" & Request.Form("EmpID") & ");'></td></tr>")
		end if
		
		
end select

%>	 

</table></form></td></tr>
	
<!-- #include file="../../footer.asp" -->
		

