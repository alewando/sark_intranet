<!--
Developer:		DMARSTON
Date:			04/07/2000
Description:	Allows an employee to be deleted from the intranet database.
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
	if (confirm("Are you sure you want to delete this employee?")){ 
		document.frmInfo.EmpID.value = EmployeeID
		document.frmInfo.action.value = "EmployeeSave"
		document.frmInfo.submit()}
	else{
		ReturnToEmployee(EmployeeID)}
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
	<form name=frmInfo action='del_emp.asp' method=post>
	<input type=hidden name='action' value=''>
	<input type=hidden name='EmpID' value=''>
	

<%
NL = chr(13) & chr(10)

'If an employee has been selected from the combo box get the form information...
select case Request.Form("Action")
	case ""
		Response.Write("<table border=0 cellpadding=5>")
		Response.Write("<tr><td colspan=2><font size=2 color=red><b>")
		Response.write("DELETE EMPLOYEE</b></font></td>")
		Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=blue><b>")
		Response.write("To delete an employee, select the employee you wish to remove and click the 'Delete' button.</b></td></tr>")
		
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

		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Title:</b></td>")
		Response.Write("<td>&nbsp</td></tr>")
	
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Username:</b></td>")
		Response.Write("<td>&nbsp</td></tr>")

		Response.Write("<tr><td></td>")
		'Response.Write("<td><input type=button class=button value='  Delete  ' OnClick='NoOneSelected();'>")
		Response.Write("<td>")
		'Response.Write("&nbsp")
		Response.Write("<input type=button class=button value=' Cancel ' OnClick='window.location.href=" & chr(34) & "default.asp" & chr(34) & "'>")
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
		Response.write("DELETE EMPLOYEE</b></td>")
		Response.Write("<tr><td colspan=2><font size=1 face='ms sans serif, arial, geneva' color=blue><b>")
		Response.write("To delete an employee, select the employee you wish to remove and click the 'Save' button.</b></td></tr>")
		
		Response.Write("<tr><td><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Employee:</b></td>")
		Response.Write("<td><font size=2 face='ms sans serif, arial, geneva' color=red><b>" & rs_emp("LastName") & ", " & rs_emp("FirstName") & "</b></font></td>")

		Response.Write("<tr><td width='20%'><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Title:</b></td>")
		Response.Write("<td><font size=2 face='ms sans serif, arial, geneva' color=red>" & rs_emp("Title_Desc") & "</font></td></tr>")
	
		Response.Write("<tr><td width='20%'><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Username:</b></td>")
		Response.Write("<td><font size=2 face='ms sans serif, arial, geneva' color=red>" & rs_emp("Username") & "</font></td></tr>")
		Response.Write("<tr><td></td>")
		Response.Write("<td><input type=button class=button value='  Delete  ' OnClick='SaveChanges(" & Request.Form("EmpID") & ");'>")
		Response.Write("&nbsp")
		Response.Write("<input type=button class=button value=' Cancel ' OnClick='BackToSelection();'>")
		Response.Write("</td>")
		rs_emp.close
	
	case "EmployeeSave"
		'Begin code added 6/1/00 by Ryan Dlugosz
		'The following statements check the Classifieds, Event, News, Employee_cert_xref,
		'Employee_industry_xref, Employee_project_xref, and Employee_tech_xref tables for
		'possible FK constraint issues before the actual DELETE is processed.
		'Classifieds and employee_* records are deleted; Event and News have the employee_id 
		'set to Chris Dolan's ID (Id = 3).
		strSQL = "DELETE FROM classifieds WHERE employee_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "UPDATE event SET contact_id = 3 WHERE contact_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "UPDATE event SET creator_id = 3 WHERE creator_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "UPDATE news SET employee_id = 3 WHERE employee_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "DELETE FROM employee_cert_xref" _
			& " WHERE employee_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "DELETE FROM employee_industry_xref" _
			& " WHERE employee_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "DELETE FROM employee_project_xref" _
			& " WHERE employee_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "DELETE FROM employee_tech_xref" _
			& " WHERE employee_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "DELETE FROM review" _
			& " WHERE employeeid = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "DELETE FROM skills_inventory" _
			& " WHERE employeeid = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		strSQL = "DELETE FROM employee_cert_numbers" _
			& " WHERE employee_id = " & Request.Form("EmpID")
		DBQuery(strSQL)
		
		'End code added by Ryan Dlugosz 6/1/00
	
		strSQL = "DELETE FROM Employee" _
			& " WHERE (EmployeeID = " & Request.Form("EmpID") & ")"
		DBQuery(strSQL)
		Response.write("<img src='../../common/images/checker.gif' height=32 width=32 border=0 align=left>")
		Response.Write("<table border=0 cellpadding=10>")
		Response.Write("<tr><td align='center'><font size=1 face='ms sans serif, arial, geneva' color=blue>")
		Response.Write("<b>Employee has been deleted.</b></font></td></tr>")
		Response.Write("<tr><td><input type=button class=button value='  OK  ' OnClick='BackToSelection();'></td></tr>")
	
		
end select

%>	 

</table></form></td></tr>
	
<!-- #include file="../../footer.asp" -->
		

