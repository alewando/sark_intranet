<!--
Developer:	  SSeissiger
Date:         08/10/2000
Description:  Allows a solution services to be added.
-->
<% @language=VBscript%>

<!-- #include file="../section.asp" -->

<%
If Request.Form("Submitted") = "True" Then
	emp_entered = Request.form("sEmpID")
	area_entered = Request.Form("sArea")
	level_entered = Request.Form("sLevel")
	
	' Make sure that null values didn't get sent into this page.
	If emp_entered <> "" And area_entered <> "" And level_entered <> ""  And Request.Form("Submitted") = "True" Then
		'------------------------------------------------------------------
		' Check to see if employee is already in the Solution Services area
		' If so, inform user and don't insert!
		'------------------------------------------------------------------
		sSQL="Select 1 From Tech_Specialists where employee_ID= " & emp_entered & " and Tech_Area_ID= " & area_entered
		Set rs = DBQuery(sSQL)
		If Not rs.eof Then
			duplicateFlag=True		' Employee already in this area, set flag for message
		Else
			duplicateFlag=False
			sSQL="Insert into Tech_Specialists (employee_ID, Tech_Area_ID, Tech_Specialist_Type_ID) " & _
			     "values ('" & emp_entered & "', '" & area_entered & "', " & level_entered & ")"
			Set rs = DBQuery(sSQL)
		End If
		
	End If
End If
%>

<HTML>
<HEAD><title>Add New Solution Services Team Member</title></HEAD>

<BODY>

<table border=0>
<tr>
	<td colspan=2><font size=2 color=red>
		<b>Add Solution Services Team Member</b></font>
	</td>
</tr><tr>
</tr></table>


<form NAME="frmAddSS" Action="add_ss.asp" method=post>
	<input type=hidden name="Submitted" value="False">
	<table border=0 width=100%>
<tr>
    <TD width=150>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><b>Employee</b> </FONT></P>
    </TD>
    <TD></TD>
    <TD><select name=sEmpID id=Title>
    	<% 
		'--------------------------------------------
		'Get a list of employees - exclude sales, 
		'administrative, intern, and recruiting staff
		'---------------------------------------------
		strSQL = "SELECT EmployeeID, LastName, FirstName FROM Employee " & _
			"WHERE EmployeeTitle_ID not in ('9','10','15','17','11','22','12','21','20') ORDER BY LastName"
		set rs_emp = DBQuery(strSQL)%>
		<!-- <option value=" & chr(34) & chr(34) & "></option> -->
		<option value=""></option>
		<%while not rs_emp.eof%>
			<option value=<%=rs_emp("EmployeeID")%>><%=rs_emp("LastName")%>,&nbsp;<%=rs_emp("FirstName")%> </option>
		<%rs_emp.movenext
		wend
		rs_emp.close
		%>
    </TD>
</tr>
<tr>
    <TD width=150>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><b>Area</b> </FONT></P>
    </TD>
    <TD></TD>
    <TD><select name=sArea id=Title>
		<% strSQL = "SELECT Tech_Area_ID, Tech_Area FROM Tech_Area " & _
			"ORDER BY Tech_Area"
		set rs_emp = DBQuery(strSQL)%>
		<!-- <option value=" & chr(34) & chr(34) & "></option> -->
		<option value=""></option>
		<%while not rs_emp.eof%>
			<option value=<%=rs_emp("Tech_Area_ID")%>><%=rs_emp("Tech_Area")%> </option>
		<%rs_emp.movenext
		wend
		rs_emp.close
		%>
    </TD>
</tr>
<tr>
    <TD width=150>
		<P align=left><FONT face="ms sans serif, arial, geneva" size=1 color=blue><b>Lead or Specialist</b> </FONT></P>
    </TD>
    <TD></TD>
    <TD><select name=sLevel id=Title >
		<% strSQL = "SELECT Tech_Specialist_Type_ID, substring(Specialist_Type,1,10) Specialist_Type FROM Tech_Specialist_Type ORDER BY Specialist_Type"
		set rs_emp = DBQuery(strSQL)%>
		<!-- <option value=" & chr(34) & chr(34) & "></option> -->
		<option value=""></option>
		<%while not rs_emp.eof%>
			<option value=<%=rs_emp("Tech_Specialist_Type_ID")%>><%=rs_emp("Specialist_Type")%></option>
		<%rs_emp.movenext
		wend
		rs_emp.close
		%>
    </TD>
</tr>
<tr height=10></tr>
<tr>
   <TD></TD>
   <TD></TD>
   <TD>	<INPUT type=button class=button onclick='CheckInfo();' value=' Insert ' id=button1 name=button1>&nbsp
		<INPUT type=button class=button value=' Cancel ' OnClick='GoHome();' id=button2 name=button2>
   </td>
</tr>
<%
If duplicateFlag Then
	'---------------------------------------------------------------------------
	' The user attempted to re-add the employee into the same Solution Services
	' practice area. We tell them about it.
	'---------------------------------------------------------------------------
%>
	<tr>
	<TD COLSPAN=2>
	&nbsp;
	</TD>
	</tr>
	<tr>
	<TD COLSPAN=2>
	<b>	<font size=2 color=red>
	Employee already in this Solution Services Area!
	</font></b>
	</TD>
	</tr>
<%
End If
%>
</form>
</table>

</BODY>

<!-- #include file="../../footer.asp" -->

</HTML>

<SCRIPT LANGUAGE=javascript>
<!--
function CheckInfo()
{	
	//alert( "Value=" + document.frmAddSS.sEmpID.options[document.frmAddSS.sEmpID.selectedIndex].value );

	//--------------------------------------------------------
	// Validates all SELECT objects. selectedIndex=0 indicates
	// the first option was selected and this is blank 
	// (not selected from user's perspective)
	//--------------------------------------------------------
	if ( document.frmAddSS.sEmpID.selectedIndex < 1 || 
	     document.frmAddSS.sArea.selectedIndex < 1 || 
	     document.frmAddSS.sLevel.selectedIndex < 1 )
	{
		alert("Selection(s) cannot be blank!")
	}
	else 
	{
		document.frmAddSS.Submitted.value = "True"
		document.frmAddSS.submit()
	}
		
}
	
function GoHome()
	{
		window.navigate("default.asp")
	}
//-->
</SCRIPT>