<!--
Developer:		SSeissiger
Date:			8/22/2000
Description:	Allows an solution services member to be edited in the intranet database.				Adapted from chg_emp.asp-->

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
	function checkChildWindow() 
	{
		var empid;
		
		if ( !window.winTech || !window.winTech.open || window.winTech.closed )
		{
			//------------------------------------------------------
			// The new window created by editDocument() has closed
			// we turn off the timer allowing us to keep checking
			// and force a refresh of the information.
			//------------------------------------------------------
			clearInterval(intervalID);

			//------------------------------------------------------
			// This is a combination of server side and client side 
			// script so we can save the employee ID for the refresh.
			//------------------------------------------------------
			<%If Request.Form("EmpID") & "x" <> "x" Then%>
				empid = <%=Request.Form("EmpID")%>;
			<%End If%>
			
			ReturnToEmployee(empid);
		}
	}
	
	function EmployeeSelected(){
		document.frmInfo.EmpID.value = document.frmInfo.EmployeeID.value
		document.frmInfo.action.value = "EmployeeSelected"
		document.frmInfo.submit()
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
	
	function editDocument(arg)
	{
		winTech=window.open("update_ss.asp?tech_Spec_ID=" + arg, "AddDocument","height=615,width=500,toolbar = 0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
		//Set up to allow automatic refresh when new window is closed
		intervalID=setInterval('checkChildWindow()', 500);
	}		
// -->
</script>

<!-- <tr><td valign=top> -->
	<form name=frmInfo action='edit_ss.asp' method=post>
	<input type=hidden name='action' value=''>
	<input type=hidden name='EmpID' value=''>
<%
NL = chr(13) & chr(10)

'----------------------------------------------------------------------------
'If an employee has been selected from the combo box get the form information
'----------------------------------------------------------------------------
select case Request.Form("Action")
	case ""
%>
		<font size=2 color=red>
		<b>Edit Solution Services Member</b>
		<P>
		<font size=1 face='ms sans serif, arial, geneva' color=black>
		<b><p>To change employee information, select the employee you wish to edit.</b></p>
		<BR><P>
		<table border=0 cellpadding=0 width='90%'>
		<%	
		'---------------------------------------------------
		'Get the names of all solution services team members
		'---------------------------------------------------
		strSQL = "SELECT distinct t.Employee_ID, e.LastName, e.FirstName " & _
			"FROM Employee e, Tech_Specialists t " & _
			"WHERE t.Employee_ID = e.EmployeeID " &_
			"ORDER BY e.LastName, e.FirstName"
		set rs_emp = DBQuery(strSQL)
		
		If Not rs_emp.EOF Then
		%>
			<tr><td width=100><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Select Employee</b>
			</font></td>
			<td width=100><select name=EmployeeID id=Title onChange='EmployeeSelected();'>
			<option value="">
			<%while not rs_emp.eof%>
				<option value='<%=rs_emp("Employee_ID")%>'><%=rs_emp("LastName")%>, <%=rs_emp("FirstName")%></option>
				<%rs_emp.movenext
			wend
			rs_emp.close
		%>
			</select></td></tr>
			<tr></tr>
		<%Else%>
			<tr><td rowspan=3 colspan=2><font size=1 face='ms sans serif, arial, geneva' color=blue><b>There are no employees in Solution Services. Click the Cancel button to return to the tools menu.</b>
			</font></td>
			</tr>
			<TR><TD colspan=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
			<TR><TD colspan=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
		<%End If%>
		<tr height=10></tr><tr><td colspan=2 align=center><input type=button class=button value=' Cancel ' OnClick='window.navigate("default.asp");' id=button1 name=button1>
		</td></tr>
		
	<%case "EmployeeSelected"
		'---------------------------------------
		'Get the info for the selected employee
		'---------------------------------------
		strSQL = "SELECT e.LastName, e.FirstName, e.EmployeeID, t.Tech_Area_ID, " &_
			"t.Tech_Specialists_ID, t.Tech_Specialist_Type_ID " &_
			"FROM Employee e, Tech_Specialists t " &_
			"WHERE e.EmployeeID = t.Employee_ID AND " &_
			"e.EmployeeID = " & Request.Form("EmpID") &_
			"ORDER BY Tech_Specialists_ID"
		
		set rs_selectedEmp=DBQuery(strSQL)
		If Not rs_selectedEmp.EOF Then
			'-----------------------------------------------------
			' The employee has rows in the Tech_Specialists table
			'-----------------------------------------------------
			empLastName = rs_selectedEmp("LastName")
			empFirstName = rs_selectedEmp("FirstName")
		
			'--------------------------------------------
			'Get the tech area for the selected employee
			'---------------------------------------------
		
			SQL_area = "SELECT ts.Employee_ID, ts.Tech_Area_ID, ta.Tech_Area " & _
			"FROM Tech_Area ta, Tech_Specialists ts " &_
			"WHERE ts.Tech_Area_ID = ta.Tech_Area_ID AND " &_
			"ts.Employee_ID = " & Request.Form("EmpID")
		
			set rs_selArea = DBQuery(SQL_area)
		
			'---------------------------------------------------------------------------------
			'Get the level (specialist or lead) for the selected employee
			'The substring function is used to remove the "(s)" off the end of "specialist(s)"
			'----------------------------------------------------------------------------------
			SQL_level = "SELECT ts.Employee_ID, ts.Tech_Area_ID, substring(s.Specialist_Type,1,10) Specialist_Type " &_
			"FROM Tech_Specialist_Type s, Tech_Specialists ts " &_
			"WHERE ts.Tech_Specialist_Type_ID = s.Tech_Specialist_Type_ID AND " &_
			"ts.Employee_ID = " & Request.Form("EmpID")
		
			set rs_selTech = DBQuery(SQL_level)
		End If
		%>	
		
			<font size=2 color=red>
			<b>Edit Solution Services Member</b>
			<P>
			<font size=1 face='ms sans serif, arial, geneva' color=black>
			<b><p>To change employee information, select the Solution Services area you wish to edit.</b></p>
			<BR><P>
			<table border=0 cellpadding=0>
			<tr><td width=70><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Employee</b>
			</font></td>
			<%If Not rs_selectedEmp.EOF Then%>
			<td width=70><input width=100 type=text name=txtLastName id=Text value=<%=empFirstname%>&nbsp;<%=empLastName%> readonly>
			<%Else%>
			<td width=70><input width=100 type=text name=txtLastName id=Text value=&nbsp; readonly>
			<%End If%>
			</td>
			</table>
		
			<table border=0 cellpadding=0 width='75%'>
			<!--
			<colgroup>
			<col width="49%">
			<col width="51%">
			</colgroup>
			-->
			<tr> <td width='49%'>&nbsp;</td><td width='51%'>&nbsp;</td></tr>
			<tr height=10></tr>	
			<tr><td><font size=2 face='ms sans serif, arial, geneva' color=black><b><u>Area</u></b></font></td>
			<td><font size=2 face='ms sans serif, arial, geneva' color=black><b><u>Specialist or Lead</u></b></font></td>
			</tr>
			<tr height=10></tr>
			<%If Not rs_selectedEmp.EOF Then%>
			<tr>
			<%
			while not rs_selectedEmp.eof
				ls_tech_spec_ID = rs_selectedEmp("Tech_Specialists_ID")
				ls_tech_area = rs_selArea("Tech_Area")
				ls_tech_level = rs_selTech("Specialist_Type")
			%>
				<tr><td align=left width='49%'>
						<input type=hidden name='tech_Spec_ID' value='<%=ls_tech_spec_ID%>'>
						<font size=1 face='ms sans serif, arial, geneva' color=black>
						
						<a href="javascript:editDocument(<%=ls_tech_spec_ID%>)"; onMouseOver="top.status='Update or Delete Solution Services area/role'; return true;" onMouseOut="top.status=''; return true;"> <%=ls_tech_area%> </a>
					</td>
				<td align=left width='51%'><font size=1 face='ms sans serif, arial, geneva' color=black><%=ls_tech_level%></td>
			
				</tr>
				<tr height=10></tr>
			<%
				rs_SelectedEmp.movenext
				rs_selTech.movenext
				rs_selArea.movenext
			wend
			rs_selTech.close
			rs_selArea.close
			Else
			'--------------------------------------------------------------------------
			' The specified employee is no longer in a Solution Services practice area
			'--------------------------------------------------------------------------
		%>
		<tr><td rowspan=3 colspan=2><font size=1 face='ms sans serif, arial, geneva' color=blue><b>Employee is no longer a Lead or Specialist in any Solution Services Practice Areas. Click the Cancel button to return to the Edit Employee menu.</b>
		</font></td>
		</tr>
		<TR><TD colspan=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
		<TR><TD colspan=2><IMG SRC="../../common/images/spacer.gif" HEIGHT=3></TD></TR>  
		<%
		End If
		rs_SelectedEmp.close	
		%>
		<tr height=10></tr><tr><td colspan=2 align=center><input type=button class=button value=' Cancel ' OnClick='BackToSelection();'>
		</td></tr>
<%
end select

%>	 

</table></form>
<!--
</td></tr>
-->
		
<!-- #include file="../../footer.asp" -->
