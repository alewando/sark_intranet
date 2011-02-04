<!--
Developer:    SSeissiger
Date:         8/24/00
Description:  Allows admin to update a solution services member
				Info sent to editPreviousSS.asp and deleteSS.asp to handle the sql statements

-->

<!-- #include file="../../script.asp" -->

<HTML>
<HEAD><TITLE>Edit Solution Services</TITLE></HEAD>
<!-- #include file="../../style.htm" -->

<%
Tech_Spec_ID = Request.QueryString("Tech_Spec_ID")

	'-----------------------------------------------
	'	Get SS member info of the selected record						
	'-----------------------------------------------
set rs = DBQuery("select * from Tech_Specialists WHERE Tech_Specialists_ID = " & Tech_Spec_ID)
	ls_tech_area_id = rs("tech_area_id")
	ls_employee_ID = rs("employee_ID")
	ls_Tech_Specialist_Type_ID = rs("Tech_Specialist_Type_ID")
rs.close
	
	'--------------------------------------------
	'Get the tech area for the selected employee
	'---------------------------------------------
SQL_emp = "SELECT e.FirstName, e.LastName " & _
		"FROM Employee e, Tech_Specialists ts " &_
		"WHERE ts.Employee_ID = e.employeeID AND " &_
		"ts.Tech_Specialists_ID = " & Tech_Spec_ID
set rs_selEmp = DBQuery(SQL_emp)

	ls_empFName = rs_selemp("FirstName")
	ls_empLName = rs_selemp("LastName")
	ls_empName= ls_empFName + " " + ls_empLname
	
	'--------------------------------------------
	'Get the tech area for the selected employee
	'---------------------------------------------
SQL_area = "SELECT ta.Tech_Area, ts.Tech_Area_ID " & _
		"FROM Tech_Area ta, Tech_Specialists ts " &_
		"WHERE ts.Tech_Area_ID = ta.Tech_Area_ID AND " &_
		"ts.Tech_Specialists_ID = " & Tech_Spec_ID
set rs_selArea = DBQuery(SQL_area)

	ls_tech_area = rs_selArea("Tech_Area")
	ls_tech_areaID = rs_selArea("Tech_Area_ID")	
	
	'-------------------------------------------------------------
	'Get the level (specialist or lead) for the selected employee
	'-------------------------------------------------------------
SQL_level = "SELECT substring(s.Specialist_Type,1,10) Specialist_Type, ts.Tech_Specialist_Type_ID " &_
	"FROM Tech_Specialist_Type s, Tech_Specialists ts " &_
	"WHERE ts.Tech_Specialist_Type_ID = s.Tech_Specialist_Type_ID AND " &_
	"ts.Tech_Specialists_ID = " & Tech_Spec_ID
set rs_selTech = DBQuery(SQL_level)

	ls_Tech_Specialist_Type = rs_selTech("Specialist_Type")
	ls_Tech_Specialist_TypeID = rs_selTech("Tech_Specialist_Type_ID")
	
	'-------------------------------------------------------------
	'Get list of all DISTINCT tech areas and Levels
	'-------------------------------------------------------------
	strSQL = "SELECT distinct Tech_Area_ID, Tech_Area " &_
				"FROM Tech_Area ORDER BY Tech_Area"
		set rs_area = DBQuery(strSQL)
	
	strSQL = "SELECT distinct Tech_Specialist_Type_ID, substring(Specialist_Type,1,10) Specialist_Type FROM Tech_Specialist_Type ORDER BY Specialist_Type"
		set rs_level = DBQuery(strSQL)
	%>	

<body bgcolor=silver onLoad="window.resizeBy(0, -200)"> 

<form name=frmInfo action='deleteSS.asp'>
	<INPUT Type=Hidden Name='DelFlag' Value=''>	<INPUT Type=Hidden Name='Tech_Spec_ID' Value=<%=Tech_Spec_ID%>>		
	</form>


<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Please Enter Solution Services Member Changes
</b></font></td></tr></table>


<form NAME="frmUpdate" ACTION="editPreviousSS.asp">
<INPUT TYPE="Hidden" NAME="Tech_Spec_ID" VALUE=<%=Tech_Spec_ID%>> 

<table border=0 height=200>
	<tr>
		<td align=left><font size=1 face="ms sans serif, arial, geneva"><b>Employee</b></font></td>
		<td><INPUT TYPE="Text" NAME="txtSS_name" VALUE="<%=ls_empName%>" size=30 maxlength=50 readonly></td>
	</tr>
<%
	'----------------------------------------------------------------------------------
	'Create the Area combo box with the above employee's area set as the initial value.
	'----------------------------------------------------------------------------------
%>
		<tr><td><font size=1 face='ms sans serif, arial, geneva'><b>Area</b></td>
		<td><select name=sArea id=Text>")
		<option selected value=<%=ls_tech_areaID%>> <%=ls_tech_area%></option>
		<%		
		while not rs_area.eof
			response.write("<option value=" & chr(34) & rs_area("Tech_Area_ID") & chr(34))
			response.write(">" & rs_area("Tech_Area") & "</option>" & NL)
			rs_area.movenext
		wend
		rs_area.close
		Response.Write("</select></td></tr>")
	
		'----------------------------------------------------------------------------------
		'Create the Level combo box with the above employee's area set as the initial value.
		'-----------------------------------------------------------------------------------
		%>
		<tr><td><font size=1 face='ms sans serif, arial, geneva'><b>Specialitst or Lead</b></td>
		<td><select name=sLevel id=Level>
		<option selected value='<%=ls_Tech_Specialist_TypeID%>'><%=ls_Tech_Specialist_Type%></option>
		<%while not rs_level.eof%>
			<option value='<%=rs_level("Tech_Specialist_Type_ID")%>'><%=rs_level("Specialist_Type")%></option>
			<%rs_level.movenext
		wend
		rs_level.close
		%>
		</select></td></tr>
</table>

	<p>
	<INPUT TYPE=button CLASS=button VALUE="Update" OnClick='subSS()' id=button1 name=button1>
	<INPUT TYPE=button CLASS=button VALUE="Cancel" OnClick='closeWindow()' id=button2 name=button2>
	<INPUT type=button class=button value='Delete ' onclick='deleteSS()' id=button3 name=button3>

</form>

</BODY>

<script language="JavaScript" type="text/javascript">
<!--

function subSS()
{
	document.frmUpdate.submit( );
}
		
function closeWindow(){
	//window.parent.close();
	//window.navigate("edit_ss.asp");
	//window.location.reload();
	window.close();
	//window.navigate("edit_ss.asp");
	}		

function deleteSS()
{
	if (confirm("Are you sure you want to delete this record?"))
	{
		document.frmInfo.DelFlag.value = "True"
		document.frmInfo.submit()
	}
}

// -->
</script>
</HTML>

