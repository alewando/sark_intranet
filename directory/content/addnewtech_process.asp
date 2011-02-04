<!--
Developer:    GTYLER
Date:         02/16/2000
Description:  Allows user to add a new technology
-->
<!-- #include file="../../script.asp" -->

<HTML>
<HEAD>
<TITLE>New Technology Submitted</TITLE>
<!-- #include file="../../style.htm" -->

<script language="JavaScript" type="text/javascript">
<!--
	function OpenWin(page, EmpID, WinTitle){
		winTech= window.open(page + "?empID=" + EmpID, WinTitle,"height=615,width=500,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
		}

// --></SCRIPT>

</HEAD>

<%
'***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
'  RIGHT NOW this page does add (unapproved) technologies, but does NOT inform the user
'  if the technology is already in the database, and it does NOT indicate if the submitter
'  considers him/herself an expert in the technology.
'***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
%>

<BODY bgcolor=silver link=blue alink=purple vlink=blue><CENTER>

<form NAME="frmReturn" ACTION="techexperts.asp?empid=<%=Request.QueryString("EmpID")%>" METHOD="POST">

	<INPUT TYPE=hidden NAME="hidUsername" VALUE="<%=Session("Username")%>">
	
	<table border=0 width='90%'>
	<%

	intTechArea = Request.querystring("selArea")
	strNewTechName = Request.querystring("txtNewTechName")
	strNewTechDesc = Request.querystring("txtNewTechDesc")
	strNewTechDesc = replace(strNewTechDesc, "'", "''")
	ls_Expert = Request.QueryString("expert")
	ls_EmpNumber = Request.QueryString("EmpID")
	
	'See if a technology belonging to area intTechArea with name strNewTechName already exists
	strSQL = "SELECT tech_id FROM tech, tech_area WHERE tech.Tech_Area_ID = tech_area.Tech_Area_ID " & _
	         "AND tech.Tech_Area_ID = " & CStr(intTechArea) & " AND Tech_Name = '" & strNewTechName & "'"
	set rs = DBQuery(strSQL)
	
	If not rs.eof Then
		'Don't add the new technology, and inform the user
		rs.close
		response.Write("<tr><td><font size=2 face='ms sans serif, arial, geneva' color=navy>")
		Response.write("<b>Your newly suggested technology has NOT been submitted. ")
		Response.write("It already exisits or has already been submitted for approval.</b></FONT></td></tr>")
	Else
		rs.close
		'Add the new technology, making sure its approved status is false
		strSQL = "INSERT INTO tech (Tech_Name, Tech_Desc, Tech_Standard, Tech_Area_ID, Employee_ID) " & _
		         "VALUES ('" & strNewTechName & "', '" & strNewTechDesc & "', 1, '" & intTechArea & "', '" & ls_EmpNumber & "')"
		DBQuery(strSQL)
		
		'If the user considers him/herself an expert, note this in the DB as well
		if 	ls_Expert=1 then
			strSQL = "SELECT Tech.Tech_ID FROM Tech" _
				& " WHERE Tech_Name = '" & strNewTechName & "' AND" _
				& " Tech_Area_ID = " & intTechArea
			set rs = DBQuery(strSQL)
			ls_TechID = rs("Tech_ID")
			rs.close
			strSQL = "INSERT INTO Employee_Tech_Xref (Employee_ID, Tech_ID)" _
				& " VALUES (" & ls_EmpNumber & ", " & ls_TechID & ")"
			DBQuery(strSQL)
		end if
		
		Response.Write("<tr><td><font size=2 face='ms sans serif, arial, geneva' color=navy><b>")
		Response.Write("Your newly suggested technology has been submitted for approval. ")
		Response.Write("If approved, it will soon be listed under technical expertise.  You will be notified if your submission is not approved.</b></FONT></td></tr>")
	End If
	%>

	</table>

	<P>&nbsp;</P>

	<INPUT TYPE=button CLASS=button VALUE="OK" OnClick='window.close();OpenWin("techexperts.asp", "<%=Request.QueryString("EmpID")%>", "EditSARKTech");'>

</form>

</CENTER></BODY>
</HTML>
