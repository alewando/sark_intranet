<!--
Developer:    KDill
Date:         11/17/1998
Description:  Welcome to the technology section.
-->


<!-- #include file="../section.asp" -->

<table border=0 cellspacing=0 cellpadding=0><tr><td valign=top><font size=1 face='ms sans serif, arial, geneva'>

<img src="images/hackanm.gif" height=48 width=48 hspace=10 vspace=5 align=left>
<b>Help</b><br>
Do you have a technical issue and don't know who to call for help?
Simply select your technology to find out which SARKs can help provide
technical support.  These experts are SARKs who have declared themselves to
be knowledgeable in their technical area, and are willing to receive your
phone calls and emailed requests for assistance.  To designate yourself as an
expert, go to <a href="../../directory/content/employee.asp">the directory</a> and
edit your profile.
<p>
<b>Solution Services</b><br>
The various technologies are listed according to the Solution Services group Practice Area / Project Support Area
that it falls under.  The Leads and Specialists within Solution Services are responsible for
supporting and promoting the various technologies within the branch.

</font></td><td width=10>&nbsp;</td><td valign=middle nowrap>

<table border=0 cellspacing=0 cellpadding=9><tr><td valign=top bgcolor=#ffffcc>
<table cellpadding=2 cellspacing=0>
	<tr><td nowrap><font size=1 face='ms sans serif, arial, geneva'>
		<b>What does Solution<br>Services do?</b><br>
		</font></td></tr>
	<tr><td><font size=1 face='ms sans serif, arial, geneva'>
		&nbsp;&nbsp;&nbsp;&nbsp;<font color=red><b>S</b></font>ales Support<br>
		</font></td></tr>
	<tr><td><font size=1 face='ms sans serif, arial, geneva'>
		&nbsp;&nbsp;&nbsp;&nbsp;<font color=red><b>T</b></font>ech Support<br>
		</font></td></tr>
	<tr><td><font size=1 face='ms sans serif, arial, geneva'>
		&nbsp;&nbsp;&nbsp;&nbsp;<font color=red><b>E</b></font>ducation<br>
		</font></td></tr>
	<tr><td><font size=1 face='ms sans serif, arial, geneva'>
		&nbsp;&nbsp;&nbsp;&nbsp;<font color=red><b>E</b></font>vangelize<br>
		</font></td></tr>
	<tr><td><font size=1 face='ms sans serif, arial, geneva'>
		&nbsp;&nbsp;&nbsp;&nbsp;<font color=red><b>R</b></font>epository
		<!-- <p>About our acronym... -->
		</font></td></tr>
</table>
</td></tr></table>

</td></tr></table>
<p>

<table width="100%" border=0 cellspacing=0 cellpadding=0><tr><td valign=top><font size=1 face='ms sans serif, arial, geneva'>
<b>Resources</b><br>
Also included in this section are links to various technical websites.  Anyone can 
add links they find useful for a new technology, and the most frequently used links
are listed at the top of the list.  Experts can also review and edit links as needed.
<p>
<b>New Technologies</b><br>
If you'd like to see a new technology included, go to your Profile and click on Edit Expertise and then 
click on the Add Technology button to submit a new technology.  Your submission will be reviewed by the 
Webmaster.  You will be notified if your submission is rejected.

</font></td><td width=10>&nbsp;</td><td valign=top></td></tr></table>
<p>
<table border=0 cellspacing=0 cellpadding=10><tr><td valign=top align=center bgcolor=#ffffcc>
<table border=0 cellspacing=0 cellpadding=0>
	<tr><td><font size=1 face='ms sans serif, arial, geneva'><b>Areas</b></font></td>
	<td width=30>&nbsp;</td><td><font size=1 face='ms sans serif, arial, geneva'><b>Lead</b></font></td>
	<td width=30>&nbsp;</td><td><font size=1 face='ms sans serif, arial, geneva'><b>Specialist</b></font></td></tr>
<%
sql = "SELECT Tech_Area.Tech_Area, Tech_Specialists.Tech_Specialist_Type_ID, Employee.FirstName, Employee.LastName, Employee.Username, Employee.EmployeeID " & _
		"FROM Employee INNER JOIN Tech_Specialists ON Employee.EmployeeID = Tech_Specialists.Employee_ID RIGHT OUTER JOIN Tech_Area ON " & _
		"Tech_Specialists.Tech_Area_ID = Tech_Area.Tech_Area_ID " & _
		"WHERE (Tech_Specialists.Tech_Specialist_Type_ID = 1) " & _
		"OR (Tech_Specialists.Tech_Specialist_Type_ID = 2) " & _
		"ORDER BY Tech_Area.Tech_Area, Tech_Specialists.Tech_Specialist_Type_ID"
set rs = DBQuery(sql)

while not rs.eof
	tmpArea = rs("Tech_Area")
	response.write "<tr><td valign=top nowrap><font size=1 face='ms sans serif, arial, geneva'>" & rs("tech_area") & "</font></td><td>&nbsp;</td><td valign=top nowrap>"
	count=1
	PLead=0
	do while rs("Tech_Area") = tmpArea
		if count=1 then
			if rs("Tech_Specialist_Type_ID") = 1 then
				response.write "<font size=1 face='ms sans serif, arial, geneva'><b><a href='../../directory/content/details.asp?EmpID=" & rs("EmployeeID") & "'>"
				response.write rs("FirstName") & " " & rs("LastName") & "</a></b><br>"
				PLead=1
			elseif rs("Tech_Specialist_Type_ID") = 2 then
				response.write "<td>&nbsp;</td><td valign=top nowrap><font size=1 face='ms sans serif, arial, geneva'><a href='../../directory/content/details.asp?EmpID=" & rs("EmployeeID") & "'>"
				response.write rs("FirstName") & " " & rs("LastName") & "</a><br>"
			end if
		elseif count>1 then
			if PLead=0 then
				Response.Write "</tr><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
			end if
			PLead=0
			if rs("Tech_Specialist_Type_ID") = 1 then
				response.write "<td valign=top nowrap><font size=1 face='ms sans serif, arial, geneva'><b><a href='../../directory/content/details.asp?EmpID=" & rs("EmployeeID") & "'>"
				response.write rs("FirstName") & " " & rs("LastName") & "</a></b><br>"
			elseif rs("Tech_Specialist_Type_ID") = 2 then
				response.write "<td>&nbsp;</td><td valign=top nowrap><font size=1 face='ms sans serif, arial, geneva'><a href='../../directory/content/details.asp?EmpID=" & rs("EmployeeID") & "'>"
				response.write rs("FirstName") & " " & rs("LastName") & "</a><br>"
			end if
			Response.Write "</td></tr>"
		end if
		
		count=count+1
		rs.movenext
		if rs.eof then
			exit do
		end if
	loop
	Response.Write ("</td></tr><tr></tr>")
wend

rs.close
%>
</table>
</td></tr></table>

<!-- #include file="../../footer.asp" -->
