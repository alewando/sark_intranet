<!-- #include file="../section.asp" -->

<!--
Developer:    Chris Dolan
Date:         4/2000
-->

<!-- KNOWLEGE SNAPSHOT -->

<script language=javascript>
function toggle(item){
	var docItem = document.all(item)
	if (docItem.style.display == ""){
		docItem.style.display = "none"
		document.all(item + "Header").innerHTML = "(collapsed)"
		}
	else {
		docItem.style.display = ""
		document.all(item + "Header").innerHTML = ""
		}
	}
</script>

<font size=2 face="ms sans serif, arial, geneva"  color=black>
<b>This section provides you with a summary of the Cincinnati branch's Industry Experience and Certifications.  Click on a hyperlink below to go directly to a Sark profile.</b>

<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''
'Code to display a list of the Industry 
'Experts sorted alphabetically by Industry by
'consultant.
''''''''''''''''''''''''''''''''''''''''''''''-->
<%
sql = "SELECT Employee.EmployeeID, Employee.FirstName, Employee.LastName, Industry.Industry_Name" _
      & " FROM Employee_Industry_Xref INNER JOIN" _
      & " Industry ON" _ 
      & " Employee_Industry_Xref.Industry_ID = Industry.Industry_ID" _
      & " INNER JOIN" _
      & " Employee ON" _ 
      & " Employee_Industry_Xref.Employee_ID = Employee.EmployeeID" _
      & " ORDER BY Industry.Industry_Name, Employee.LastName" 
set rs = DBQuery(sql)

prev_ind = ""
%>
<table class=tableShadow width=225 cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc>
<tr><td><font size=2.5 face= "ms sans serif, arial, geneva"  color = navy>
	<b>Industry Experts</b>
</td></tr>
<tr><td colspan=2 bgcolor=gray height=1>
</td></tr><br>
<%while not rs.eof
	
	id = (rs("EmployeeID"))
	first_name = rs("FirstName")
	last_name = rs("LastName")
	ind = trim(rs("Industry_Name"))

	if prev_ind = "" OR prev_ind <> ind then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = navy>
		<b><span onClick="toggle('<%=ind%>')" alt="Expand or collapse this section." style="cursor:hand"><%=ind%></b>
		<span id='<%=ind & "Header"%>' style="">(collapsed)</span>
		<!--<b><%=ind%></b>-->
		</td>	</tr>
	<%end if%>	<%prev_ind = ind%>		<tr><td colspan=2>
		<span id='<%=ind%>' style="display:none">			<table>									<% while prev_ind = ind%>
				<tr>
				<td valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=id%>'><%=first_name%>&nbsp;<%=last_name%></a><br>				</td></tr>
			 				<%	rs.movenext										if not rs.eof then
						id = (rs("EmployeeID"))
						first_name = rs("FirstName")
						last_name = rs("LastName")
						ind = trim(rs("Industry_Name"))
					else						prev_ind=""
					end if			   wend%>	
						
			</table>
		</span>		
	</td></tr>
	

<%wend
	
  rs.close%>
</td>
</tr></table>
<td>&nbsp;&nbsp;</td>

<td ALIGN=LEFT VALIGN=top>

<!--''''''''''''''''''''''''''''''''''''''''''''''
'Code to display a list of the Certifications 
'sorted alphabetically by Cert by consultant.
''''''''''''''''''''''''''''''''''''''''''''''-->
<%	
sql = "SELECT Employee.EmployeeID, Employee.FirstName, Employee.LastName, Certifications.Cert_Name" _
      & " FROM Employee_Cert_Xref INNER JOIN" _
      & " Certifications ON" _ 
      & " Employee_Cert_Xref.Cert_ID = Certifications.Cert_ID" _
      & " INNER JOIN" _
      & " Employee ON" _ 
      & " Employee_Cert_Xref.Employee_ID = Employee.EmployeeID" _
      & " ORDER BY Certifications.Cert_Name, Employee.LastName" 
set rs = DBQuery(sql)

prev_ind = ""%>
<table class=tableShadow width=225 cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc>
<tr><td><font size=2.5 face= "ms sans serif, arial, geneva" color = navy>
	<b>Certifications</b>
</td></tr>
<tr><td colspan=2 bgcolor=gray height=1>
</td></tr><br>
<%while not rs.eof
	
	id = (rs("EmployeeID"))
	first_name = rs("FirstName")
	last_name = rs("LastName")
	ind = rs("Cert_Name")
		
	if prev_ind = "" OR prev_ind <> ind then%>
		<tr>
			<td><font size=1 face= "ms sans serif, arial, geneva"  color = navy>
			<b><span onClick="toggle('<%=ind%>')"  alt="Expand or collapse this section." style="cursor:hand"><%=ind%></b>
			<span id='<%=ind & "Header"%>' style="">(collapsed)</span>
	    <!--<b><%=ind%></b>-->
			</td>		</tr>
<%	end if
	prev_ind = ind%>		<tr><td colspan=2>
		<span id='<%=ind%>' style="display:none">			<table>									<% while prev_ind = ind%>
				<tr>
				<td valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=id%>'><%=first_name%>&nbsp;<%=last_name%></a><br>				</td></tr>
			 				<%	rs.movenext										if not rs.eof then
						id = (rs("EmployeeID"))
						first_name = rs("FirstName")
						last_name = rs("LastName")
						ind = trim(rs("Cert_Name"))
					else						prev_ind=""
					end if		   			   			   wend%>	
							
			</table>				</span>
	</td></tr>
<%
wend
	
rs.close%>
</td></tr>
</table>
	


<!-- #include file="../../footer.asp" -->
