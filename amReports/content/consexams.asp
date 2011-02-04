<!--
Developer:    DMARTIN
Date:         9/2000

			
-->
<%if (hasRole("AccountManager") or hasRole("Sales")) then %>
<!-- #include file="../section.asp" -->

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
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Exams Listed by Consultants</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>

<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''-->
<%

'sql="SELECT Employee.EmployeeID, Employee.FirstName, Exam_list.Exam_id, Exam_list.cert, " &_
 '   "Employee.EmployeeTitle_ID, " &_
  '  "Employee.LastName, Exams_passed.Exam_date, Exam_List.Exam_Name, Exam_List.cert " &_
   ' "FROM Employee LEFT OUTER JOIN Exams_passed INNER JOIN Exam_List ON " &_
   ' "Exams_passed.Exam_ID = Exam_List.Exam_ID ON " &_
   ' "Employee.EmployeeID = Exams_passed.EmployeeID " &_
   ' "AND Exam_list.cert = 0 " &_
   ' "AND Employee.EmployeeTitle_ID in (1, 2, 3, 4, 5, 6, 7, 8, 23) " &_
   ' "ORDER BY Employee.LastName, Employee.FirstName"
      
sql="SELECT Employee.EmployeeID, Employee.FirstName, Exam_list.Exam_id, Exam_list.cert, " &_ 
    "Employee.EmployeeTitle_ID, " &_
    "Employee.LastName, Exams_passed.Exam_date, Exam_List.Exam_Name, Exam_List.cert " &_ 
    "FROM Employee LEFT OUTER JOIN Exams_passed INNER JOIN Exam_List ON " &_
    "Exams_passed.Exam_ID = Exam_List.Exam_ID ON " &_ 
    "Employee.EmployeeID = Exams_passed.EmployeeID WHERE " &_
    "Employee.EmployeeTitle_ID in (1, 2, 3, 7, 8, 23) " &_
  
    "ORDER BY Employee.LastName, Employee.FirstName"
      
      '  "AND Exam_List.cert = 0 " &_
      
      
      'Response.Write "(" & sql & ")"
      
set rs = DBQuery(sql)

name =""
%>
<table width=475 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>

<%while not rs.eof
	
	id = (rs("EmployeeID"))
	examid = rs("exam_id")
	full_name = rs("LastName")&", "& rs("FirstName")
	exam = rs("exam_name")
	passed = rs("exam_date")
	ls_cert = rs("cert")
	
	name = full_name%>
	
	
	<% if name = "" OR name <> lastname then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = blue>
		<b><span onClick="toggle('<%=full_name%>')" alt="Expand or collapse this section." style="cursor:hand"><%=full_name%></b>
		<span id='<%=full_name & "Header"%>' style="">(collapsed)</span>
		<!--<b><%=title%></b>-->
		</td>	</tr>	
	<%end if%>		<tr><td colspan=2>
		<span id='<%=full_name%>' style="display:none">			<table>				<%				Response.write("<table><tr><font size=1 face='ms sans serif, arial, geneva'>")
				Response.write("<td colspan=1 width=400 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
				Response.Write("<b>Exams Passed:</b>")
				Response.write("<td colspan=1 width=50 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
				Response.Write("<b>Date Passed: </b>")%>					<% while name = full_name%>				<tr>
				<% if ls_cert="False"  then%>
				 <td colspan=1 width=400 valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				 <%=exam%></a><br>
								 <td colspan=1 width=50 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'><%=passed%>				 </td></tr>			    <% end if%>
			    				<%	rs.movenext										if not rs.eof then
						id = (rs("EmployeeID"))
						examid = rs("exam_id")
						full_name = rs("LastName") & ", " & rs("FirstName")
						exam = rs("exam_name")						passed = rs("exam_date")						ls_cert = rs("cert")					else						name=""
					end if
							   wend%>	
						
			</table>
		</span>		
	</td></tr>
	

<%wend
rs.close%>
</td>
</tr></table>
</tr>
</table>
	
<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager or sales representative,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>
<%end if%>

