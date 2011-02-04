<!--
Developer:		Adam Lewandowski
Date:			10/2/2000
Description:	Shows a report of exam categories, with drill-down to 
                individual exams and consultants that have passed the exam-->

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

.Employee {
 text-indent : 35pt;
 width : 50%;
 color : #0000AA;
 font-family : sans serif, arial, geneva;
 font-size: 10pt;
}
</STYLE>

<b>
<font size=3 face='ms sans serif, arial, geneva' color=black>
Microsoft Certified Professionals</font>
<p>
<font size=2 face='ms sans serif, arial, geneva' color=red>This report lists all Sarks that
 have obtained a Microsoft Certified Professional certification.</font></b>
<br>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
<TABLE WIDTH=100%>
<TR><TD>
<%
      ' Get list of employees that have passed this exam
      sSQL="SELECT distinct firstname, lastname, e.employeeID, cert_number " & _
           "FROM Employee e, Employee_Cert_Numbers ecn " & _
           "WHERE e.employeeid = ecn.employee_id " & _
           "AND ecn.category_id = 1" & _
           "ORDER BY lastname"
      set rsEmps = DBQuery(sSQL)
      Response.Write("<TABLE border=0 width=100% >")
      while not rsEmps.eof
       %><TR><TD align=left><span class="Employee"><%
       Response.Write("<A HREF='../../directory/content/details.asp?EmpID=" & rsEmps("employeeID") & "'>" & rsEmps("firstname") & "&nbsp;" & rsEmps("lastname") & "</A><BR>")
       %></span></TD><TD align=left><span class="Employee"><%=rsEmps("cert_number")%>
       </span></TD></TR><%
       rsEmps.movenext
      wend
     %></TABLE>
</TD></TR></TABLE>
		
<!-- #include file="../../footer.asp" -->
