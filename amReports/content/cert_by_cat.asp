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


.Cert {
 width : 75%;
 color : blue;
 font-family : sans serif, arial, geneva;
 cursor : hand;
 font-size : 10pt;
}

.CertCount {
 color : black;
 font-family : sans serif, arial, geneva;
 font-size : 10pt;
}

.Employee {
 text-indent : 35pt;
 width : 50%;
 color : #0000AA;
 font-family : sans serif, arial, geneva;
}
</STYLE>

<script language=javascript>
// Expand/collapse a section
function toggle(item) {
 var docItem = document.all(item);
 if (docItem.style.display=="") {
  docItem.style.display = "none";
 } else {
  docItem.style.display="";
 }
}
</script>


<b>
<font size=3 face='ms sans serif, arial, geneva' color=black>
Certifications by Category</font>
<p>
<font size=2 face='ms sans serif, arial, geneva' color=red>This report lists the total number of each certifications held.  Click on a 
certification to view the Sarks that hold that certification.</font></b>
<br>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
<TABLE WIDTH=100%>
<TR><TD>
<%
 ' Get list of categories
 set rsCerts=DBQuery("SELECT cert_id, cert_name FROM Certifications ORDER BY cert_name")
 do while not rsCerts.eof 
  %>
   <%
   ' Get count of exams passed in each category
   sSQL="SELECT count(employee_cert_xref) total FROM Employee_Cert_Xref ecx WHERE " & _
        "cert_id=" & rsCerts("cert_id")
   set rsCertCount=DBQuery(sSQL)
   if rsCertCount("total") >0 then
    %>
    <span class="Cert" onClick="toggle('<%="Cert" & rsCerts("cert_id")%>')" alt="Expand or collapse this section">
    <%
    Response.Write("<B>" & rsCerts("cert_name") & "</B>")
    %>
    </span>
    <span class="CertCount">
    <%=rsCertCount("total")%> 
    </span>
    <BR>
    <span id="Cert<%=rsCerts("cert_id")%>" style="display:none">
     <%
      ' Get list of employees that have passed this exam
      sSQL="SELECT distinct firstname, lastname, e.employeeID " & _
           "FROM Employee e, Employee_Cert_Xref ecx " & _
           "WHERE e.employeeid = ecx.employee_id " & _
           "AND ecx.cert_id=" & rsCerts("cert_id") & " " &_
           "ORDER BY lastname"
      set rsEmps = DBQuery(sSQL)
      while not rsEmps.eof
       %><span class="Employee"><%
       Response.Write("<A HREF='../../directory/content/details.asp?EmpID=" & rsEmps("employeeID") & "'>" & rsEmps("firstname") & " " & rsEmps("lastname") & "</A><BR>")
       %></span><BR><%
       rsEmps.movenext
      wend
     %>
    </span>
   <% 
  end if
  rsCerts.movenext 
 loop 
%>
</TD></TR></TABLE>
		
<!-- #include file="../../footer.asp" -->
