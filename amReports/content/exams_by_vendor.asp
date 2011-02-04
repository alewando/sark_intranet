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


.Category {
 width : 50%;
 color : blue;
 font-family : sans serif, arial, geneva;
 cursor : hand;
 font-size : 10pt;
}

.CategoryCount {
 width : 50%;
 color : black;
 font-family : sans serif, arial, geneva;
 font-size : 10pt;
}

.Exam {
 text-indent:25pt;
 width : 80%;
 color : black;
 font-family : sans serif, arial, geneva;
 cursor : hand;
 font-size : 10pt;
}

.ExamCount {
 color : black;
 font-family : sans serif, arial, geneva;
 font-size : 10pt;
}

.Employee {
 text-indent : 27pt; 
 font-size : 10pt;
 width : 100%;
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


<font size=3 face='ms sans serif, arial, geneva' color=black><b>Exams by Vendor</font><p>
<font size=2 face='ms sans serif, arial, geneva' color=red>This report lists the total number of exams passed by vendor.  Click on the vendor
name to view the total number of specific exams passed.  Then click on an exam to see who has passed that specific exam.</font><br></b>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
<TABLE WIDTH=100%>
<TR><TD>
<%
 ' Get list of categories
 set rsCats=DBQuery("SELECT category_id, category_description FROM Exam_Categories ORDER BY Category_Description")
 while not rsCats.eof 
  %>
   <span class="Category" onClick="toggle('<%="Cat" & rsCats("category_id")%>')" alt="Expand or collapse this section">
   <%
   Response.Write("<B>" & rsCats("category_description") & "</B>")
   %>
   </span>
   <%
   ' Get count of exams passed in each category
   sSQL="SELECT count(exams_passed_id) total FROM Exams_Passed ep, Exam_List el WHERE " & _
        "el.exam_id = ep.exam_id AND el.category_id=" & rsCats("category_ID") & " " &_
        "AND el.cert=0"
   set rsCatCount=DBQuery(sSQL)
   %>
   <span class="CategoryCount">
   <%=rsCatCount("total")%> 
   </span>
   <BR>
    <span id="Cat<%=rsCats("category_id")%>" style="display:none">
    <%
     ' Get list of exams and passed count for each exam in this category
     sSQL="SELECT el.exam_name, el.exam_id, COUNT(exams_passed_id) exam_total " & _
          "FROM Exam_List el, exams_passed ep " &_
          "WHERE el.exam_id = ep.exam_id AND el.category_id = " & rsCats("category_id") & " " & _
          "AND el.cert=0 " &_
          "GROUP BY el.exam_name, el.exam_id " &_
          "ORDER BY el.exam_name"
     set rsExams=DBQuery(sSQL)
     while not rsExams.eof
      %>
       <span class="Exam" onClick="toggle('<%="Exam" & rsExams("exam_id")%>')" >
        <%=rsExams("exam_name")%> 
       </span>
    <span class="ExamCount"><%=rsExams("exam_total")%></span><BR>
    <span id="<%="Exam" & rsExams("exam_id")%>" style="display:none;">
       <!-- List of employees that have passed the exam -->
       <%
        ' Get list of employees that have passed this exam
        sSQL="SELECT distinct firstname, lastname, e.employeeID, exam_date " & _
             "FROM Employee e, Exams_Passed ep " & _
             "WHERE e.employeeid = ep.employeeid " & _
             "AND exam_id=" & rsExams("exam_id") & " " &_
             "ORDER BY lastname"
        set rsEmps = DBQuery(sSQL)
        while not rsEmps.eof
         %><span class="Employee" width=100%>
          <span style='width:75%; '>
         <%
         Response.Write("<A HREF='../../directory/content/details.asp?EmpID=" & rsEmps("employeeID") & "'>" & rsEmps("firstname") & " " & rsEmps("lastname") & "</A>")
         Response.Write("</span><span >")
         Response.Write(rsEmps("exam_date") & "<BR>")
         Response.Write("</span>")
         %></span><BR><%
         rsEmps.movenext
        wend
       %>
    </span>
      <%
      rsExams.movenext
     wend
    %>
 </span>
  <% 
  rsCats.movenext 
 wend
%>
</TD></TR></TABLE>
		
<!-- #include file="../../footer.asp" -->
