<!--
Developer:    DEVELOPER
Date:         MM/DD/YYYY
Description:  About this section
-->


<!-- #include file="../section.asp" -->

<img src="images/paper_and_ink.jpg" height=82 width=64 border=0 align=left>
The following news articles have been submitted for approval.  To approve an article,
select it from the list and click on the 'Approve' check box.
<p>

Unapproved Articles...


<ul>
<%
dim count
sql = "SELECT News.News_ID, News.Title,Employee.EmployeeID, News.CreationDate, Employee.FirstName,Employee.LastName " & _
		"FROM News INNER JOIN Employee ON News.Employee_ID = Employee.EmployeeID " & _
		"WHERE News.Approved=0 AND News.Description LIKE 'Article'" & _
		"ORDER BY News.CreationDate DESC"
set rs = DBQuery(sql)
while count < 5 and not rs.eof 
	response.write("<li>" & chr(34) & "<a href='update.asp?News_ID=" & rs("News_ID") & "' onMouseOver=" & chr(34) & "statmsg('View article.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("Title") & "</a>" & chr(34) & "<br>")
	response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;by <a href='../../directory/content/details.asp?Empid=" & rs("EmployeeID") & "' onMouseOver=" & chr(34) & "statmsg('View author.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("FirstName") & " " & rs("LastName") & "</a>&nbsp;&nbsp;(" & rs("CreationDate") & ")<p>" & NL)
	rs.movenext
	count = count + 1
	wend
rs.close
%>
</ul>

<!-- #include file="../../footer.asp" -->


