<!--
Developer:    Michele Wallace & Mathangi Srinivasan
Date:         11/03/1998
Description:  Display a list of articles for selection.
-->


<!-- #include file="../section.asp" -->



Choose an article from the list...

<ul>
<%
sql = "SELECT News.News_ID, News.Title,Employee.EmployeeID, News.CreationDate, Employee.FirstName,Employee.LastName " & _
		"FROM News INNER JOIN Employee ON News.Employee_ID = Employee.EmployeeID " & _
		"WHERE News.Description LIKE 'Article' AND News.Approved=1" & _
		"ORDER BY News.CreationDate DESC"
set rs = DBQuery(sql)
while not rs.eof
	response.write("<li>" & chr(34) & "<a href='article.asp?NewsID=" & rs("News_ID") & "' onMouseOver=" & chr(34) & "statmsg('View article.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("Title") & "</a>" & chr(34) & "<br>")
	response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;by <a href='../../directory/content/details.asp?Empid=" & rs("EmployeeID") & "' onMouseOver=" & chr(34) & "statmsg('View author.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("FirstName") & " " & rs("LastName") & "</a>&nbsp;&nbsp;(" & rs("CreationDate") & ")<p>" & NL)
	rs.movenext
	wend
rs.close
%>
</ul>


<!-- #include file="../../footer.asp" -->

