<!--
Developer:    DEVELOPER
Date:         MM/DD/YYYY
Description:  About this section
-->


<!-- #include file="../section.asp" -->

<img src="../../common/images/hackanm.gif" height=48 width=48 hspace=10 vspace=5 align=left>
The following news articles and training articles have been submitted for approval and posting.  
To approve a posting select it from the list and click on the 'Approve' check box.
<p>Unapproved Postings...<br>&nbsp<br>

<%
dim count
sql = "SELECT News.News_ID, News.Title,Employee.EmployeeID, News.CreationDate, Employee.FirstName, Employee.LastName " & _
		"FROM News INNER JOIN Employee ON News.Employee_ID = Employee.EmployeeID " & _
		"WHERE News.Approved=0 AND News.Description='Article' " & _
		"ORDER BY News.CreationDate DESC"
set rs = DBQuery(sql)
Response.Write("<font size=1 face='ms sans serif, arial, geneva' color=red><b>News Article Submissions:</b></font>")
Response.Write("<ul>")
while count < 10 and not rs.eof 
	response.write("<li>" & chr(34) & "<a href='../../news/content/update.asp?News_ID=" & rs("News_ID") & "' onMouseOver=" & chr(34) & "statmsg('View article.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("Title") & "</a>" & chr(34) & "<br>")
	response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;by <a href='../../directory/content/details.asp?Empid=" & rs("EmployeeID") & "' onMouseOver=" & chr(34) & "statmsg('View author.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("FirstName") & " " & rs("LastName") & "</a>&nbsp;&nbsp;(" & rs("CreationDate") & ")<p>" & NL)
	rs.movenext
	count = count + 1
wend
Response.Write("</ul>")
Response.Write("<br>")
rs.close

count=0
sql = "SELECT News.News_ID, News.Title,Employee.EmployeeID, News.CreationDate, Employee.FirstName, Employee.LastName " & _
		"FROM News INNER JOIN Employee ON News.Employee_ID = Employee.EmployeeID " & _
		"WHERE News.Approved=0 AND News.Description='Training' " & _
		"ORDER BY News.CreationDate DESC"
set rs = DBQuery(sql)
Response.Write("<font size=1 face='ms sans serif, arial, geneva' color=red><b>Training Activity Submissions:</b></font>")
Response.Write("<ul>")
while count < 10 and not rs.eof 
	response.write("<li>" & chr(34) & "<a href='../../training/content/update.asp?News_ID=" & rs("News_ID") & "' onMouseOver=" & chr(34) & "statmsg('View article.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("Title") & "</a>" & chr(34) & "<br>")
	response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;by <a href='../../directory/content/details.asp?Empid=" & rs("EmployeeID") & "' onMouseOver=" & chr(34) & "statmsg('View author.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("FirstName") & " " & rs("LastName") & "</a>&nbsp;&nbsp;(" & rs("CreationDate") & ")<p>" & NL)
	rs.movenext
	count = count + 1
wend
Response.Write("</ul>")
Response.Write("<br>&nbsp")
rs.close

%>
</ul>

<!-- #include file="../../footer.asp" -->


