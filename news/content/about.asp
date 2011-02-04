<!--
Developer:    DEVELOPER
Date:         MM/DD/YYYY
Description:  About this section
-->


<!-- #include file="../section.asp" -->

<img src="images/paper_and_ink.jpg" height=82 width=64 border=0 align=left>
Check out this section to learn more about what is going on in the 
Cincinnati branch.
<p>
Here you can read Current Articles about things that are happening in the branch, 
or if you missed an article, be sure to check out the Sark <a href="article_archive.asp">Archives</a>.


<p>

Current Articles...

<ul>
<%
dim count
sql = "SELECT News.News_ID, News.Title,Employee.EmployeeID, News.CreationDate, Employee.FirstName,Employee.LastName " & _
		"FROM News INNER JOIN Employee ON News.Employee_ID = Employee.EmployeeID " & _
		"WHERE News.Description LIKE 'Article' AND News.Approved=1 " & _
		"ORDER BY News.CreationDate DESC"
set rs = DBQuery(sql)
while count < 5 and not rs.eof 
	response.write("<li>" & chr(34) & "<a href='article.asp?NewsID=" & rs("News_ID") & "' onMouseOver=" & chr(34) & "statmsg('View article.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("Title") & "</a>" & chr(34) & "<br>")
	response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;by <a href='../../directory/content/details.asp?Empid=" & rs("EmployeeID") & "' onMouseOver=" & chr(34) & "statmsg('View author.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("FirstName") & " " & rs("LastName") & "</a>&nbsp;&nbsp;(" & rs("CreationDate") & ")<p>" & NL)
	rs.movenext
	count = count + 1
	wend
rs.close
%>
</ul>
<p>
<br>
...For a complete listing of all current and past articles check out the Sark <a href="article_archive.asp">Archives</a>.
<p>
If you would like to post an article about News in the Cincinnati branch go to the
<a href='../../news/content/submit.asp'>Submit</a> page.

<!-- #include file="../../footer.asp" -->


