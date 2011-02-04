<%
if Application("dateArticle") <> Date then
	graphic = "../../common/images/spacer.gif"
	if Application("debug") then response.write(NL & "<!-- Build article text and store in Application obj -->" & NL)
	
	' GET CURRENT ARTICLE FROM DATABASE
	sql = "SELECT e.EmployeeID, n.News_ID, n.Title, n.CreationDate, e.FirstName, e.LastName FROM News n, Employee e WHERE n.Employee_ID = e.EmployeeID And n.Approved=1 And n.Description LIKE 'Article' ORDER BY n.CreationDate desc"
	set rs = DBQuery(sql)
	htmlArticle = chr(34) & "<a href='../../news/content/article.asp?NewsID=" & rs("News_ID") & "' onMouseOver=" & chr(34) & "statmsg('View article.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("Title") & "</a>" & chr(34) & " by <a href='../../directory/content/details.asp?Empid=" & rs("EmployeeID") & "' onMouseOver=" & chr(34) & "statmsg('View author.'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & rs("FirstName") & " " & rs("LastName") & "</a>&nbsp;&nbsp;(" & rs("CreationDate") & ")<br>" & NL
	
	' CHECK CREATION DATE
	CreationDate = rs("CreationDate")
	if (Month(CreationDate) = Month(Date)) and (Year(CreationDate) = Year(Date)) then
		graphic = "../../common/images/tiny/new2.gif"
		end if
	
	' GET ARTICLE FILE FROM SERVER
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	path = Server.MapPath("../../news/content/articles/") & "\article" & rs("News_ID") & ".txt"
	set ts = fso.OpenTextFile(path)
	line = ts.ReadLine
	if left(ucase(line),3) = "<P>" then line = mid(line, 4)
	if line = "" then line = ts.ReadLine
	htmlArticle = htmlArticle & line & " " & ts.ReadLine & "..."
	
	' CLEANUP
	ts.close
	rs.close
	
	' SET VALUES IN APPLICATION OBJECT
	Application.Lock
	Application("dateArticle") = Date
	Application("htmlArticle") = htmlArticle
	Application("iconArticle") = graphic
	Application.Unlock
	end if
%>

<table width="100%" border=0 cellpadding=0 cellspacing=0>
<tr>
	<td width=28 valign=top><img src="<%=Application("iconArticle")%>" height=9 width=25></td>
	<td><font size=1 face="ms sans serif, arial, geneva"><%=Application("htmlArticle")%></font></td>
</tr>
</table>
