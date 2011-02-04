<!--
Developer:    Mark Apgar
Date:         04/25/2000
Description:  Display a list of articles for selection divided by category.
-->

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
<font size=4 face="ms sans serif, arial, geneva"  color=black>
<b>Article Archive</b>
<p>
<font size=1 face="ms sans serif, arial, geneva"  color=blue>
<b>This page displays the article archive ordered by category by date.  Select a category to see the archived articles for that category.</b>

<TABLE BORDER=0 WIDTH=480>
<tr><td ALIGN=LEFT VALIGN=top>
<%
sql = "SELECT News.News_ID, News.Title, Employee.EmployeeID, News.CreationDate, Employee.FirstName, NewsCategories.Category_Title, Employee.LastName " & _
		"FROM News INNER JOIN Employee ON News.Employee_ID = Employee.EmployeeID " & _
		"INNER JOIN NewsCategories ON News.Category_ID = NewsCategories.Category_ID " & _
		"WHERE News.Description LIKE 'Article' AND News.Approved=1" & _
		"ORDER BY NewsCategories.Category_Title, News.CreationDate DESC"
set rs = DBQuery(sql)

prev_Cat = ""
%>
<table class=tableShadow width=450 cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc>
<tr><td><font size=2.5 face= "ms sans serif, arial, geneva"  color = navy>
	<b>Archive Categories</b>
</td></tr>
<tr><td colspan=2 bgcolor=gray height=1>
</td></tr><br>
<%while not rs.eof

	ArtID=rs("News_ID")
	ArtTitle=rs("Title")
	EEID=rs("EmployeeID")
	CreationDate=rs("CreationDate")
	first_name = rs("FirstName")
	last_name = rs("LastName")
	Cat = rs("Category_Title")

	if prev_Cat = "" OR prev_Cat <> Cat then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = navy>
		<b><span onClick="toggle('<%=Cat%>')" alt="Expand or collapse this section." style="cursor:hand"><%=Cat%></b>
		<span id='<%=Cat & "Header"%>' style=""> (collapsed)</span>
		</td>	</tr>
	<%end if%>	<%prev_Cat = Cat%>		<tr><td colspan=2>
		<span id='<%=Cat%>' style="display:none">			<table>									<% do while prev_Cat = Cat%>
				<tr>
				<td valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				<li><a href='article.asp?NewsID=<%=ArtID%>'><%=ArtTitle%>&nbsp; - <%=CreationDate%></a></li><br>
				&nbsp;&nbsp; - &nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=EEID%>'><%=first_name%>&nbsp;<%=last_name%></a><br>				</td></tr>
			 				<%	rs.movenext										if not rs.eof then
						Cat = rs("Category_Title")
						ArtTitle=rs("Title")
						CreationDate=rs("CreationDate")
						first_name = rs("FirstName")
						last_name = rs("LastName")						ArtID=rs("News_ID")						EEID=rs("EmployeeID")
					else						prev_Cat=""
						exit do
					end if			   loop%>	
						
			</table>
		</span>		
	</td></tr>

<%wend
	
  rs.close%>
</td>
</tr></table>

<!-- #include file="../../footer.asp" -->

