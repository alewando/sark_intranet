<!--
Developer:    Michele Wallace & Mathangi Srinivasan
Date:         11/03/1998
Description:  Displays the current article or an archive that was selected
-->


<!-- #include file="../section.asp" -->


<%

'---------------------------
'  Retrieve the article...  
'---------------------------
dim id
If Request("NewsID") = "" Then
	sql = "Select n.News_ID, n.Title, n.CreationDate, e.EmployeeID, e.FirstName, e.LastName, e.Username From News n, Employee e Where n.Employee_ID = e.EmployeeID And n.Description LIKE 'Training' And n.Approved=1 Order By n.CreationDate desc"
Else
	sql = "Select n.News_ID, n.Title, n.CreationDate, e.EmployeeID, e.FirstName, e.LastName, e.Username From News n, Employee e Where n.Employee_ID = e.EmployeeID And n.Description LIKE 'Training' And n.Approved=1 And n.News_ID = " & request("NewsID")
End If

set rs = DBQuery(sql)
id = rs("News_ID")
Username = rs("Username")

response.write("<table border=0 width='100%'><tr><td><font size=1 face='ms sans serif, arial, geneva'>")
response.write("<b>" & chr(34) & rs("Title") & chr(34) & "</b><br>")
response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;by <a href='../../directory/content/details.asp?Empid=" & rs("EmployeeID") & "'>" & rs("FirstName") & " " & rs("LastName") & "</a>")
response.write("</font></td><td align=right valign=top nowrap><font size=1 color=navy face='ms sans serif, arial, geneva'>")
'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;" & rs("CreationDate"))
response.write("&nbsp;&nbsp;&nbsp;&nbsp;" & MonthName(DatePart("m", rs("CreationDate"))) & " " & DatePart("d", rs("CreationDate")) & ", " & DatePart("yyyy", rs("CreationDate")))
response.write("</font></td></tr></table>")
rs.close

'---------------------------
'  Open and read file...    
'---------------------------
set fso = Server.CreateObject("Scripting.FileSystemObject")
path = Server.MapPath("../../news/content/articles/") & "\article" & id & ".txt"
set ts = fso.OpenTextFile(path)
response.write(ts.ReadAll)
ts.close
Response.Write("<br><br><center>")
%>
<input type=button class=button name=btnCancel value="Back to Training" onClick="window.location='../../training/content/default.asp'">
<%
if hasRole("WebMaster") or Ucase(Username) = Ucase(Session("Username")) then
	Response.Write("<br><br><center>")
	Response.Write("<a href='../../training/content/update.asp?News_ID=" & id & "'><img src='../../common/images/btnUpdateArticle.jpg' border=0></a>")
	Response.Write("</center>")
End if
%>


<!-- #include file="../../footer.asp" -->

