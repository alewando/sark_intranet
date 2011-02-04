<!--
Developer:    Barry Behrmann
Date:         12/20/1999
Description:  Submit new news articles
-->


<!-- #include file="../section.asp" -->

<%

'SetWebMaster "cdolan"
TitleValue = ""
BodyValue = ""

'-------------------------------------
'  Get Reference to FileSystemObject  
'-------------------------------------
Set fso = Server.CreateObject("Scripting.FileSystemObject")


' ***************************************
' Create a new News article with the info submitted
' ***************************************
If Request.Form.Item("Title") <> "" then
	'---------------------------------------------
	'	Execute Database Query for Employee Info  
	'---------------------------------------------	strTitle = Request.Form.Item("Title")	strTitle = Replace(strTitle, "'", "''")
	ls_sql = "INSERT INTO News (Title, Employee_ID, CreationDate, Description) " _
         	& "VALUES ('" & strTitle & "', " & Session("ID") & ", '" & Date() & "', 'Article')"

	set rs = DBQuery(ls_sql)

	' *** DISCLAIMER: The following code is not optimal, but I couldn't get the MoveLast to work
	' properly.  If you can do so, be my guest, but please let me know how you did it.
	' Thanks.  bbehrmann@sark.com
		
	ls_sql = "SELECT * FROM News order by News_ID"
	
	Response.Write(vbNewLine & vbNewLine & "<!--" & ls_sql & "-->" & vbNewLine & vbNewLine)
	
	set rs = DBQuery(ls_sql)
		
	while not rs.eof
		NewArticleNumber = rs("News_ID")
		rs.MoveNext
	wend
'	rs.MoveLast
		
	' *** END DISCLAIMER	

'	filename = "d:\inetpub\intranet\news\content\articles\article" & NewArticleNumber & ".txt"
	filename = Server.MapPath("articles/") & "\article" & trim(NewArticleNumber) & ".txt"

	if fso.FileExists(filename) then
		Response.Write("ERROR: File already exists.  Contact System Administrator.")

		' Delete table entry
		ls_sql = "DELETE FROM News WHERE News_ID = '" & NewArticleNumber & "'"
		rs = DBQuery(ls_sql)		
		
	Else
		Set a = fso.CreateTextFile(filename, true)

		a.WriteLine(Request.Form.Item("Body"))
		a.close
							
		Response.Write("<BR>Your article has been submitted for review.<BR>")
		Response.Write("Once reviewed and approved it will be posted to the News section.<BR>")
		Response.Write("Please contact your Account Manager if you have any questions.<BR>")
		Response.Write("<BR><A Href='../content/default.asp'>Return to News</A>")
	
	End if
		
' ***************************************
' Get information for new News article from user.
' ***************************************
Else
    if not Session("isGuest") then
	 ' Enter article details
	 Response.Write("<form action='submit.asp' method='post'>")
	 Response.Write("<font size=2 face='geneva, arial, ms sans serif'>") 
 
 	 Response.Write("<B>Submitted By:</B><I> " & UCase(Session("Username")) & " on ")
 	 Response.Write(Now() & "</I><BR><BR>")
	 Response.Write("<B>Title:</B> <INPUT size=50 maxlength=255 type=text id=title name=title value='" & TitleValue & "'><BR><BR>")
	 'Response.Write("<font color=red>Please remove all quotes (single & double) from the title.</font><BR><BR>")
	 Response.Write("<B>Article Body:</B> (Include HTML tags for formatting.)<BR><textarea name=body rows=15 cols=50 wrap=off>" & BodyValue & "</textarea><BR>")

   %>
   <BR>

  <CENTER>
  <INPUT type="submit" class=button value="Submit" id=submit name=submit>
  <INPUT type=button class=button value="Cancel" onClick="window.location='about.asp'" id=button name=button>
  </font>
  </form>
  </CENTER>
 <% else %>
  <CENTER>Guest users cannot submit news articles</CENTER><BR>
 <% end if%>
<% End If %>


<!-- #include file="../../footer.asp" -->

