<!--
Developer:    Barry Behrmann
Date:         12/20/1999
Description:  Submit new news articles
History:      Redeveloped by Drew J. Marston on 4/13/00
-->


<!-- #include file="../section.asp" -->

<script language=javascript>
<!--

	function DeleteArticle(){
	if (confirm("Are you sure you want to delete this article?")){
		document.frmInfo.DelFlag.value = "True" 
		document.frmInfo.submit()}
	}

// -->
</script>

	<form name=frmInfo action='update.asp' method=post>
	<INPUT Type=Hidden Name='DelFlag' Value=''>

<%

	TitleValue = ""
	BodyValue = ""

	'-------------------------------------------------------------------------
	'  Set Reference to FileSystemObject  
	'-------------------------------------------------------------------------
	Set fso = Server.CreateObject("Scripting.FileSystemObject")

	'-------------------------------------------------------------------------
	' Load page with selected News article
	'-------------------------------------------------------------------------
	UpdateArticleNumber = Request("News_ID")
	ls_sql = "Select * from News WHERE News_ID = '" & UpdateArticleNumber & "'"
	set rs = DBQuery(ls_sql)

	TitleValue = rs("Title")
	EmployeeID = rs("Employee_ID")	
	CreationDate = rs("CreationDate")
	NewsType = rs("Description")
	ApproveStat = rs("Approved")
	
	If ApproveStat then
		CheckedValue = " checked"
	Else
		CheckedValue = ""
	End if
	
	rs.Close
	
	'-------------------------------------------------------------------------
	'Obtain Employee's name from Employee table.
	'-------------------------------------------------------------------------
	ls_sql = "Select Username from Employee Where EmployeeID=" & EmployeeID
	set rs = DBQuery(ls_sql)
	CreatedBy = UCase(rs("Username"))
	rs.Close
	
	'-------------------------------------------------------------------------
	'Obtain text file containing the article from server.
	'-------------------------------------------------------------------------
	'filename = "d:\inetpub\intranet\news\content\articles\article" & NewArticleNumber & ".txt"
	filename = Server.MapPath("/intranet/news/content/articles") & "\article" & UpdateArticleNumber & ".txt"
	'set fso = Server.CreateObject("Scripting.FileSystemObject")
	set a = fso.OpenTextFile(filename)
	BodyValue = a.ReadAll
	a.close

	'-------------------------------------------------------------------------
	'Check to see if form is loaded with an article already.
	'-------------------------------------------------------------------------
	If Request.Form.Item("Title") <> "" then 
		If Request.Form("DelFlag") <> "True" then
			UpdateArticleNumber = Request.Form.Item("News_ID")
			IsApproved = Request.Form.Item("approved")
			ApprovedStr = ""
			'-------------------------------------------------------------------------
			'Check to see if user is WebMaster. If so, obtain Approval checkbox value.
			'-------------------------------------------------------------------------
			if hasRole("WebMaster") then
				ApprovedStr = ", Approved='" & IsApproved & "'"
			End If
			
			'-------------------------------------------------------------------------
			'Update DB with applicable form info.
			'-------------------------------------------------------------------------			strTitle = Request.Form.Item("Title")			strTitle = Replace(strTitle, "'", "''")
			ls_sql = "UPDATE News Set Title='" & strTitle & "'" & ApprovedStr _
					& " WHERE News_ID=" & UpdateArticleNumber 
			set rs = DBQuery(ls_sql)
			
			'-------------------------------------------------------------------------
			'Update the article text file on the server with form info.
			'-------------------------------------------------------------------------
			'filename = "d:\inetpub\intranet\news\content\articles\article" & NewArticleNumber & ".txt"
			filename = Server.MapPath("/intranet/news/content/articles") & "\article" & UpdateArticleNumber & ".txt"
			Set a = fso.CreateTextFile(filename, true)
			a.WriteLine(Request.Form.Item("Body"))
			a.close
									
			Response.write("<img src='../../common/images/checker.gif' height=32 width=32 border=0 align=left>")
			Response.Write("<table border=0 cellpadding=10>")
			Response.Write("<tr><td align='center'><font size=1 face='ms sans serif, arial, geneva' color=blue>")
			Response.Write("<b>Article has been updated.</b></font></td></tr>")
			Response.Write("<tr><td><input type=button class=button value='  OK  ' OnClick='window.location=" & chr(34) & "../../tools/content/approve_postings.asp" & chr(34) & "' id=button1 name=button1></td></tr></table>")
			
		else '(DelFlag <> "True")

			'-------------------------------------------------------------------------
			'Delete the article from the database.
			'-------------------------------------------------------------------------
			strSQL = "DELETE FROM News WHERE News_ID='" & UpdateArticleNumber & "'"
			DBQuery(strSQL)

			'-------------------------------------------------------------------------
			'Delete the article text file from the server.
			'-------------------------------------------------------------------------		
			fso.DeleteFile filename, true
			Response.write("<img src='../../common/images/checker.gif' height=32 width=32 border=0 align=left>")
			Response.Write("<table border=0 cellpadding=10>")
			Response.Write("<tr><td align='center'><font size=1 face='ms sans serif, arial, geneva' color=blue>")
			Response.Write("<b>Article has been deleted.</b></font></td></tr>")
			Response.Write("<tr><td><input type=button class=button value='  OK  ' OnClick='window.location=" & chr(34) & "../../tools/content/approve_postings.asp" & chr(34) & "' id=button1 name=button1></td></tr></table>")
		end if
		
	else
	
		'-------------------------------------------------------------------------
		'Load the page with the article details.
		'-------------------------------------------------------------------------
		Response.Write("<B>Submitted By: <font size=1 face='ms sans serif, arial, geneva' color=blue>" & CreatedBy & " on ")
		Response.Write(CreationDate & "</font></b>")
		Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
		Response.Write("<b>Submission Type: <font size=1 face='ms sans serif, arial, geneva' color=blue>")
		Response.Write(NewsType & "</font></b><BR><BR>" & vbCRLF)
			
		'-------------------------------------------------------------------------
		'If the user is the WebMaster then show the approved checkbox.
		'-------------------------------------------------------------------------
		if hasRole("WebMaster") then
			Response.Write("<B>Approved? </B><INPUT type=checkbox id=approved name=approved value=1" & CheckedValue & "><BR>" & vbCRLF & "<BR>")
		End If
	
	    If Not Request("DoEdit") Then
		   Response.Write("</Select><BR><BR>")
		   Response.Write("<B>Title:</B> <INPUT size=54 maxlength=255 type=text id=title name=title value=" & chr(34) & TitleValue & chr(34) & "'><BR><BR>" & vbCRLF)
		   Response.Write("<B>Article Body:</B><BR>") '<textarea name=body rows=15 cols=50 wrap=off>" & BodyValue & "</textarea><BR>")
		   Response.Write(BodyValue)
		   Response.Write("<textarea name=body rows=15 cols=50 wrap=off style='display:none;'>" & BodyValue & "</textarea>")		   
		   Response.Write("<INPUT type=hidden id=News_ID name=News_ID value='" & Request.QueryString("News_ID") & "'>")
		   Response.Write("<CENTER><BR>")
		   %>
		   <INPUT type=submit class=button value='Save' id=submit1 name=submit1>&nbsp		   
		   <input type=button class=button value="  Edit  " onClick="window.location.href='update.asp?News_ID=<%=Request.QueryString("News_ID")%>&DoEdit=true'" id=button1 name=button1>
		   <%
		   Response.Write("<INPUT type=button class=button value='Delete ' onclick='DeleteArticle();'>&nbsp&nbsp")
        Else		
		   Response.Write("<B>Title:</B> <INPUT size=54 maxlength=255 type=text id=title name=title value=" & chr(34) & TitleValue & chr(34) & "'><BR><BR>" & vbCRLF)
		   Response.Write("<B>Article Body:</B> (Include HTML tags for formatting.)<BR><textarea name=body rows=15 cols=50 wrap=off>" & BodyValue & "</textarea><BR>")
		   Response.Write("<INPUT size=5 maxlength=5 type=hidden id=News_ID name=News_ID value='" & Request.QueryString("News_ID") & "'>")
		   Response.Write("<CENTER><BR><INPUT type=submit class=button value='Update'>&nbsp&nbsp")
		   Response.Write("<INPUT type=button class=button value='Delete ' onclick='DeleteArticle();'>&nbsp&nbsp")
		End If
		
		if ApproveStat = False then
			Response.Write("<INPUT type=button class=button value=' Back ' onClick='javascript:window.history.back(1)' id=button2 name=button2></CENTER>")
		else
			Response.Write("<INPUT type=button class=button value=' Back ' onClick='javascript:window.history.back(1)' id=button2 name=button2></CENTER>")	
		end if

	end if
	%>
	
	</form>

<!-- #include file="../../footer.asp" -->

