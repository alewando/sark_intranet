<!--
Developer:    Barry Behrmann
Date:         12/20/1999
Description:  Submit new news articles
History:      Redeveloped by Drew J. Marston on 4/13/00
			  Redeveloped by Mark Apgar on 4/25/00 - Category field added
-->


<!-- #include file="../section.asp" -->

<script language=javascript>
<!--

	function DeleteArticle(){
	if (confirm("Are you sure you want to delete this article?")){
		document.frmInfo.DelFlag.value = "True"
		document.frmInfo.submit()
		}
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
	
	if rs("Category_ID")<>"" then
		CatID=rs("Category_ID")
	else
		CatID="8"  'Default value of miscellaneous if nothing is there
	end if
	
	CategoryID=Request.Form("Category_ID")
	
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
	filename = Server.MapPath("articles/") &"\article" & UpdateArticleNumber & ".txt"
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
			If hasRole("WebMaster") then 
				ApprovedStr = ", Approved='" & IsApproved & "'"
			End If
			
			'-------------------------------------------------------------------------
			'Update DB with applicable form info.
			'-------------------------------------------------------------------------			strTitle = Request.Form.Item("Title")			strTitle = Replace(strTitle, "'", "''")
			ls_sql = "UPDATE News Set Title='" & strTitle & "'" & ApprovedStr & ", Category_ID='" & CategoryID & "'" _
					& " WHERE News_ID=" & UpdateArticleNumber 
			set rs = DBQuery(ls_sql)
			
			'-------------------------------------------------------------------------
			'Update the article text file on the server with form info.
			'-------------------------------------------------------------------------
			filename = Server.MapPath("articles/") & "\article" & UpdateArticleNumber & ".txt"
			Set a = fso.CreateTextFile(filename, true)
			a.WriteLine(Request.Form.Item("Body"))
			a.close
									
			Response.write("<img src='../../common/images/checker.gif' height=32 width=32 border=0 align=left>")
			Response.Write("<table border=0 cellpadding=10>")
			Response.Write("<tr><td align='center'><font size=1 face='ms sans serif, arial, geneva' color=blue>")
			Response.Write("<b>Article has been updated.</b></font></td></tr>")
			Response.Write("<tr><td><input type=button class=button value='  Back to Approve New Postings  ' OnClick='window.location=" & chr(34) & "../../tools/content/approve_postings.asp" & chr(34) & "' id=button1 name=button1>")
			Response.Write("<input type=button class=button value='  Back to News  ' OnClick='window.location=" & chr(34) & "../../News/content/about.asp" & chr(34) & "' id=button1 name=button1></td></tr></table>")
			
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
		Response.Write("<B>Submitted By: <font size=1 face='ms sans serif, arial, geneva' color=blue><a href='../../directory/content/details.asp?EmpID=" & EmployeeID & "'>" & CreatedBy & "</a> on ")
		Response.Write(CreationDate & "</font></b>")
		Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
		Response.Write("<b>Submission Type: <font size=1 face='ms sans serif, arial, geneva' color=blue>")
		Response.Write(NewsType & "</font></b><BR><BR>" & vbCRLF)
			
		'-------------------------------------------------------------------------
		'If the user is the WebMaster then show the approved checkbox.
		'-------------------------------------------------------------------------
		if hasRole("WebMaster") then
			Response.Write("<B>Approved? </B><INPUT type=checkbox id=approved name=approved value=1" & CheckedValue & ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp")' & vbCRLF & "<BR>")
		End If
		
		if CatID <> "" then
			ls_Selected="Select Category_Title from NewsCategories where Category_ID='" & CatID & "'"
			set SelRs=dbQuery(ls_Selected)
			CatName=SelRs("Category_Title")
			SelRs.close
		end if
		
		ls_Cat="Select * from NewsCategories"
		set CatRs=DBQuery(ls_Cat)
		
		Response.Write("<B>Categories: </B><Select NAME='Category_ID'>")
		if not CatRs.eof and not CatRs.bof then
			Response.Write("<Option value=" & chr(34) & chr(34) & "></option>") 
			do while not CatRs.EOF
				Response.Write("<Option Value='" & CatRs("Category_ID") & "'>" & CatRs("Category_Title") & "</Option>")
				CatRs.movenext
			loop
			CatRs.close
			Response.Write ("<Option Selected Value='" & CatID & "'>" & CatName & "</Option>")
		end if
		
		'We want to display articles as they will be seen first so DoEdit is set to false on calling
		'page. We set it to true and refresh the page if we click the edit button, so it will be displayed
		'in the text box to be edited.
		If Not Request("DoEdit") then
		   Response.Write("</Select><BR><BR>")
		   Response.Write("<B>Title:</B> <INPUT size=54 maxlength=255 type=text id=title name=title value=" & chr(34) & TitleValue & chr(34) & "'><BR><BR>" & vbCRLF)
		   Response.Write("<B>Article Body:</B><BR>") 
		   Response.Write(BodyValue)
		   Response.Write("<textarea name=body rows=15 cols=50 wrap=off style='display:none;'>" & BodyValue & "</textarea>")
		   Response.Write("<INPUT type=hidden id=News_ID name=News_ID value='" & Request.QueryString("News_ID") & "'>")
		   Response.Write("<CENTER><BR>")
		   %>
		   <INPUT type=submit class=button value='Save'>&nbsp
		   <input type=button class=button value="  Edit  " onClick="window.location.href='update.asp?News_ID=<%=Request.QueryString("News_ID")%>&DoEdit=true'">
		   <%
		   Response.Write("<INPUT type=button class=button value='Delete ' onclick='DeleteArticle();'>&nbsp&nbsp")
		Else
		   Response.Write("</Select><BR><BR>")
		   Response.Write("<B>Title:</B> <INPUT size=54 maxlength=255 type=text id=title name=title value=" & chr(34) & TitleValue & chr(34) & "'><BR><BR>" & vbCRLF)
		   Response.Write("<B>Article Body:</B> (Include HTML tags for formatting.)<BR><textarea name=body rows=15 cols=50 wrap=off>" & BodyValue & "</textarea><BR>")
		   Response.Write("<INPUT size=5 maxlength=5 type=hidden id=News_ID name=News_ID value='" & Request.QueryString("News_ID") & "'>")
		   Response.Write("<CENTER><BR><INPUT type=submit class=button value='Save'>&nbsp&nbsp")
		End If		
		
		if ApproveStat = False then
			Response.Write("<INPUT type=button class=button value=' Back ' onClick='javascript:window.history.back(1)'></CENTER>")
		else
			Response.Write("<INPUT type=button class=button value=' Back ' onClick='javascript:window.history.back(1)'></CENTER>")	
		end if

	end if
	%>
	
	</form>

<!-- #include file="../../footer.asp" -->

