<!--
Developer:    Shannon Smith
Date:         07/11/2000
Description:  Update Announcements
-->


<!-- #include file="../section.asp" -->

<form name=frmInfo action='updateannouncement.asp' method=post>

<%

	BodyValue = ""

	'-------------------------------------------------------------------------
	'  Set Reference to FileSystemObject  
	'-------------------------------------------------------------------------
	Set fso = Server.CreateObject("Scripting.FileSystemObject")

	'-------------------------------------------------------------------------
	' Load page with selected News article
	'-------------------------------------------------------------------------
	UpdateAnnouncementNumber = Request("Announcement_ID")
	'-------------------------------------------------------------------------
	'Obtain text file containing the article from server.
	'-------------------------------------------------------------------------
	
	filename = Server.MapPath("../../welcome/announcements/") & "\announcements" & UpdateAnnouncementNumber & ".txt"
		
	'filename = Application("WebRootDir") & Application("web") & "welcome\announcements\announcements" & UpdateAnnouncementNumber & ".txt"
	'filename = "announcements/announcements" & UpdateAnnouncementNumber & ".txt"
	'Response.Write(filename & "<BR>")
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set a = fso.OpenTextFile(filename)
	BodyValue = a.ReadAll
	a.close

	'-------------------------------------------------------------------------
	'Check to see if form is loaded with an article already.
	'-------------------------------------------------------------------------
	If Request.Form.Item("Body") <> "" then 
			UpdateArticleNumber = Request.Form.Item("Announcement_ID")
						'-------------------------------------------------------------------------
			'Update the article text file on the server with form info.
			'-------------------------------------------------------------------------
			filename = Server.MapPath("../../welcome/announcements/") & "\announcements" & UpdateAnnouncementNumber & ".txt"
			'filename = Application("WebRootDir") & Application("web") & "welcome/announcements/announcements" & UpdateAnnouncementNumber & ".txt"
			Set a = fso.CreateTextFile(filename, true)
			a.WriteLine(Request.Form.Item("Body"))
			a.close
									
			Response.write("<img src='../../common/images/checker.gif' height=32 width=32 border=0 align=left>")
			Response.Write("<table border=0 cellpadding=10>")
			Response.Write("<tr><td align='center'><font size=1 face='ms sans serif, arial, geneva' color=blue>")
			Response.Write("<b>Announcement" & UpdateAnnouncementNumber & " has been updated.</b></font></td></tr>")
			Response.Write("<tr><td align='center'><input type=button class=button value='OK' OnClick='window.location=" & chr(34) & "../../welcome/content/default.asp" & chr(34) & "'></td></tr></table>")
			
		else
	
		'-------------------------------------------------------------------------
		'Load the page with the article details.
		'-------------------------------------------------------------------------
		Response.Write("<font size=2 color=red><b>EDIT ANNOUNCEMENTS</b></font><p>")
	
		Response.Write("<b>Title: <font size=1 face='ms sans serif, arial, geneva' color=blue>")
		Response.Write("Announcement" & UpdateAnnouncementNumber & "</font></b><BR><BR><BR>")
		
		Response.Write("<B>Announcement Body:</B> (Include HTML tags for formatting.)<BR><textarea name=body rows=15 cols=50 wrap=off>" & BodyValue & "</textarea><BR>")
		Response.Write("<INPUT size=5 maxlength=5 type=hidden id=Announcement_ID name=Announcement_ID value='" & Request.QueryString("Announcement_ID") & "'>")
		
		if UpdateAnnouncementNumber = 1 then
				Response.Write("<CENTER><BR><INPUT TYPE=button class=button VALUE = 'Next >' OnClick='window.location=" & chr(34) & "updateannouncement.asp?Announcement_ID=2" & chr(34)& "'  id=button1 name=button1>")
		else
				Response.Write("<CENTER><BR><INPUT TYPE=button class=button VALUE = '< Prev' OnClick='window.location=" & chr(34) & "updateannouncement.asp?Announcement_ID=1" & chr(34)& "'  id=button1 name=button1>")
		end if
		Response.Write("<INPUT type=submit class=button value='Update' id=submit1 name=submit1>")
		Response.Write("<INPUT type=button class=button value='Cancel' onClick='window.location=" & chr(34) & "../../welcome/content/default.asp" & chr(34) & "'></CENTER>")
		
	end if
	%>
	
	</form>
<!-- #include file="../../footer.asp" -->


