<!--
Developer:    Steven Smith
Date:         11/25/1998
Description:  Include file for technology section to allow update of tech links
-->


<script language=javascript>
<!--
function AddLink(){
	winTech=window.open("links_add.asp?empname=<%=Session("Username")%>&TechName=<%=Server.URLEncode(curTitle)%>&Recipients=<%=Recipients%>&Tech_ID=<%=techid%>", "AddLink","height=250,width=580,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")
	}
function EditLink(tech_links_id){
	winTech=window.open('links_edit.asp?empname=<%=Session("Username")%>&TechName=<%=Server.URLEncode(curTitle)%>&Tech_ID=<%=techid%>&Recipients=<%=Recipients%>&Link_ID=' + tech_links_id, 'EditLink','height=250,width=580,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1')
	}
// -->
</script>


<table width="100%" cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc class=tableShadow>
	<tr>
		<td><font size=1 face="ms sans serif, arial, geneva" color=navy><b>Web Links</b></font></td>
		<td align=right><font size=1 face="ms sans serif, arial, geneva" color=black>
			[<a href="javascript:AddLink();" onMouseOver="top.status='Add link.'; return true;" onMouseOut="top.status=''; return true;">Add</a>]
		</FONT></td>
	</tr>
	<tr><td colspan=2 bgcolor=gray height=1></td></tr>

<%
	'-----------------------------------
	'  DISPLAY LINKS FOR THIS TECHNOLOGY
	'-----------------------------------
	set rs = DBQuery("select url, tech_links_id, active, num_hits, text, url_title, Employee_ID, username, LastModified, timestamp from tech_links where Tech_ID = " & techid & " AND active=1 order by num_hits DESC")
	if rs.eof then
		response.write("<tr><td colspan=2><font size=1 color=gray face='ms sans serif, arial, geneva'>[None]</td></tr>")
	else
		while not rs.eof
			if isnull(rs("Employee_ID")) then ShowGuest = true
%>
	<tr>
		<td><font size=1 face="ms sans serif, arial, geneva" color=black>
			<a href="/intranet/technology/content/hitcount.asp?url=<%=rs("url")%>&LinkID=<%=rs("tech_links_id")%>" target="Links"
			 onMouseOver="top.status='<%=Server.HTMLEncode(rs("url"))%>';return true;"
			 onMouseOut="top.status=''; return true;">
			<b><%=Server.HTMLEncode(rs("url_title"))%></b></a>
<%
			if not isnull(rs("timestamp")) then
				DiffDate = DateDiff("d", rs("timestamp"), Date)
				select case DiffDate
					case 0: daysOld = "Added today!"
					case 1: daysOld = "Added yesterday!"
					case else: daysOld = "Added " & DiffDate & " days ago."
				end select
				if DiffDate < 7 then
					response.write("<img src='../../common/images/tiny/new.gif' alt='" & daysOld & "' height=11 width=28 border=0>")
					end if
				end if
%>
			<br><%=Server.HTMLEncode(rs("text"))%><br>
			<font color=navy>&nbsp;&nbsp;-&nbsp;&nbsp;<%=rs("num_hits")%> hits since <%=FormatDateTime(rs("timestamp"),2)%>,</font><br>
			<font color=navy>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
			if (isnull(rs("LastModified"))) then
				response.write("created by " & lcase(rs("username")))
				if isnull(rs("Employee_ID")) then response.write("<b>*</b>")
				response.write(".")
			else
				response.write("last modified on " & FormatDateTime(rs("LastModified"),2) & " by " & lcase(rs("username")) & ".")
			end if
%>
		</font></td>

<%
			If fIsExpert Then
%>
		<td align=right valign=top nowrap><font size=1 face="ms sans serif, arial, geneva" color=black>
		[<a href="javascript: EditLink(<%=rs("tech_links_id")%>)" onMouseOver="top.status='Edit or delete link.';return true;" onMouseOut="top.status='';return true;">Edit</a>]
		</font></td>
<%
			Else
%>
		<td>&nbsp;</td>
<%
			End If
%>
	</tr>
<%			
			rs.movenext
			wend
		end if
	rs.close
	set rs = nothing

	if ShowGuest then%>
	<tr><td><font size=1 face="ms sans serif, arial, geneva" color=black>
		<sup><b>*</b></sup>Guest account
		</font></td><td></td></tr>
	<%end if%>
</table>
