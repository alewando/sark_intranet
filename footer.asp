

<!-- END OF BODY-->
</font></td></tr></table>


<!-- FOOTER -->
<hr size=1><font size=1 face="ms sans serif, arial, geneva"><center>

<a href="../../welcome/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Welcome")%>'); return true;"
 onMouseOut="statmsg('')">Welcome</a> |
<a href="../../preferences/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Preferences")%>'); return true;"
 onMouseOut="statmsg('')">Preferences</a> |
<a href="../../directory/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Directory")%>'); return true;"
 onMouseOut="statmsg('')">Directory</a> |
<a href="../../events/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Events")%>'); return true;"
 onMouseOut="statmsg('')">Events</a> |
<a href="../../news/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_News")%>'); return true;"
 onMouseOut="statmsg('')">News</a> 
 <% if not Session("isGuest") then %>
 |
<a href="../../email/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Email")%>'); return true;"
 onMouseOut="statmsg('')">Email</a>
 <% end if %>
<%
if hasRole("SolutionServices") or hasRole("WebMaster") then
%>
 &nbsp;|
<a href="../../solutionservices/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Repository")%>'); return true;"
 onMouseOut="statmsg('')">Repository</a>
<%end if%>
<br>
<a href="../../technology/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Technology")%>'); return true;"
 onMouseOut="statmsg('')">Technology</a> |
<a href="../../training/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Training")%>'); return true;"
 onMouseOut="statmsg('')">Training & Certifications</a> |
<a href="../../office/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Office")%>'); return true;"
 onMouseOut="statmsg('')">Office</a> |
<a href="../../recruiting/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Recruiting")%>'); return true;"
 onMouseOut="statmsg('')">Recruiting</a> |
<a href="../../projects/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_Projects")%>'); return true;"
 onMouseOut="statmsg('')">Projects</a> |
<a href="../../sports/content/default.asp"
 onMouseOver="statmsg('<%=Application("LinkDesc_sports")%>'); return true;"
 onMouseOut="statmsg('')">Sports</a>
<%if Application("Web") = "intranetdev" then%>
| <a href="../../dev/content/default.asp"
 onMouseOver="statmsg('Check out our cool development tools!'); return true;"
 onMouseOut="statmsg('')">Tools</a>
<%end if%>

<p><font color=gray>--- Software Architects confidential and proprietary ---</font><br>
<font size=5 color=white>CONFIDENTIAL</font>

</center></font>


</td></tr></table></td>


<!-- COLUMN SPACER -->
<td width=20><img src="/<%=Application("Web")%>/common/images/spacer.gif" height=1 width=20></td>


<!-- RIGHT MARGIN -->
<td width=140 valign=top nowrap><font size=1 color=white face="ms sans serif, arial, geneva">
<%=SectionBtns%>
</font></td>

</tr>

</table>


</body>
</html>


<script language=javascript>
<!--
function SendFeedback()
{
	if (curSection.substring(1,4) != "Wel") {
		winTech=window.open("../../email.asp?recipient=<%=sectionWebMaster%>&cc=<%=ccWebMaster%>&from=<%=session("Username")%>&name=<%=Server.URLEncode(session("name"))%>&subject=<%=Server.URLEncode("FEEDBACK: " & sectionTitle & " section ")%>&footer=<%=Server.URLEncode(UCase(Application("web")) & ": " & curTitle)%>", "SendEmail","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")
	}
	else {
		winTech=window.open("../../email.asp?recipient=<%=sectionWebMaster%>&cc=<%=ccWebMaster%>&from=<%=session("Username")%>&name=<%=Server.URLEncode(session("name"))%>&subject=<%=Server.URLEncode("FEEDBACK: " & mid(sectionTitle,1,instr(1,sectionTitle," <")) & "section ")%>&footer=<%=Server.URLEncode(UCase(Application("web")) & ": " & curTitle)%>", "SendEmail","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")
	}
}
// -->
</script>


<%
'--------------------------------
'  Cleanup database connections  
'--------------------------------
on error resume next
DataConn.close
set DataConn = nothing
if Application("debug") then Response.Write("<!-- END OF PAGE (" & round((now-startTime)*100000000)/100 & " ms) -->" & NL & NL)
%>
