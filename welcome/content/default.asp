<!--
Developer:    Nick Yates
Date:         02/17/2000
Description:  Main welcome page first displayed to user.
-->


<!-- #include file="../section.asp" -->


<%
Response.Write vbNewLine & vbNewLine & "<!-- Your employee ID is '" & Session("ID") & "',-->" & vbNewLine
Response.Write "<!-- and you are logged in as '" & Session("Username") & "'.-->" & vbNewLine & vbNewLine
%>


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

<%


dim thisHour, thisAMPM

thisAMPM = "AM"
thisHour = hour(now)
if thisHour >= 12 then thisAMPM = "PM"
if not (thisHour < 13) then thisHour = thisHour - 12
	
'--------------------------------------------------------
'  If calendar (or other) hasn't been built today,       
'  build it and store it in the Application object.      
'  We assume that the web server will be stopped and     
'  restarted at least once per year.                     
'--------------------------------------------------------




'-------------------------------------------
'  Function to build individual event item  
'-------------------------------------------
function buildItem(link, textval, imageloc, imagealt)
	dim item
	if not isnull(imageloc) then item = "<img src=" & chr(34) & imageloc & chr(34) & " alt=" & chr(34) & imagealt & chr(34) & " height=16 width=16 border=0>"
	if not isnull(textval) then item = item & textval
	if not isnull(link) then item = "<a href=" & chr(34) & link & chr(34) & " onMouseOver=" & chr(34) & "window.status='View detail.'" & chr(34) & " onMouseOut=" & chr(34) & "window.status=''" & chr(34) & ">" & item & "</a>"
	buildItem = item
end function


'-------------------------------------------
'  Procedure to build annoucement item      
'-------------------------------------------
dim AnnouncementText

sub buildAnnouncement(announce_date, textval)
	dim dt, result
	
	result = ""
	if isdate(announce_date) then
		dt = DateValue(announce_date)
		if DateAdd("d", 45, dt) > Date then
			
			result = result & "		<tr><td colspan=2>" & NL
			result = result & "			<table width='100%' border=0 cellpadding=0 cellspacing=0>" & NL
			result = result & "			<tr><td colspan=2><img src='../../common/images/spacer.gif' height=10 width=5></td></tr>" & NL
			result = result & "			<tr>" & NL
			result = result & "				<td width=28 valign=top>"
			
			if DateAdd("d", 15, dt) > Date then
				result = result & 	"<img src='../../common/images/tiny/new2.gif' height=9 width=25>"
				end if
			
			result = result & "</td>" & NL
			result = result & "				<td><font size=1 face='ms sans serif, arial, geneva'>" & NL
			result = result & "				<u>" & announce_date & "</u> - " & textval & NL
			result = result & "				</font></td>" & NL
			result = result & "			</tr>" & NL
			result = result & "			</table>" & NL
			result = result & "			</td></tr>" & NL
			end if
		end if
	
	AnnouncementText = AnnouncementText & result
end sub



'sectionBtns = Application("WelcomeMargin")
sectionBtns = Session("WelcomeMargin")
if Application("tmpHour") <> hour(now) then
	
	'---------------------------------
	'  Build annoucement html code    
	'---------------------------------


	

	'---------------------------------
	'  Build calendar html code       
	'---------------------------------
	if Application("debug") then response.write(NL & "<!-- Build monthly calendar and store in Application obj -->" & NL)
	dim Calendar
	dim FirstDayOfMonth, CurDate, rdate
	rdate = Date
	'rdate = DateAdd("m",1,date)
	FirstDayOfMonth = DateAdd("d",-(DatePart("d", date)-1), rdate)
	CurDate = DateAdd("d",-(DatePart("w", FirstDayOfMonth)-1), FirstDayOfMonth)
	Calendar = Calendar & "<table class=plate width='100%' border=0 cellpadding=1 cellspacing=0 bgcolor=silver>"
	Calendar = Calendar & "<tr><td colspan=7 align=center><font size=1 face='ms sans serif, arial, geneva'>"
	Calendar = Calendar & "<a href='../../events/content/calendar.asp?date=" & date & "' onMouseOver=" & chr(34) & "statmsg('View calendar'); return true;" & chr(34) & " onMouseOut=" & chr(34) & "statmsg('')" & chr(34) & ">" & MonthName(Month(rdate)) & "</a>"
	Calendar = Calendar & "</font></td></tr>"
	do
		CurWeekday = DatePart("w", CurDate)
		if CurWeekday = 1 then Calendar = Calendar & "<tr>"
		Calendar = Calendar & "<td align=center"
		if CurDate = Date then Calendar = Calendar & " bgcolor=#eeeeee"
		Calendar = Calendar & "><font color="
		if DatePart("m",CurDate) <> Month(rdate) then
			Calendar = Calendar & "gray"
		else
			if CurWeekday = 1 or CurWeekday = 7 then
				Calendar = Calendar & "maroon"
			else
				Calendar = Calendar & "navy"
			end if
		end if
		Calendar = Calendar & " size=1>" & DatePart("d",CurDate) & "</font></td>"
		if CurWeekday = 7 then Calendar = Calendar & "</tr>"
		PrevDate = CurDate
		CurDate = DateAdd("d",1,CurDate)
	loop until ((DatePart("m",PrevDate) > DatePart("m",rdate) or DatePart("yyyy",PrevDate) > DatePart("yyyy",rdate)) and (DatePart("w",PrevDate) = 7))
	Calendar = Calendar & "</table>"
	
	
	'---------------------------------
	'  Build events html code         
	'---------------------------------
	Dim rs, sql, StartDate, EndDate, LoopDate, CalendarDate(7), htmlEvents, eventImage, EventText
	StartDate = Date
	EndDate = DateAdd("d", 6, StartDate)
	response.write("<!-- StartDate = " & Month(StartDate) & "/" & Day(StartDate) & "/" & Year(StartDate) & " -->" & NL)
	response.write("<!-- EndDate = " & Month(EndDate) & "/" & Day(EndDate) & "/" & Year(EndDate) & " -->" & NL)
	
	' Build Birthday list...
	sql =	"SELECT		EmployeeID, FirstName, LastName, Birthday " & _
			"FROM		Employee " & _
			"WHERE      (Birthday IS NOT NULL) AND (Branch_ID = 1) AND (DateTermination IS NULL) AND "
	
	' Modify WHERE clause if time period overlaps months...
	if Month(StartDate) = Month(EndDate) then
		sql = sql & "(DatePart(month, Birthday) = " & Month(StartDate) & " AND DatePart(day, Birthday) >= " & Day(StartDate) & " AND DatePart(day, Birthday) <= " & Day(EndDate) & ") "
	else
		sql = sql & "(DatePart(month, Birthday) = " & Month(StartDate) & " AND DatePart(day, Birthday) >= " & Day(StartDate) & ") OR (DatePart(month, Birthday) = " & Month(EndDate) & " AND DatePart(day, Birthday) <= " & Day(EndDate) & ") "
	end if
	
	' Take into account Jan -> Dec overlap!...
	if (Year(StartDate) <> Year(EndDate)) then
		sql = sql & "ORDER BY DatePart(month, Birthday) DESC, DatePart(day, Birthday), Lastname"
	else
		sql = sql & "ORDER BY DatePart(month, Birthday), DatePart(day, Birthday), Lastname"
	end if
	
	response.write("<!-- birthday sql = " & sql & " -->" & NL)
	set rs = DBQuery(sql)
	i = 1
	LoopDate = StartDate
	LoopCount = 0
	while (not rs.eof) and (i < 7)
		LoopCount = LoopCount + 1
		while (DatePart("d", LoopDate) <> DatePart("d", rs("Birthday")) and (i < 7))
			i = i + 1
			LoopDate = DateAdd("d", LoopDate, 1)
			wend
		CalendarDate(i) = CalendarDate(i) & buildItem("../../directory/content/details.asp?EmpID=" & rs("EmployeeID"),  " Birthday:&nbsp;&nbsp;" & rs("FirstName") & " " & rs("LastName"), "../../events/content/images/calendar_birthday.gif", "") & "<br>" & NL
		rs.movenext
		wend
	rs.close
	set rs = nothing
	response.write("<!-- records: " & LoopCount & " -->" & NL)
	
	' Build Event list
	sql =	"SELECT Sub_Type, ImageLoc, Event_ID, Title, StartTime, EndTime, EventDate " &_
		"FROM v_events_by_date " &_
		"WHERE "
	
	' Modify WHERE clause if time period overlaps months...
	if Month(StartDate) = Month(EndDate) then
		sql = sql & "(DatePart(year, EventDate) = " & Year(StartDate) & " AND DatePart(month, EventDate) = " & Month(StartDate) & " AND DatePart(day, EventDate) >= " & Day(StartDate) & " AND DatePart(day, EventDate) <= " & Day(EndDate) & ") "
	else
		sql = sql & "(DatePart(year, EventDate) = " & Year(StartDate) & " AND DatePart(month, EventDate) = " & Month(StartDate) & " AND DatePart(day, EventDate) >= " & Day(StartDate) & ") OR (DatePart(year, EventDate) = " & Year(EndDate) & " AND DatePart(month, EventDate) = " & Month(EndDate) & " AND DatePart(day, EventDate) <= " & Day(EndDate) & ") "
	end if
	
	' Take into account Jan -> Dec overlap!...
	if (Year(StartDate) <> Year(EndDate)) then
		sql = sql & "ORDER BY DATEPART(month, EventDate) DESC, DATEPART(day, EventDate), Type, Sub_Type"
	else
		sql = sql & "ORDER BY DATEPART(month, EventDate), DATEPART(day, EventDate), Type, Sub_Type"
	end if
	
	response.write("<!-- event sql = " & sql & " -->" & NL)
	set rs = DBQuery(sql)
	i = 1
	LoopDate = StartDate
	LoopCount = 0
	while (not rs.eof) and (i < 7)
		LoopCount = LoopCount + 1
		while (DatePart("d", LoopDate) <> DatePart("d", rs("EventDate")) and (i < 7))
			i = i + 1
			LoopDate = DateAdd("d", LoopDate, 1)
			wend
		if isnull(rs("ImageLoc")) then
			EventImage = "../../events/content/images/calendar_temp.gif"
		else
			EventImage = "../../events/content/" & rs("ImageLoc")
		end if
		EventText = " " & rs("Sub_Type") & ":&nbsp;&nbsp;" & rs("Title")
		if (rs("StartTime") <> " ") or (rs("EndTime") <> " ") then EventText = EventText & " (" & rs("StartTime") & " - " & rs("EndTime") & ")"
		CalendarDate(i) = CalendarDate(i) & buildItem("../../events/content/details.asp?id=" & rs("Event_ID") & "&lastpage=calendar&date=" & server.URLEncode(date),  EventText, EventImage, "") & "<br>" & NL
		rs.movenext
		wend
	rs.close
	set rs = nothing
	response.write("<!-- records: " & LoopCount & " -->" & NL)
	
	' Build HTML...
	LoopDate = StartDate
	for i = 1 to 7
		if (CalendarDate(i) <> "") then
			htmlEvents = htmlEvents & "<tr><td valign=top width=26 nowrap><font size=1 color="
			if (WeekDay(LoopDate) = WeekDay(Date)) then
				htmlEvents = htmlEvents & "red"
			else
				htmlEvents = htmlEvents & "navy"
				end if
			htmlEvents = htmlEvents & " face='ms sans serif, arial, geneva'>" &_
				Left(WeekDayName(WeekDay(LoopDate)), 3) & "</font></td>" &_
				"<td valign=top width='100%'><font size=1 face='ms sans serif, arial, geneva'>" &_
				CalendarDate(i) & "</font></td></tr>"
			end if
		LoopDate = DateAdd("d", LoopDate, 1)
		next
	htmlEvents = "<table border=0 width='100%' cellpadding=0 cellspacing=1>" & htmlEvents & "</table>"
		
	'-----------------------------------------
	'  Store html code in application object  
	'  for quick access and display           
	'-----------------------------------------
	Application.Lock
	Application("tmpHour") = hour(now)
	
'***********************************************************************************************
	Application("htmlAnnouncements") = session("sAnnouncements")
'***********************************************************************************************
	Application("htmlCalendar") = Calendar & "<br>"
	Application("htmlEvents") = htmlEvents
	Application.Unlock
end if
%>


<table width="470" border=0 cellpadding=0 cellspacing=0><tr>
	
	<td valign=top><table width="340" border=0 cellpadding=0 cellspacing=0>

<%
 '*********************************************************************************************
 '  Create instances of MAPI session	and login, storing session in the users ASP Session Object	
 '  This code is here to get the # of unread messages and display it to the user on the Welcome page			
 '*********************************************************************************************
 
 on error goto 0
 if Session("id") <> "" then
  dim objOMSession
  if Session("MAPIsession") is nothing then 
 	set objRenderApp = Application("RenderApplication")
	set objOMSession = Server.CreateObject("MAPI.Session")
	on error resume next
	
	objOMSession.Logon "", "", false, true, 0, false, Application("ExchangeServer") & vbLF & Session("Username")
	'Response.Write("Logon Err=" & Err & "<BR>")
	'Response.Write("objOMSession = " & objOMSession.Session & "<BR>")
	Session("UseMAPI") = false
	if Err = 0 then
		set Session("MAPIsession") = objOMSession
		if Err = 0 then
		 Session("hImp") = objRenderApp.ImpID
		 if Err =0 then
		  set Session("/Inbox") = objOMSession.Inbox
		  if Err = 0 then
		   set Session("/") = objOMSession.GetFolder(objOMSession.Inbox.FolderID)
		   if Err = 0 then
		    Session("UseMAPI") = true 
		    'Response.Write("USING MAPI<BR>")
		   end if
		  end if
		 end if
		end if
	else
		Session("UseMAPI") = false
		Response.Write("<!-- MAPI Logon Error: " & Err.Description & " -->" & vbCRLF)
	end if
  end if
	on error goto 0

 if Session("UseMAPI") then	
  '---------------------------------------
  '  Log on to MAPI session				
  '---------------------------------------

  set objOMSession = Session("MAPIsession")
  set objFolder = Session("/Inbox")
  set objRootFolder = Session("/")

  dim EmailUnreadCnt, objMessages
  if Session("id") <> "" then 
 	EmailUnreadCnt = 0
	Set objMessages = Session("/Inbox").Messages
		for i = objMessages.count to 1 step -1
			set objMessage = objMessages.Item(i)
			if objMessage.Unread then
				EmailUnreadCnt = EmailUnreadCnt + 1
			end if
		next
  end if


%>
<font color=red size=2>
<b>You have <%=EmailUnreadCnt%> unread e-mails in your <a href="../../email/content/email.asp">InBox</a>.</b>
</font><p>

<font color=navy size=1 face="ms sans serif, arial, geneva"><b><span>Quote of the Month</span></b>
<tr><td bgcolor=black colspan=2><img src="../../common/images/spacer.gif" width=10 height=2></td></tr>
<tr><td colspan=2><font color=maroon size=2><b>"Build a system that even a fool can use and <br>
only a fool	will use it" - Unknown<b>
</td></tr></font>
<%


  set objFolder = nothing
  set objOMSession = nothing
  set objRootFolder = nothing
 end if
else
 Response.Write("NOT USING MAPI<BR>")
end if 'End if Session("UseMaPI")

'********************************************************************************************
%>

		<!-- ANNOUNCEMENTS -->
<%'******************************************************************************************
 if Application("htmlAnnouncements") <> "" then%>
		<tr><td nowrap><font color=navy size=1 face="ms sans serif, arial, geneva"><b>
			<span onClick="toggle('Announcements')" alt="Expand or collapse this section." style="cursor:hand">Announcements</span></b>
			<span id="AnnouncementsHeader" style=""></span>
			</b></font></td><td>&nbsp;</td></tr>
			<tr><td bgcolor=black colspan=2><img src="../../common/images/spacer.gif" width=10 height=2></td></tr>
			<tr><td colspan=2><img src="../../common/images/spacer.gif" width=10 height=5></td></tr>

		<tr><td colspan=2>
			<span id=Announcements>
			<table width="100%" border=0 cellpadding=0 cellspacing=0>
			<tr>
			<td width=28 valign=top><img src="../../common/images/tiny/new2.gif" height=9 width=25></td>
			<td>
			<font size=1><%=Application("htmlAnnouncements")%></font>
			</td>
			</tr>
			</table>
			</span>
			</td></tr>

		<tr><td colspan=2>&nbsp;</td></tr>
<%end if%>



		<!-- EVENTS -->
<% if Application("htmlEvents") <> "" then%>
		<tr><td nowrap><font color=navy size=1 face="ms sans serif, arial, geneva"><b>
			<span onClick="toggle('Events')" alt="Expand or collapse this section." style="cursor:hand">Upcoming Events</span></b>
			<span id="EventsHeader" style=""></span>
			</b></font></td><td>&nbsp;</td></tr>
			<tr><td bgcolor=black colspan=2><img src="../../common/images/spacer.gif" width=10 height=2></td></tr>
			<tr><td colspan=2><img src="../../common/images/spacer.gif" width=10 height=5></td></tr>

		<tr><td colspan=2>
			<span id=Events>
			<%=Application("htmlEvents")%>
			</span>
			</td></tr>

		<tr><td colspan=2>&nbsp;</td></tr>
<%end if%>

		
		<!-- NEWSLETTER -->

		<tr><td nowrap><font color=navy size=1 face="ms sans serif, arial, geneva"><b>
			<span onClick="toggle('CurrentNewsletter')" alt="Expand or collapse this section." style="cursor:hand">Current News</span></b>
			<span id="CurrentNewsletterHeader" style=""></span>
			</b></font></td><td width="100%"></td></tr>
			<tr><td bgcolor=black colspan=2><img src="../../common/images/spacer.gif" width=10 height=2></td></tr>
			<tr><td colspan=2><img src="../../common/images/spacer.gif" width=10 height=5></td></tr>

		<tr><td colspan=2>
			<span id="CurrentNewsletter">
			<!-- #include file="incNewsletter.asp" -->
			</span>
			</td></tr>

		<tr><td colspan=2>&nbsp;</td></tr>


		
	</td></tr></table>
	
		
		<!-- CNET HEADLINES -->
	
<table width=340 border=0 cellpadding=0 cellspacing=0>
		<tr>
			<td nowrap><font color=navy size=1 face="ms sans serif, arial, geneva">
				<b><span onClick="toggle('CNetHeadlines')" alt="Expand or collapse this section." style="cursor:hand">Headlines</span></b> by <a target=_blank href="http://www.cnet.com">CNet</a>
				<span id="CNetHeadlinesHeader" style=""></span></font></td>
			<td align=right nowrap><font color=silver size=1 face="ms sans serif, arial, geneva">
				&nbsp;&nbsp;<%=thisHour & ":00 " & thisAMPM & "&nbsp;&nbsp " & date%>
				</font></td>
		</tr>
		<tr><td colspan=2 bgcolor=black><img src="../../common/images/spacer.gif" width=10 height=2></td></tr>
		<tr><td colspan=2><img src="../../common/images/spacer.gif" width=10 height=5></td></tr>

		<tr><td colspan=2>
			<table border=0 cellspacing=0 cellpadding=0><tr>
				<td width=28>&nbsp;</td>
				<td><font size=1 face="ms sans serif, arial, geneva">
					<span id="CNetHeadlines">

<table>
	<tr>
		<td>
		<%=session("sCNet_HTML")%>
		</td>
	</tr>
</table>
<br>

</span>
				</td></tr>
			</table>
		</tr>
</table>



		<!-- WEATHER CHANNEL -->

<table width=340 border=0 cellpadding=0 cellspacing=0>
		<tr>
			<td nowrap><font color=navy size=1 face="ms sans serif, arial, geneva">
				<b><span onClick="toggle('Weather')" alt="Expand or collapse this section." style="cursor:hand">Weather</span></b> by <a target=_blank href="http://www.weather.com">The Weather Channel</a>
				<span id="WeatherHeader" style=""></span></font></td>
			<td align=right nowrap><font color=silver size=1 face="ms sans serif, arial, geneva">
				&nbsp;&nbsp;<%=thisHour & ":00 " & thisAMPM & "&nbsp;&nbsp; " & date%>
				</font></td>
		</tr>
		<tr><td colspan=2 bgcolor=black><img src="../../common/images/spacer.gif" width=10 height=2></td></tr>
		<tr><td colspan=2><img src="../../common/images/spacer.gif" width=10 height=5></td></tr>

		<tr><td colspan=2>
			<table border=0 cellspacing=0 cellpadding=0><tr>
				<td width=14>&nbsp;</td>
				<td><font size=1 face="ms sans serif, arial, geneva">
					<span id="Weather">

						<table border=0 cellpadding=5 cellspacing=0>
							<tr>
								<a target=_blank href="http://oap.weather.com/fcgi-bin/oap/redirect_magnet?loc_id=USOH0188&par=internal&site=magnet&code=232070&promo=english">
								<img border=0 width=270 height=140 SRC="http://oap.weather.com/fcgi-bin/oap/generate_magnet?loc_id=USOH0188&code=232070"></a>
							</tr>
						</table>

					</span>
				</td></tr>
			</table>
		</tr>
</table>

<% 'END %>
	
	
	


	</td><td width=20>&nbsp;</td>

	<td width=130 valign=top><%=Application("htmlCalendar")%>
    <!--	<a href='http://www.microsoft.com/mcsp/' target='_blank'><img src='mcspsmall.gif' height=29 width=115 border=0></a>-->
	<a href='http://www.microsoft.com/mcsp/' target='_blank'><img src='mcsp_anim_logo.gif' height=40 width=135 border=0></a>
	<p>
    <a href='http://207.152.96.9/a4roott/a4home.htm' target='_blank'><img src='acct4.gif' alt='Account4 Homepage' height=50 width=135 border=0></a>
	<center>
	<a href='http://www.my401k.com' target='_blank'><img src='metlifelogosnoopy.gif' alt='MetLife 401k Homepage' height=50 width=85 border=0></a>
	</center>
	<table width="100%" border=0 cellpadding=5 cellspacing=0><tr><td bgcolor=#ffffcc><font size=1 face="ms sans serif, arial, geneva">


<!-- HERE ARE THE FEATURED LINKS!! -->


		<b><font color=navy>Featured Links</font></b><font color=gray><br>
<!--		&nbsp;&nbsp;-&nbsp;<a href="../../extras/education/default.htm" target="_blank">Certification Notes</a><br> -->
<!--		&nbsp;&nbsp;-&nbsp;<a href="../../sports/content/Golf_League.asp">Golf Handicapper Spreadsheet</a><br>-->

		&nbsp;-&nbsp;<a href="../../office/content/album.asp">Photo Album</a><br>
		&nbsp;-&nbsp;<a target=body href="../../wellness/content/default.asp">Wellness Program</a><br>
		&nbsp;-&nbsp;<a href="javascript:SlideShow('Weblogic', 'WebLogicPresentation')">Weblogic Presentation</a><br>
	    &nbsp;-&nbsp;<a href="javascript:SlideShow('2001 Business Plan', 'BusinessPlan2001')">2001 Business Plan</a><br>
		&nbsp;-&nbsp;<a href="../../welcome/content/Sark_News_Flash.htm" target="_blank">Sark News Flash</a><br>
		&nbsp;-&nbsp;<a href="../../welcome/content/B2B_Primer.htm" target="_blank">B2B Primer</a><br>
		&nbsp;-&nbsp;<a href="http://www.sark.com"target="_blank">New Sark.com Site</a><br>
		&nbsp;-&nbsp;<a href="../../office/content/corpintranet.asp">Corporate Intranet</a><br>
<!--		&nbsp;&nbsp;-&nbsp;<a href="../../news/content/About.asp">Branch News</a><br>-->
<!--		&nbsp;&nbsp;-&nbsp;<a href="../../news/content/Classifieds.asp">Classifieds</a><br>-->
<!--		&nbsp;-&nbsp;<a href="http://www.my401k.com"target="_blank">401k web site</a><br>-->
<!--		&nbsp;&nbsp;-&nbsp;<a href="http://207.152.96.9/a4roott/a4home.htm"target="_blank">Account4</a><br>-->
	    &nbsp;-&nbsp;<a href="../../office/content/documentviewers.asp">Document Viewers</a><br>
	    &nbsp;-&nbsp;<a href="../../events/content/aroundtown.asp">Links Around Town</a><br>
		
		<b><font color=navy>New Offices</font></b><font color=gray><br>
		&nbsp;-&nbsp;<a href="../../welcome/content/SarkIndy.htm" target="_blank">Indianapolis</a><br>
		&nbsp;-&nbsp;<a href="../../welcome/content/SarkSanDiego.htm" target="_blank">San Diego</a><br>
		
<!--		<font color=navy>Certification Links</font><font color=gray><br>
		&nbsp;&nbsp;-&nbsp;<a href="../../news/content/article.asp?NewsID=7">Linux Certification</a><br>
		&nbsp;&nbsp;-&nbsp;<a href="../../Certifications/Notes/Linux_Level1_Exam_Tips.doc"target="_blank">Linux Exam Tips</a><br>		
		&nbsp;&nbsp;-&nbsp;<a href="../../Certifications/Notes/Visual_Interdev_Exam_Tips.doc"target="_blank">Interdev Exam Tips</a><br>
		&nbsp;&nbsp;-&nbsp;<a href="../../Certifications/Notes/sql_server_exam_tips.doc"target="_blank">SQL Server 7.0 Exam Tips</a><br>		
		<p>
		<font color=navy>Training</font><font color=gray><br>
		&nbsp;&nbsp;-&nbsp;<a href="../../training/content/training.asp?NewsID=75">Intro to Data Warehousing</a><br>
		&nbsp;&nbsp;-&nbsp;<a href="../../training/content/training.asp?NewsID=76">Intro to XML</a><br>
		&nbsp;&nbsp;-&nbsp;<a href="../../training/content/training.asp?NewsID=88">Advanced Java</a><br>
-->		
		<b><font color=navy>Technology</font></b><font color=gray><br>
<!--		&nbsp;&nbsp;-&nbsp;<a href="javascript:SlideShow('Microsoft Transaction Server', 'MSTransactionServer')">MTS Presentation</a><br>-->
		&nbsp;-&nbsp;<a href="../../Training/Content/cert_study.asp">Java 2 Prog Exam Tips</a><br>
		&nbsp;-&nbsp;<a href="../../welcome/content/DevDays.htm" target="_blank">DevDays Info</a><br>
		&nbsp;-&nbsp;<a href="../../technology/content/tech.asp?techid=21&">XML Links</a><br>
		&nbsp;-&nbsp;<a href="../../technology/content/tech.asp?techid=80&">XSL Links</a><br>
		&nbsp;-&nbsp;<a href="../../technology/content/tech.asp?techid=84&">Websphere Links</a><br>
		&nbsp;-&nbsp;<a href="../../technology/content/tech.asp?techid=87&">WebLogic Links</a><br>
<!--		&nbsp;&nbsp;-&nbsp;<a href="../../Certifications/Notes/Corba_presentation_notes.doc"target="_blank">Corba Presentation Notes</a><br>-->
		
		</font>

<%
'--------------------------
'  Build new hire listing  
'--------------------------
rdate = DateAdd("m", -1, Date)
set rs = DBQuery("SELECT employeeid, firstname, lastname FROM employee WHERE StartDate > '" & rdate & "' ORDER BY lastname, firstname")
if not rs.eof then
	response.write("<font color=navy>New Hires:</font><br>")
	while not rs.eof
		response.write("&nbsp;&nbsp;-&nbsp;<a href='/" & application("web") & "/directory/content/details.asp?EmpId=" & rs("EmployeeID") & "'>" & rs("FirstName") & " " & rs("LastName") & "</a><br>" & NL)
		rs.movenext
		wend
	end if
rs.close
%>

<%
'----------------------------
'Build New Tech links listing
'-----------------------------

'set rs = DBQuery("SELECT url, substring(url, 8, CASE WHEN charindex('/', substring(url, 8, len(url))) = 0 THEN len(url) ELSE charindex('/', substring(url, 8, len(url))) - 1 END) urlstripped, timestamp FROM Tech_Links WHERE timestamp > '5/1/00' ORDER BY timestamp")

'if not rs.eof then
	
'	Response.Write("<p>")
'	Response.Write("<font color=navy>New Tech Links:</font><br>")
'	while not rs.eof
'	ls_urlstripped = rs("urlstripped")
'   ls_url
'	response.write("&nbsp;-&nbsp;<a href='" & ls_url & "' target='_blank'>" & ls_url & "</a><br>" & NL)
'		rs.movenext
'		wend
'	end if
'rs.close
		
%>		
	</font></td></tr></table><br>

<!-- Display Session Hit Counter -->
	<table class=tableShadow width="100%" height=125 border=0 cellspacing=0 cellpadding=3 bgcolor=#ffffcc>
		<tr><td valign=top><font size=1 face="ms sans serif, arial, geneva" color=navy>
			Number of User Visits<br>&nbsp;&nbsp;in the last 30 days:
		</td></tr>

		<tr><td align=center valign=bottom nowrap><img src='../../common/images/ruler.gif' height=75 width=13 border=0 alt="Please don't touch the meter, it's fragile.">
		
<%
	
	if Application("SessionCountDate") < DateAdd("d", -1, Date()) then
		' If first visitor of the day, get Session count for last 30 days and
		' store in the application variable 'SessionCountDate' and update
		' application variable 'SessionCountDate'

		Dim TempArray(30, 2)
		set rs = DBQuery("GetLast30DaysSessionCount")
		Application("SessionCountDate") = DateAdd("d", -1, Date())

		for i=30 to 1 Step -1
			rdate = DateAdd("d", -i, Date()) ' Date looking for
			TempArray(i, 0) = rdate
					
			DailyCount = 0 ' Initialize
			
			if Cdate(rs("SessionDate")) = rdate and not rs.eof then
				DailyCount = rs("Count")
				rs.MoveNext
			End if
			TempArray(i,1) = DailyCount			
		next

		Application("SessionCount") = TempArray
	End if	
	
	CountArray = Application("SessionCount")

	for i=30 to 1 step -1
		If WeekDay(Cdate(CountArray(i,0))) = vbSunday or WeekDay(Cdate(CountArray(i,0))) = vbSaturday then
			'Display a red dot for weekends
			Response.write("<img src='../../common/images/reddot.gif' height=" & CountArray(i,1)/4 & " width=3 border=0 alt='" & CountArray(i,0) & ": " & CountArray(i,1) & " hits'>")
		Else
			Response.write("<img src='../../common/images/greendot.gif' height=" & CountArray(i,1)/4 & " width=3 border=0 alt='" & CountArray(i,0) & ": " & CountArray(i,1) & " hits'>")
		End If
	next
%>
		</td></tr>
	</table><br>

</td>
	

</table>

<%
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Added by Jason Moore (10/23/00) 
	' Prompts user to update his/her skills inventory if it hasn't been updated in last thirty
	' days.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	If Not Session("UpdatedSkills") Then
		Dim skillsSQL, skillsRS, dtLastModified, intTitleID, blnShowUpdatePrompt
		skillsSQL =		"SELECT DateSkillsModified, EmployeeTitle_ID " & _
						"FROM Employee " & _
						"WHERE EmployeeID = " & Session("ID")
		Set skillsRS = DBQuery(skillsSQL)
		
		If IsNull(skillsRS("DateSkillsModified")) Then
			dtLastModified = CDate("01/01/70")
		Else
			dtLastModified = CDate(skillsRS("DateSkillsModified"))
		End If
		intTitleID = CInt(skillsRS("EmployeeTitle_ID"))
		
		Set skillsRS = Nothing
		
		If intTitleID >= 1 And intTitleID <= 4 Then		
			If DatePart("m", Date()) - DatePart("m", dtLastModified) >= 3 Then
				If DatePart("m", Date()) - DatePart("m", dtLastModified) = 3 Then
					If DatePart("d", Date()) >= DatePart("d", dtLastModified) Then
						blnShowUpdatePrompt = True
					End If
				Else
					blnShowUpdatePrompt = True
				End If
			End If
		End If
		
		'Note: blnShowUpdatePrompt defaults to False
		If blnShowUpdatePrompt Then
%>
			<script language=javascript>
				window.open('skillsInvUpdate.asp', '', 'height=300,width=400,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0,scrollbars=1');
			</script>
<%		
		End If
		Session("UpdatedSkills") = True
	End If 
%>
	


<script language=javascript>
<%If Not Session("Surveyed") Then%>
window.open('Add_Survey.asp','SurveyWindow','height=600,width=600,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0,scrollbars=1');
<%End If%>
</script>


	
<!-- #include file="../../footer.asp" -->


<script language=javascript>
<!--
var wIndex = 0
var supportAnimateMap = false
var weatherIndex = 0
var maxWeatherIndex = 2

function getNextMapIndex(index){
	index++
	if (index > maxWeatherIndex) index = 0
	return index
	}

function getMapTitle(index){
	if (index == 0) result = "East Central Conditions"
	else if (index == 1) result = "Local Doppler"	
	else if (index == 2) result = "U.S. Conditions"	
	return result
	}

function getMapImage(index){
	if (index == 0) result = "http://maps.weather.com/images/maps/current/cur_ec_720x486.jpg"
	else if (index == 1) result = "http://maps.weather.com/images/radar/regions/east_cen_rad_450x284.jpg"
	else if (index == 2) result = "http://maps.weather.com/images/maps/current/curwx_720x486.jpg"
	return result
	}

function getMapThumbnail(index){
	if (index == 0) result = "http://maps.weather.com/images/maps/current/cur_ec_277x187.jpg"
	else if (index == 1) result = "http://maps.weather.com/images/radar/regions/east_cen_rad_300x187.jpg"
	else if (index == 2) result = "http://maps.weather.com/images/maps/current/curwx_277x187.jpg"
	return result
	}

function timerWeatherThumbnailRefresh(){
	wIndex = getNextMapIndex(wIndex)
	document.map.src = img[wIndex].src
	i = setTimeout("timerWeatherThumbnailRefresh()", 5000)
	}

function showWeather(){
	
	winWeather = window.open("displayMap.asp?url=" + getMapImage(wIndex) ,"weatherMap","height=480,width=640,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0")
	}
function statusMap(){
	window.status = "Display maps."
	}
function statusNone(){
	window.status = ""
	}

if (browserVersion > 3){
	supportAnimateMap = true
	var img = new Array(maxWeatherIndex)
	for (x=0; x < maxWeatherIndex+1; x++){
		img[x] = new Image(83,52)
		img[x].src = getMapThumbnail(x)
		}
	}

if (supportAnimateMap) i = setTimeout("timerWeatherThumbnailRefresh()", 7000)
// -->
</script>
