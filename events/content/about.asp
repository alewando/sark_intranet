<!--
Developer:    Kevin Dill
Date:         11/18/1998
Description:  About this section
-->

<!-- #include file="../section.asp" -->

<img src="images/calset_image.gif" width=50 height=38 hspace=15 vspace=5 align=left>
Looking for something to do, or wondering when the next training class is being offered?
Check out the SARK <a href="calendar.asp?date=">Calendar</a> for a list of SARK events.  
Here you will find activities such as SARK Adventures, Sporting Events, Recruiting Trips,
Training Courses, and many more!  
<p>
This section will also keep you up to date on important SARK dates, such as birthdays,
holidays and timesheet days.  You can also view the events in the section
<a href="bydate.asp?date=">by date</a> or <a href="bytype.asp?date=">by type</a>.<p>

<img src='../../common/images/tiny/new2.gif' height=9 width=25></img> Also included in this section are <a href="aroundtown.asp">Links Around Town</a>.
Here you can find links to various sites devoted to Cincinnati and it's surrounding area,
including the Cincinnati Zoo, Kings Island, and the Cincinnati Post.
<!--
<p>
Don't forget this year's <a href="http://splash.sarkcincinnati.com/auction" target="_blank">Charity Auction</a> will be held Thursday, December 9 directly<br>
following the company meeting.
-->
<br><br>

<table width="470" border=0 cellpadding=0 cellspacing=0><tr>

		<!-- EVENTS -->
<% if Application("htmlEvents") <> "" then%>
		<tr><td nowrap><font color=navy size=1 face="ms sans serif, arial, geneva"><b>
			Upcoming Events
			</b></font></td><td>&nbsp;</td></tr>
			<tr><td bgcolor=black colspan=2><img src="../../common/images/spacer.gif" width=10 height=1></td></tr>
			<tr><td colspan=2><img src="../../common/images/spacer.gif" width=10 height=5></td></tr>

		<tr><td colspan=2>
			<%=Application("htmlEvents")%>
			</td></tr>

		<tr><td colspan=2>&nbsp;</td></tr>
<%end if%>

</table>

<!-- #include file="../../footer.asp" -->
