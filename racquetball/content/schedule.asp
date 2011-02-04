<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--
Developer:    Julie Walters
Date:         11/14/2000
Description:  League Schedules
Modifications: Adam Lewandowski (11/21/00) 
               Pull League info from database
--><!-- #include file="../section.asp" --><HTML><HEAD>
</HEAD>
<BODY>

<center><FONT face=Arial size=3><b> Sark Racquetball - Fall 2000 </b></FONT></center>

<p>
Important Information regarding scheduling matches:
<br>
<OL>
<li>Players set up their own matches during allotted week.
<li>Please call the Recreation Center up to one week in advance to reserve a court.  Please indicate that you play in the league and also tell the attendant who your opponent is.
<li>Reserved court time must be cancelled 24 hours in advance or a $2 fee will be charged to players.
<li>Match results should be turned in at the Control Desk 
  immediately following play. This will help us keep the standings up to date. 
  The Recreation Department will assume a match was not played if they do not get 
  a result sheet before the next scheduled week.</li></OL>

<% if hasRole("racquetballAdmin") then %>
<P>
<A HREF="editSchedule.asp">Add/Edit/Remove Leagues</A><BR>
<% end if %>

<P>Click on the league you belong to for the schedule.<BR>
<UL>
<%
 sql = "SELECT * FROM Racquetball_League ORDER BY League_ID"
 set rs = DBQuery(sql)
 do while not rs.eof
  Response.Write("<LI><A HREF='showSchedule.asp?leagueID=" & rs("League_ID") & "'>" & _
	rs("Description") & "</A><BR>")
  rs.moveNext
 loop 
%>
</UL>
<P><!-- #include file="../../footer.asp" --></P>
</BODY></HTML>
