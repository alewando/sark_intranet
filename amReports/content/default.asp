<!--
Developer:    SSeissiger
Date:         8-16-00
Description:  Default Page for AM Reports
-->

<!-- #include file="../section.asp" -->

<p><font size = 2 color = blue><b>This section builds reports for account managers.</b>
<p><b>Reports...</b></font><br><br><br>

<%if (hasRole("AccountManager")) then%>

<TABLE border=0 WIDTH=480><font size=1>

<TR><TD><font size=1 color = blue>Vacation</TD></font>
<TD><b><font size=1>Vacation - Next 30 Days</b> - Vacation scheduled for the next 30 days</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Vacation - Next 12 Months</b> - Vacation scheduled for the next 12 months</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Vacation - Past 2 Weeks</b> - Vacation taken in the past 2 weeks</TD></font>
</TR>
<TR><TD><font size=1 color = blue>Reviews</TD></font>
<TD><font size=1><b>Review by Date</b> - Reviews listed by date</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Review by A/M</b> - Reviews listed by account manager</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Reviews by A/M - Next 90 Days</b> - Reviews listed by A/M - 1 to 90 Days</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Reviews Past Due</b> - Reviews listed by Consultant</TD></font>
</TR>
<TR><TD><font size=1>&nbsp;</TD></font>
<TD><font size=1><b>Sponsor/Sponsoree</b> - Sponsors and sponsorees with review date</TD></font>
</TR>
<TR><TD><font size=1 color = blue>Exams & Classes</TD>
<TD><font size=1><b>Exams by SARKs</b> - Lists passed exams by consultant</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>SARKs by Exams</b> - Lists consultants by passed exams</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Exams 0-60d old</b> - 0-60 day aging report for exams</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Exams by Vendor</b> - Total exams passed by vendor</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Certs by Cat.</b> - Lists certifications by category</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>MCP Listing</b> - List of MCP #s for all Sarks</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Classes by SARKs</b> - List of consultants and classes taken</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>SARKs by Classes</b> - List of classes and consultants who have taken them</TD></font>
</TR>
<TR><TD><font size=1 color = blue>Skills Inventory</TD></font>
<TD><font size=1><b>Skills Inv: By Date</b> - All SARKs and Date modified</TD></font>
</TR>
<TR><TD><font size=1>&nbsp;</TD></font>
<TD><font size=1><b>Skills Inv: Tardy</b> - Sarks who haven't entered skill inventories</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Skills Inv: Aging</b> - Sarks with skill inventories older than 3 months </TD></font>
</TR>
<TR><TD><font size=1 color = blue>Other Reports</TD></font>
<TD><font size=1><b>SARKs by Title</b> - List of SARKs by employee title</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Clients: S/R & A/M</b> - Clients with Sales Rep and Acct. Manager</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Clients by A/M</b> - Client list with Sales Rep by Acct. Manager</TD></b></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>No Techs Listed</b> - List of Consultants who have not entered a technology </TD></b></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Intranet Survey</b> - Intranet Survey Results</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Employee Security</b> - Security roles assigned to Sarks</TD></font>
</TR>

<%
'--------------------------------------------------
'Added By SSeissiger - Allows TBreuer to see the
'Review and sponsor reports for the 
'Sponsorship Program
'--------------------------------------------------
	elseif hasRole("SponsorAdmin") then %>
<TR><TD>&nbsp;</TD>
<TD><font size=1><TR><TD><b>Review by Date</b> - Reviews listed by date</TD></font>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Review by A/M</b> - Reviews listed by account manager</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Reviews by A/M- Next 45 Days</b> - Reviews listed by A/M - 45 Days</TD></font>
</TR>
<TR><TD>&nbsp;</TD>
<TD><font size=1><b>Sponsor/Sponsoree</b> - Sponsors and sponsorees with review date</TD></font>
</TR>
</TABLE>
		
		
		
<%end if%>



<!-- #include file="../../footer.asp" -->
