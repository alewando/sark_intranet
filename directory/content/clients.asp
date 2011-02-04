<!--
Developer:    GINGRICE
Date:         08/18/1998
Description:  Displays listing of consultants by client
-->


<!-- #include file="../section.asp" -->


<% 
lb_group1 = false
lb_group2 = false
lb_group3 = false
lb_group4 = false
lb_group5 = false

'----------------------------
'	Execute Database Query   
'   to get all employees at  
'   clients or on beach      
'----------------------------
set rs = DBQuery("select c.sortoverride, c.client_id, c.clientname, c.website, e.firstname, e.lastname, e.employeeid, e.client_phone from client c INNER JOIN employee e on e.client_id = c.client_id where e.client_id <> 1 and c.sortoverride <> 2 and e.Branch_ID = " & Application("DefaultBranchID") & " and e.DateTermination is null order by c.sortoverride, c.clientname, e.lastname")  
Response.Write("<TABLE width=465>")
Response.Write("<tr><td ALIGN=LEFT VALIGN=top>")
Response.write("<TABLE width=225 BORDER=0>")
while not rs.eof
	ls_clientid = rs("client_id")
	ls_clientphone = rs("client_phone")
	If ls_clientid <> ls_previd then
		Response.Write("<TR><TD COLSPAN=3><FONT SIZE=1 face='ms sans serif, arial, geneva'><B><U>")
		client = UCase(rs.fields("clientname"))
		Response.Write("<br>")
		if ((left(client,1) = "A") or (left(client,1) = "B") or (left(client,1) = "C") or  (left(client,1) = "D") or (left(client,1) = "E"))  and not lb_group1 then
			Response.Write("<a name=Group1></a>")
			lb_group1 = true
		else 
			if ((left(client,1) = "F") or (left(client,1) = "G") or (left(client,1) = "H") or  (left(client,1) = "I") or (left(client,1) = "J"))  and not lb_group2 then
				Response.Write("<a name=Group2></a>")
				lb_group2 = true
			else 
				if ((left(client,1) = "K") or (left(client,1) = "L") or (left(client, 1) = "M") or  (left(client,1) = "N") or (left(client,1) = "O"))  and not lb_group3 then
					Response.Write("<a name=Group3></a>")
					lb_group3 = true
				else 
					if ((left(client,1) = "P") or (left(client,1) = "Q") or (left(client,1) = "R") or  (left(client,1) = "S") or (left(client,1) = "T"))  and not lb_group4 then
						Response.Write("<a name=Group4></a>")
						lb_group4 = true
					else 
						if ((left(client,1) = "U") or (left(client,1) = "V") or (left(client,1) = "W") or  (left(client,1) = "X") or (left(client,1) = "Y") or (left(client,1) = "Z"))  and not lb_group5 then
							Response.Write("<a name=Group5></a>")
							lb_group5 = true
						end if 
					end if
				end if
			end if
		end if
		if not isnull(rs("website")) and rs("website") <> "" then client = "<a href='http://" & rs("website") & "' target='_blank'>" & client & "</a>"
		Response.Write(client)
		Response.Write("</U></B></FONT></TD></TR>")
	end if 
	Response.Write("<TR><TD><FONT SIZE=1 face='ms sans serif, arial, geneva'>&nbsp;<a href='details.asp?EmpId=" & rs("employeeid") & "'>" & rs("firstname") & " " & rs("lastname") & "</a>") 
	If IsNull(ls_clientphone) and rs("sortoverride") = 0 then
		ls_clientphone = "N/A"
	End If
	Response.Write("</FONT></TD><td><font size=1>&nbsp;&nbsp;&nbsp;</font></td><TD><FONT SIZE=1 face='ms sans serif, arial, geneva'>" & ls_clientphone & "</FONT></TD></TR>")
	ls_previd = ls_clientid
	rs.movenext
wend
rs.close
Response.Write("</TABLE><br>")

Response.Write("<td>&nbsp;&nbsp;</td>")
Response.Write("<td ALIGN=LEFT VALIGN=top>")


'------------------------
'  Display Office Staff  
'------------------------
set rs = DBQuery("select e.firstname, e.lastname, e.employeeid, e.cellphone, e.pagernumber, e.pageremail, e.voicemail from employee e where e.voicemail is not null and e.voicemail between 2810 and 2825 and DateTermination is null order by e.lastname")  
Response.Write("<br>")
Response.Write("<TABLE border=0 class=tableShadow width=245 cellspacing=0 cellpadding=2 bgcolor=#ffffcc>")
Response.Write("<TR valign=bottom>")
Response.Write("<TD ALIGN=CENTER><FONT SIZE=1 face='ms sans serif, arial, geneva'><B><a name=OfficeStaff></a>OFFICE STAFF</B></FONT></TD>")
Response.Write("<TD ALIGN=CENTER><FONT SIZE=1 face='ms sans serif, arial, geneva'><B>MOBILE/<br>PAGER</B></FONT></TD>")
Response.Write("<TD ALIGN=CENTER><FONT SIZE=1 face='ms sans serif, arial, geneva'><B>VM</B></FONT></TD>")
Response.Write("</TR>")
Response.Write("<TR valign=top><TD COLSPAN=3><b><hr size=1></b></TD></TR>")
while not rs.eof
	ls_carphone = rs("cellphone")
	ls_pagernumber = rs("pagernumber")
	ls_pageremail = rs("pageremail")
	ls_voicemail = rs("voicemail")
	Response.Write(chr(13) & "<TR valign=top><TD ALIGN=LEFT>")
	Response.Write("<FONT SIZE=1 face='ms sans serif, arial, geneva'>" & "&nbsp;<a href='details.asp?EmpId=" & rs("employeeid") & "'>" & rs("firstname") & "&nbsp;" & rs("lastname") & "</a>&nbsp;") 
	Response.Write("</FONT></TD>")
	Response.Write("<TD ALIGN=RIGHT><FONT SIZE=1 face='ms sans serif, arial, geneva'>")
	Response.Write( ls_carphone )
	if not IsNull(ls_pageremail) and ls_pageremail <> "" then
	 if not IsNull(ls_carphone) and ls_carphone <> "" then
	  Response.Write("<BR>")
	 end if
	 Response.Write("<A HREF='javascript:SendPage(" & chr(34) & ls_pageremail & chr(34) & ")'>" & replace(ls_pagernumber," ","&nbsp;") & "</A>&nbsp;&nbsp;(Pager)")
	end if
	Response.Write("</FONT></TD>")
	Response.Write("<TD><FONT SIZE=1 face='ms sans serif, arial, geneva'>" & ls_voicemail & " " & "</FONT></TD></TR>")
	rs.movenext
wend
rs.close
Response.Write("</TABLE>")
Response.Write("</TD>")
Response.Write("</TR>")
Response.Write("</TD>")
Response.Write("</TR>")
Response.Write("</TABLE>")



SectionBtns = SectionBtns & _ 
	"<TABLE BORDER=0 ALIGN=CENTER>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group1>A-E</B></FONT></a></TD></TR>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group2>F-J</B></FONT></a></TD></TR>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group3>K-O</B></FONT></a></TD></TR>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group4>P-T</B></FONT></a></TD></TR>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group5>U-Z</B></FONT></a></TD></TR></TABLE>" 
%>

<SCRIPT LANGUAGE="JavaScript">
 function SendEmail(userID){winTech= window.open("../../email.asp?editto=yes&recipient="+userID+"&footer=<%=Server.URLEncode(UCase(Application("web")) & ": " & curTitle)%>", "SendEmail","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")}
 function SendPage(pagerEmail){winTech= window.open("../../pager.asp?recipient="+pagerEmail+"&from=<%=session("Username")%>&name=<%=Server.URLEncode(session("name"))%>&footer=<%=Server.URLEncode(UCase(Application("web")) & ": " & curTitle)%>", "SendPage","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")}

</SCRIPT>
<!-- #include file="../../footer.asp" -->
