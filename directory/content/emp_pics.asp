<!--
Developer:    GINGRICE
Date:         08/18/1998
Description:  Displays listing of employees with pictures
-->


<!-- #include file="../section.asp" -->

<%
'----------------------------------------
'	Determines if Employee Picture exists
'----------------------------------------
Set fso = Server.CreateObject("Scripting.FileSystemObject")

'-------------------------
'	Execute Database Query
'-------------------------
set rs = DBQuery("select * from employee where DateTermination is NULL and Branch_ID = " & Application("DefaultBranchID") & " order by lastname")  

'-----------------------------------
'	Display Employee Info & Pictures
'-----------------------------------
lb_group1 = false
lb_group2 = false
lb_group3 = false
lb_group4 = false
lb_group5 = false

maxcols = 2
Response.Write("<TABLE WIDTH='100%' BORDER=0>" & NL)
path = Server.MapPath("images/") & "\"
while not rs.eof
	Response.Write("<TR VALIGN=TOP>")
	for col = 1 to maxcols
		Response.Write("<TD VALIGN=TOP>")
		if not rs.eof then  
			imgEmployee = "images/" & rs("username") & ".jpg"
			if not fso.FileExists(path & rs("username") & ".jpg") then imgEmployee = "images/no_photo.jpg"
			if NOT rs("unlistedpersonal") then
				ls_address = rs("address") & "<br>" & rs("city") & "," & rs("state") & "  " & rs("zip") & "<br>" & rs("homephone") 
			else
				ls_address = "<font color=gray>UNLISTED</font>"
			end if
			ls_name = rs("firstname") & " " & rs("lastname")
			if rs("spousename") <> "" then ls_name = ls_name & " (" & rs("spousename") & ")"
			'if NOT IsNull(rs("spousename")) then ls_name = ls_name & " (" & rs("spousename") & ")"
			
			if ((left(rs("lastname"),1) = "A") or (left(rs("lastname"),1) = "B") or (left(rs("lastname"),1) = "C") or  (left(rs("lastname"),1) = "D") or (left(rs("lastname"),1) = "E"))  and not lb_group1 then
				Response.Write("<a name=Group1></a>")
				lb_group1 = true
			else 
				if ((left(rs("lastname"),1) = "F") or (left(rs("lastname"),1) = "G") or (left(rs("lastname"),1) = "H") or  (left(rs("lastname"),1) = "I") or (left(rs("lastname"),1) = "J"))  and not lb_group2 then
					Response.Write("<a name=Group2></a>")
					lb_group2 = true
				else 
					if ((left(rs("lastname"),1) = "K") or (left(rs("lastname"),1) = "L") or (left(rs("lastname"),1) = "M") or  (left(rs("lastname"),1) = "N") or (left(rs("lastname"),1) = "O"))  and not lb_group3 then
						Response.Write("<a name=Group3></a>")
						lb_group3 = true
					else 
						if ((left(rs("lastname"),1) = "P") or (left(rs("lastname"),1) = "Q") or (left(rs("lastname"),1) = "R") or  (left(rs("lastname"),1) = "S") or (left(rs("lastname"),1) = "T"))  and not lb_group4 then
							Response.Write("<a name=Group4></a>")
							lb_group4 = true
						else 
							if ((left(rs("lastname"),1) = "U") or (left(rs("lastname"),1) = "V") or (left(rs("lastname"),1) = "W") or  (left(rs("lastname"),1) = "X") or (left(rs("lastname"),1) = "Y") or (left(rs("lastname"),1) = "Z"))  and not lb_group5 then
								Response.Write("<a name=Group5></a>")
								lb_group5 = true
							end if 
						end if
					end if
				end if
			end if
			Response.Write("<img src='" & imgEmployee & "' height=64 width=64 align=left></TD>" & NL)
			Response.Write("<TD><FONT SIZE=1 face='ms sans serif, arial, geneva'><a href='details.asp?EmpId=" & rs("employeeid") & "'>" & ls_name & "</a><br>" & ls_address & "<br>&nbsp;")
			rs.MoveNext
		end if
		Response.Write("</TD>" & NL)
	next     
	Response.Write("</TR>" & NL)
	wend
Response.Write("</TABLE>" & NL)
rs.close

SectionBtns = SectionBtns & _ 
	"<TABLE BORDER=0 ALIGN=CENTER>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group1>A-E</B></FONT></a></TD></TR>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group2>F-J</B></FONT></a></TD></TR>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group3>K-O</B></FONT></a></TD></TR>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group4>P-T</B></FONT></a></TD></TR>" & _
	"<TR><TD><FONT SIZE='-1'><B><a href=#Group5>U-Z</B></FONT></a></TD></TR></TABLE>" 
%>

<!-- #include file="../../footer.asp" -->
