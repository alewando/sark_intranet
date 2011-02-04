<!--
Developer:    GINGRICE
Date:         08/18/1998
Description:  Displays listing of branches
-->

<!-- #include file="../section.asp" -->

<CENTER>
<% 
'-------------------------
'	Execute Database Query
'-------------------------
set rs = DBQuery("select * from branch order by branch_name")  
while not rs.eof
	
	'-------------------------------
	'	Display Branch Information
	'-------------------------------
	response.write("<b><a href='http://" & rs("BranchURL") & "' target='_blank'>" & UCase(rs("branch_name")) & "</a></b><br>" & NL)
	response.write(rs("address") & "<br>")
	response.write(rs("city") & ", " & rs("state") & "  " & rs("zip") & "<br>")
	response.write("Phone " & rs("phone") & " / Fax " & rs("fax") & "<br>")	if not IsNull(rs("tollfree")) and rs("tollfree") <> "" then
	 Response.Write("Toll Free " & rs("tollfree") & "<BR>")
	end if	Response.Write("<br>")
	rs.movenext
	wendrs.close
%>
</CENTER>

<!-- #include file="../../footer.asp" -->


