<!--
Developer:		DMARSTON, MAPGAR
Date:			03/28/2000
Description:	Allows the webmaster the ability to approve new technology submissions.
History:		Created - 03/28/2000, 04/19/2000-->

<!-- #include file="../section.asp" -->

<script language = javascript>
<!--
	function IsAccepted(KeyVal){
		if (KeyVal==1){
			document.frmInfo.accepted.value=1
			document.frmInfo.action="ApproveProc.asp"
			//document.frmInfo.action="techapproval.asp"
			}
		else {
			document.frmInfo.accepted.value=0
			document.frmInfo.action="ApproveProc.asp"
			//document.frmInfo.action="techapproval.asp"
			}
		document.frmInfo.submit()
		}
	function AdjCounter(Val) {
		i1 = Number(document.frmInfo.Counter.value)
	    if (i1==0) {
			i1=1
		}
	    i2=Number(document.frmInfo.Total.value)
	    if (i1+Val<=1) {
			document.frmInfo.file.value = "Begin"
			i1=1
			}
		else 
			if (i1+Val>=i2) {
			document.frmInfo.file.value = "End"
			i1=i2
			}
		else {
			i1 += Val
			document.frmInfo.file.value = ""
		} 
		document.frmInfo.Counter.value = i1
		document.frmInfo.action="techapproval.asp"
		document.frmInfo.submit()
	}
// -->
</script>

<tr><td valign=top>
	<form NAME="frmInfo" ACTION="" method=post>
	<INPUT type=hidden name="accepted" value="">	
	<INPUT type=hidden name=Counter value="<%=Request.Form("Counter")%>"> <!--Holds the record number-->
	<INPUT type=hidden name="file" value=""><!--Used to determine BOF or EOF-->
	
<table border=0><tr><td colspan=2>
<font size=1 face="ms sans serif, arial, geneva" color=maroon>
<b>Note:</b> Please select either the 'Approve' or 'Decline' checkbox
for each new technology submission.  Click the 'Done' button when finished.
</font></td></tr>

<%
NL = chr(13) & chr(10)

ls_sql = "SELECT Tech_Area.Tech_Area, Tech.Tech_Name, Tech.Tech_Desc, Tech.Approved," _
	& " Tech.Tech_ID, Tech.Tech_Area_ID" _
	& " FROM Tech INNER JOIN Tech_Area ON" _
	& " Tech.Tech_Area_ID = Tech_Area.Tech_Area_ID" _
	& " WHERE Tech.Approved <> 1"

ls_Count="SELECT Count(*) as Total" _
	& " FROM Tech INNER JOIN Tech_Area ON" _
	& " Tech.Tech_Area_ID = Tech_Area.Tech_Area_ID" _
	& " WHERE Tech.Approved <> 1"
			
set rs = DBQuery(ls_sql)
set MyCountRs=DBQuery(ls_Count)
MyCount=MyCountRs("Total")
MyCountRs.close
set MyCountRs=nothing

ls_ID=Request.Form("Counter")

if ls_ID>1 then
	for i=2 to ls_ID
		rs.movenext
	next
end if

if Request.form("File")="End" then
	Response.Write("<font size=1 face='ms sans serif, arial, geneva' color=red><strong>End of File Reached!....</strong></font>")
elseif Request.Form("File")="Begin" then
	Response.Write("<font size=1 face='ms sans serif, arial, geneva' color=red><strong>Beginning of File Reached!....</strong></font>")
end if

if not rs.eof then
	ls_techID = rs("Tech_ID")
	ls_techAreaID = rs("Tech_Area_ID")
	ls_techAREA = rs("Tech_Area")
	ls_techNAME = rs("Tech_Name")
	ls_techDESC = rs("Tech_Desc")
	Response.Write("<tr><td colspan=2>&nbsp</td></tr>")
	Response.Write("<tr><td width='18%' align=right><font size=1 face='ms sans serif, arial, geneva'>Technology Area: </font></td>")
	Response.Write("<td><INPUT type='text' name='txtTechArea' value='" & ls_techAREA & "'></td></tr>")
	Response.Write("<tr><td align=right><font size=1 face='ms sans serif, arial, geneva'>Technology Name: </font></td>")
	Response.Write("<td><INPUT type='text' name='txtTechName' value='" & ls_techNAME & "'></td></tr>")	
	Response.Write("<tr><td valign='top' align=right><font size=1 face='ms sans serif, arial, geneva'>Description: </font></td>")
	Response.Write("<td><textarea name='txtaTechDesc' rows='10' cols='50' wrap='on'>" & ls_techDESC)
	Response.Write("</textarea></td></tr>")
	if MyCount>1 then
		Response.Write("<tr><td>&nbsp</td><td><input type=button class=button value='Accept' OnClick='javascript:IsAccepted(1);'>")
		Response.Write("&nbsp;&nbsp<input type=button class=button value='Decline' OnClick='javascript:IsAccepted(0);'>")
		Response.Write("&nbsp;&nbsp<input type=button class=button value='Cancel' OnClick='window.location.href=" & chr(34) & "default.asp" & chr(34) & "'>")
		Response.Write("&nbsp;&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<input type=button class=button value=' << ' OnClick='javascript:AdjCounter(-1);'>")
		Response.Write("&nbsp;&nbsp<input type=button class=button value=' >> ' OnClick='javascript:AdjCounter(1);'></td>")
	else
		Response.Write("<tr><td>&nbsp</td><td><input type=button class=button value='Accept' OnClick='IsAccepted(1);'>")
		Response.Write("&nbsp;&nbsp<input type=button class=button value='Decline' OnClick='IsAccepted(0);'>")
		Response.Write("&nbsp;&nbsp<input type=button class=button value='Cancel' OnClick='window.location.href=" & chr(34) & "default.asp" & chr(34) & "'>")
	end if
else
	Response.Write("<tr><td>&nbsp</td></tr><tr><td>No Current Submissions......</td></tr>")
	Response.Write("<tr><td>&nbsp</td></tr>")
	Response.Write("<tr><td><input type=button class=button value='Cancel' OnClick='window.location.href=" & chr(34) & "default.asp" & chr(34) & "'></td></tr>")
end if
rs.close
set rs=nothing
%>
<INPUT type=hidden name=Total value="<%=MyCount%>">  <!--Gets the total number of records-->
<INPUT type=hidden name=tech_id value="<%=ls_techID%>">
<INPUT type=hidden name=tech_area_id value="<%=ls_techAreaID%>">

</form></TABLE>

<!-- #include file="../../footer.asp" -->
