<%@ Language=VBScript %>

<!-- #include file="../../script.asp" -->

<%
ls_TechID=Request.Form("tech_id")
ls_TechAreaID=Request.Form("tech_area_id")
ls_TechName=Request.Form("txtTechName")
ls_TechDesc=Request.Form("txtaTechDesc")
ls_Accepted=Request.Form("accepted")

ls_TechDesc=server.HTMLEncode(ls_TechDesc)
ls_TechDesc=replace(ls_TechDesc, "'", "")

if ls_Accepted=1 then
	ud_sql = "UPDATE Tech" _
		& " SET Tech_Name = '" & ls_TechName & "'," _
		& " Tech_Desc = '" & ls_TechDesc & "', Tech_Area_ID = " & ls_TechAreaID & "," _
		& " Approved = 1" _
		& " WHERE Tech_ID = " & ls_TechID 
else
	ud_sql = "DELETE FROM Employee_Tech_Xref" _
		& " WHERE Tech_ID = " & ls_TechID
	DBQuery(ud_sql)
	ud_sql = "DELETE FROM Tech" _
		& " WHERE Tech_ID = " & ls_TechID 
end if

DBQuery(ud_sql)		

Response.Redirect("techapproval.asp")

%>