
<!-- #include file="../script.asp" -->

<html>
<head>
<title>Restore Link</title>
</head>

<body bgcolor=silver>

<%
dim valid, tech_id, link_url, link_title, link_desc

valid = true
tech_id = request("id")
if tech_id = "" then valid = false
if not isnumeric(tech_id) then valid = false

if valid then
	set rs = DBQuery("select * from tech_links where active=0 and tech_links_id=" & tech_id)
	if rs.eof then
		valid = false
	else
		link_url = rs("url")
		link_title = rs("url_title")
		link_desc = rs("text")
		rs.close
		set rs = DBQuery("update tech_links set active=1 where tech_links_id=" & tech_id)
		end if
	DataConn.close
	set DataConn = nothing
	end if

if valid then
%>

<center><h2>Link Restored!</h2></center>

URL: <%=link_url%><br>
Title: <%=link_title%><br>
Description: <%=link_desc%>

<%else%>

Error!

<%end if%>

</body></html>
