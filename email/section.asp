<!-- #include file="../script.asp" -->
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
if sectionTitle = "" then sectionTitle = "Email"
sectionDir = "email"

buildNavItem "About this section", "about.asp"
if Session("Username") <> "sarkguest" and Session("UseMAPI") then
	buildNavItem "Mailbox", "email.asp"
	end if
%>
<!-- #include file="../header.asp" -->
