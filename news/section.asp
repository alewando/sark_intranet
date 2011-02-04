<!-- #include file="../script.asp" -->
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
sectionTitle = "News"
sectionDir = "news"

'SetWebMaster "cdolan"

buildNavItem "About this section", "about.asp"
buildNavItem "News Article, Current", "article.asp"
buildNavItem "News Article, Archive", "article_archive.asp"
buildNavItem "Classifieds", "classifieds.asp"
buildNavItem "Submit Article", "submit.asp"
%>
<!-- #include file="../header.asp" -->
