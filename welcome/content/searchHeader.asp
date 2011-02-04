<!--
Developer:    Kevin Dill
Date:         08/17/1998
Description:  Search header with search info and engine.
-->


<html>

<head>
<title>Search the Web</title>
</head>

<!-- #include file="../../style.htm" -->


<body bgcolor=black link=red vlink=red alink=red><center>

<table border=0 cellpadding=0 cellspacing=0><tr>

<form name=frmInfo method=get action="searchResults.asp" target="SearchResults"><td valign=center>
<font size=1 face='ms sans serif, geneva' size=2 color=white><b>Search the Web:</b></font>
<input type=text class=fldSearch name=search size=25 maxlength=250 value="<%=request.querystring("search")%>">
<select name="searchEngine" class=fldSearch onChange="document.frmInfo.submit()">
	<option>--- Search Engines ---</option>
	<option<%if request.Cookies("SearchEngine")="ALTAVISTA" then response.write(" selected")%>>AltaVista</option>
	<option<%if request.Cookies("SearchEngine")="EXCITE" then response.write(" selected")%>>Excite</option>
	<option<%if request.Cookies("SearchEngine")="HOTBOT" then response.write(" selected")%>>HotBot</option>
	<option<%if request.Cookies("SearchEngine")="INFOSEEK" then response.write(" selected")%>>Infoseek</option>
	<option<%if request.Cookies("SearchEngine")="LOOKSMART" then response.write(" selected")%>>LookSmart</option>
	<option<%if request.Cookies("SearchEngine")="LYCOS" then response.write(" selected")%>>Lycos</option>
	<option<%if request.Cookies("SearchEngine")="NORTHERNLIGHT" then response.write(" selected")%>>NorthernLight</option>
	<option<%if request.Cookies("SearchEngine")="WEBCRAWLER" then response.write(" selected")%>>WebCrawler</option>
	<option<%if request.Cookies("SearchEngine")="YAHOO!" then response.write(" selected")%>>Yahoo!</option>
</select>
<input type=submit class=fldSearch value="Go"><input type=button class=fldSearch value="Back" onClick="parent.location.href='default.asp'">
</td></form>
</tr></table>

</center>
</body>
</html>