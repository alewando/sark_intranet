<!--
Developer:    Kevin Dill
Date:         08/17/1998
Description:  Main search frame containing a header and the results.
-->


<html>

<head>
<title>SARK Intranet</title>
</head>


<frameset rows="34,*" frameborder=0 framespacing=0>
	<frame name="SearchHeader" src="searchHeader.asp?search=<%=request.form("search")%>" marginheight="2" marginwidth="0" frameborder="0" scrolling="no" noresize>
	<frame name="SearchResults" src="searchResults.asp?search=<%=request.form("search")%>&searchEngine=<%=request.cookies("SearchEngine")%>" frameborder="0">
</frameset>


</html>