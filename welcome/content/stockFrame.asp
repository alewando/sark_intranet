<!--
Developer:    Kevin Dill
Date:         08/17/1998
Description:  Main stock frame containing a header and the results.
-->


<html>

<head>
<title>SARK Intranet</title>
</head>


<frameset rows="34,*" frameborder=0 framespacing=0>
	<frame name="StockHeader" src="stockHeader.asp?stock=<%=request.form("stock")%>" marginheight="2" marginwidth="0" frameborder="0" scrolling="no" noresize>
	<frame name="StockResults" src="stockResults.asp?stock=<%=request.form("stock")%>" marginheight="0" marginwidth="0" frameborder="0">
</frameset>


</html>