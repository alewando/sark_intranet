<%'response.buffer = true
Const ACTION_SEARCH = "SEARCH"
Const DelimiterString = " "
'The Skills tables use char data types instead of varchar and therefore alwasy have the
' column values space filled. This means if a value is not supplied. it has 60 spaces.
Const emptySkillValue ="                                                            "

%>
<!--
Developer:    Dave Podnar
Date:         9/5/2000
Description: Main script for searching the consultant skills inventory information.

Scripts: Search_Skills.asp = Search skills inventory for consultants
-->

<!-- #include file="../section.asp" -->

<%
sql = "select employeeid, " _
& "sum(CASE skillid WHEN 3 THEN skillranking ELSE 0 END) as r3," _
& "sum(CASE skillid WHEN 4 THEN skillranking ELSE 0 END) as r4," _
& "sum(CASE skillid WHEN 5 THEN skillranking ELSE 0 END) as r5," _
& "sum(CASE skillid WHEN 6 THEN skillranking ELSE 0 END) as r6," _
& "sum(CASE skillid WHEN 7 THEN skillranking ELSE 0 END) as r7," _
& "sum(CASE skillid WHEN 8 THEN skillranking ELSE 0 END) as r8," _
& "sum(CASE skillid WHEN 9 THEN skillranking ELSE 0 END) as r9," _
& "sum(CASE skillid WHEN 10 THEN skillranking ELSE 0 END) as r10," _
& "sum(CASE skillid WHEN 11 THEN skillranking ELSE 0 END) as r11," _
& "sum(CASE skillid WHEN 12 THEN skillranking ELSE 0 END) as r12," _
& "sum(CASE skillid WHEN 13 THEN skillranking ELSE 0 END) as r13," _
& "sum(CASE skillid WHEN 14 THEN skillranking ELSE 0 END) as r14," _
& "sum(CASE skillid WHEN 15 THEN skillranking ELSE 0 END) as r15," _
& "sum(CASE skillid WHEN 16 THEN skillranking ELSE 0 END) as r16," _
& "sum(CASE skillid WHEN 17 THEN skillranking ELSE 0 END) as r17," _
& "sum(CASE skillid WHEN 18 THEN skillranking ELSE 0 END) as r18," _
& "sum(CASE skillid WHEN 19 THEN skillranking ELSE 0 END) as r19," _
& "sum(CASE skillid WHEN 20 THEN skillranking ELSE 0 END) as r20," _
& "sum(CASE skillid WHEN 21 THEN skillranking ELSE 0 END) as r21," _
& "sum(CASE skillid WHEN 22 THEN skillranking ELSE 0 END) as r22," _
& "sum(CASE skillid WHEN 23 THEN skillranking ELSE 0 END) as r23," _
& "sum(CASE skillid WHEN 24 THEN skillranking ELSE 0 END) as r24," _
& "sum(CASE skillid WHEN 25 THEN skillranking ELSE 0 END) as r25," _
& "sum(CASE skillid WHEN 26 THEN skillranking ELSE 0 END) as r26," _
& "sum(CASE skillid WHEN 27 THEN skillranking ELSE 0 END) as r27," _
& "sum(CASE skillid WHEN 28 THEN skillranking ELSE 0 END) as r28," _
& "sum(CASE skillid WHEN 29 THEN skillranking ELSE 0 END) as r29," _
& "sum(CASE skillid WHEN 30 THEN skillranking ELSE 0 END) as r30," _
& "sum(CASE skillid WHEN 31 THEN skillranking ELSE 0 END) as r31," _
& "sum(CASE skillid WHEN 32 THEN skillranking ELSE 0 END) as r32," _
& "sum(CASE skillid WHEN 33 THEN skillranking ELSE 0 END) as r33," _
& "sum(CASE skillid WHEN 34 THEN skillranking ELSE 0 END) as r34," _
& "sum(CASE skillid WHEN 35 THEN skillranking ELSE 0 END) as r35," _
& "sum(CASE skillid WHEN 36 THEN skillranking ELSE 0 END) as r36," _
& "sum(CASE skillid WHEN 37 THEN skillranking ELSE 0 END) as r37," _
& "sum(CASE skillid WHEN 38 THEN skillranking ELSE 0 END) as r38," _
& "sum(CASE skillid WHEN 39 THEN skillranking ELSE 0 END) as r39," _
& "sum(CASE skillid WHEN 40 THEN skillranking ELSE 0 END) as r40," _
& "sum(CASE skillid WHEN 41 THEN skillranking ELSE 0 END) as r41," _
& "sum(CASE skillid WHEN 42 THEN skillranking ELSE 0 END) as r42," _
& "sum(CASE skillid WHEN 43 THEN skillranking ELSE 0 END) as r43," _
& "sum(CASE skillid WHEN 44 THEN skillranking ELSE 0 END) as r44," _
& "sum(CASE skillid WHEN 45 THEN skillranking ELSE 0 END) as r45," _
& "sum(CASE skillid WHEN 46 THEN skillranking ELSE 0 END) as r46," _
& "sum(CASE skillid WHEN 47 THEN skillranking ELSE 0 END) as r47," _
& "sum(CASE skillid WHEN 48 THEN skillranking ELSE 0 END) as r48," _
& "sum(CASE skillid WHEN 49 THEN skillranking ELSE 0 END) as r49," _
& "sum(CASE skillid WHEN 50 THEN skillranking ELSE 0 END) as r50," _
& "sum(CASE skillid WHEN 51 THEN skillranking ELSE 0 END) as r51," _
& "sum(CASE skillid WHEN 52 THEN skillranking ELSE 0 END) as r52," _
& "sum(CASE skillid WHEN 53 THEN skillranking ELSE 0 END) as r53," _
& "sum(CASE skillid WHEN 54 THEN skillranking ELSE 0 END) as r54," _
& "sum(CASE skillid WHEN 55 THEN skillranking ELSE 0 END) as r55," _
& "sum(CASE skillid WHEN 56 THEN skillranking ELSE 0 END) as r56," _
& "sum(CASE skillid WHEN 57 THEN skillranking ELSE 0 END) as r57," _
& "sum(CASE skillid WHEN 58 THEN skillranking ELSE 0 END) as r58," _
& "sum(CASE skillid WHEN 59 THEN skillranking ELSE 0 END) as r59," _
& "sum(CASE skillid WHEN 60 THEN skillranking ELSE 0 END) as r60," _
& "sum(CASE skillid WHEN 61 THEN skillranking ELSE 0 END) as r61," _
& "sum(CASE skillid WHEN 62 THEN skillranking ELSE 0 END) as r62," _
& "sum(CASE skillid WHEN 63 THEN skillranking ELSE 0 END) as r63," _
& "sum(CASE skillid WHEN 64 THEN skillranking ELSE 0 END) as r64," _
& "sum(CASE skillid WHEN 65 THEN skillranking ELSE 0 END) as r65," _
& "sum(CASE skillid WHEN 66 THEN skillranking ELSE 0 END) as r66," _
& "sum(CASE skillid WHEN 67 THEN skillranking ELSE 0 END) as r67," _
& "sum(CASE skillid WHEN 68 THEN skillranking ELSE 0 END) as r68," _
& "sum(CASE skillid WHEN 69 THEN skillranking ELSE 0 END) as r69," _
& "sum(CASE skillid WHEN 70 THEN skillranking ELSE 0 END) as r70," _
& "sum(CASE skillid WHEN 71 THEN skillranking ELSE 0 END) as r71," _
& "sum(CASE skillid WHEN 72 THEN skillranking ELSE 0 END) as r72," _
& "sum(CASE skillid WHEN 73 THEN skillranking ELSE 0 END) as r73," _
& "sum(CASE skillid WHEN 74 THEN skillranking ELSE 0 END) as r74," _
& "sum(CASE skillid WHEN 75 THEN skillranking ELSE 0 END) as r75," _
& "sum(CASE skillid WHEN 76 THEN skillranking ELSE 0 END) as r76," _
& "sum(CASE skillid WHEN 77 THEN skillranking ELSE 0 END) as r77," _
& "sum(CASE skillid WHEN 78 THEN skillranking ELSE 0 END) as r78," _
& "sum(CASE skillid WHEN 79 THEN skillranking ELSE 0 END) as r79," _
& "sum(CASE skillid WHEN 80 THEN skillranking ELSE 0 END) as r80," _
& "sum(CASE skillid WHEN 81 THEN skillranking ELSE 0 END) as r81," _
& "sum(CASE skillid WHEN 82 THEN skillranking ELSE 0 END) as r82," _
& "sum(CASE skillid WHEN 83 THEN skillranking ELSE 0 END) as r83," _
& "sum(CASE skillid WHEN 84 THEN skillranking ELSE 0 END) as r84," _
& "sum(CASE skillid WHEN 85 THEN skillranking ELSE 0 END) as r85," _
& "sum(CASE skillid WHEN 86 THEN skillranking ELSE 0 END) as r86," _
& "sum(CASE skillid WHEN 87 THEN skillranking ELSE 0 END) as r87," _
& "sum(CASE skillid WHEN 88 THEN skillranking ELSE 0 END) as r88," _
& "sum(CASE skillid WHEN 89 THEN skillranking ELSE 0 END) as r89," _
& "sum(CASE skillid WHEN 90 THEN skillranking ELSE 0 END) as r90," _
& "sum(CASE skillid WHEN 91 THEN skillranking ELSE 0 END) as r91," _
& "sum(CASE skillid WHEN 92 THEN skillranking ELSE 0 END) as r92," _
& "sum(CASE skillid WHEN 93 THEN skillranking ELSE 0 END) as r93," _
& "sum(CASE skillid WHEN 94 THEN skillranking ELSE 0 END) as r94," _
& "sum(CASE skillid WHEN 95 THEN skillranking ELSE 0 END) as r95," _
& "sum(CASE skillid WHEN 96 THEN skillranking ELSE 0 END) as r96," _
& "sum(CASE skillid WHEN 97 THEN skillranking ELSE 0 END) as r97," _
& "sum(CASE skillid WHEN 98 THEN skillranking ELSE 0 END) as r98," _
& "sum(CASE skillid WHEN 99 THEN skillranking ELSE 0 END) as r99," _
& "sum(CASE skillid WHEN 100 THEN skillranking ELSE 0 END) as r100," _
& "sum(CASE skillid WHEN 101 THEN skillranking ELSE 0 END) as r101," _
& "sum(CASE skillid WHEN 102 THEN skillranking ELSE 0 END) as r102," _
& "sum(CASE skillid WHEN 103 THEN skillranking ELSE 0 END) as r103," _
& "sum(CASE skillid WHEN 104 THEN skillranking ELSE 0 END) as r104," _
& "sum(CASE skillid WHEN 105 THEN skillranking ELSE 0 END) as r105," _
& "sum(CASE skillid WHEN 106 THEN skillranking ELSE 0 END) as r106," _
& "sum(CASE skillid WHEN 107 THEN skillranking ELSE 0 END) as r107," _
& "sum(CASE skillid WHEN 108 THEN skillranking ELSE 0 END) as r108," _
& "sum(CASE skillid WHEN 109 THEN skillranking ELSE 0 END) as r109," _
& "sum(CASE skillid WHEN 110 THEN skillranking ELSE 0 END) as r110," _
& "sum(CASE skillid WHEN 111 THEN skillranking ELSE 0 END) as r111," _
& "sum(CASE skillid WHEN 112 THEN skillranking ELSE 0 END) as r112," _
& "sum(CASE skillid WHEN 113 THEN skillranking ELSE 0 END) as r113," _
& "sum(CASE skillid WHEN 114 THEN skillranking ELSE 0 END) as r114," _
& "sum(CASE skillid WHEN 115 THEN skillranking ELSE 0 END) as r115," _
& "sum(CASE skillid WHEN 116 THEN skillranking ELSE 0 END) as r116," _
& "sum(CASE skillid WHEN 117 THEN skillranking ELSE 0 END) as r117," _
& "sum(CASE skillid WHEN 118 THEN skillranking ELSE 0 END) as r118," _
& "sum(CASE skillid WHEN 119 THEN skillranking ELSE 0 END) as r119," _
& "sum(CASE skillid WHEN 120 THEN skillranking ELSE 0 END) as r120," _
& "sum(CASE skillid WHEN 121 THEN skillranking ELSE 0 END) as r121," _
& "sum(CASE skillid WHEN 122 THEN skillranking ELSE 0 END) as r122," _
& "sum(CASE skillid WHEN 123 THEN skillranking ELSE 0 END) as r123," _
& "sum(CASE skillid WHEN 124 THEN skillranking ELSE 0 END) as r124," _
& "sum(CASE skillid WHEN 125 THEN skillranking ELSE 0 END) as r125," _
& "sum(CASE skillid WHEN 126 THEN skillranking ELSE 0 END) as r126," _
& "sum(CASE skillid WHEN 127 THEN skillranking ELSE 0 END) as r127," _
& "sum(CASE skillid WHEN 128 THEN skillranking ELSE 0 END) as r128," _
& "sum(CASE skillid WHEN 129 THEN skillranking ELSE 0 END) as r129," _
& "sum(CASE skillid WHEN 130 THEN skillranking ELSE 0 END) as r130," _
& "sum(CASE skillid WHEN 131 THEN skillranking ELSE 0 END) as r131," _
& "sum(CASE skillid WHEN 132 THEN skillranking ELSE 0 END) as r132 " _
& "from skills_inventory  group by employeeid order by 1"
	  
set rs = DBQuery(sql)
first=True
%>
<TABLE BORDER=0 WIDTH=100%>

<%
While Not rs.EOF
	If first Then
		first=False
		Response.Write "<TR>"
		For i = 0 to rs.Fields.Count - 1
			Response.Write "<TD>" & rs(i).Name & ",</TD>"
		Next
		Response.Write "</TR>"
	End If
	
	Response.Write "<TR>"
	For i = 0 to rs.Fields.Count - 1
        Response.Write "<TD VALIGN=TOP>" & rs(i) & ",</TD>"
	Next
	Response.Write "</TR>"

	rs.movenext	
wend
rs.close
%>	
</TABLE>
	<!-- #include file="../../footer.asp" -->


