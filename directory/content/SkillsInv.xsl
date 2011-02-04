<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
	<xsl:template match="/">
		<html>
			<form id="skills" name="skills" action="SkillsSubmit.asp" method="POST">
				<head>
					<title>Skills Inventory</title>
					<STYLE>
					    H3 { page-break-before: always }
					</STYLE>
				</head>
				<font face="ms sans serif,arial,helvetica" size="3">
					<body bgcolor="silver">
						<center>
							<p>
								<b>Skills Inventory</b>
							</p>
							<table border="0" cellspacing="0" cellpadding="0" width="460">
								<xsl:apply-templates select="Employee"/>
							</table>
						</center>
						<p>This document provides an assessment of consultant technical skills and should be used to facilitate discussion between the reviewer and consultant concerning the progress of technical education and in setting consultant goals.</p>
						<p>A preliminary assessment should be filled out by the consultant prior to the review. The reviewer should note any areas of change from previous reviews, discern areas for future development and make any rating adjustments. The reviewer should be prepared to discuss these items with the consultant at the time of the review.</p>
						<b><font color="maroon">Warning:  Do not use the "&amp;" character when entering an optional skill.</font></b>
						
						<p><b>Explanation of Ratings</b></p>
						<p>
							<table border="1" cellspacing="1" cellpadding="1" width="90%">
								<tr>
									<td width="30%" valign="top">
										<p>N/A</p>
									</td>
									<td width="60%" valign="top">
										<p>Consultant has had no exposure to the technology.</p>
									</td>
								</tr>
								<tr>
									<td width="30%" valign="top">
										<p>Basic Education</p>
									</td>
									<td width="60%" valign="top">
										<p>Consultant has attended a class in which the technology was taught or is self-taught, but has no actual work experience.</p>
									</td>
								</tr>
								<tr>
									<td width="30%" valign="top">
										<p>Some Experience</p>
									</td>
									<td width="60%" valign="top">
										<p>Consultant has had some exposure to the technology as it is applied in a working environment.</p>
									</td>
								</tr>
								<tr>
									<td width="30%" valign="top">
										<p>Productive w/minimal assistance</p>
									</td>
									<td width="60%" valign="top">
										<p>Consultant can be immediately productive in a working environment, with infrequent assistance.</p>
									</td>
								</tr>
								<tr>
									<td width="30%" valign="top">
										<p>Productive w/o assistance</p>
									</td>
									<td width="60%" valign="top">
										<p>Consultant can be immediately productive in a working environment, with no assistance.</p>
									</td>
								</tr>
								<tr>
									<td width="30%" valign="top">
										<p>Expert</p>
									</td>
									<td width="60%" valign="top">
										<p>Consultant has extensive, detailed working knowledge of the technology and can assist others with complex technical issues.</p>
									</td>
								</tr>
							</table>
						</p>

						<table border="0" cellspacing="0" cellpadding="0" width="460">
							<xsl:apply-templates select="Employee/SkillGroup"/>
						</table>
						<br/>
						<center>
							<xsl:element name="input">
								<xsl:attribute name="type">SUBMIT</xsl:attribute>
								<xsl:attribute name="name">Submit</xsl:attribute>
								<xsl:attribute name="class">button</xsl:attribute>
								<xsl:attribute name="value">Update Skills</xsl:attribute>
								<xsl:if test=".[/Employee/CanUpdate='False']">
									<xsl:attribute name="disabled">true</xsl:attribute>
								</xsl:if>
							</xsl:element>

							<xsl:element name="input">
								<xsl:attribute name="type">button</xsl:attribute>
								<xsl:attribute name="name">Close</xsl:attribute>
								<xsl:attribute name="class">button</xsl:attribute>
								<xsl:attribute name="value">Cancel</xsl:attribute>
								<xsl:attribute name="onclick">window.close()</xsl:attribute>
							</xsl:element>
						</center>
					</body>
				</font>
			</form>
		</html>
	</xsl:template>

	<xsl:template match="Employee">
		<font face="ms sans serif, arial, helvetica" font-size="3">
			<tr>
				<td width="160" valign="top">
					<b> Employee Name </b>
				</td>
				<td width="300" valign="top">
					<xsl:value-of select="FirstName"/>
					<xsl:value-of select="LastName"/>
				</td>
			</tr>
			<tr>
				<td width="160" valign="top">
					<b> Date Modified </b>
				</td>
				<td width="300" valign="top">
					<xsl:value-of select="DateSkillsModified"/>
				</td>
			</tr>
		</font>
	</xsl:template>

	<xsl:template match="SkillGroup">
		<font face="ms sans serif, arial, helvetica" font-size="3">
			<tr>
				<tr/>
				<table border="0" width="460" cellpadding="0">
					<xsl:if match=".[PageBreak='1']">
      					<td>
							<H3> </H3>
						</td>
					</xsl:if>
				</table>
				<td valign="top" width="460">				
					<br/>
					<b>
					<xsl:value-of select="GroupName"/>
					</b>
				</td>
				<tr>
					<td valign="top" width="460">
						<table border="1" width="460" cellpadding="0">
							<xsl:apply-templates select="./Skill/SkillName"/>
						</table>
					</td>
				</tr>
			</tr>
		</font>
	</xsl:template>

	<xsl:template match="SkillName">
		<font face="ms sans serif, arial, helvetica" font-size="3">
			<tr>
				<td valign="top" width="380">
					<xsl:choose>
						<xsl:when test="..[isOtherSkill = 'True']">
							<xsl:element name="input">
								<xsl:attribute name="type">text</xsl:attribute>
								<xsl:attribute name="name">OtherSkill</xsl:attribute>
								<xsl:attribute name="size">60</xsl:attribute>
								<xsl:attribute name="value"><xsl:value-of select="../SkillValue" /></xsl:attribute>
								<!--../../.. = the Employee node-->
								<xsl:if test="../../..[./CanUpdate='False']">
									<xsl:attribute name="disabled">true</xsl:attribute>
								</xsl:if>
							</xsl:element>
						</xsl:when>
						<xsl:otherwise>
							<xsl:element name="input">
								<xsl:attribute name="type">Hidden</xsl:attribute>
								<xsl:attribute name="name">OtherSkill</xsl:attribute>
								<xsl:attribute name="size">60</xsl:attribute>
								<xsl:attribute name="value"><xsl:value-of select="../SkillValue" /></xsl:attribute>
								<!--../../.. = the Employee node-->
								<xsl:if test="../../..[./CanUpdate='False']">
									<xsl:attribute name="disabled">true</xsl:attribute>
								</xsl:if>
							</xsl:element>
							<xsl:value-of />
						</xsl:otherwise>
					</xsl:choose>
				</td>
				<td valign="top" width="80">
					<xsl:apply-templates select="../SkillID"/>
				</td>
			</tr>
		</font>
	</xsl:template>

	<xsl:template match="SkillID">
		<input>
			<xsl:attribute name="type">hidden</xsl:attribute>
			<xsl:attribute name="name">SkillID</xsl:attribute>
			<xsl:attribute name="value"><xsl:value-of /></xsl:attribute>
			<!--../../.. = the Employee node-->
			<xsl:if test="../../..[./CanUpdate='False']">
				<xsl:attribute name="disabled">true</xsl:attribute>
			</xsl:if>
			<select>
				<xsl:attribute name="name">SkillRanking</xsl:attribute>
				<!--../../.. = the Employee node-->
				<xsl:if test="../../..[./CanUpdate='False']">
					<xsl:attribute name="disabled">true</xsl:attribute>
				</xsl:if>
				<option>
					<xsl:attribute name="value">1</xsl:attribute>
					<xsl:if test="..[SkillRanking='1']">
						<xsl:attribute name="selected">True</xsl:attribute>
					</xsl:if>
					N/A
				</option>
				<option>
					<xsl:attribute name="value">2</xsl:attribute>
					<xsl:if test="..[SkillRanking='2']">
						<xsl:attribute name="selected">True</xsl:attribute>
					</xsl:if>
					Basic Education
				</option>
				<option>
					<xsl:attribute name="value">3</xsl:attribute>
					<xsl:if test="..[SkillRanking='3']">
						<xsl:attribute name="selected">True</xsl:attribute>
					</xsl:if>
					Some Experience
				</option>
				<option>
					<xsl:attribute name="value">4</xsl:attribute>
					<xsl:if test="..[SkillRanking='4']">
						<xsl:attribute name="selected">True</xsl:attribute>
					</xsl:if>
					Productive w/minimal assistance
				</option>
				<option>
					<xsl:attribute name="value">5</xsl:attribute>
					<xsl:if test="..[SkillRanking='5']">
						<xsl:attribute name="selected">True</xsl:attribute>
					</xsl:if>
					Productive w/o assistance
				</option>
				<option>
					<xsl:attribute name="value">6</xsl:attribute>
					<xsl:if test="..[SkillRanking='6']">
						<xsl:attribute name="selected">True</xsl:attribute>
					</xsl:if>
					Expert
				</option>
			</select>
		</input>
	</xsl:template>
</xsl:stylesheet>
