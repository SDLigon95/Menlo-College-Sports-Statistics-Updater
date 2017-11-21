<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:template match="/">
		<html>
			<head>
				<title>Activity Details Report</title>
				<link rel="stylesheet" type="text/css" href="reports.css"/>
			</head>
			<body>
				<xsl:apply-templates/>
			</body>
		</html>
	</xsl:template>
	<xsl:template match="report">
		<h1>Activity Details Report</h1>
		<table >
			<thead>
				<tr>
					<th></th>
					<th>Activity</th>
					<th>Date Time</th>
					<th class="numeric">Goal WPM</th>
					<th class="numeric">WPM</th>
					<th class="numeric">AWPM</th>
					<th class="numeric">Peak WPM</th>
					<th class="numeric">Acc.</th>
				</tr>
			</thead>
			<tbody>
				<xsl:apply-templates/>
			</tbody>
		</table>
	</xsl:template>
	<xsl:template match="user">
		<tr>
			<td colspan="8" class="user">
				<xsl:value-of select="@fullname"/>
			</td>
		</tr>
		<xsl:apply-templates/>
		<tr>
			<td colspan="8" >&#160;</td>
		</tr>
	</xsl:template>
	<xsl:template match="act">
		<tr>
			<td>&#160;</td>
			<td class="celldata">
				<xsl:choose>
				<xsl:when test="starts-with(@name, '$!')">
					<xsl:value-of select="@lu_name"/>&#160;&#160;
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="@name"/>&#160;&#160;
				</xsl:otherwise>
				</xsl:choose>
			</td>
			<td class="celldata">
				<xsl:value-of select="@completetime"/>
			</td>
			<td class="numeric celldata">
				<xsl:value-of select="@goalwpm"/>
			</td>
			<td class="numeric celldata">
				<xsl:value-of select="@wpm"/>
			</td>
			<td class="numeric celldata">
				<xsl:value-of select="@awpm"/>
			</td>
			<td class="numeric celldata">
				<xsl:value-of select="@peakwpm"/>
			</td>
			<td class="numeric celldata">
				<xsl:value-of select="@acc"/>%
			</td>
		</tr>
	</xsl:template>
</xsl:stylesheet>