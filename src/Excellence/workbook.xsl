<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">

	<!-- activate this to get well formated xml -->
	<!--<xsl:output omit-xml-declaration="yes" indent="yes"/>-->

	<xsl:template match="sheets">

		<!-- get id attribute and store it in $id -->
		<xsl:variable name="dimension">
			<xsl:value-of select="@dimension"/>
		</xsl:variable>

		<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
			<!-- set dimensions -->
			<dimension ref="{$dimension}"/>

			<!-- sheet view - activate first sheet -->
			<sheetViews>
				<sheetView tabSelected="1" workbookViewId="0" />
			</sheetViews>
			<sheetFormatPr baseColWidth="10" defaultRowHeight="16" x14ac:dyDescent="0"/>

			<!-- sheet data -->
			<sheetData>

				<!-- apply matching templates -->
				<xsl:apply-templates/>

			</sheetData>
		</worksheet>
	</xsl:template>

	<!-- rows -->
	<xsl:template match="row">

		<!-- get id attribute and store it in $id -->
		<xsl:variable name="id">
			<xsl:value-of select="@id"/>
		</xsl:variable>

		<row r="{$id}">

			<!-- apply matching templates -->
			<xsl:apply-templates/>

		</row>
	</xsl:template>

	<!-- columns -->
	<xsl:template match="column">

		<!-- get id attribute and store it in $id -->
		<xsl:variable name="id">
			<xsl:value-of select="@id"/>
		</xsl:variable>

		<!-- get type attribute and store it in $type -->
		<xsl:variable name="type">
			<xsl:value-of select="@type"/>
		</xsl:variable>

		<!-- string values -->
		<xsl:if test="$type='1'">
			<c r="{$id}" t="inlineStr">
				<is><t><xsl:value-of select="text()" /></t></is>
			</c>
		</xsl:if>

		<!-- number / float values -->
		<xsl:if test="$type='2'">
			<c r="{$id}" t="n">
				<v><xsl:value-of select="text()" /></v>
			</c>
		</xsl:if>

		<!-- number / float values -->
		<xsl:if test="$type='4'">
			<c r="{$id}">
				<f><xsl:value-of select="text()" /></f>
			</c>
		</xsl:if>


	</xsl:template>

</xsl:stylesheet>