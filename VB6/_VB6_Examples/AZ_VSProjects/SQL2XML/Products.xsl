<?xml version="1.0" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

<xsl:template match="/">
	<xsl:apply-templates select="Root" />
</xsl:template>

<xsl:template match="Root">
	<TABLE WIDTH="75%" ALIGN="LEFT" BORDER="1">
		<TR>
			<TD ALIGN="LEFT" BGCOLOR="#0033CC">
				<FONT SIZE="3" COLOR="#FFFFFF">Product Name</FONT>
			</TD>
			<TD ALIGN="CENTER" BGCOLOR="#0033CC">
				<FONT SIZE="3" COLOR="#FFFFFF">Qty Per Unit</FONT>
			</TD>
			<TD ALIGN="CENTER" BGCOLOR="#0033CC">
				<FONT SIZE="3" COLOR="#FFFFFF">Units In Stock</FONT>
			</TD>
		</TR>
		<xsl:apply-templates select="Products" />
	</TABLE>
</xsl:template>

<xsl:template match="Products">
	<TR>
			<TD ALIGN="LEFT">
				<FONT SIZE="2">
					<xsl:value-of select="ProductName" />
				</FONT>
			</TD>
			<TD ALIGN="CENTER">
				<FONT SIZE="2">
					<xsl:value-of select="QuantityPerUnit" />
				</FONT>
			</TD>
			<TD ALIGN="CENTER">
				<FONT SIZE="2">
					<xsl:value-of select="UnitsInStock" />
				</FONT>
			</TD>
	</TR>
</xsl:template>

</xsl:stylesheet>

