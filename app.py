<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:output method="text" omit-xml-declaration="yes" encoding="ISO-8859-1"/>

  <!-- Prevent whitespace nodes from turning into blank output -->
  <xsl:strip-space elements="*"/>
  <xsl:template match="text()"/>

  <!-- ========================= -->
  <!-- CSV DEFAULT EXTENSION -->
  <!-- ========================= -->
  <xsl:variable name="fileExt" select="'csv'"/>

  <!-- ========================= -->
  <!-- Decimal formatting -->
  <!-- ========================= -->
  <xsl:variable name="DecPt" select="'.'"/>
  <xsl:decimal-format name="Standard"
                      decimal-separator="."
                      grouping-separator=","
                      infinity="Infinity"
                      minus-sign="-"
                      NaN=""
                      percent="%"
                      per-mille="&#2030;"
                      zero-digit="0"
                      digit="#"
                      pattern-separator=";" />

  <xsl:variable name="DecPl0" select="'#0'"/>
  <xsl:variable name="DecPl1" select="concat('#0', $DecPt, '0')"/>
  <xsl:variable name="DecPl2" select="concat('#0', $DecPt, '00')"/>
  <xsl:variable name="DecPl3" select="concat('#0', $DecPt, '000')"/>
  <xsl:variable name="DecPl4" select="concat('#0', $DecPt, '0000')"/>
  <xsl:variable name="DecPl5" select="concat('#0', $DecPt, '00000')"/>
  <xsl:variable name="DecPl8" select="concat('#0', $DecPt, '00000000')"/>

  <!-- User fields (kept as-is, even though we now sort by TimeStamp) -->
  <xsl:variable name="userField1" select="'sortType|Point name sort|stringMenu|2|Numerical|Alphabetical'"/>
  <xsl:variable name="sortType" select="'Numerical'"/>

  <!-- Same key you are using in the working file -->
  <xsl:key name="obsID-search" match="/JOBFile/FieldBook/PointRecord" use="@ID"/>

  <!-- Environment -->
  <xsl:variable name="DistUnit"   select="/JOBFile/Environment/DisplaySettings/DistanceUnits" />
  <xsl:variable name="CoordOrder" select="/JOBFile/Environment/DisplaySettings/CoordinateOrder" />

  <!-- ========================= -->
  <!-- FORCE US SURVEY FEET OUTPUT -->
  <!-- ========================= -->
  <!-- JobXML values are metres; convert to US Survey Feet -->
  <xsl:variable name="DistConvFactor" select="3.2808333333357"/>

  <!-- Coordinate order boolean -->
  <xsl:variable name="NECoords">
    <xsl:choose>
      <xsl:when test="$CoordOrder='North-East-Elevation'">true</xsl:when>
      <xsl:when test="$CoordOrder='X-Y-Z'">true</xsl:when>
      <xsl:otherwise>false</xsl:otherwise>
    </xsl:choose>
  </xsl:variable>

  <!-- ========================= -->
  <!-- New line (Trimble already normalizes this; DO NOT output CRLF here) -->
  <!-- ========================= -->
  <xsl:template name="NewLine">
    <xsl:text>&#10;</xsl:text>
  </xsl:template>

  <!-- ========================= -->
  <!-- CSV-safe quoting -->
  <!-- ========================= -->
  <xsl:template name="CsvQuote">
    <xsl:param name="s"/>
    <xsl:text>"</xsl:text>
    <xsl:call-template name="ReplaceQuotes">
      <xsl:with-param name="s" select="$s"/>
    </xsl:call-template>
    <xsl:text>"</xsl:text>
  </xsl:template>

  <xsl:template name="ReplaceQuotes">
    <xsl:param name="s"/>
    <xsl:choose>
      <xsl:when test="contains($s, '&quot;')">
        <xsl:value-of select="substring-before($s, '&quot;')"/>
        <xsl:text>""</xsl:text>
        <xsl:call-template name="ReplaceQuotes">
          <xsl:with-param name="s" select="substring-after($s, '&quot;')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$s"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- ========================= -->
  <!-- MAIN: Header row ONLY -->
  <!-- ========================= -->
  <xsl:template match="/">
    <xsl:choose>
      <xsl:when test="$NECoords">
        <xsl:text>Point,North,East,Elev,Code,Hz Prec,Vt Prec,PDOP,Satellites</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>Point,East,North,Elev,Code,Hz Prec,Vt Prec,PDOP,Satellites</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:call-template name="NewLine"/>

    <xsl:apply-templates select="JOBFile/Reductions"/>
  </xsl:template>

  <!-- ========================= -->
  <!-- Reductions (NOW SORTED BY SHOT TIME) -->
  <!-- ========================= -->
  <xsl:template match="Reductions">
    <xsl:apply-templates select="Point">
      <!-- Primary sort: observation timestamp from linked PointRecord -->
      <xsl:sort data-type="text" order="ascending" select="key('obsID-search', ID)[1]/@TimeStamp"/>
      <!-- Secondary sort to keep stable output when timestamps match -->
      <xsl:sort data-type="text" order="ascending" select="Name"/>
    </xsl:apply-templates>
  </xsl:template>

  <!-- ========================= -->
  <!-- Point -> CSV Row (skip blank rows) -->
  <!-- ========================= -->
  <xsl:template match="Point">

    <!-- FINAL PROCESSED COORDINATES:
         Prefer ComputedGrid; if missing, fall back to Grid -->
    <xsl:variable name="FinalGrid" select="(ComputedGrid | Grid)[1]"/>

    <!-- Skip rows that have no point name AND no FINAL grid coordinates -->
    <xsl:if test="normalize-space(Name) != '' or normalize-space(string($FinalGrid/North)) != '' or normalize-space(string($FinalGrid/East)) != '' or normalize-space(string($FinalGrid/Elevation)) != ''">

      <xsl:variable name="HorizPrec">
        <xsl:for-each select="key('obsID-search', ID)">
          <xsl:value-of select="Precision/Horizontal"/>
        </xsl:for-each>
      </xsl:variable>

      <xsl:variable name="VertPrec">
        <xsl:for-each select="key('obsID-search', ID)">
          <xsl:value-of select="Precision/Vertical"/>
        </xsl:for-each>
      </xsl:variable>

      <xsl:variable name="tempPDOP">
        <xsl:for-each select="key('obsID-search', ID)">
          <xsl:value-of select="QualityControl1/PDOP"/>
        </xsl:for-each>
      </xsl:variable>

      <xsl:variable name="tempNbrSat">
        <xsl:for-each select="key('obsID-search', ID)">
          <xsl:value-of select="QualityControl1/NumberOfSatellites"/>
        </xsl:for-each>
      </xsl:variable>

      <!-- Point (quoted) -->
      <xsl:call-template name="CsvQuote">
        <xsl:with-param name="s" select="Name"/>
      </xsl:call-template>
      <xsl:text>,</xsl:text>

      <!-- North/East order based on NECoords -->
      <xsl:choose>
        <xsl:when test="$NECoords">
          <xsl:value-of select="format-number($FinalGrid/North * $DistConvFactor, $DecPl3, 'Standard')"/>
          <xsl:text>,</xsl:text>
          <xsl:value-of select="format-number($FinalGrid/East * $DistConvFactor, $DecPl3, 'Standard')"/>
          <xsl:text>,</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="format-number($FinalGrid/East * $DistConvFactor, $DecPl3, 'Standard')"/>
          <xsl:text>,</xsl:text>
          <xsl:value-of select="format-number($FinalGrid/North * $DistConvFactor, $DecPl3, 'Standard')"/>
          <xsl:text>,</xsl:text>
        </xsl:otherwise>
      </xsl:choose>

      <!-- Elev -->
      <xsl:value-of select="format-number($FinalGrid/Elevation * $DistConvFactor, $DecPl3, 'Standard')"/>
      <xsl:text>,</xsl:text>

      <!-- Code (quoted) -->
      <xsl:call-template name="CsvQuote">
        <xsl:with-param name="s" select="Code"/>
      </xsl:call-template>
      <xsl:text>,</xsl:text>

      <!-- Hz Prec, Vt Prec, PDOP, Satellites -->
      <xsl:value-of select="format-number($HorizPrec * $DistConvFactor, $DecPl3, 'Standard')"/>
      <xsl:text>,</xsl:text>
      <xsl:value-of select="format-number($VertPrec * $DistConvFactor, $DecPl3, 'Standard')"/>
      <xsl:text>,</xsl:text>
      <xsl:value-of select="format-number($tempPDOP, $DecPl1, 'Standard')"/>
      <xsl:text>,</xsl:text>
      <xsl:value-of select="format-number($tempNbrSat, $DecPl0, 'Standard')"/>

      <xsl:call-template name="NewLine"/>

    </xsl:if>
  </xsl:template>

</xsl:stylesheet>
