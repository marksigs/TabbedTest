<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="node()|@*">
		<xsl:copy>
			<xsl:apply-templates select="node()|@*"/>
		</xsl:copy>
	</xsl:template>
	<xsl:template match="@GENDER">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='M'">1</xsl:when><xsl:when test=".='F'">2</xsl:when><xsl:when test=".='X'">4</xsl:when></xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@ADDRESSTYPE">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='H'">1</xsl:when><xsl:when test=".='M'">2</xsl:when><xsl:when test=".='P'">3</xsl:when></xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@COUNTRY">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='UK'">1</xsl:when><xsl:when test=".=''">1</xsl:when></xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@OTHERSYSTEMTYPE">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='10'">10</xsl:when></xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@USAGE">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='R'">1</xsl:when><xsl:when test=".='B'">2</xsl:when><xsl:when test=".='C'">3</xsl:when><xsl:when test=".='F'">4</xsl:when></xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@CUSTOMERCATEGORY">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='DUP'">10</xsl:when><xsl:when test=".='DEC'">20</xsl:when><xsl:when test=".='BC'">30</xsl:when><xsl:when test=".='BNK'">40</xsl:when></xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@TITLE">
		<xsl:attribute name="{name()}">
			<xsl:choose>
				<xsl:when test=".='Mr'">1</xsl:when>
				<xsl:when test=".='Mr.'">1</xsl:when>
				<xsl:when test=".='Mrs'">2</xsl:when>
				<xsl:when test=".='Mrs.'">2</xsl:when>
				<xsl:when test=".='Miss'">3</xsl:when>
				<xsl:when test=".='Ms'">4</xsl:when>
				<xsl:when test=".='Ms.'">4</xsl:when>
				<xsl:when test=".='Dr'">5</xsl:when>
				<xsl:when test=".='Dr.'">5</xsl:when>
				<xsl:when test=".='Rev.'">53</xsl:when>
				<xsl:when test=".='Rev'">53</xsl:when>
				<xsl:when test=".='Adml.'">6</xsl:when>
				<xsl:when test=".='Adml'">6</xsl:when>
				<xsl:when test=".='Air_Cm'">7</xsl:when>
				<xsl:when test=".='Air_M.'">8</xsl:when>
				<xsl:when test=".='Air_M'">8</xsl:when>
				<xsl:when test=".='Air_Vm.'">9</xsl:when>
				<xsl:when test=".='Air_Vm'">9</xsl:when>
				<xsl:when test=".='R/Adml'">52</xsl:when>
				<xsl:when test=".='Prof'">51</xsl:when>
				<xsl:when test=".='Private'">50</xsl:when>
				<xsl:when test=".='Princess'">49</xsl:when>
				<xsl:when test=".='Prince'">48</xsl:when>
				<xsl:when test=".='Plt_Off'">47</xsl:when>
				<xsl:when test=".='Mraf'">46</xsl:when>
				<xsl:when test=".='Master'">45</xsl:when>
				<xsl:when test=".='Marquis'">44</xsl:when>
				<xsl:when test=".='Marquess'">43</xsl:when>
				<xsl:when test=".='Major'">42</xsl:when>
				<xsl:when test=".='Madam'">41</xsl:when>
				<xsl:when test=".='Lt_Col.'">40</xsl:when>
				<xsl:when test=".='Lt_Col'">40</xsl:when>
				<xsl:when test=".='Lt.'">39</xsl:when>
				<xsl:when test=".='Lord'">38</xsl:when>
				<xsl:when test=".='Lady'">37</xsl:when>
				<xsl:when test=".='L/Cpl.'">36</xsl:when>
				<xsl:when test=".='L/Cpl'">36</xsl:when>
				<xsl:when test=".='Judge'">35</xsl:when>
				<xsl:when test=".='HRH'">34</xsl:when>
				<xsl:when test=".='Hon_Mrs.'">33</xsl:when>
				<xsl:when test=".='Hon_Mrs'">33</xsl:when>
				<xsl:when test=".='Hon_Mr.'">32</xsl:when>
				<xsl:when test=".='Hon_Mr'">32</xsl:when>
				<xsl:when test=".='Grp_Capt.'">31</xsl:when>
				<xsl:when test=".='Grp_Capt'">31</xsl:when>
				<xsl:when test=".='Gen.'">30</xsl:when>
				<xsl:when test=".='Gen'">30</xsl:when>
				<xsl:when test=".='Flt_Lt.'">29</xsl:when>
				<xsl:when test=".='Flt_Lt'">29</xsl:when>
				<xsl:when test=".='Fg_Off.'">28</xsl:when>
				<xsl:when test=".='Fg_Off'">28</xsl:when>
				<xsl:when test=".='Father'">27</xsl:when>
				<xsl:when test=".='F/Sgt.'">26</xsl:when>
				<xsl:when test=".='F/Sgt'">26</xsl:when>
				<xsl:when test=".='Earl'">25</xsl:when>
				<xsl:when test=".='Duke'">24</xsl:when>
				<xsl:when test=".='Duchess'">23</xsl:when>
				<xsl:when test=".='Dame'">22</xsl:when>
				<xsl:when test=".='Cpl.'">21</xsl:when>
				<xsl:when test=".='Cpl'">21</xsl:when>
				<xsl:when test=".='Countess'">20</xsl:when>
				<xsl:when test=".='Count'">19</xsl:when>
				<xsl:when test=".='Col.'">18</xsl:when>
				<xsl:when test=".='Col'">18</xsl:when>
				<xsl:when test=".='Cmdr.'">17</xsl:when>
				<xsl:when test=".='Cmdr'">17</xsl:when>
				<xsl:when test=".='Cdr.'">16</xsl:when>
				<xsl:when test=".='Cdr'">16</xsl:when>
				<xsl:when test=".='Capt.'">15</xsl:when>
				<xsl:when test=".='Capt'">15</xsl:when>
				<xsl:when test=".='Canon'">14</xsl:when>
				<xsl:when test=".='C/Sgt.'">13</xsl:when>
				<xsl:when test=".='C/Sgt'">13</xsl:when>
				<xsl:when test=".='Brig.'">12</xsl:when>
				<xsl:when test=".='Brig'">12</xsl:when>
				<xsl:when test=".='Baroness'">11</xsl:when>
				<xsl:when test=".='Baron'">10</xsl:when>
				<xsl:when test=".='Wo.'">15</xsl:when>
				<xsl:when test=".='Wo'">15</xsl:when>
				<xsl:when test=".='Wg._Cmdr.'">68</xsl:when>
				<xsl:when test=".='Wg Cmdr'">68</xsl:when>
				<xsl:when test=".='Viscount'">67</xsl:when>
				<xsl:when test=".='V/Adml.'">66</xsl:when>
				<xsl:when test=".='V/Adml'">66</xsl:when>
				<xsl:when test=".='The_Rt_Hon'">65</xsl:when>
				<xsl:when test=".='The_Hon'">64</xsl:when>
				<xsl:when test=".='Sqn._Ldr.'">63</xsl:when>
				<xsl:when test=".='Sqn_Ldr'">63</xsl:when>
				<xsl:when test=".='Sister'">62</xsl:when>
				<xsl:when test=".='Sir'">61</xsl:when>
				<xsl:when test=".='Sheriff'">60</xsl:when>
				<xsl:when test=".='Sheikh'">59</xsl:when>
				<xsl:when test=".='Sgt.'">58</xsl:when>
				<xsl:when test=".='Sgt'">58</xsl:when>
				<xsl:when test=".='Sgn._Capt.'">57</xsl:when>
				<xsl:when test=".='Sac.'">56</xsl:when>
				<xsl:when test=".='S/Sgt.'">55</xsl:when>
				<xsl:when test=".='S/Sgt'">55</xsl:when>
				<xsl:when test=".='Capt'">15</xsl:when>
				<xsl:when test=".='Rev._Dr.'">54</xsl:when>
				<xsl:when test=".='Rev_Dr'">54</xsl:when>				
				<xsl:when test=".='Other'">99</xsl:when>
				</xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@SPECIALNEEDS">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='0'">1</xsl:when><xsl:when test=".='L'">2</xsl:when><xsl:when test=".='A'">3</xsl:when><xsl:when test=".='B'">4</xsl:when></xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@PROSPECTPASSWORDEXISTS">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='Y'">1</xsl:when><xsl:when test=".='N'">0</xsl:when></xsl:choose></xsl:attribute>
	</xsl:template>
	<xsl:template match="@MAILSHOTREQUIRED">
		<xsl:attribute name="{name()}"><xsl:choose><xsl:when test=".='true'">0</xsl:when><xsl:when test=".='false'">1</xsl:when><xsl:otherwise/></xsl:choose></xsl:attribute>
	</xsl:template>
</xsl:stylesheet>
