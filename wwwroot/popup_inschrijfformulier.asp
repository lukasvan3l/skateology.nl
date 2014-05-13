<%	@LANGUAGE = VBScript %>
<%
Option Explicit
Response.Buffer = TRUE
Response.Expires = 0
session.lcid = 1043
%><!-- #include file="inc/functions.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>Inschrijfformulier Skateology</TITLE>
	<STYLE type=text/css>
	BODY, TD, TABLE {
		FONT-WEIGHT: normal; 
		FONT-SIZE: 10px; 
		COLOR: black; 
		FONT-STYLE: normal; 
		FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; 
		TEXT-DECORATION: none
	}
	INPUT, SELECT, TEXTAREA {
		BORDER-RIGHT: black 1px solid; 
		BORDER-TOP: black 1px solid; 
		PADDING-LEFT: 3px; 
		FONT-WEIGHT: normal; 
		FONT-SIZE: 10px; 
		BORDER-LEFT: black 1px solid; 
		COLOR: black; 
		BORDER-BOTTOM: black 1px solid; 
		FONT-STYLE: normal; 
		FONT-FAMILY: Verdana,Arial,Tahoma; 
		BACKGROUND-COLOR: white; 
		TEXT-DECORATION: none
	}
	H1 {
		FONT-WEIGHT: bold; 
		FONT-SIZE: 16px; 
		COLOR: #000000; 
		FONT-STYLE: normal; 
		FONT-FAMILY: Arial, Tahoma; 
		TEXT-DECORATION: none
	}
	IMG {
		BORDER-TOP-WIDTH: 0px; 
		BORDER-LEFT-WIDTH: 0px; 
		BORDER-BOTTOM-WIDTH: 0px; 
		BORDER-RIGHT-WIDTH: 0px
	}
	</STYLE>
</HEAD>
<BODY bottomMargin=10 topMargin=10 marginheight="10" marginwidth="10">
<H1>Inschrijfformulier Skateology <%
set rs = adocon.execute("SELECT begindatum, einddatum, contributie FROM jaartal WHERE jaartal_id="&skatejaar)
	jaar = year(rs("begindatum"))
	response.write(jaar & "-" & jaar+1)
%></H1>

<FORM name=inschrijfformulier action=inschrijfformulier.asp?a=inschrijven method=post>
<p>Welkom bij Skateology, de vereniging om gezellig te skaten in Leiden en omstreken. De vereniging organiseert op iedere dinsdag en zondag rondritten in de omgeving van Leiden, waarbij de nadruk ligt op gezelligheid met een groep skaters. We verzamelen dinsdag om 20.00 uur en zondag om 14:00 uur op de Beestenmarkt in het centrum van Leiden. Tevens organiseert Skateology soms activiteiten in het weekend.</p>
<p>Om ieder lid te registreren en te kunnen voorzien van informatie zijn de volgende gegevens nodig: </p>
<TABLE width=100% border=1 align="center" cellPadding=0 cellSpacing=0 frame="box" rules="none">
  <TBODY>
  <TR>
    <TD width="45%"><div align="right">Voornaam</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2> <INPUT maxLength=50 size=50 name=cv1></TD></TR>
  <TR>
    <TD width="45%"><div align="right">Achternaam</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><INPUT maxLength=50 size=50 name=cv3></TD></TR>
  <TR>
    <TD width="45%"><div align="right">Adres</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><INPUT maxLength=50 size=50 name=cv4></TD></TR>
  <TR>
    <TD width="45%"><div align="right">Postcode, woonplaats</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><INPUT maxLength=6 size=8 name=cv5> <INPUT 
      maxLength=39 size=39 name=cv6></TD></TR>
  <TR>
    <TD width="45%"><div align="right">Telefoon</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><INPUT maxLength=50 size=50 name=cv7></TD></TR>
  <TR>
    <TD width="45%"><div align="right">Telefoon Mobiel</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><INPUT maxLength=50 size=50 name=cv8></TD></TR>
  <TR>
    <TD width="45%"><div align="right">E-mail Adres*</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><INPUT maxLength=50 size=50 name=cv9></TD></TR>
  <TR>
    <TD width=45%><div align="right">Bank/Girorekening</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><INPUT maxLength=50 size=50 name=cv1></TD></TR>
  <TR>
    <TD width="45%"><div align="right">Geboortedatum</div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><INPUT maxLength=50 size=50 name=cv2></TD></TR>
  <TR>
    <TD width="45%"><div align="right">Soort skates </div></TD>
    <TD width="3%">&nbsp;</TD>
    <TD colSpan=2><select name="select">
	<option selected>Inline Skates</option>
	<option>Rollerskates</option>
	<option>Skeelers</option>
	<option>Off-the-Roads</option>
    </select></TD></TR>
</TBODY></TABLE>

<p>
	Als je een e-mail adres hebt ingevuld, ontvang je de twee-maandelijkse Skateology Nieuwsbrief per e-mail.
</p>
<TABLE width=100% border=1 align="center" cellPadding=0 cellSpacing=0 frame="box" rules="none">
  <TBODY>
    <TR>
      <TD colspan="4">Hieronder kun je je gegevens invullen wie we kunnen waarschuwen in geval van calamiteiten (bijvoorbeeld een ongeval) </TD>
      </TR>
    <TR>
      <TD width="45%"><div align="right">Naam</div></TD>
      <TD width="3%">&nbsp;</TD>
      <TD colSpan=2><INPUT maxLength=50 size=50 name=cv12></TD>
    </TR>
    <TR>
      <TD width="45%"><div align="right">Telefoon</div></TD>
      <TD width="3%">&nbsp;</TD>
      <TD colSpan=2><INPUT maxLength=50 size=50 name=cv22></TD>
    </TR>
    <TR>
      <TD width="45%"><div align="right">Relatie (bv moeder, vriend) </div></TD>
      <TD width="3%">&nbsp;</TD>
      <TD colSpan=2><INPUT maxLength=50 size=50 name=cv32></TD>
    </TR>
  </TBODY>
</TABLE>
<p>Het lidmaatschap van Skateology geeft je recht mee te skaten met de tochten en deel te nemen aan activiteiten. Tevens zal je op de hoogte worden gehouden door middel van nieuwsbrieven, ledenlijsten en tussentijdse e-mailtjes. Daarnaast ben je gelijk lid van de Skatebond Nederland. 
<%
if isnull(rs("contributie")) or rs("contributie") = 0 or rs("contributie") = "" then
	%>De contributie voor dit skatejaar is nog niet vastgesteld. Zodra deze is vastgesteld krijg je hier bericht van, en de wijze waarop je dit bedrag aan de vereniging kan betalen.<%
else
	%>Het lidmaatschap bedraagt &euro; <%=formatnumber(rs("contributie"))%> (lopend tot september <%=jaar+1%>) en dient te worden overgemaakt op Postbanknr. 9151332 t.n.v. Skateology te Leiden. Je krijgt kort nadat je het formulier ingeleverd hebt bericht hoe je dit bedrag kunt betalen.<%	
end if
%> Wil je meer weten voordat je lid wordt, bel dan even met: </p>
<table border=0>
<%set rs = nothing
set rs = adocon.execute("SELECT * FROM functie LEFT JOIN lid ON functie.lid_id=lid.lid_id WHERE jaartal_id="&skatejaar&" AND bestuur=true ORDER BY functie DESC")
do until rs.EOF%>
	<tr>
		<td width="112"><%=rs("functie")%></td>
		<td	width="70"><%=rs("voornaam")%></td>
		<td width="162"><%		
		    If rs("functie") = "Voorzitter" then
				response.write(rs("telefoonnummer"))
			End If
	  %></td>
	</tr>
<%rs.movenext : loop%>
  <tr>
    <td>E-mail</td>
    <td>&nbsp;</td>
    <td><a href="javascript:mailto('info')">E-mail</a></td>
  </tr>
  <tr>
    <td>Internet</td>
    <td>&nbsp;</td>
    <td><a href="http://www.skateology.nl">http://www.skateology.nl</a></td>
  </tr>
</table>
<hr noshade>
<p>De vereniging of het bestuur is op geen enkele wijze aansprakelijk voor schade en / of letsel aan ondergetekende of voor schade en / of letsel aan derden. Bij ondertekening van dit formulier geeft de deelnemer aan van bovenstaande op de hoogte te zijn. </p>
<TABLE width=100% border=1 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR><TD width="200" height="67">&nbsp;</TD>
      <TD rowspan="2" align="right"><p>Print dit ingevulde formulier,<br>
		onderteken hem<br>
		en vestuur  naar: </p>
<!--
    <p>Lukas van Driel<br />
    Acacialaan 58<br />
    2351CD leiderdorp</p>-->
<%rs.movefirst
i = false
if i = false then
	do until rs.EOF or i = true
		if instr(lcase(rs("functie")), "leden") then
			%>
			<p><%call naam(rs)%><br>
			<%=rs("adres")%><br>
			<%=rs("postcode") & " " & rs("woonplaats")%></p>
			<%
			i = true
		end if
	rs.movenext : loop
end if
rs.movefirst
if i = false then
	do until rs.EOF or i = true
		if instr(lcase(rs("functie")), "secretaris") then
			%>
			<p><%call naam(rs)%><br>
			<%=rs("adres")%><br>
			<%=rs("postcode") & " " & rs("woonplaats")%></p>
			<%
			i = true
		end if
	rs.movenext : loop
end if
rs.movefirst
if i = false then
	do until rs.EOF or i = true
		if instr(lcase(rs("functie")), "voorzitter") then
			%>
			<p><%call naam(rs)%><br>
			<%=rs("adres")%><br>
			<%=rs("postcode") & " " & rs("woonplaats")%></p>
			<%
			i = true
		end if
	rs.movenext : loop
end if
rs.close : set rs = nothing%>
		</TD>
    </TR>
    <TR>
      <TD width="200" height="15">(handtekening)</TD>
      </TR>
</tbody>
</table>

<P align=right><INPUT onclick="print()" type=button value=Print>&nbsp;&nbsp;</P></FORM></BODY></HTML>

<!-- #include file="inc/functionsclose.asp" -->
