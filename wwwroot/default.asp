<%	@LANGUAGE = VBScript %>
<%
Option Explicit
Response.Buffer = TRUE
Response.Expires = 0
session.lcid = 1043
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- #include file="inc/adovbs.inc" -->
<!-- #include file="inc/functions.asp" -->
<%
'referrer toevoegen aan db
dim referrer
referrer = Request.ServerVariables("HTTP_REFERER")
if not (trim(referrer) = "") AND not (instr(referrer, "skateology.nl/") <> 0) then
	adocon.execute("INSERT INTO referrer(datum, referer, ipadres) VALUES(FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) &"','dd-mm-yyyy hh:mm:ss'), '"& changequotes(referrer) &"', '"& request.servervariables("REMOTE_ADDR") &"')")
end if
%><HTML>
<HEAD>
	<TITLE>Skateology - In-line skaten in Leiden en omstreken</TITLE>
	<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
	<link rel="stylesheet" href="css/frontend.css" type="text/css">
	<script src="inc/javascript.js" type="text/javascript"></script>
</HEAD>
<BODY>
<TABLE cellPadding=0 cellSpacing=0 width=750 BORDER=0>
<tr> 
	<td class="logo">
		<a href="default.asp"><img src="images/skateologyweb.gif" width="147" height="35" border="0" alt="Skateology home"></a>
		</td>
	<td class="datum">
		<%= dagvanweek(DatePart("w", Date())) & " " & day(now()) & " " & maandvanjaar(month(now())) & " " & year(now())%>
		</td>
</tr>
<tr>
	<td colspan="2" class="menu" nowrap>
		<a href="default.asp" 					class="menulink">Home</a> |
<!--	<a href="default.asp?a=nieuwsbrieven" 	class="menulink">Nieuwsbrieven</a> |
		<a href="default.asp?a=agenda" 			class="menulink">Agenda</a> | -->
		<a href="default.asp?a=fotos" 			class="menulink">Fotos</a> |
		<a href="default.asp?a=routes" 			class="menulink">Skate routes</a>
		<!--<a href="default.asp?a=spelregels" 		class="menulink">Spelregels</a> |
		<a href="default.asp?a=vereniging" 		class="menulink">Vereniging</a>-->
		</td>
</tr>
<%response.flush%>
<tr>
	<td colspan="2" height=100 width=750><img src="kopfotos/<%
	set rs2 = adocon.execute("SELECT * FROM kopfoto_datum WHERE day(datum)=day(now) AND month(datum)=month(now)")
	if not rs2.EOF then
		set rs = adocon.execute("SELECT * FROM kopfoto WHERE kopfoto_id="&rs2("kopfoto_id"))
	else
		set rs = ExecQuery("SELECT * FROM kopfoto WHERE actief=true ORDER BY kopfoto_id ASC",adocon)
		randomize()
		j = (Int((rs.recordcount) * Rnd+1))
		for i = 2 to j
			rs.movenext
		next
	end if
	response.write(rs("url"))
	set rs2 = nothing
	set rs = nothing
	%>" alt="Skateology - In-line skaten in Leiden en omstreken" width="750" height="100"></td>
</tr>
<tr>
	<td class="links" width="540">
		<!-- #include file="links.asp" -->
		<%
		select case request.querystring("a")
		'case "nieuwsbrieven"
		'	call l_nieuwsbrief()
		case "agenda"
			call l_agenda()
		case "fotos"
			call l_fotos()
		case "fotos."
			call l_fotos()
		case "routes"
			call l_routes()
		'case "spelregels"
		'	call l_spelregels()
		'case "vereniging"
		'	call l_vereniging()
		'case "inschrijven"
		'	call l_inschrijven()
		case "ervaringen"
			call l_ervaringen()
		'case "3House"
			'call l_3House()
		case "skateweekend2007"
			call l_skateweekend2007()
		case "skateweekend"
			call l_skateweekend()
		case "skateweekend_save"
			call l_skateweekend_save()
		case else
			call l_welkom()
		end select
		%>
		</td>
	<td class="rechts" width="210">
		<!-- #include file="rechts.asp" -->
		<%
		select case request.querystring("a")
		'case "nieuwsbrieven"
		'	call r_nieuwsbrief(request.querystring("id"))
		'	call r_nieuwsbrieven(100)
		case "agenda"
			call r_agenda(0)
		case "fotos"
			call r_fotos()
			call r_logo()
		case "fotos."
			call r_fotos()
			call r_logo()
		case "routes"
			call r_agenda(3)
		'	call r_inschrijven()
		'case "spelregels"
		'	call r_spelregels()
		'	call r_logo()
		'case "vereniging"
		'	call r_ervaringen()
		'	call r_inschrijven()
		'case "inschrijven"
		'	call r_ervaringen()
		'	call r_agenda(0)
		case "ervaringen"
			call r_ervaringen()
		case "3House"
			call r_3House()
		case "skateweekend", "skateweekend_save"
			call r_agenda(0)
			call r_skateweekend()
		case else
			call r_logo()
			'call r_twitter()
			'call r_ervaringen()
			call r_banners()
		end select
		%>
		</td>
</tr>
</TABLE>


	<script type="text/javascript">
		var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
		document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
	</script>
	<script type="text/javascript">
		try {
		var pageTracker = _gat._getTracker("UA-10406828-1");
		pageTracker._trackPageview();
		} catch(err) {}
	</script>

</BODY>
</HTML>
<!-- #include file="inc/functionsclose.asp" -->
