<h1>Ledenadministratie</h1>

<%select case request.querystring("a")
case "updLid"
	sql = "UPDATE lid SET "
	sql = sql & "voornaam='"& changequotes(request.form("voornaam")) &"', "
	sql = sql & "tussenvoegsel='"& changequotes(request.form("tussenvoegsel")) &"', "
	sql = sql & "achternaam='"& changequotes(request.form("achternaam")) &"', "
	sql = sql & "adres='"& changequotes(request.form("adres")) &"', "
	sql = sql & "postcode='"& changequotes(request.form("postcode")) &"', "
	sql = sql & "woonplaats='"& changequotes(request.form("woonplaats")) &"', "
	sql = sql & "telefoon='"& telefoon(request.form("telefoon")) &"', "
	sql = sql & "mobiel='"& telefoon(request.form("mobiel")) &"', "
	sql = sql & "email='"& changequotes(request.form("email")) &"', "
	sql = sql & "bankrekening='"& changequotes(request.form("bankrekening")) &"', "
	sql = sql & "em_naam='"& changequotes(request.form("em_naam")) &"', "
	sql = sql & "em_telefoon='"& telefoon(request.form("em_telefoon")) &"', "
	sql = sql & "em_relatie='"& changequotes(request.form("em_relatie")) &"', "
	sql = sql & "sbn='"& changequotes(request.form("sbn")) &"', "
	sql = sql & "website='"& website(request.form("website")) &"', "
	sql = sql & "opmerkingen='"& changequotes(request.form("opmerkingen")) &"', "
	sql = sql & "lid_type_id="& request.form("lid_type_id") &", "
	if request.form("email_publiek") = "on" then sql = sql & "email_publiek=true, " else sql = sql & "email_publiek=false, " end if
	if request.form("telefoon_publiek") = "on" then sql = sql & "telefoon_publiek=true, " else sql = sql & "telefoon_publiek=false, " end if
	if request.form("mobiel_publiek") = "on" then sql = sql & "mobiel_publiek=true, " else sql = sql & "mobiel_publiek=false, " end if
	if request.form("adres_publiek") = "on" then sql = sql & "adres_publiek=true, " else sql = sql & "adres_publiek=false, " end if
	if request.form("actief") = "on" then sql = sql & "actief=true, " else sql = sql & "actief=false, " end if
	if request.form("post") = "on" then sql = sql & "post=true, " else sql = sql & "post=false, " end if
	if request.form("geboortedatum") <> "" then
		sql = sql & "geboortedatum=FORMAT('" & request.form("geboortedatum") & "','dd-mm-yyyy'), "
	end if
	if request.form("EHBO") = "" or isnumeric(left(request.form("ehbo"),1)) = false then
		sql = sql & "EHBO=null "
	else
		sql = sql & "EHBO=FORMAT('" & request.form("EHBO") & "','dd-mm-yyyy') "
	end if
	sql = sql & "WHERE lid_id="& request.querystring("id")
	response.write(sql)
	adocon.execute(sql)
	response.redirect(huidigepagina(2) & "&a=bewerk&alert=opgeslagen")


case "insLid"
	sql = "INSERT INTO lid(sbn, voornaam, tussenvoegsel, achternaam, adres, postcode, woonplaats, telefoon, mobiel, email, bankrekening, em_naam, em_telefoon, em_relatie, website, opmerkingen, lid_type_id, email_publiek, telefoon_publiek, mobiel_publiek, adres_publiek, actief, post, geboortedatum, ehbo, ingevoerd_door) values("
	if not(request.form("sbn")=null) then
		sql = sql & changequotes(request.form("sbn")) &", "
	else
		sql = sql & "null, "
	end if
	sql = sql & "'"& changequotes(request.form("voornaam")) &"', "
	sql = sql & "'"& changequotes(request.form("tussenvoegsel")) &"', "
	sql = sql & "'"& changequotes(request.form("achternaam")) &"', "
	sql = sql & "'"& changequotes(request.form("adres")) &"', "
	sql = sql & "'"& changequotes(request.form("postcode")) &"', "
	sql = sql & "'"& changequotes(request.form("woonplaats")) &"', "
	sql = sql & "'"& telefoon(request.form("telefoon")) &"', "
	sql = sql & "'"& telefoon(request.form("mobiel")) &"', "
	sql = sql & "'"& changequotes(request.form("email")) &"', "
	sql = sql & "'"& changequotes(request.form("bankrekening")) &"', "
	sql = sql & "'"& changequotes(request.form("em_naam")) &"', "
	sql = sql & "'"& telefoon(request.form("em_telefoon")) &"', "
	sql = sql & "'"& changequotes(request.form("em_relatie")) &"', "
	sql = sql & "'"& website(request.form("website")) &"', "
	sql = sql & "'"& changequotes(request.form("opmerkingen")) &"', "
	sql = sql & request.form("lid_type_id") &", "
	if request.form("email_publiek") = "on" then sql = sql & "true, " else sql = sql & "false, " end if
	if request.form("telefoon_publiek") = "on" then sql = sql & "true, " else sql = sql & "false, " end if
	if request.form("mobiel_publiek") = "on" then sql = sql & "true, " else sql = sql & "false, " end if
	if request.form("adres_publiek") = "on" then sql = sql & "true, " else sql = sql & "false, " end if
	if request.form("actief") = "on" then sql = sql & "true, " else sql = sql & "false, " end if
	if request.form("post") = "on" then sql = sql & "true, " else sql = sql & "false, " end if
	if request.form("geboortedatum") <> "" then
		sql = sql & "FORMAT('" & request.form("geboortedatum") & "','dd-mm-yyyy'), "
	else
		sql = sql & "null, "
	end if
	if request.form("EHBO") = "" or isnumeric(left(request.form("ehbo"),1)) = false then
		sql = sql & "null "
	else
		sql = sql & "FORMAT('" & request.form("EHBO") & "','dd-mm-yyyy') "
	end if
	sql = sql & ","&session("functie_id")&")"
	response.write(sql)
	adocon.execute(sql)
	response.redirect(huidigepagina(1) & "&a=welkomemail&ace=true")


case "meldopnieuwaan"
	set rs = adocon.execute("SELECT opmerkingen, datum_aangemeld, datum_afgemeld FROM lid WHERE lid_id="&request.querystring("id"))
	sql = "UPDATE lid SET opmerkingen='"&changequotes( rs("opmerkingen") & vbcrlf & "Eerder lid geweest van "&date2datum(rs("datum_aangemeld"),false)&" t/m "& date2datum(rs("datum_afgemeld"),false) )&"', datum_afgemeld=null, datum_aangemeld=FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) &"','dd-mm-yyyy hh:mm:ss') WHERE lid_id="&request.querystring("id")
	response.write(sql)
	set rs = nothing
	adocon.execute(sql)
	response.redirect(huidigepagina(2) & "&alert=opgeslagen")


case "nieuw"%>
	<h2>Lid aanmaken</h2>
	<form action="<%=huidigepagina(2)%>&a=insLid" method="post" name="form" id="form">
	<h3>Inschrijfformulier</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="30%">Gegeven</td>
		<td width="50%">Waarde</td>
		<td width="20%">Weergeven</td>
	</tr>
	<tr><td>Voornaam</td><td><input type="text" maxlength="50" size="40" name="Voornaam" value=""></td></tr>
	<tr><td>Tussenvoegsel</td><td><input type="text" maxlength="10" size="10" name="Tussenvoegsel" value=""></td></tr>
	<tr><td>Achternaam</td><td><input type="text" maxlength="50" size="40" name="Achternaam" value=""></td></tr>
	<tr><td>Adres</td><td><input type="text" maxlength="50" size="40" name="Adres" value=""></td>
		<td><input type="checkbox" name="adres_publiek" checked></td></tr>
	<tr><td>Postcode</td><td><input type="text" maxlength="50" size="40" name="Postcode" value=""></td></tr>
	<tr><td>Woonplaats</td><td><input type="text" maxlength="50" size="40" name="Woonplaats" value=""></td></tr>
	<tr><td>Telefoon</td><td><input type="text" maxlength="15" size="20" name="Telefoon" value=""></td>
		<td><input type="checkbox" name="telefoon_publiek" checked></td></tr>
	<tr><td>Mobiel</td><td><input type="text" maxlength="15" size="20" name="Mobiel" value="06-"></td>
		<td><input type="checkbox" name="mobiel_publiek" checked></td></tr>
	<tr><td>E-mail</td><td><input type="text" maxlength="100" size="40" name="Email" value=""></td>
		<td><input type="checkbox" name="email_publiek" checked></td></tr>
	<tr><td>Bankrekening</td><td><input type="text" maxlength="15" size="20" name="Bankrekening" value=""></td></tr>
	<tr><td>Geboortedatum</td><td><input type="text" maxlength="10" size="10" name="Geboortedatum" value=""> (dd-mm-jjjj)</td></tr>
	</table>
	<h3>Bestuurlijke gegevens</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr><td width="30%">Aangemeld</td><td width="50%"><input type="text" maxlength="50" size="10" name="aangemeld" value="<%=date2datum(now(), false)%>" disabled></td><td width="20%"></td></tr>
	<tr><td>Actief</td><td><input type="checkbox" name="actief" checked></td></tr>
	<tr><td>Post ipv E-mail</td><td><input type="checkbox" name="post"></td></tr>
	<tr><td>Website</td><td><input type="text" maxlength="100" size="40" name="website" value=""></td></tr>
	<tr><td>EHBO geldig tot</td><td><input type="text" maxlength="10" size="10" name="ehbo" value=""></td></tr>
	<tr><td>Lid type</td><td><%call ddlbLid_type(1)%></td></tr>
	<tr><td>SBN Nummer</td><td><input type="text" maxlength="10" size="10" name="sbn" value=""></td></tr>
	<tr><td valign="top">Opmerkingen</td><td><textarea name="Opmerkingen" cols="40" rows="5"></textarea></td></tr>
	</table>
	<h3>Emergency gegevens</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr><td width="30%">Naam</td><td width="50%"><input type="text" maxlength="50" size="40" name="em_naam" value=""></td><td width="20%"></td></tr>
	<tr><td>Telefoon</td><td><input type="text" maxlength="50" size="40" name="em_telefoon" value=""></td></tr>
	<tr><td>Relatie</td><td><input type="text" maxlength="50" size="40" name="em_relatie" value=""></td></tr>
	</table>
	<h3>Opslaan</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr><td width="30%"></td><td width="50%"><input type="submit" value="Opslaan"></td><td width="20%"></td></tr>
	</table>
	</form>
	<%


case "welkomemail"
	set rs = adocon.execute("select top 1 * from lid order by lid_id desc")%>
	<h2>Lid aanmaken</h2>
	<form action="<%=huidigepagina(1)%>&a=sendmail" method="post" name="Form">
		Onderstaande mail wordt verstuurd naar <%call ddlblid(rs("lid_id"))%>.<br><br>
	<%set rs = nothing%>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr>
		<td>
		<table width="100%" cellspacing=0 cellpadding=2 border=0>
		<tr>
			<td>Onderwerp</td>
			<td><input type="text" name="onderwerp" size="80" maxlength="100" value="Welkom bij Skateology"></td>
		</tr>
		<tr>
			<td colspan=2><strong>Bijlagen</strong> Er is een maximum van 3 bijlagen, maar uiteraard hoeven ze niet gebruikt te worden.</td>
		</tr>
		<%for teller = 1 to 3%>
			<tr>
				<td>Bijlage <%=teller%></td>
				<td><%
					if teller = 1 then
						set rs = adocon.execute("SELECT TOP 1 * FROM bestand WHERE type=1 and verwijderd=false ORDER BY jaar desc, maand desc")
							call ddlbBijlage("bijlage"&teller,rs("bestand_id"))
						set rs = nothing
					else
						call ddlbBijlage("bijlage"&teller,0)
					end if%></td>
			</tr>
		<%next%>
		</table>
		</td>
	</tr>
	<tr>
		<td valign="top" align="center"><%call ACEJavaScript("obj1", "email")%>
		<p align="right"><input type="submit" value="Verzenden" onclick="SubmitForm()"></p></td>
	</tr>
	</table>
	</form>
	<TEXTAREA rows=7 cols=40 ID="idTextarea" NAME="idTextarea" style="display:none"><%
	set rs = adocon.execute("SELECT * FROM tekst WHERE tekst_id=6")
		response.write(rs("tekst"))
	set rs = nothing
	%></TEXTAREA>
	<%


case "sendmail"
	foutmeldingen = ""
	
	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
	
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.skateology.nl"
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objCDOSYSCon.Fields.Update
	Set objCDOSYSMail.Configuration = objCDOSYSCon
	
	objCDOSYSMail.From = "nieuwsbrief@skateology.nl"
	
	%><h2>Geadresseerde</h2><%
	set rs = adocon.execute("SELECT email FROM lid WHERE lid_id="&request.form("lid_id"))
	if IsValidEmail(rs("email")) then
		objCDOSYSMail.To = rs("email")
		response.write(rs("email"))
	else
		response.write("<font color=""red"">"&rs("email")&" is geen geldig emailadres.</font>")
	End if
	set rs = nothing
	
	'TEKST en BIJLAGEN
	objCDOSYSMail.HTMLBody = request.form("txtContent")
	objCDOSYSMail.Subject = request.form("onderwerp")
	
	for i = 1 to 3
	if request.form("bijlage"&i) <> "" and request.form("bijlage"&i) <> "null" then
		%><h2>Bijlagen</h2><%
		set rs = adocon.execute("SELECT url FROM bestand WHERE bestand_id="&request.form("bijlage"&i))
			objCDOSYSMail.AddAttachment ("C:\domains\skateology.nl\wwwroot\bestanden\" & rs("url"))
			response.write("C:\domains\skateology.nl\wwwroot\bestanden\" & rs("url") & "<br>")
		set rs = nothing
	end if
	next
	set rs = nothing
	
	'VERZENDEN
	On Error Resume Next
	objCDOSYSMail.Send
	If Err <> 0 Then
		foutmeldingen = foutmeldingen & "<p>" & Err.Description &"</p>"
		response.write("<font color=""red"">"&rs("email")&" heeft de mailing niet ontvangen (foutmelding " & teller & ").</font><br>")
	else
		response.write(rs("email") & "<br>")
	End If
	Set objCDOSYSMail = Nothing
	Set objCDOSYSCon = Nothing
	
	response.write("<br><strong>Verzonden aan 1 lid.</strong><br>")
	
	if not foutmeldingen = "" then
		%><h2>Foutmeldingen</h2><%
		response.write(foutmeldingen)
	end if

case "bewerk"
	set rs = adocon.execute("SELECT * FROM lid WHERE lid_id="&request.querystring("id"))%>
	<h2>Lid wijzigen</h2>
	<form action="<%=huidigepagina(2)%>&a=updLid" method="post" name="form" id="form">
	<h3>Inschrijfformulier</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="30%">Gegeven</td>
		<td width="50%">Waarde</td>
		<td width="20%">Weergeven</td>
	</tr>
	<tr><td>Lid id</td><td><input type="text" maxlength="50" size="10" name="lid_id" value="<%=rs("lid_id")%>" disabled></td></tr>
	<tr><td>Voornaam</td><td><input type="text" maxlength="50" size="40" name="Voornaam" value="<%=rs("Voornaam")%>"></td></tr>
	<tr><td>Tussenvoegsel</td><td><input type="text" maxlength="10" size="10" name="Tussenvoegsel" value="<%=rs("Tussenvoegsel")%>"></td></tr>
	<tr><td>Achternaam</td><td><input type="text" maxlength="50" size="40" name="Achternaam" value="<%=rs("Achternaam")%>"></td></tr>
	<tr><td>Adres</td><td><input type="text" maxlength="50" size="40" name="Adres" value="<%=rs("Adres")%>"></td>
		<td><input type="checkbox" name="adres_publiek"<%if rs("adres_publiek")=true then response.write(" checked")%>></td></tr>
	<tr><td>Postcode</td><td><input type="text" maxlength="50" size="40" name="Postcode" value="<%=rs("Postcode")%>"></td></tr>
	<tr><td>Woonplaats</td><td><input type="text" maxlength="50" size="40" name="Woonplaats" value="<%=rs("Woonplaats")%>"></td></tr>
	<tr><td>Telefoon</td><td><input type="text" maxlength="15" size="20" name="Telefoon" value="<%=telefoon(rs("Telefoon"))%>"></td>
		<td><input type="checkbox" name="telefoon_publiek"<%if rs("telefoon_publiek")=true then response.write(" checked")%>></td></tr>
	<tr><td>Mobiel</td><td><input type="text" maxlength="15" size="20" name="Mobiel" value="<%=telefoon(rs("Mobiel"))%>"></td>
		<td><input type="checkbox" name="mobiel_publiek"<%if rs("mobiel_publiek")=true then response.write(" checked")%>></td></tr>
	<tr><td>E-mail</td><td><input type="text" maxlength="100" size="40" name="Email" value="<%=rs("Email")%>"></td>
		<td><input type="checkbox" name="email_publiek"<%if rs("email_publiek")=true then response.write(" checked")%>></td></tr>
	<tr><td>Bankrekening</td><td><input type="text" maxlength="15" size="20" name="Bankrekening" value="<%=rs("Bankrekening")%>"></td></tr>
	<tr><td>Geboortedatum</td><td><input type="text" maxlength="10" size="10" name="Geboortedatum" value="<%=date2datum(rs("Geboortedatum"),false)%>"> (dd-mm-jjjj)</td></tr>
	</table>
	<h3>Bestuurlijke gegevens</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<%if rs("datum_afgemeld") <> "" then%>
		<tr>
			<td colspan="3">Deze persoon is inmiddels geen lid meer van Skateology.<br>
			<a href="<%=huidigepagina(2)%>&a=meldopnieuwaan">Klik hier om hem/haar opnieuw aan te melden als lid.</a></td>
		</tr>
	<%end if%>

	<tr><td width="30%">Aangemeld</td><td width="50%"><input type="text" maxlength="50" size="10" name="aangemeld" value="<%=date2datum(rs("datum_aangemeld"),false)%>" disabled></td><td width="20%"></td></tr>
	<tr><td>Afgemeld</td><td><input type="text" maxlength="10" size="10" name="afgemeld" value="<%=date2datum(rs("datum_afgemeld"),false)%>" disabled></td></tr>
	<tr><td>Actief</td><td><input type="checkbox" name="actief" <%if rs("actief") = true then response.write("checked")%>></td></tr>
	<tr><td>Post ipv E-mail</td><td><input type="checkbox" name="post" <%if rs("post") = true then response.write("checked")%>></td></tr>
	<tr><td>Website</td><td><input type="text" maxlength="100" size="40" name="website" value="<%=rs("website")%>"></td></tr>
	<tr><td>EHBO geldig tot</td><td><input type="text" maxlength="10" size="10" name="ehbo" value="<%=rs("ehbo")%>"></td></tr>
	<tr><td>Lid type</td><td><%call ddlbLid_type(rs("lid_type_id"))%></td></tr>
	<tr><td>SBN Nummer</td><td><input type="text" maxlength="10" size="10" name="sbn" value="<%=rs("sbn")%>"></td></tr>
	<tr><td valign="top">Opmerkingen</td><td><textarea name="Opmerkingen" cols="40" rows="5"><%=rs("Opmerkingen")%></textarea></td></tr>
	</table>
	<h3>Emergency gegevens</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr><td width="30%">Naam</td><td width="50%"><input type="text" maxlength="50" size="40" name="em_naam" value="<%=rs("em_naam")%>"></td><td width="20%"></td></tr>
	<tr><td>Telefoon</td><td><input type="text" maxlength="50" size="40" name="em_telefoon" value="<%=telefoon(rs("em_telefoon"))%>"></td></tr>
	<tr><td>Relatie</td><td><input type="text" maxlength="50" size="40" name="em_relatie" value="<%=rs("em_relatie")%>"></td></tr>
	</table>
	<h3>Opslaan</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr><td width="30%"></td><td width="50%"><input type="submit" value="Opslaan"></td><td width="20%"></td></tr>
	</table>
	</form>
	<%set rs = nothing
end select%>