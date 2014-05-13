<h1>E-mail verzenden</h1>

<%select case request.querystring("a")
case "delmail"
	sql = "DELETE FROM mailing WHERE mailing_id=" & request.querystring("id")
	adocon.execute(sql)
	response.redirect(huidigepagina(1)&"&alert=verwijderd")

case "instekst"
	sql = "INSERT INTO mailing(type, datum, tekst, ingevoerd_door, onderwerp, bijlage1) VALUES("
	sql = sql & request.form("type") & ","
	sql = sql & "FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) &"','dd-mm-yyyy hh:mm:ss'),"
	sql = sql & "'" & changequotes(request.form("txtContent")) & "',"
	sql = sql & session("functie_id") & ","
	sql = sql & "'" & changequotes(request.form("onderwerp")) & "',"
	if request.form("bijlage") <> "" then
		sql = sql & request.form("bijlage") & ")"
	else
	sql = sql & "null)"
	end if
	response.write(sql)
	adocon.execute(sql)
	response.redirect(huidigepagina(1))


case "updmailing"
	sql = "UPDATE mailing SET onderwerp='"&changequotes(request.form("onderwerp"))&"', tekst='"&changequotes(request.form("txtContent"))&"', bijlage1="&request.form("bijlage1")&",bijlage2="&request.form("bijlage2")&", bijlage3="&request.form("bijlage3")&", datum=FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) &"','dd-mm-yyyy hh:mm:ss') WHERE mailing_id="&request.form("mailing_id")
	response.write(sql)
	adocon.execute(sql)
	response.redirect(huidigepagina(1))


case "nieuwsbrief", "flitsmet"%>
	<h2>Werkwijze</h2>
	Om de nieuwsbrief te verzenden worden de volgende stappen achtereenvolgens doorlopen:
	<ul>
		<li>Bijlagen uploaden</li>
		<li>Tekst schrijven</li>
		<li>Tekst controleren</li>
		<li>Verzenden</li>
		<li>Nieuwsbrief per post verzenden</li>
	</ul>
	<a href="default.asp?m=bestand&a=nieuw&redir=<%=request.querystring("a")%>2">Klik hier voor stap 1</a>
	<%


case "nieuwsbrief2", "flitsmet2", "flitszonder"
	select case request.querystring("a")
	case "nieuwsbrief2"
		teller = 1	'mailing_type
		teller2 = 4	'standaard tekst
	case "flitsmet2"
		teller = 2
		teller2 = 5
	case "flitszonder"
		teller = 2
		teller2 = 5
	end select
	
	if request.querystring("a") = "flitszonder" then%>
		<h2>Tekst schrijven</h2>
	<%else%>
		<ul>
			<li>Bijlagen uploaden</li>
			<li><strong>Tekst schrijven</strong></li>
			<li>Tekst controleren</li>
			<li>Verzenden</li>
			<li>Nieuwsbrief per post verzenden</li>
		</ul>
	<%end if%>
	<form action="<%=huidigepagina(1)%>&a=instekst" method="post" name="Form">
	<input type="hidden" name="type" value="<%=teller%>">
	<input type="hidden" name="bijlage" value="<%=request.querystring("bijlage")%>">
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr>
		<td>Onderwerp: <input type="text" name="onderwerp" size="80" maxlength="100" value="Skateology Nieuws">
		<br>&nbsp;</td>
	</tr>
	<tr>
		<td valign="top" align="center"><%call ACEJavaScript("obj1", "email")%>
		<p align="right"><input type="submit" value="Opslaan" onclick="SubmitForm()"></p></td>
	</tr>
	</table>
	</form>
	<TEXTAREA rows=7 cols=40 ID="idTextarea" NAME="idTextarea" style="display:none"><%
	set rs = adocon.execute("SELECT tekst FROM tekst WHERE tekst_id="&teller2)
		response.write rs("tekst")
	set rs = nothing
	%></TEXTAREA>
	<%

case "editmail"
	set rs = adocon.execute("SELECT * FROM mailing WHERE mailing_id="&request.querystring("id"))
	%><h2>Mailing bewerken</h2>
	<form action="<%=huidigepagina(1)%>&a=updmailing" method="post" name="Form">
	<input type="hidden" name="mailing_id" value="<%=request.querystring("id")%>">
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr>
		<td>
		<table width="100%" cellspacing=0 cellpadding=2 border=0>
		<tr>
			<td>Onderwerp</td>
			<td><input type="text" name="onderwerp" size="80" maxlength="100" value="<%=rs("onderwerp")%>"></td>
		</tr>
		<tr>
			<td colspan=2><strong>Bijlagen</strong> Er is een maximum van 3 bijlagen, maar uiteraard hoeven ze niet gebruikt te worden.</td>
		</tr>
		<%for teller = 1 to 3%>
			<tr>
				<td>Bijlage <%=teller%></td>
				<td><%if rs("bijlage"&teller) = "" or isnull(rs("bijlage"&teller)) then
						call ddlbBijlage("bijlage"&teller,0)
					else
						call ddlbBijlage("bijlage"&teller,rs("bijlage"&teller))
					end if
					%></td>
			</tr>
		<%next%>
		</table>
		</td>
	</tr>
	<tr>
		<td valign="top" align="center"><%call ACEJavaScript("obj1", "email")%>
		<p align="right"><input type="submit" value="Opslaan" onclick="SubmitForm()"></p></td>
	</tr>
	</table>
	</form>
	<TEXTAREA rows=7 cols=40 ID="idTextarea" NAME="idTextarea" style="display:none"><%response.write rs("tekst")%></TEXTAREA>
	<%set rs = nothing


case "sendmail"
	set rs = adocon.execute("SELECT * FROM mailing WHERE mailing_id="&request.querystring("id"))%>
	<h2>Mailing versturen</h2>
	Controleer onderstaande gegevens, selecteer de geadresseerde(n) en klik op "verzenden".
	<form action="<%=huidigepagina(1)%>&a=verzenden&id=<%=request.querystring("id")%>" method="post" name="Form">
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td colspan=2>E-mail</td>
	</tr>
	<tr>
		<td valign="top" width="30%"><strong>Geadresseerden</strong></td>
		<td width="70%"><%call ddlbGeadresseerden(rs("ontvangers"))%></td>
	</tr>
	<tr>
		<td valign="top"><strong>Onderwerp</strong></td>
		<td><%=rs("onderwerp")%></td>
	</tr>
	<tr>
		<td valign="top"><strong>Bijlagen</strong></td>
		<td><%for i = 1 to 3 
				call bestandnaam(rs("bijlage"&i))
			next%></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Verstuur"></td>
	</tr>
	</table>
	<%set rs = nothing

case "verzenden"
	Dim objCDOSYSCon, objCDOSYSMail, foutmeldingen, rs3
	foutmeldingen = ""
	teller = 0
	
	%><h2>Geadresseerden</h2><%
	select case request.form("ontvangers")
	case 2 'alle leden
		set rs = adocon.execute("SELECT email FROM lid WHERE not isnull(email) and not email='' AND actief=true AND (isnull(datum_afgemeld)) AND post=false ORDER BY email ASC")
	case 1 'Bestuursleden
		set rs = adocon.execute("SELECT email FROM functie WHERE not isnull(email) and not email='' AND bestuur=true AND jaartal_id=" & skatejaar&" ORDER BY email ASC")
	case 3 'mijzelf
		set rs = adocon.execute("SELECT email FROM functie WHERE not isnull(email) and not email='' AND functie_id="&session("functie_id")&" ORDER BY email ASC")
	end select
		
	do until rs.EOF
	if IsValidEmail(rs("email")) then

		teller = teller + 1
		
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
		
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost" ' was smtp.skateology.nl
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			objCDOSYSCon.Fields.Update
		Set objCDOSYSMail.Configuration = objCDOSYSCon
		
		objCDOSYSMail.From = "nieuwsbrief@skateology.nl"
		objCDOSYSMail.To = rs("email")
		
		'TEKST en BIJLAGEN
		set rs3 = adocon.execute("SELECT * FROM mailing WHERE mailing_id="&request.querystring("id"))
		objCDOSYSMail.HTMLBody = rs3("tekst")
		objCDOSYSMail.Subject = rs3("onderwerp")
		
		for i = 1 to 3
			if rs3("bijlage"&i) <> "" and isnumeric(rs3("bijlage"&i)) then
				%><h2>Bijlagen</h2><%
				set rs2 = adocon.execute("SELECT url FROM bestand WHERE bestand_id="&rs3("bijlage"&i))
					objCDOSYSMail.AddAttachment ("C:\domains\skateology.nl\wwwroot\bestanden\" & rs2("url"))
				set rs2 = nothing
			end if
		next
		set rs3 = nothing
		
		'VERZENDEN
		On Error Resume Next
		objCDOSYSMail.Send
		If Err <> 0 Then
			foutmeldingen = foutmeldingen & "<p>"& teller & ": " & Err.Description &"</p>"
			response.write("<font color=""red"">"&rs("email")&" heeft de mailing niet ontvangen (foutmelding " & teller & ").</font><br>")
		else
			response.write(rs("email") & "<br>")
		End If
		Set objCDOSYSMail = Nothing
		Set objCDOSYSCon = Nothing
	
	else
		response.write("<font color=""red"">"&rs("email")&" is geen geldig emailadres.</font><br>")
	End if
	rs.movenext : loop
	rs.close : set rs = nothing

	response.write("<br><strong>Verzonden aan " & teller & " leden.</strong><br>")
	
	if not foutmeldingen = "" then
		%><h2>Foutmeldingen</h2><%
		response.write(foutmeldingen)
	end if
	adocon.execute("UPDATE mailing SET verzonden_datum=FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) &"','dd-mm-yyyy hh:mm:ss'), ontvangers="&request.form("ontvangers")&" WHERE mailing_id="&request.querystring("id"))


case else%>
	<h2>Nieuwsbrieven</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="20%">Datum</td>
		<td width="50%">Titel</td>
		<td width="15%">Verzonden</td>
		<td width="15%"></td>
	</tr>
	<%set rs = adocon.execute("SELECT * FROM mailing WHERE type=1 ORDER BY datum DESC")
	do until rs.EOF%>
		<tr onmouseover="highlightRow(this, 'over')" onmouseout="highlightRow(this, 'out')">
			<td><%=date2datum(rs("datum"),false)%></td>
			<td><%=rs("onderwerp")%> <%call aantalbestanden(rs)%></td>
			<td><%=date2datum(rs("verzonden_datum"),false)%></td>
			<td><a href="<%=huidigepagina(1)%>&a=editmail&ace=true&id=<%=rs("mailing_id")%>"><img src="images/edit.gif" alt="Bewerk deze mailing" border="0"></a>
				<a href="<%=huidigepagina(1)%>&a=delmail&id=<%=rs("mailing_id")%>"><img src="images/delete.gif" alt="Verwijder deze mailing" border="0"></a>
				<a href="<%=huidigepagina(1)%>&a=sendmail&id=<%=rs("mailing_id")%>"><img src="images/send.gif" alt="Verstuur deze mailing" border="0"></a>
				</td>
		</tr>
	<%rs.movenext : loop
	rs.close : set rs =nothing%>
	</table>
	
	
	<h2>Nieuwsflitsen</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="20%">Datum</td>
		<td width="50%">Titel</td>
		<td width="15%">Verzonden</td>
		<td width="15%"></td>
	</tr>
	<%set rs = adocon.execute("SELECT * FROM mailing WHERE type=2 ORDER BY datum DESC")
	do until rs.EOF%>
		<tr onmouseover="highlightRow(this, 'over')" onmouseout="highlightRow(this, 'out')">
			<td><%=date2datum(rs("datum"),false)%></td>
			<td><%=rs("onderwerp")%> <%call aantalbestanden(rs)%></td>
			<td><%=date2datum(rs("verzonden_datum"),false)%></td>
			<td><a href="<%=huidigepagina(1)%>&a=editmail&ace=true&id=<%=rs("mailing_id")%>"><img src="images/edit.gif" alt="Bewerk deze mailing" border="0"></a>
				<a href="<%=huidigepagina(1)%>&a=delmail&id=<%=rs("mailing_id")%>"><img src="images/delete.gif" alt="Verwijder deze mailing" border="0"></a>
				<a href="<%=huidigepagina(1)%>&a=sendmail&id=<%=rs("mailing_id")%>"><img src="images/send.gif" alt="Verstuur deze mailing" border="0"></a>
				</td>
		</tr>
	<%rs.movenext : loop
	rs.close : set rs =nothing%>
	</table>
	<%

end select
%>