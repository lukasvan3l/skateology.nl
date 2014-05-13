<h1>Ledenadministratie</h1>
<%select case request.querystring("a")
case "dellid"
	sql = "UPDATE lid SET datum_afgemeld=FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) &"','dd-mm-yyyy hh:mm:ss') WHERE lid_id="&request.querystring("id")
	response.write(sql)
	adocon.execute(sql)
	response.redirect(huidigepagina(2) & "&zoek=&alert=verwijderd")

case "nieuwsbrief"%>
	<H2>Ledenlijst voor in de Nieuwsbrief</H2>
	<table width=95% cellspacing=0 class="tabledata" style="font: 13px times new roman;">
	<tr class="trhead">
		<td><strong>Naam<br>
		SBN Lidnummer</strong></td>
		<td><strong>Adres<br>
		Postcode Woonplaats</strong></td>
		<td><strong>Telefoon<br>
		Mobiel</strong></td>
		<td align="right"><strong>Geboortedatum<br>
		E-mailadres</strong></td>
	</tr>
	<%
	dim rowbg
	rowbg = "#dddddd"
	set rs = adocon.execute("SELECT * FROM lid WHERE actief=true AND isnull(datum_afgemeld) ORDER BY voornaam asc")
	do until rs.EOF
		if rowbg = "#dddddd" then rowbg = "#ffffff" else rowbg="#dddddd"%>
		<tr bgcolor="<%=rowbg%>">
			<td nowrap valign="top">
				<%call naam(rs)%><br />
				<%=rs("sbn")%>
      </td>
			<td nowrap valign="top">
				<%if rs("adres_publiek") = true then%>
					<%=rs("adres")%>&nbsp;<br>
					<%=rs("postcode")%>&nbsp;<%=rs("woonplaats")%>
				<%end if%></td>
			<td nowrap valign="top">
				<%if rs("telefoon_publiek") = true then%>
					<%=telefoon(rs("telefoon"))%>&nbsp;
				<%end if%>
				<br>
				<%if rs("mobiel_publiek") = true then%>
					<%=telefoon(rs("mobiel"))%>
				<%end if%></td>
			<td nowrap valign="top" align="right">
				<%= Date2Datum(rs("geboortedatum"), false) %><br>
				<%if rs("email_publiek") = true then%>
					<%=rs("email")%>
				<%end if%></td>
		</tr>	
	
		<%
		rs.movenext : loop
	rs.close : set rs = nothing%>
	</table>

<%case "ehbo"
	sql = "SELECT * FROM lid WHERE actief=true AND isnull(datum_afgemeld) AND (lid_type_id=1 OR lid_type_id=4 OR lid_type_id=6) ORDER BY voornaam ASC"
	set rs = ExecQuery(sql,adocon)
	%>
	<h2>EHBO Ledenlijst (<%=rs.recordcount%> leden)</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td valign="top">Voor- en achternaam</td>
		<td valign="top">Bellen in nood</td>
		<td valign="top">Relatie</td>
	</tr>
	<%do until rs.EOF%>
		<tr>
			<td valign="top"><%call naam(rs)%></td>
			<td><%=rs("em_telefoon")%></td>
			<td><%=rs("em_naam")%> (<%=rs("em_relatie")%>)</td>
		</tr>	
	<%	rs.movenext : loop
	rs.close : set rs = nothing%>
	</table><%

case "betaald"
	adocon.execute("UPDATE contributie SET betaald=true WHERE lid_id="&request.querystring("l_id")&" AND jaartal_id="&skatejaar)
	response.redirect("default.asp?m=ledenlijst&a=contributie")

case "nietbetaald"
	adocon.execute("UPDATE contributie SET betaald=false WHERE lid_id="&request.querystring("l_id")&" AND jaartal_id="&skatejaar)
	response.redirect("default.asp?m=ledenlijst&a=contributie")

case "contributie"
	set rs = adocon.execute("SELECT COUNT(*) as aantal FROM lid WHERE actief=true AND isnull(datum_afgemeld)")
		Response.Write(rs("aantal") & " leden waarvan ")
	rs.close : set rs = nothing
	set rs = adocon.execute("SELECT count(*) as aantal FROM lid l LEFT JOIN contributie c ON l.lid_id = c.lid_id WHERE l.actief=true AND isnull(l.datum_afgemeld) AND c.betaald=true AND c.jaartal_id="&skatejaar)
		Response.Write(rs("aantal") & " betaald hebben.<br /><br />")
	rs.close : set rs = nothing
	
	%><table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0><%
	set rs = adocon.execute("SELECT * FROM lid WHERE actief=true AND isnull(datum_afgemeld) ORDER BY voornaam asc")
	i = ""
	do until rs.EOF
		set rs2 = adocon.execute("SELECT * FROM contributie WHERE lid_id="&rs("lid_id")&" AND jaartal_id="&skatejaar)
		if rs2.EOF then
			set rs2 = nothing
			adocon.execute("INSERT INTO contributie(lid_id, jaartal_id, ingevoerd_door) VALUES("&rs("lid_id")&", "&skatejaar&", "&session("functie_id")&")")
			set rs2 = adocon.execute("SELECT * FROM contributie WHERE lid_id="&rs("lid_id")&" AND jaartal_id="&skatejaar)
		end if%>
		<tr<%if rs2("betaald") = false then response.write(" bgcolor=""#FFCDCD""")end if%>>
			<td><a href="default.asp?m=lid&id=<%=rs("lid_id")%>&a=bewerk"><%call naam(rs)%></a></td>
			<td><%if not rs("email") = "" then
					response.write("<a href=""mailto:"&rs("email")&""">e-mail</a></td>")
				end if
				response.write("<td>")
				if rs("mobiel") <> "" then
					response.write telefoon(rs("mobiel"))
				else
					response.write telefoon(rs("telefoon"))
				end if%></td>
			<td align="right"><%
				if rs2("betaald") = true then
					response.write("<a href="""&huidigepagina(1)&"&a=nietbetaald&l_id="&rs("lid_id")&""">markeer als niet-betaald</a>")
				else
					response.write("<a href="""&huidigepagina(1)&"&a=betaald&l_id="&rs("lid_id")&""">markeer als betaald</a>")
					i = i & rs("email") & "; "
				end if
				%></td>
		</tr>
		<%set rs2 = nothing
	rs.movenext : loop
	rs.close : set rs = nothing%>
	<tr>
		<td colspan=4><h2>E-mailadressen van nog niet-betaalde leden</h2><%=i%></td>
	</tr>
	</table><%

case "sbn"
	sql = "SELECT * FROM lid WHERE actief=true AND isnull(datum_afgemeld) AND (lid_type_id=1 OR lid_type_id=4 OR lid_type_id=6) ORDER BY achternaam ASC"
	set rs = ExecQuery(sql,adocon)
	%>
	<h2>Ledenlijst (<%=rs.recordcount%> leden)</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td valign="top">Voor- en achternaam<br />
			SBN lidnummer
			</td>
		<td valign="top">Geslacht<br />
			Geboortedatum</td>
		<td valign="top">Adres<br />
			PC Woonplaats</td>
		<td valign="top"></td>
	</tr>
	<%do until rs.EOF%>
		<tr>
			<td valign="top"><%call naam(rs)%><br />
				<%=rs("sbn")%></td>
			<td><%=rs("geslacht")%>&nbsp;<br>
				<%=rs("geboortedatum")%>&nbsp;</td>
			<td><%=rs("adres")%>&nbsp;<br>
				<%=rs("postcode") & " " & rs("woonplaats")%>&nbsp;</td>
			<td></td>
		</tr>	
	<%	rs.movenext : loop
	rs.close : set rs = nothing%>
	</table><%

case "skateweekend"%>
	<h2>Aanmeldingen voor het skateweekend</h2>
	<%set rs = adocon.execute("SELECT * FROM skateweekend WHERE skatejaar="&skatejaar&" ORDER BY id DESC")
	if not rs.EOF then
	%><table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<%do until rs.EOF%>
	<tr>
		<td valign="top" nowrap><%=rs("naam")%></td>
		<td nowrap>Van: <%=rs("aankomst")%><br>Tot: <%=rs("vertrek")%></td>
		<td><%=rs("vervoer")%>
			<%if rs("vervoer") = "Auto / eigen vervoer" then
				Response.write(" ("&rs("vrijeplek")&")")
			end if%><br><%=rs("contactpersoon")%></td>
		<td><%=rs("eetwensen")%><br><%=rs("overigen")%></td>
	</tr>
	<%rs.movenext : loop
	rs.close%>
	</table>
	<%
	end if
	set rs = nothing
	
case "ontvangers"%>
	Onderstaande mensen wensen de nieuwsbrief per e-mail te ontvangen:
	<br><HR><br>
	<%
	teller = 0
	set rs = adocon.execute("SELECT email FROM lid WHERE isnull(datum_afgemeld)AND actief=true AND post=false ORDER BY email asc")
	do until rs.EOF
		if rs("email") <> "" then
			teller = teller+1
			if teller <> 1 then
				response.write("; ")
			end if
			response.write(rs("email"))
		end if
		rs.MoveNext
		loop
	rs.close
	set rs = nothing
	%>
	<br><br><hr><br>
	<br>
	Onderstaande mensen wensen de nieuwsbrief per post te ontvangen
	<br><HR><br>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td><strong>Naam</strong></td>
		<td><strong>Adres</strong></td>
		<td><strong>Postcode</strong></td>
		<td><strong>Woonplaats</strong></td>
	</tr>
	<%
	dim teller2
	teller2 = 0
	set rs = adocon.execute("SELECT * FROM lid WHERE isnull(datum_afgemeld) AND actief=true AND post=true ORDER BY voornaam asc")
	do until rs.EOF
		teller2 = teller2 + 1%>
		<tr>
			<td><%call naam(rs)%></td>
			<td><%=rs("adres")%></td>
			<td><%=rs("postcode")%></td>
			<td><%=rs("woonplaats")%></td>
		</tr>
		<%rs.movenext
		loop
	rs.close
	set rs = nothing%>
	</table>
	<br><HR><br>
	Totaal: <%=teller%> e-mailtjes en <%=teller2%> postzegels.
	<%

case else
	sql = "SELECT * FROM lid WHERE actief=true "
	if request.querystring("zoek") <> "" then
		sql = sql & "AND (voornaam like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "achternaam like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "adres like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "woonplaats like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "telefoon like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "mobiel like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "email like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "bankrekening like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "em_naam like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "em_telefoon like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "em_relatie like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "website like '%"&request.querystring("zoek")&"%' OR "
		sql = sql & "opmerkingen like '%"&request.querystring("zoek")&"%') "
	end if
	if request.querystring("a") <> "oudleden" then
		sql = sql & "AND isnull(datum_afgemeld) "
	end if
	if trim(request.querystring("orderby")) = "datum_aangemeld" then
		sql = sql & "ORDER BY datum_aangemeld DESC"
	else
		sql = sql & "ORDER BY voornaam ASC"
	end if
	set rs = ExecQuery(sql,adocon)
	if rs.recordcount = 1 then
		response.redirect(huidigepagina(0) & "?m=lid&id="&rs("lid_id")&"&a=bewerk")
	end if
	%>
	<table cellspacing=0 cellpadding=0 width="100%">
	<tr>
		<td valign="top"><h2>Ledenlijst (<%=rs.recordcount%> leden)</h2></td>
		<td align="right" valign="top">
			<form action="<%=huidigepagina(3)%>" method="post" name="zoekForm" onsubmit="document.location.href='<%=huidigepagina(3)%>&zoek='+document.forms['zoekForm'].zoek.value; return false">
			<input type="text" name="zoek" size=20 maxlength=40 value="<%=request.querystring("zoek")%>">
			<input type="submit" value="zoek">
			</form>
			</td>
	</tr>
	</table>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td valign="top">Voor- en achternaam<br />
      Aanmeldingsdatum</td>
		<td valign="top">Telefoonnummer<br>
			Mobiel nummer</td>
		<td valign="top">Adres<br>
			PC Woonplaats</td>
		<td valign="top">Bewerking</td>
	</tr>
	<%do until rs.EOF%>
		<tr onmouseover="highlightRow(this, 'over')" onmouseout="highlightRow(this, 'out')">
			<td valign="top">
				<%if rs("email") <> "" then%>
					<a href="mailto:<%=rs("email")%>" title="<%=rs("email")%>"><%call naam(rs)%></a>
          <br /><%=rs("datum_aangemeld")%>
				<%else%>
					<%call naam(rs)%>
				<%end if%></td>
			<td><%=telefoon(rs("telefoon"))%>&nbsp;<br>
				<%=telefoon(rs("mobiel"))%>&nbsp;</td>
			<td><%=rs("adres")%>&nbsp;<br>
				<%=rs("postcode") & " " & rs("woonplaats")%>&nbsp;</td>
			<td>
				<a href="<%=huidigepagina(0)%>?m=lid&id=<%=rs("lid_id")%>&a=bewerk"><img src="images/edit.gif" alt="Wijzig" border="0"></a>
				<%if request.querystring("a") <> "oudleden" then%>
					<a href="<%=huidigepagina(1)%>&id=<%=rs("lid_id")%>&a=dellid"><img src="images/delete.gif" alt="Verplaats naar oud-leden" border="0"></a>
				<%end if%>
				</td>
		</tr>	
	<%	rs.movenext : loop
	rs.close : set rs = nothing%>
	</table>
<%end select%>
