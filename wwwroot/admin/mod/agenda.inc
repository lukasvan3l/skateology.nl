<h1>Agenda</h1>
<%
select case request.querystring("a")
case "insAgenda"
	sql = "INSERT INTO agenda(titel, datum, bericht, contactpersoon, ingevoerd_door, prijs) VALUES("
	sql = sql & "'" & changequotes(request.form("titel")) & "',"
	sql = sql & "FORMAT('" & request.form("begindatum") & "','dd-mm-yyyy')"
	sql = sql & ",'" & changequotes(request.form("omschrijving")) & "',"
	sql = sql & request.form("lid_id") & ","
	sql = sql & session("functie_id")
	sql = sql & "," & replace(request.form("prijs"),",",".")
	sql = sql & ")"
	adocon.execute(sql)
	response.redirect(huidigepagina(1) & "&alert=opgeslagen")

case "updAgenda"
	sql = "UPDATE agenda SET "
	sql = sql & "titel='" & changequotes(request.form("titel")) & "',"
	sql = sql & "datum=FORMAT('" & request.form("begindatum") & "','dd-mm-yyyy'),"
	sql = sql & "bericht='" & changequotes(request.form("omschrijving")) & "',"
	sql = sql & "contactpersoon="&request.form("lid_id") & ","
	sql = sql & "prijs=" & replace(request.form("prijs"),",",".")
	sql = sql & " WHERE agenda_id="&request.querystring("id")
	response.write(sql)
	adocon.execute(sql)
	response.redirect(huidigepagina(1) & "&alert=opgeslagen")

case "nieuw"%>
	<form action="<%=huidigepagina(2)%>&a=insAgenda" method="post" name="form" id="form">
	<h3>Nieuw agendapunt</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="30%">Gegeven</td>
		<td width="70%">Waarde</td>
	</tr>
	<tr>
		<td>Titel</td>
		<td><input type="text" maxlength="50" size="40" name="titel" value=""></td>
	</tr>
	<tr>
		<td>(Begin-) Datum</td>
		<td><input type="text" maxlength="10" size="15" name="begindatum" value=""> <em>dd-mm-jjjj</em></td>
	</tr>
	<tr>
		<td>Contactpersoon</td>
		<td><%call ddlblid(0)%></td>
	</tr>
	<tr>
		<td valign="top">Omschrijving</td>
		<td><textarea cols=60 rows=8 name="omschrijving"></textarea></td>
	</tr>
	<tr>
		<td>Prijs</td>
		<td>€ <input type="text" maxlength="5" size="10" name="prijs" value="<%=formatnumber(0)%>"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>
	<%

case "bewerk"
	set rs = adocon.execute("SELECT * FROM agenda WHERE agenda_id="&request.querystring("id"))%>
	<form action="<%=huidigepagina(2)%>&a=updAgenda" method="post" name="form" id="form">
	<h2>Agendapunt bewerken</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="30%">Gegeven</td>
		<td width="70%">Waarde</td>
	</tr>
	<tr>
		<td>Titel</td>
		<td><input type="text" maxlength="50" size="40" name="titel" value="<%=rs("titel")%>"></td>
	</tr>
	<tr>
		<td>(Begin-) Datum</td>
		<td><input type="text" maxlength="10" size="15" name="begindatum" value="<%=date2datum(rs("datum"),false)%>"> <em>dd-mm-jjjj</em></td>
	</tr>
	<tr>
		<td>Contactpersoon</td>
		<td><%call ddlblid(rs("contactpersoon"))%></td>
	</tr>
	<tr>
		<td valign="top">Omschrijving</td>
		<td><textarea cols=60 rows=8 name="omschrijving"><%=rs("bericht")%></textarea></td>
	</tr>
	<tr>
		<td>Prijs</td>
		<td>€ <input type="text" maxlength="5" size="10" name="prijs" value="<%=formatnumber(rs("prijs"))%>"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>
	<%set rs = nothing

case else
	set rs = adocon.execute("SELECT agenda.*,lid.voornaam FROM agenda LEFT JOIN lid ON lid.lid_id=agenda.contactpersoon WHERE datum>now ORDER BY datum ASC")%>
	<h2>Overzicht agendapunten</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="20%">Datum</td>
		<td width="50%">Titel</td>
		<td width="20%">Contactpersoon</td>
		<td width="10%">Prijs</td>
	</tr>
	<%do until rs.EOF
		response.write("<tr><td>"&date2datum(rs("datum"),false)&"</td>")
		response.write("<td><a href="""&huidigepagina(1)&"&a=bewerk&id="&rs("agenda_id")&""">"&rs("titel")&"</a></td>")
		response.write("<td>"&rs("voornaam")&"</td>")
		response.write("<td>&euro; "&formatnumber(rs("prijs"))&"</td></tr>")
	rs.movenext : loop
	rs.close : set rs = nothing
	response.write("</table>")
end select
%>