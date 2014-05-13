<h1>Nieuwsbrieven</h1>
<%
select case request.querystring("a")
case "insBrief"
	sql = "INSERT INTO nieuwsbrief(jaar, maand, ingevoerd_door, ingevoerd_datum) VALUES("
	sql = sql & request.form("jaar") & ","
	sql = sql & request.form("maand") & ","
	sql = sql & session("functie_id") & ","
	sql = sql & "FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) &"','dd-mm-yyyy hh:mm:ss')"
	sql = sql &")"
	adocon.execute(sql)
	response.redirect(huidigepagina(1) & "&alert=opgeslagen")

case "updArtikel"
	sql = "UPDATE artikel SET titel='"&changequotes(request.form("titel"))&"', tekst='"&changequotes(request.form("tekst"))&"' WHERE artikel_id="&request.querystring("id")
	adocon.execute(sql)
	response.redirect(huidigepagina(1) & "&alert=opgeslagen")

case "insArtikel"
	sql = "INSERT INTO artikel(nieuwsbrief_id,volgorde, titel, tekst, ingevoerd_door) VALUES("
	sql = sql & request.form("nieuwbrief_id") & ","
	set rs = adocon.execute("SELECT top 1 * FROM artikel WHERE nieuwsbrief_id="&request.form("nieuwbrief_id")&" ORDER BY volgorde DESC")
	if rs.EOF then
		sql = sql & "1,"
	else
		sql = sql & rs("volgorde")+1 & ","
	end if
	set rs = nothing
	sql = sql & "'" & changequotes(request.form("titel")) & "',"
	sql = sql & "'" & changequotes(request.form("tekst")) & "',"
	sql = sql & session("functie_id") & ")"
	adocon.execute(sql)
	response.redirect(huidigepagina(1) & "&alert=opgeslagen")

case "nieuw"%>
	<form action="<%=huidigepagina(2)%>&a=insBrief" method="post" name="form" id="form">
	<h2>Nieuwe Nieuwsbrief</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="30%">Gegeven</td>
		<td width="70%">Waarde</td>
	</tr>
	<tr>
		<td>Jaar</td>
		<td><select name="jaar">
			<%for i = 1997 to year(now)+1
				response.write("<option value="&i&"")
				if i=year(now) then
					response.write(" selected")
				end if
				response.write(">"&i&"</option>")
			next%>
			</select></td>
	</tr>
	<tr>
		<td>Maand</td>
		<td><select name="maand">
			<%for i = 1 to 12
				response.write("<option value="&i&"")
				if i=month(now) then
					response.write(" selected")
				end if
				response.write(">"&maandvanjaar(i)&"</option>")
			next%>
			</select></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>
	<%set rs = nothing


case "nieuwartikel"%>
	<form action="<%=huidigepagina(2)%>&a=insArtikel" method="post" name="form" id="form">
	<input type="hidden" name="nieuwbrief_id" value="<%=request.querystring("id")%>">
	<h2>Nieuw artikel</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="30%">Gegeven</td>
		<td width="70%">Waarde</td>
	</tr>
	<tr>
		<td>Nieuwsbrief</td>
		<td><input type="text" maxlength="50" size="40" name="titel" value="<%
		set rs = adocon.execute("SELECT maand, jaar FROM nieuwsbrief WHERE nieuwsbrief_id="&request.querystring("id"))
		response.write(maandvanjaar(rs("maand")) & " " & rs("jaar"))
		set rs = nothing
		%>" disabled></td>
	</tr>
	<tr>
		<td>Titel</td>
		<td><input type="text" maxlength="50" size="40" name="titel" value=""></td>
	</tr>
	<tr>
		<td valign="top">Tekst</td>
		<td><textarea cols=60 rows=30 name="tekst"></textarea></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>
	<%set rs = nothing

case "bewerk"
	set rs = adocon.execute("SELECT * FROM artikel WHERE artikel_id="&request.querystring("id"))%>
	<form action="<%=huidigepagina(2)%>&a=updArtikel" method="post" name="form" id="form">
	<h2>Bewerk artikel</h2>
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
		<td valign="top">Tekst</td>
		<td><textarea cols=60 rows=30 name="tekst"><%=rs("tekst")%></textarea></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>
	<%set rs = nothing

case else
	%><h2>Overzicht</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="30%">Nieuwsbrief</td>
		<td width="50%">Artikel</td>
		<td width="20%"></td>
	</tr>
	<%
	set rs = adocon.execute("SELECT top 10 * FROM nieuwsbrief ORDER BY jaar DESC, maand DESC")
	do until rs.EOF
		response.write("<tr><td valign=top><strong>"&maandvanjaar(rs("maand")) & " " & rs("jaar") & "</strong></td>")
		response.write("<td>")
		
		set rs2 = adocon.execute("SELECT * FROM artikel WHERE nieuwsbrief_id="&rs("nieuwsbrief_id")&" ORDER BY volgorde ASC")
		if not rs2.EOF then
		do until rs2.EOF
			response.write("<a href="""&huidigepagina(1)&"&a=bewerk&id="&rs2("artikel_id")&""">"&rs2("titel")&"</a><br>")
		rs2.movenext : loop
		end if
		rs2.close : set rs2 = nothing
		
		response.write("</td><td align=""right"" valign=""top""><a href="""&huidigepagina(1)&"&a=nieuwartikel&id="&rs("nieuwsbrief_id")&""">Artikel toevoegen</a></td></tr>")
	rs.movenext : loop
	rs.close : set rs = nothing
	response.write("</table>")
end select
%>