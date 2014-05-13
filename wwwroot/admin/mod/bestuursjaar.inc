<h1>Bestuursfuncties</h1>

<%select case request.querystring("a")
case "insbestuurslid"
	if request.form("bestuur") = "true" then
		set rs = adocon.execute("SELECT email, telefoon, mobiel FROM lid WHERE lid_id="&request.form("lid_id"))
		sql = "INSERT INTO functie(jaartal_id, lid_id, functie, bestuur, email, telefoonnummer, wachtwoord, username) VALUES("
		sql = sql & request.querystring("id") & ","
		sql = sql & request.form("lid_id") & ",'"
		sql = sql & changequotes(request.form("functie")) & "',"
		sql = sql & request.form("bestuur") & ",'"
		sql = sql & changequotes(rs("email")) & "','"
		if rs("mobiel") <> "" then	
			sql = sql & telefoon(rs("mobiel")) & "','"
		else
			sql = sql & telefoon(rs("mobiel")) & "','"
		end if
		sql = sql & changequotes(request.form("wachtwoord")) & "','"
		sql = sql & changequotes(request.form("username")) & "'"
		sql = sql & ")"
		set rs = nothing
	else
		sql = "INSERT INTO functie(jaartal_id, lid_id, functie, bestuur) VALUES("
		sql = sql & request.querystring("id") & ","
		sql = sql & request.form("lid_id") & ",'"
		sql = sql & changequotes(request.form("functie")) & "',false)"
	end if
	response.write(sql)
	adocon.execute(sql)
	response.redirect(huidigepagina(2) & "&a=edit&alert=toegevoegd")

case "editbestuurslid"
	adocon.execute("UPDATE functie SET functie='"&changequotes(request.form("functie"))&"', wachtwoord='"&changequotes(request.form("wachtwoord"))&"', username='"&changequotes(request.form("username"))&"' WHERE functie_id=" + request.form("functie_id"))
	response.redirect(huidigepagina(2) & "&a=edit&alert=opgeslagen")

case "delbestuurslid"
	adocon.execute("DELETE FROM functie WHERE functie_id="&request.querystring("f_id"))
	response.redirect(huidigepagina(2) & "&a=edit&alert=verwijderd")

case "updskateweekend"
	sql = "UPDATE jaartal SET "
	if request.form("vertrekdatum") <> "" then
		sql = sql & "vertrekdatum=FORMAT('" & request.form("vertrekdatum") & "','dd-mm-yyyy'),"
	end if
	if request.form("retourdatum") <> "" then
		sql = sql & "retourdatum=FORMAT('" & request.form("retourdatum") & "','dd-mm-yyyy'),"
	end if
	if request.form("prijs") <> "" then
		sql = sql & "prijs="& replace(request.form("prijs"),",",".") & ","
	end if
	sql = sql & "bestemming='" & changequotes(request.form("bestemming"))& "'"
	sql = sql & " WHERE jaartal_id="&request.querystring("id")
	response.write(sql)
	adocon.execute(sql)
	response.redirect(huidigepagina(2) & "&a=edit&alert=opgeslagen")

case "nieuw"
	


case "edit"
	set rs = ExecQuery("SELECT * FROM jaartal WHERE jaartal_id="&request.querystring("id"), adocon)%>
	
	<h2><%=date2datum(rs("begindatum"),false)%> t/m <%
		if rs("einddatum") = "" or isnull(rs("einddatum")) then
			response.write(year(rs("begindatum"))+1)
		else
			response.write(date2datum(rs("einddatum"),false))
		end if
		%></h2>
	
	<h3>Functies</h3>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td valign="top">Functie</td>
		<td valign="top">Lid<br>E-mail</td>
		<td valign="top">Inlognaam<br>Telefoonnummer</td>
		<td valign="top">Bewerking</td>
	</tr>
	<%set rs2 = ExecQuery("SELECT * FROM functie LEFT JOIN lid ON functie.lid_id=lid.lid_id WHERE jaartal_id=" & rs("jaartal_id")& " ORDER BY bestuur ASC", adocon)
	do until rs2.EOF%>
		<tr onmouseover="highlightRow(this, 'over')" onmouseout="highlightRow(this, 'out')">
			<td valign="top"><%=rs2("functie")%></td>
			<td valign="top"><%call naam(rs2)%><br>
				<a href="mailto:<%=rs2("functie.email")%>"><%=rs2("functie.email")%></a></td>
			<td valign="top"><%=rs2("username")%><br>
				<%=rs2("telefoonnummer")%></td>
			<td valign="top">
				<a href="<%=huidigepagina(3)%>&f_id=<%=rs2("functie_id")%>"><img src="images/edit.gif" alt="Wijzig" border="0"></a>
				<a href="<%=huidigepagina(2)%>&f_id=<%=rs2("functie_id")%>&a=delbestuurslid"><img src="images/delete.gif" alt="Verwijder" border="0"></a></td>
		</tr>
		<%
	rs2.movenext : loop
	rs2.close : set rs2 = nothing%>
	</table>
	

	<% if request.querystring("f_id") = "" then %>
		<h3>Nieuwe functie</h3>
		<form action="<%=huidigepagina(2)%>&a=insbestuurslid" method="post">
		<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
		<tr>
			<td>Functie</td>
			<td><input name="functie" type="text" value="" size="40" maxlength="100"></td>
		</tr>
		<tr>
			<td>(Oud-)Lid:</td>
			<td><%call ddlbOudLid(0)%></td>
		</tr>
		<tr>
			<td>Bestuursfunctie</td>
			<td><input type="checkbox" name="bestuur" value="true"></td>
		</tr>
		<tr>
			<td>Inlognaam</td>
			<td><input type="text" name="username" size="20" maxlength="20"></td>
		</tr>
		<tr>
			<td>Wachtwoord</td>
			<td><input type="password" name="wachtwoord" size="10" maxlength="10"></td>
		</tr>
		<tr>
			<td></td>
			<td><input type="submit" value="toevoegen"></td>
		</tr>
		</table>
		</form><br>
	<%else
		set rs3 = adocon.execute("SELECT jaartal_id, lid_id, functie, bestuur, email, telefoonnummer, wachtwoord, username FROM functie WHERE functie_id=" & request.querystring("f_id"))
		%>

		<h3>Bewerk functie</h3>
		<form action="<%=huidigepagina(2)%>&a=editbestuurslid" method="post">
		<input type="hidden" name="functie_id" value="<%=request.querystring("f_id")%>" />
		<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
		<tr>
			<td>Functie</td>
			<td><input type="text" name="functie" size="40" maxlength="100" value="<%=rs3("functie")%>"></td>
		</tr>
		<tr>
			<td>Inlognaam</td>
			<td><input type="text" name="username" size="20" maxlength="20" value="<%=rs3("username")%>"></td>
		</tr>
		<tr>
			<td>Wachtwoord</td>
			<td><input type="text" name="wachtwoord" size="10" maxlength="10" value="<%=rs3("wachtwoord")%>"></td>
		</tr>
		<tr>
			<td></td>
			<td><input type="submit" value="toevoegen"></td>
		</tr>
		</table>
		</form><br>
		<%
		set rs3 = nothing
	end if%>
	
	<h3>Skateweekend</h3>
	<form action="<%=huidigepagina(2)%>&a=updskateweekend" method="post">
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr>
		<td>Bestemming</td>
		<td><input type="text" name="bestemming" maxlength="200"  size="40" value="<%
			if rs("bestemming") <> "" then
				response.write(rs("bestemming"))
			end if%>"></td>
	</tr>
	<tr>
		<td>Vertrekdatum</td>
		<td><input type="text" name="vertrekdatum" maxlength="10"  size="10" value="<%
			if rs("vertrekdatum") <> "" then
				response.write(date2datum(rs("Vertrekdatum"),false))
			end if%>"></td>
	</tr>
	<tr>
		<td>Retourdatum</td>
		<td><input type="text" name="retourdatum" maxlength="10"  size="10" value="<%
			if rs("Retourdatum") <> "" then
				response.write(date2datum(rs("Retourdatum"),false))
			end if%>"></td>
	</tr>
	<tr>
		<td>Prijs per persoon</td>
		<td>€ <input type="text" name="prijs" maxlength="6"  size="8" value="<%
			if rs("prijs") <> "" and rs("prijs") <> 0 then
				response.write(formatnumber(rs("prijs")))
			end if%>"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>
	<%set rs = nothing

case else

	set rs = ExecQuery("SELECT * FROM jaartal ORDER BY begindatum DESC", adocon)
	do until rs.EOF
		%>
		<table width="100%" cellpadding=0 cellspacing=0><tr>
		<td><h2><%=year(rs("begindatum"))%> t/m <%
		if rs("einddatum") = "" or isnull(rs("einddatum")) then
			response.write(year(rs("begindatum"))+1)
		else
			response.write(year(rs("einddatum")))
		end if
		%></h2></td>
		<td align="right"><a href="default.asp?m=<%=request.querystring("m")%>&a=edit&id=<%=rs("jaartal_id")%>"><img src="images/edit.gif" alt="" border="0"></a></td>
		</tr></table>
		<table width="90%" border=0 align="center">
		<tr><td valign="top">
		<fieldset align="center">
		<legend>Functies</legend>
		<table width="280">
		<%set rs2 = ExecQuery("SELECT * FROM functie LEFT JOIN lid ON functie.lid_id=lid.lid_id WHERE jaartal_id=" & rs("jaartal_id")& " ORDER BY bestuur, functie ASC", adocon)
		do until rs2.EOF%>
			<tr>
				<td><%=rs2("functie")%></td>
				<td><%call naam(rs2)%></td>
			</tr>
		<%rs2.movenext : loop
		rs2.close : set rs2 = nothing%>
		</table>
		</fieldset>
		</td><td valign="top">
		<fieldset>
		<legend>Skateweekend</legend>
		<%if rs("bestemming") = "" or isnull(rs("bestemming")) then%>
			<p>Er is (nog) geen skateweekend gehouden dit jaar.</p>
		<%else%>
		<table width="280">
		<tr>
			<td>Bestemming</td>
			<td><%=rs("bestemming")%></td>
		</tr>
		<tr>
			<td>Vertrekdatum</td>
			<td><%=date2datum(rs("vertrekdatum"), false)%></td>
		</tr>
		<tr>
			<td>Retourdatum</td>
			<td><%=date2datum(rs("retourdatum"), false)%></td>
		</tr>
		<tr>
			<td>Prijs</td>
			<td>€ <%=formatnumber(rs("prijs"))%></td>
		</tr>
		</table>
		<%end if%>
		</fieldset>
		</td></tr></table><br><br>
		<%
	rs.movenext : loop
	rs.close : set rs = nothing

end select%>