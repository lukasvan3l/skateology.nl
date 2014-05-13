<h1>DB Omzetten van oud naar nieuw</h1>

<%
dim hoppa
hoppa = false
set rs = oudCon.execute("SELECT * FROM [Oud Leden] ORDER BY lidnummer ASC")
do until rs.EOF

response.write("<strong>" & rs("voornaam") & "</strong><br>")

sql = "INSERT INTO lid(voornaam, tussenvoegsel, achternaam, adres, postcode, woonplaats, telefoon, mobiel, email, bankrekening, geboortedatum, em_naam, em_telefoon, em_relatie, datum_aangemeld, datum_afgemeld, actief, post, lid_type_id, opmerkingen) VALUES("
sql = sql & "'" & changeQuotes(rs("voornaam")) & "', "
sql = sql & "'" & changeQuotes(rs("tussenvoegsel")) & "', "
sql = sql & "'" & changeQuotes(rs("achternaam")) & "', "
sql = sql & "'" & changeQuotes(rs("adres")) & "', "
sql = sql & "'" & changeQuotes(rs("postcode")) & "', "
sql = sql & "'" & changeQuotes(rs("woonplaats")) & "', "
sql = sql & "'" & changeQuotes(rs("telefoonnummer")) & "', "
sql = sql & "'" & changeQuotes(rs("mobiel_nummer")) & "', "
sql = sql & "'" & changeQuotes(rs("email")) & "', "
sql = sql & "'" & changeQuotes(rs("bankgiro")) & "', "
if rs("geboortedatum") <> "" then
	sql = sql & "FORMAT('" & tweegetallen(day(rs("geboortedatum"))) & "-" & tweegetallen(month(rs("geboortedatum"))) & "-" & year(rs("geboortedatum")) & "','dd-mm-yyyy'), "
else
	sql = sql & "null, "
end if
sql = sql & "'" & changeQuotes(rs("emergency_naam")) & "', "
sql = sql & "'" & changeQuotes(rs("emergency_nummer")) & "', "
sql = sql & "'" & changeQuotes(rs("emergency_relatie")) & "', "
sql = sql & "FORMAT('" & tweegetallen(day(rs("aanmelddatum"))) & "-" & tweegetallen(month(rs("aanmelddatum"))) & "-" & year(rs("aanmelddatum")) & "','dd-mm-yyyy'), "
if rs("afmeld_jaar") <> "" then 
	sql = sql & "FORMAT('01-09-" & year(rs("afmeld_jaar")) & "','dd-mm-yyyy'), "
else
	sql = sql & "null, "
end if
sql = sql & "false, "
if rs("nieuwsbriefid") = "1" then
	sql = sql & "true, "
else
	sql = sql & "false, "
end if
sql = sql & rs("soortlidid") & ", "
sql = sql & "'" & changeQuotes(rs("opmerkingen")) & "'"
sql = sql & ")"
response.write(sql & "<br>")
'adocon.execute(sql)
'response.write("<strong>Committed</strong>")
response.write("<br><br><br>")

rs.movenext : loop
rs.close : set rs = nothing
%>