<h1>Teksten</h1>

<%select case request.querystring("a")
case "updTekst"
	adocon.execute("UPDATE tekst SET tekst='"& changeQuotes(request.form("txtContent")) &"', datum=FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) &"','dd-mm-yyyy hh:mm:ss'), gewijzigd_door="& session("functie_id")&" WHERE tekst_id="&request.querystring("id"))
	response.redirect(huidigepagina(1) & "&alert=opgeslagen")

case ""%>
	<h2>Teksten overzicht</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td valign="top">Titel</td>
		<td valign="top">Laatst gewijzigd</td>
	</tr>
	<%
	set rs = adocon.execute("SELECT * FROM tekst LEFT JOIN (functie LEFT JOIN lid ON lid.lid_id=functie.lid_id) ON tekst.gewijzigd_door=functie.functie_id ORDER BY titel ASC")
	do until rs.EOF%>
		<tr>
			<td><a href="<%=huidigepagina(1)%>&a=<%=rs("tekst_id")%>"><%=rs("titel")%></a></td>
			<td>op <%=rs("datum")%> door <%call naam(rs)%></td>
		</tr>
		<%rs.movenext : loop
	rs.close : set rs = nothing
	%></table><%

case else
	set rs = adocon.execute("SELECT * FROM tekst WHERE tekst_id="&request.querystring("a"))
	%>
	<h2><%=rs("titel")%> aanpassen</h2>
	<form action="<%=huidigepagina(1)%>&a=updTekst&id=<%=request.querystring("a")%>" method="post" name="Form">
	<TEXTAREA rows=7 cols=40 ID="txtContent" NAME="txtContent" style="width:100%;"><%=rs("tekst")%></TEXTAREA>
	<input type="submit" value="Opslaan" />
	</form>
	<%set rs = nothing%>
	
<%end select%>
