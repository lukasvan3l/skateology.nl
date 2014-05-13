<%if request.querystring("redir") <> "" then%>
	<h1>Nieuwsbrief</h1>
<%else%>
	<h1>Bestanden archief</h1>
<%end if


select case request.querystring("a")
case "uplFile"
	dim Upload, FileName, File
	Set Upload = Server.CreateObject ("Dundas.Upload.2")
	Upload.Save("C:\domains\skateology.nl\wwwroot\bestanden\")
	Upload.MaxFileSize = 3000000
	Upload.UseUniqueNames = true
	
	For Each File in Upload.Files
		FileName = replace(file.path,"C:\domains\skateology.nl\wwwroot\bestanden\", "")
		sql = "INSERT INTO bestand(type,url,jaar,maand,titel,ingevoerd_door) VALUES("&upload.form("type")&", '"&FileName&"', "&upload.form("jaar")&", "&upload.form("maand")&", '"&changequotes(upload.form("titel"))&"', "&session("functie_id")&")"
		response.write(sql)
		adocon.execute(sql)
		if request.querystring("redir") <> "" then
			set rs = adocon.execute("SELECT max(bestand_id) as bestand FROM bestand")
			response.redirect(huidigepagina(0) & "?m=email&ace=true&bijlage=" & rs("bestand") & "&a=" & request.querystring("redir"))
			set rs = nothing
		else
			response.redirect(huidigepagina(2))
		end if
	Next

case "delFile"
	set rs = adocon.execute("SELECT url FROM bestand WHERE bestand_id="&request.querystring("id"))
		Set fs=Server.CreateObject("Scripting.FileSystemObject")
		if fs.FileExists("C:\domains\skateology.nl\wwwroot\bestanden\" & rs("url")) then
			fs.DeleteFile("C:\domains\skateology.nl\wwwroot\bestanden\" & rs("url"))
		end if
	set rs = nothing
	adocon.execute("UPDATE bestand SET verwijderd=true WHERE bestand_id="&request.querystring("id"))
	response.redirect(huidigepagina(1) & "&alert=verwijderd")

case "nieuw"
	if request.querystring("redir") <> "" then%>
		<ul>
			<li><strong>Bijlagen uploaden</strong></li>
			<li>Tekst schrijven</li>
			<li>Tekst controleren</li>
			<li>Verzenden</li>
			<li>Nieuwsbrief per post verzenden</li>
		</ul>
	<%else%>
		<h2>Bestand uploaden</h2>
	<%end if%>
	<form action="<%=huidigepagina(2)%>&a=uplFile&redir=<%=request.querystring("redir")%>" method="post" name="uploadForm" id="uploadForm" enctype="multipart/form-data">
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="30%">Gegeven</td>
		<td width="70%">Waarde</td>
	</tr>
	<tr>
		<td>Titel</td>
		<td><input type="text" name="titel" size="40" maxlength="150" value="<%if request.querystring("redir") = "nieuwsbrief2" then response.write( "Skateology Nieuwsbrief "&maandvanjaar( month( now() ) ) ) end if%>"></td>
	</tr>
	<tr>
		<td>Type</td>
		<td><%if request.querystring("redir") = "nieuwsbrief2" then
				call ddlbBestand_type(1)
			else
				call ddlbBestand_type(5)
			end if%></td>
	</tr>
	<tr>
		<td>Datum<br>
		(in geval van Nieuwsbrief)</td>
		<td><%call ddlbNieuwsbrief_datum()%></td>
	</tr>
	<tr>
		<td>Bestand</td>
		<td><input type="file" name="Path" size="27"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>
	<%if request.querystring("redir") <> "" then
		%><a href="<%=huidigepagina(0) & "?m=email&ace=true&a=" & request.querystring("redir")%>">Als er geen bestand geupload hoeft te worden, klik hier.</a><%
	end if%>

<%case else%>
	<h2>Nieuwsbrieven</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="20%">Datum</td>
		<td width="50%">Titel</td>
		<td width="20%">Grootte</td>
		<td width="10%"></td>
	</tr>
	<%set rs = adocon.execute("SELECT * FROM bestand WHERE type=1 AND verwijderd=false ORDER BY jaar DESC, maand DESC")
	do until rs.EOF%>
		<tr onmouseover="highlightRow(this, 'over')" onmouseout="highlightRow(this, 'out')">
			<td><%=rs("jaar") & "-" & rs("maand")%></td>
			<td><%=rs("titel")%><br>
				<a href="http://www.skateology.nl/bestanden/<%=rs("url")%>" target=_blank><%
				if len(rs("url")) > 40 then
					response.write right(rs("url"),len(rs("url"))-39)
				else
					response.write rs("url")
				end if%></a></td>
			<td><%=filesize(rs("url"))%> Kb</td>
			<td><a href="<%=huidigepagina(2)%>&a=delFile&id=<%=rs("bestand_id")%>"><img src="images/delete.gif" alt="Verwijder dit bestand" border="0"></a></td>
		</tr>
	<%rs.movenext : loop
	rs.close : set rs =nothing%>
	</table>
	
	<h2>Overige documenten</h2>
	<table width="95%" class="tabledata" cellspacing=0 cellpadding=2 border=0>
	<tr class="trhead">
		<td width="20%">Type</td>
		<td width="50%">Titel</td>
		<td width="20%">Grootte</td>
		<td width="10%"></td>
	</tr>
	<%set rs = adocon.execute("SELECT * FROM bestand LEFT JOIN bestand_type ON bestand_type.bestand_type_id=bestand.type WHERE bestand.type<>1 and verwijderd=false ORDER BY bestand_type.type ASC, titel ASC")
	do until rs.EOF%>
		<tr onmouseover="highlightRow(this, 'over')" onmouseout="highlightRow(this, 'out')">
			<td valign="top"><%=rs("bestand_type.type")%></td>
			<td valign="top"><%=rs("titel")%><br>
				<a href="http://www.skateology.nl/bestanden/<%=rs("url")%>" target=_blank><%
				if len(rs("url")) > 40 then
					response.write right(rs("url"),len(rs("url"))-39)
				else
					response.write rs("url")
				end if%></a></td>
			<td valign="bottom"><%=filesize(rs("url"))%> Kb</td>
			<td><a href="<%=huidigepagina(2)%>&a=delFile&id=<%=rs("bestand_id")%>"><img src="images/delete.gif" alt="Verwijder dit bestand" border="0"></a></td>
		</tr>
	<%rs.movenext : loop
	rs.close : set rs =nothing%>
	</table>
	
<%end select%>