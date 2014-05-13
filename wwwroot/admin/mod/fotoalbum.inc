<h1>Fotoalbum</h1>

<%if request.querystring("a") = "nieuw" then%>
	<form method="POST" action="?m=fotoalbum&a=toevoegen">
	<table>
	<tr>
		<td>Titel</td>
		<td><input type="text" name="titel" size=40></td>
	</tr>
	<tr>
		<td>Link</td>
		<td><input type="text" name="link" size=40 value="http://"></td>
	</tr>
	<tr>
		<td>Thumbnail</td>
		<td><input type="text" name="thumbnail" size=40 value="http://"></td>
	</tr>
	<tr>
		<td>Aantal</td>
		<td><input type="text" name="aantal" size=4></td>
	</tr>
	<tr>
		<td>Startdatum (dd-mm-yyyy)</td>
		<td><input type="text" name="startdatum" size=10></td>
	</tr>
	<tr>
		<td>Toondatum</td>
		<td><input type="text" name="toondatum" size=40></td>
	</tr>
	<tr>
		<td>Fotograaf</td>
		<td><select name="fotograaf_id">
			<option value="">- kies -</option>
			<%
			set rs = adocon.execute("SELECT * FROM fotograaf ORDER BY naam ASC")
			do until rs.EOF
				response.write("<option value="&rs("fotograaf_id")&">"&rs("naam")&"</option>")
			rs.movenext : loop
			rs.close : set rs = nothing
			%>
			</select></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="opslaan"></td>
	</tr>
	</table>
	</form>

<%elseif request.querystring("a") = "toevoegen" then
	dim query
	query = "INSERT INTO fotoalbum(titel, link, thumbnail, aantal, fotograaf_id, toondatum, startdatum) VALUES ("
	query = query & "'" & request.form("titel") & "',"
	query = query & "'" & request.form("link") & "',"
	query = query & "'" & request.form("thumbnail") & "',"
	query = query & request.form("aantal") & ","
	query = query & request.form("fotograaf_id") & ","
	query = query & "'" & request.form("toondatum") & "',"
	query = query & "FORMAT('" & request.form("startdatum") & "','dd-mm-yyyy'))"
	
	response.write(query)
	adocon.execute(query)
	response.redirect("?m=fotoalbum&alert=opgeslagen")

elseif request.querystring("a") = "verwijderen" then
	query = "DELETE FROM fotoalbum WHERE fotoalbum_id=" & request.querystring("id")
	adocon.execute(query)
	response.redirect("?m=fotoalbum&alert=verwijderd")

elseif request.querystring("a") = "bewerken" then
	set rs = adocon.execute("SELECT * FROM  fotoalbum WHERE fotoalbum_id=" & request.querystring("id"))
	%>
	<form method="POST" action="?m=fotoalbum&a=editsave">
	<input type="hidden" value="<%=rs("fotoalbum_id")%>" name="id" />
	<table>
	<tr>
		<td>Titel</td>
		<td><input type="text" name="titel" size=40 value="<%=rs("titel")%>"></td>
	</tr>
	<tr>
		<td>Link</td>
		<td><input type="text" name="link" size=40 value="<%=rs("link")%>"></td>
	</tr>
	<tr>
		<td>Thumbnail</td>
		<td><input type="text" name="thumbnail" size=40 value="<%=rs("thumbnail")%>"></td>
	</tr>
	<tr>
		<td>Aantal</td>
		<td><input type="text" name="aantal" size=4 value="<%=rs("aantal")%>"></td>
	</tr>
	<tr>
		<td>Startdatum (dd-mm-yyyy)</td>
		<td><input type="text" name="startdatum" size=10 value="<%=rs("startdatum")%>"></td>
	</tr>
	<tr>
		<td>Toondatum</td>
		<td><input type="text" name="toondatum" size=40 value="<%=rs("toondatum")%>"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="opslaan"></td>
	</tr>
	</table>
	</form>
	<%
	set rs = nothing
elseif request.querystring("a") = "editsave" then
	
	query = "UPDATE fotoalbum SET "
	query = query & "titel='" & changequotes(request.form("titel")) & "',"
	query = query & "link='" & changequotes(request.form("link")) & "',"
	query = query & "thumbnail='" & changequotes(request.form("thumbnail")) & "',"
	query = query & "aantal=" & changequotes(request.form("aantal")) & ","
	query = query & "toondatum='" & changequotes(request.form("toondatum")) & "',"
	query = query & "startdatum=FORMAT('" & request.form("startdatum") & "','dd-mm-yyyy') "
	query = query & "WHERE fotoalbum_id=" & request.form("id")

	response.write(query)
	adocon.execute(query)
	response.redirect("?m=fotoalbum&alert=opgeslagen")

else
	response.write("<table width='100%'>")
	set rs = adocon.execute("SELECT fotoalbum.*, fotograaf.naam FROM fotoalbum, fotograaf WHERE fotoalbum.fotograaf_id = fotograaf.fotograaf_id ORDER BY startdatum DESC,titel ASC")
	do until rs.EOF
	%>
	
	<tr>
		<td><%=rs("toondatum")%></td>
		<td><%=rs("titel")%></td>
		<td><%=rs("naam")%></td>
		<td><%=rs("aantal")%></td>
		<td><a href="?m=fotoalbum&a=verwijderen&id=<%=rs("fotoalbum_id")%>" title="Verwijder dit fotoalbum">delete</a></td>
		<td><a href="?m=fotoalbum&a=bewerken&id=<%=rs("fotoalbum_id")%>" title="Bewerk dit fotoalbum">edit</a></td>
	</tr>
	
	<%
	rs.movenext : loop
	rs.close : set rs = nothing
	response.write("</table>")
end if%>