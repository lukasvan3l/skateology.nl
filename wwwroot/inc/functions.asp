<%
'//////////////////////
' Variabelen declareren
'//////////////////////
dim rs, rs2, teller, i, j, sql, rsFunctions, referer
dim oudCon, adoCon, skatejaar, fs, f, jaar, maand




'//////////////////////
' databaseconnecties
' worden in functionsclose.asp weer afgesloten
'//////////////////////

Set adoCon = Server.CreateObject ("ADODB.Connection")
adoCon.Mode = 3
adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\domains\skateology.nl\db\skateology.mdb"

set rs = adocon.execute("SELECT jaartal_id FROM jaartal WHERE begindatum < date() AND isnull(einddatum)")
	skatejaar = rs("jaartal_id")
rs.close : set rs = nothing




'//////////////////////
' rechten per persoon
'//////////////////////

function checkUser()
	if session("username") = "" or session("functie_id") = "" or session("lid_id") = "" then
		response.redirect("login.asp")
	end if
end function


'//////////////////////
' tekst netjes weergeven
'//////////////////////

function displayText(text)
  text = replaceEmailLinks(text)
  Response.Write(text)
end function

function replaceEmailLinks(text)
  dim RegularExpressionObject 
  Set RegularExpressionObject = New RegExp
  
  With RegularExpressionObject
    .Pattern = "mailto:(.*?)@(.*?)\.nl"
    .IgnoreCase = True
    .Global = True
  End With
  
  replaceEmailLinks = RegularExpressionObject.Replace(text, "javascript:mailto('$1','$2')")
  Set RegularExpressionObject = nothing
end function


'//////////////////////
' skateology functies
'//////////////////////


function IsValidEmail(email)
	if trim(email) = "" then
		IsValidEmail = false
		exit function
	end if
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
		IsValidEmail = false
		exit function
	end if
	for each name in names
		if Len(name) <=  0 then
			IsValidEmail = false
			exit function
		end if
		for i = 1 to Len(name)
			c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
				IsValidEmail = false
				exit function
			end if
		next
		if Left(name, 1) = "." or Right(name, 1) = "." then
			IsValidEmail = false
			exit function
		end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
		exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
		IsValidEmail = false
		exit function
	end if
	if InStr(email, "..") > 0 then
		IsValidEmail = false
	end if
end function

sub uploadmeuk()
	Dim UploadProgress, PID, barref
	Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
	PID = "PID=" & UploadProgress.CreateProgressID()
	barref = "framebar.asp?to=10&" & PID
	%>
	<SCRIPT LANGUAGE="JavaScript">
	function ShowProgress()
	{
		strAppVersion = navigator.appVersion;
		if (document.uploadForm.Path.value != "")
		{
			if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
			{
				winstyle = "dialogWidth=375px; dialogHeight:130px; center:yes";
				window.showModelessDialog('<% = barref %>&b=IE',null,winstyle);
			}
			else
			{
			window.open('<% = barref %>&b=NN','','width=370,height=115', true);
			}
		}
		return true;
	}
	</SCRIPT><%
end sub

sub naam(rsFunctions)
	response.write(rsFunctions("voornaam") & " ")
	if rsFunctions("tussenvoegsel") <> "" then
		response.write(rsFunctions("tussenvoegsel") & " ")
	end if
	response.write(rsFunctions("achternaam"))
end sub

Function telefoon(telnrFunction)
	if telnrFunction <> "" then
		telnrFunction = trim(replace(telnrFunction,"-",""))
		if len(telnrFunction) <> 10 then
			telefoon = telnrFunction
		else
			if left(telnrFunction,2) = "06" then 'mobiel nummer
				telefoon = left(telnrFunction,2) & "-" & right(telnrFunction,8)
			else 'vaste telefoon
				telefoon = left(telnrFunction,3) & "-" & right(telnrFunction,7)
			end if
		end if
	end if
end function

Function website(websiteFunction)
	if left(websiteFunction,7) = "http://" then
		website = websiteFunction
	else
		if len(websiteFunction) = 0 then
		website = ""
		elseif left(websiteFunction,4) = "www." then
			website = "http://" & websiteFunction
		else 'eigenlijk even tellen hoeveel punten er in voorkomen...
			website = "http://" & websiteFunction
		end if
	end if
end function

sub aantalBestanden(rsFunction)
	teller = 0
	for i = 1 to 3
		if rsFunction("bijlage"&i) <> "" then
			teller = teller+1
		end if
	next
	if teller = 1 then
		response.write(" (" & teller & " bijlage)")
	elseif teller>1 then
		response.write(" (" & teller & " bijlagen)")
	end if
end sub

function filesize(urlFunction)
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	set f=fs.GetFile(server.mappath("../bestanden/") & "\" & urlFunction)
		filesize = formatnumber(f.Size/1000,0)
	set f=nothing
	set fs=nothing
end function

sub ddlblid(geselecteerd)
	%><select name="lid_id">
	<option value="null">- kies een lid -</option>
	<%
	set rsFunctions = ExecQuery("SELECT * FROM lid WHERE actief=true AND (datum_afgemeld>date() or isnull(datum_afgemeld)) ORDER BY voornaam ASC, achternaam ASC", adocon)
	do until rsFunctions.EOF%>
		<option value="<%=rsFunctions("lid_id")%>"<%
			if Cint(rsFunctions("lid_id")) = geselecteerd then
				response.write(" selected")
			end if%>><%call naam(rsFunctions)%></option><%
	rsFunctions.movenext : loop
	rsFunctions.close : set rsFunctions = nothing
	%></select><%
end sub

sub ddlbOudLid(geselecteerd)
	%><select name="lid_id">
	<option value="null">- kies een (oud-)lid -</option>
	<%
	set rsFunctions = ExecQuery("SELECT * FROM lid WHERE actief=true ORDER BY voornaam ASC, achternaam ASC", adocon)
	do until rsFunctions.EOF%>
		<option value="<%=rsFunctions("lid_id")%>"<%
			if Cint(rsFunctions("lid_id")) = geselecteerd then
				response.write(" selected")
			end if%>><%call naam(rsFunctions)%></option><%
	rsFunctions.movenext : loop
	rsFunctions.close : set rsFunctions = nothing
	%></select><%
end sub

sub ddlbLid_type(geselecteerd)
	%><select name="lid_type_id">
	<option value="null">- kies een type -</option>
	<%
	set rsFunctions = ExecQuery("SELECT * FROM lid_type ORDER BY soortlid ASC", adocon)
	do until rsFunctions.EOF%>
		<option value="<%=rsFunctions("soortlidid")%>"<%
			if Cint(rsFunctions("soortlidid")) = geselecteerd then
				response.write(" selected")
			end if%>><%=rsFunctions("soortlid")%></option><%
	rsFunctions.movenext : loop
	rsFunctions.close : set rsFunctions = nothing
	%></select><%
end sub

sub ddlbBestand_type(geselecteerd)
	%><select name="type">
	<option value="null">- kies een type -</option>
	<%
	set rsFunctions = ExecQuery("SELECT * FROM bestand_type ORDER BY type ASC", adocon)
	do until rsFunctions.EOF%>
		<option value="<%=rsFunctions("bestand_type_id")%>"<%
			if Cint(rsFunctions("bestand_type_id")) = geselecteerd then
				response.write(" selected")
			end if%>><%=rsFunctions("type")%></option><%
	rsFunctions.movenext : loop
	rsFunctions.close : set rsFunctions = nothing
	%></select><%
end sub

sub ddlbNieuwsbrief_datum()
	%><select name="jaar"><%
	for i = year(now)+1 to 1997 step -1
		%><option value="<%=i%>"<%if i=year(now) then response.write(" selected")%>><%=i%></option><%
	next
	%></select>&nbsp;<%
	
	%><select name="maand"><%
	for i = 1 to 12
		%><option value="<%=i%>"<%if i=month(now) then response.write(" selected")%>><%=maandvanjaar(i)%></option><%
	next
	%></select>&nbsp;<%
end sub

sub ddlbBijlage(selectName, selectedValue)
	%><select name="<%=selectName%>">
	<option value="null"> - Selecteer een bijlage - </option><%
	i = ""
	set rsFunctions = adocon.execute("SELECT * FROM bestand LEFT JOIN bestand_type ON bestand_type.bestand_type_id = bestand.type WHERE verwijderd=false ORDER BY bestand_type.type asc, jaar desc,maand desc")
	do until rsFunctions.EOF
		if i <> rsFunctions("bestand_type_id") then
			%><option value="null"></option><option value="null"><%=rsFunctions("bestand_type.type")%></option><%
			i = rsFunctions("bestand_type_id")
		end if
		%><option value="<%=rsFunctions("bestand_id")%>"<%if rsFunctions("bestand_id") = selectedValue then response.write(" selected") end if%>>&nbsp;&nbsp;<%=rsFunctions("titel")%></option><%
	rsFunctions.Movenext : loop
	rsFunctions.close : set rsFunctions = nothing
	%></select><%
end sub

sub ddlbGeadresseerden(selectedValue)
	%><select name="ontvangers">
	<%set rsFunctions = adocon.execute("SELECT * FROM mailing_ontvangers ORDER BY ontvangers ASC")
	do until rsFunctions.EOF
		%><option value="<%=rsFunctions("mailing_ontvangers_id")%>"<%if selectedValue =rsFunctions("mailing_ontvangers_id") then response.write(" selected") end if%>><%=rsFunctions("ontvangers")%></option><%
	rsFunctions.Movenext : loop
	rsFunctions.close : set rsFunctions = nothing
	%></select><%
end sub

sub bestandnaam(idee)
	if isnumeric(idee) and not idee = "" then
	set rsFunctions = adocon.execute("SELECT titel FROM bestand WHERE bestand_id="&idee)
	response.write(rsFunctions("titel") & "<br>")
	set rsFunctions = nothing
	end if
end sub

'//////////////////////
' global functions
'//////////////////////

'Vult variabele maandvanjaar met de maand voluit in het Nederlands

function maandvanjaar(maandnr)
	Select case maandnr
	case "1"  maandvanjaar = "Januari"
	case "2"  maandvanjaar = "Februari"
	case "3"  maandvanjaar = "Maart"
	case "4"  maandvanjaar = "April"
	case "5"  maandvanjaar = "Mei"
	case "6"  maandvanjaar = "Juni"
	case "7"  maandvanjaar = "Juli"
	case "8"  maandvanjaar = "Augustus"
	case "9"  maandvanjaar = "September"
	case "10"  maandvanjaar = "Oktober"
	case "11"  maandvanjaar = "November"
	case "12"  maandvanjaar = "December"
	End Select
end function

function dagvanweek(dagnr)
	SELECT CASE dagnr
	Case "1" dagvanweek = "Zondag"
	Case "2" dagvanweek = "Maandag"
	Case "3" dagvanweek = "Dinsdag"
	Case "4" dagvanweek = "Woensdag"
	Case "5" dagvanweek = "Donderdag"
	Case "6" dagvanweek = "Vrijdag"
	Case "7" dagvanweek = "Zaterdag"
	END SELECT
end function

function changeQuotes(string)
	if not string = "" then
		changequotes = replace(string, "'", "''")
		changequotes = trim(changequotes)
	end if
end function

Function ExecQuery (strQ, objCon)
	Dim objRS
	Set objRS = Server.CreateObject("ADODB.RecordSet")
	
	objRS.CursorLocation = adUseClient
	objRS.CursorType = adOpenForwardOnly
	objRs.LockType = adLockReadOnly
	
	objRS.Open strQ, objCon, , , adCmdText
	Set objRS.ActiveConnection = Nothing
	
	Set ExecQuery = objRS
End Function

Function huidigePagina(tm)
	huidigePagina = "http://"
	huidigePagina = huidigePagina & request.servervariables("HTTP_HOST")
	huidigePagina = huidigePagina & request.servervariables("SCRIPT_NAME")
	if tm >= 1 then
		huidigePagina = huidigePagina & "?m=" & request.querystring("m")
	end if
	if tm >= 2 AND request.querystring("id") <> "" then
		huidigePagina = huidigePagina & "&id=" & request.querystring("id")
	end if
	if tm >= 3 AND request.querystring("a") <> "" then
		huidigePagina = huidigePagina & "&a=" & request.querystring("a")
	end if
	if tm >= 4 AND request.querystring("zoek") <> "" then
		huidigePagina = huidigePagina & "&zoek=" & request.querystring("zoek")
	end if
End Function

Sub ACEJavaScript(obj, Functype)%>
	<input type="hidden" name="txtContent"  value="" ID="Hidden1">
	<script>
	var <%=obj%> = new ACEditor("<%=obj%>")
	<%if Functype = "tekst" then%>
		<%=obj%>.isFullHTML = false
		<%=obj%>.width = "475"
		<%=obj%>.height = 500
		<%=obj%>.useImage = true
		<%=obj%>.PageStyle = "frontend.css"
		<%=obj%>.PageStylePath_RelativeTo_EditorPath = "../css/";
	<%elseif Functype = "email" then%>
		<%=obj%>.isFullHTML = true
		<%=obj%>.width = "475"
		<%=obj%>.height = 500
		<%=obj%>.useImage = false
	<%end if%>
	
	<%=obj%>.usePrint = false
	<%=obj%>.useParagraph  = false
	<%=obj%>.useFontName = false
	<%=obj%>.useSize = true
	<%=obj%>.useText = false
	<%=obj%>.useSelectAll = true
	<%=obj%>.useCut = true
	<%=obj%>.useCopy = true
	<%=obj%>.usePaste = true
	<%=obj%>.useUndo = true
	<%=obj%>.useRedo = true
	<%=obj%>.useBold = true
	<%=obj%>.useItalic = true
	<%=obj%>.useUnderline = true
	<%=obj%>.useStrikethrough = true
	<%=obj%>.useSuperscript = false
	<%=obj%>.useSubscript = false
	<%=obj%>.useSymbol = true
	<%=obj%>.useJustifyLeft = true
	<%=obj%>.useJustifyCenter = true
	<%=obj%>.useJustifyRight = true
	<%=obj%>.useJustifyFull = true
	<%=obj%>.useNumbering = true
	<%=obj%>.useBullets = true
	<%=obj%>.useIndent = true
	<%=obj%>.useOutdent = true
	<%=obj%>.useForeColor = true
	<%=obj%>.useBackColor = true
	<%=obj%>.useExternalLink = true
	<%=obj%>.useTable = true
	<%=obj%>.useShowBorder = true
	<%=obj%>.useAbsolute = false
	<%=obj%>.useClean = false
	<%=obj%>.useAsset = false
	<%=obj%>.useLine = true
	<%=obj%>.usePageProperties = false
	<%=obj%>.useWord = false
	<%=obj%>.useSave = false
	<%=obj%>.useZoom = false
	<%=obj%>.useInternalLink = false
	<%=obj%>.useStyle  = true
	<%=obj%>.StyleSelection = "ace_selection.css";
	<%=obj%>.StyleSelectionPath_RelativeTo_EditorPath = "../css/";
	<%=obj%>.RUN()
	</script>
	<%
End Sub


function tweegetallen(input)
	if len(input) = 1 then
		tweegetallen = Cstr("0") & Cstr(input)
	elseif input = "" or isnull(input) then
		tweegetallen = ""
	else
		tweegetallen = Cstr(input)
	end if
end function


function date2datum(input, time)
	if len(trim(input)) = 0 or len(trim(input)) = "" or isnull(len(trim(input))) then
		date2datum = ""
	else
		date2datum = tweegetallen(day(input)) & "-" & _
			tweegetallen(month(input))  & "-" & _
			tweegetallen(year(input)) 
		if time = true then
			date2datum = date2datum & " " & _
			tweegetallen(hour(input))   & ":" & _
			tweegetallen(minute(input)) & ":" & _
			tweegetallen(second(input))
		end if
	end if
end function


Sub writePageCountLink(link, tekst, qstr)
	Response.Write(	"<A HREF=""" &request.servervariables("SCRIPT_NAME")& "?pagina=" & link )
	for each i in qstr
		if not request.querystring(i) = "" then
			response.write( "&" & i & "=" &request.querystring(i) )
		end if
	next
	response.write( """" )
	response.write( " title=""Naar pagina " )
	response.write( link & """>" & tekst & "</A> " )
End Sub


Sub Print_Navigation(newsRS, nPage, Geschied)
	Dim nRecCount	' Number of records found
	Dim nPageCount	' Number of pages of records we have
	Dim lokatie     ' Geeft aan waar vandaan deze sub is aangeroepen (blok of content)
	Dim p
	
	nRecCount = newsRS.RecordCount
	nPageCount = newsRS.PageCount
	
	'qstr bevat alle querystrings die meegegeven worden als je op volgende of een nummertje klikt.
	'om zelf eentje toe te voegen, verhoog qstr(int) met ééntje, en voeg je eigen querystring toe
	'aan de array hieronder.
	Dim qstr(12)		' Mee te geven querystrings
	qstr(0) = "id"
	qstr(1) = "m"
	qstr(2) = "r_id"
	qstr(3) = "m_id"
	qstr(4) = "action"
	qstr(5) = "actie"
	qstr(6) = "zoek"
	qstr(7) = "f_id"
	qstr(8) = "keywords"
	qstr(9) = "l_id"
	qstr(10) = "sub"
	qstr(11) = "c_id"
	qstr(12) = "orderby"
	
	If nPage < 1 Or nPage > nPageCount Then
		nPage = 1
	End If
	
	dim vorigetekst, volgendetekst
	vorigetekst = "vorige"
	volgendetekst = "volgende"
	
	If Not (nPageCount = 1) Then
		
		if nPageCount < (2*geschied)+1 then
			If Not nPage = 1 Then
				call writePageCountLink(nPage-1, vorigetekst, qstr)
			else
				response.write(vorigetekst)
			End If
			
			response.write(" | ")
			
			For p = 1 To nPageCount
				If Not Cint(nPage) = Cint(p) then
					if Cint(nPageCount) > 9 then
						call writePageCountLink(p,tweegetallen(p),qstr)
					else
						call writePageCountLink(p,p,qstr)
					end if
				else
					if Cint(nPageCount) > 9 then
						response.write( "<strong><u>" & tweegetallen(p) & "</u></strong> " )
					else
						response.write( "<strong><u>" & p & "</u></strong> " )
					end if
				end if
			Next
			
			response.write("| ")
			
			If Not Cint(nPage) = Cint(nPageCount) Then
				call writePageCountLink(nPage+1, volgendetekst, qstr)
			else
				response.write(volgendetekst)
			End If
		
		else
		
			If Not nPage = 1 Then
				call writePageCountLink(1, "<<", qstr)
				response.write(" ")
				call writePageCountLink(nPage-1, "<", qstr)
			else
				response.write("<< <")
			End If
			
			response.write(" | ")
			
			For p = 1 to nPageCount
				if (p >= (nPage-geschied) AND p <= (nPage+geschied)) then
					If Not Cint(nPage) = Cint(p) then
						if Cint(nPageCount) > 9 then
							call writePageCountLink(p,tweegetallen(p),qstr)
						else
							call writePageCountLink(p,p,qstr)
						end if
					else
						if Cint(nPageCount) > 9 then
							response.write( "<strong><u>" & tweegetallen(p) & "</u></strong> " )
						else
							response.write( "<strong><u>" & p & "</u></strong> " )
						end if
					end if
				end if
			Next
			if nPage < (nPageCount-geschied) then
				'response.write(" >")
			end if
			
			response.write("| ")
			
			If Not Cint(nPage) = Cint(nPageCount) Then
				call writePageCountLink(nPage+1, ">", qstr)
				response.write(" ")
				call writePageCountLink(nPageCount, ">>", qstr)
			else
				response.write("> ")
				response.write(">>")
			End If
		
		end if
	End If
End Sub

Function deleteFile(file)
	'file = "\content_images\fotos\bla.jpg"
	dim fs
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists( Server.MapPath("/") & file ) then
		fs.DeleteFile( Server.MapPath("/") & file )
	else
		response.write("Het bestand kon niet worden gevonden en is niet verwijderd.")
	end if
	set fs = nothing
End Function


Function URLDecode(What)
	'URL decode Function
	'2001 Antonin Foller, PSTRUH Software, http://www.motobit.com
	 Dim Pos, pPos
	
	 'replace + To Space
	 What = Replace(What, "+", " ")
	
	 on error resume Next
	 Dim Stream: Set Stream = CreateObject("ADODB.Stream")
	 If err = 0 Then 'URLDecode using ADODB.Stream, If possible
	   on error goto 0
	   Stream.Type = 2 'String
	   Stream.Open
	
	   'replace all %XX To character
	   Pos = InStr(1, What, "%")
	   pPos = 1
	   on error resume next
	   Do While Pos > 0
		 Stream.WriteText Mid(What, pPos, Pos - pPos) + _
	       Chr(CLng("&H" & Mid(What, Pos + 1, 2)))
	     pPos = Pos + 3
	     Pos = InStr(pPos, What, "%")
	   Loop
	   Stream.WriteText Mid(What, pPos)
	
	   'Read the text stream
	   Stream.Position = 0
	   URLDecode = Stream.ReadText
	
	   'Free resources
	   Stream.Close
	 Else 'URL decode using string concentation
	   on error goto 0
	   'UfUf, this is a little slow method. 
	   'Do Not use it For data length over 100k
	   Pos = InStr(1, What, "%")
	   Do While Pos>0 
	     What = Left(What, Pos-1) + _
	       Chr(Clng("&H" & Mid(What, Pos+1, 2))) + _
	       Mid(What, Pos+3)
	     Pos = InStr(Pos+1, What, "%")
	   Loop
	   URLDecode = What
	 End If
End Function

Function replaceHTMLTags(input)
'laat <strong> zien ipv dat de tekst vetgedrukt wordt
 	replaceHTMLTags = changeQuotes(input)
	replaceHTMLTags = Replace(replaceHTMLTags, """","&quot;")
	replaceHTMLTags = Replace(replaceHTMLTags, "'", "&quot;")
	replaceHTMLTags = Replace(replaceHTMLTags, "<", "&lt;")
	replaceHTMLTags = Replace(replaceHTMLTags, ">", "&gt;")
	replaceHTMLTags = Replace(replaceHTMLTags, "€", "&euro;")
End Function
%>