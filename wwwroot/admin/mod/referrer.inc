<h1>Referrers</h1>

Via onderstaande pagina's is op www.skateology.nl gekomen. Soms zijn dit persoonlijke websites, soms zijn het zoekmachines. Deze laatste categorie is te herkennen aan een icoontje aan het begin van de regel. Het veld 'herkomst' toont het woord waarop is gezocht in de zoekmachine.<br><br>

<% funcReferrers 100, 100

function funcReferrers(aantalrefs, aantalchars)
	dim icon, bitSearch, bitsearchtot, j, refererAfgekort
	set rs = adocon.execute("SELECT * FROM referrer ORDER BY id desc")
	if not rs.EOF then
		response.write("<table width=""100%"">")
		j = 0 'j = aantalrefs
		do until rs.EOF or j > aantalrefs
			j = j+1
			response.write("<tr><td width=150>" & rs("datum") & "</td><td>")
			writeReferrer rs("referer"),aantalchars,rs("ipadres")
			response.write("</td></tr>")
		rs.movenext : loop
		rs.close
		response.write("</table>")
	end if
	set rs = nothing
end function

function writeReferrer(referrer, aantalchars, ipadres)
	dim icon, bitSearch, bitsearchtot, j, refererAfgekort
	referer = trim(lcase(referrer))
	icon = ""
	bitSearch = ""
	bitsearchtot = "&"
	if left(referer,7) = "http://" then
		referer = right(referer,len(referer)-7)
	end if
	
	if referrer <> "" then
		'IMAGES pagina's
		'YAHOO
		if instr(referer,"search.yahoo.") <> 0 and instr(referer,"images") <> 0 then
			bitsearch = "back="
		end if
		if bitsearch <>"" then
			i = instr(referer, bitsearch)+len(bitsearch)-1
			referer = right(referer,len(referer)-i)
			i = instr(referer, bitsearchtot)-1
			if i > 0 then
				referer = left(referer,i)
				referer = URLDecode(referer)
			end if
		end if
		
		'GOOGLE
		if instr(referer,".google.") and (instr(referer,"q=cache") or instr(referer,"images.google.") or instr(referer,"/custom?")) then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_g.png"" alt=""Google"" width=16 height=16> "
		elseif instr(referer,".google.") and instr(referer,"images.google") = 0 and (instr(referer,"search") <> 0) then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_g.png"" alt=""Google"" width=16 height=16> "
				bitsearch = "q="
'				end if
		'MSN SEARCH
		elseif instr(referer,"search.msn.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_msn2.gif"" alt=""MSN Search"" width=16 height=16> "
			bitsearch = "q="
		'PLANET GOOGLE
		elseif instr(referer,"googleq=q") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_planet.gif"" alt=""Planet Search"" width=16 height=16> "
			bitsearch = "q="
			bitsearchtot = "%26"
			referer = replace(referer,"q%3d","")
		'TISCALI GOOGLE
		elseif instr(referer,"search-dyn.tiscali.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_tiscali.png"" alt=""Tiscali Search"" width=16 height=16> "
			bitsearch = "key="
		'YAHOO
		elseif instr(referer,"search.yahoo.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_yahoo.gif"" alt=""Yahoo"" width=16 height=16> "
			bitsearch = "p="
		'MYWAY
		elseif instr(referer,"mysearch.myway") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_mysearch.gif"" alt=""MySearch"" width=16 height=16> "
			bitsearch = "searchfor="
		'ALTAVISTA
		elseif instr(referer,".altavista.") <> 0 and instr(referer,"randomlink") =0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_altavista.gif"" alt=""Altavista"" width=16 height=16> "
			bitsearch = "q="
		'DOWNLOADS.COM
		elseif instr(referer,".downloads.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_downloads.gif"" alt=""Downloads.com"" width=16 height=16> "
			bitsearch = "qry="
		'LYCOS.COM
		elseif instr(referer,".lycos.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_lycos.gif"" alt=""Lycos.com"" width=16 height=16> "
			bitsearch = "query="
		'ILSE.NL
		elseif instr(referer,".ilse.nl") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_ilse.gif"" alt=""Ilse.com"" width=16 height=16> "
			bitsearch = "search_for="
		'vinden.NL
		elseif instr(referer,".vinden.nl") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_vinden.gif"" alt=""Vinden.nl"" width=16 height=16> "
			bitsearch = "query="
		'AOL.de
		elseif instr(referer,"sucheaol.aol.de") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_aol.gif"" alt=""AOL.nl"" width=16 height=16> "
			bitsearch = "q="
		'WANADOO SEARCH
		elseif instr(referer,".wanadoo.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_wanadoo.gif"" alt=""Wanadoo.nl"" width=16 height=16> "
			bitsearch = "keywords="
		'ZOEKINALLES
		elseif instr(referer,".zoekinalles.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_zoekinalles.gif"" alt=""Wanadoo.nl"" width=16 height=16> "
			bitsearch = "keywords="
		'COMET SYSTEMS
		elseif instr(referer,".cometsystems.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_comet.gif"" alt=""Wanadoo.nl"" width=16 height=16> "
			bitsearch = "qry="
		'IXQUICK
		elseif instr(referer,"ixquick.com") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_ixquick.gif"" alt=""IXQuick.nl"" width=16 height=16> "
			bitsearch = "query="
		'AOL
		elseif instr(referer,"buscador.aol.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_aol.gif"" alt=""AOL"" width=16 height=16> "
			bitsearch = "query="
		'SKYNET
		elseif instr(referer,"search.skynet.") <> 0 then
			icon = "<img src=""http://www.3l.nl/inc/img/ico_skynet.gif"" alt=""Skynet"" width=16 height=16> "
			bitsearch = "keywords="
		end if
		
		if bitSearch = "" then
			i = instr(referer, "/")
			if i > 5 then
				referer = left(referer,i)
			end if
			if right(referer,1) = "/" then
				referer = left(referer,len(referer)-1)
			end if
			if left(referer,4) = "www." then
				referer = right(referer,len(referer)-4)
			end if
		else
			i = instr(referer, bitsearch)+len(bitsearch)-1
			referer = right(referer,len(referer)-i)
			i = instr(referer, bitsearchtot)-1
			if i > 0 then
				referer = left(referer,i)
			end if
			referer = URLDecode(referer)
		end if
		
		if icon = "" then
			icon = "<img src=""http://www.3l.nl/inc/img/spacer.gif"" height=16 width=16> "
		end if
		
		if len(referer) > aantalchars then
			refererAfgekort = left(referer,aantalchars) & "..."
		else
			refererAfgekort = referer
		end if%>
		<%=icon%><a href="<%=replaceHTMLTags(referrer)%>" target=_blank title="<%=ipadres%>"><%=replaceHTMLTags(refererAfgekort)%></a><%
	end if
end function

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
%>