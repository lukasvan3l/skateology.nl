
<%sub r_ervaringen()%>
	<div class="rechtsdiv" style="background-color:#c6c">
	<h3>Skateology Ervaringen</h3>
	<%set rs = adocon.execute("SELECT * FROM ervaring ORDER BY datum DESC")
	if not rs.EOF then
	do until rs.EOF%>
		<a href="default.asp?a=ervaringen&id=<%=rs("ervaring_id")%>"><%=rs("titel")%> (<%=rs("auteur")%>)</a><br>
	<%rs.movenext : loop
	rs.close
	end if
	set rs = nothing%>
	</div>
	<br>
<%end sub%>

<%
sub r_banners()
  call r_rollerwave()
  call r_inlineskate()
end sub
%>

<%sub r_rollerwave()%>
	<div class="rechtsdiv" style="background-color:#fff" align="center">
    <a href="http://www.inlineskateshop.nl" target="_blank">
    	<img src="images/banner-rollerwave.gif" alt="Rollerwave" border="0">
    </a>
	</div>
	<br>
<%end sub%>

<%sub r_inlineskate()%>
	<div class="rechtsdiv" style="background-color:#fff" align="center">
    <a href="http://www.inline-skate.nl" target="_blank" title="inline-skate.nl">
      <img src="http://www.inline-skate.nl/promotie/inline-skate_01.gif" name="inline-skate.nl" alt="inline-skate.nl" border="0" height="60px" width="160px">
    </a>
	</div>
	<br>
<%end sub%>


<%sub r_logo()%>
	<div class="rechtsdiv" style="background-color:#fff" align="center">
	  <img src="images/logo-skateology.gif" alt="Skateology logo" border="0">
	</div>
	<br>
<%end sub%>

<%Sub r_twitter()%>	
	<div class="rechtsdiv twitterdiv" style="background-color:#c6c">
	  <h3><a href="http://www.twitter.com/skateology">Twitter</a></h3>

<%
		dim twitterurl, twitterxml, xmlHttp, twitterxmldom, tweets, tweet, tweettext, tweetdate, linkRegExp, linkRegExpMatches, linkRegExpMatch, tweettitle
		twitterurl = "http://twitter.com/statuses/user_timeline/skateology.xml"

		Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		xmlHttp.Open "GET", twitterurl, false
		xmlHttp.Send()
		twitterxml = xmlHttp.ResponseText

		Set twitterxmldom = Server.CreateObject("MSXML2.DomDocument.3.0")
		twitterxmldom.async = False
		twitterxmldom.validateOnParse = False
		twitterxmldom.resolveExternals = False

		If twitterxmldom.LoadXml(twitterxml) Then

		  Set tweets = twitterxmldom.getElementsByTagName("status")

		  For i = 0 To 6
			Set tweet = tweets.Item(i)
			tweettext = tweet.getElementsByTagName("text").Item(0).text
			If Left(tweettext,1) <> "@" then
				tweetdate = tweet.getElementsByTagName("created_at").Item(0).text

				'insert links where http:// or www.
				Set linkRegExp = New regexp 
				linkRegExp.Pattern = "http://[^\s]+"
				linkRegExp.Global = True 
				linkRegExp.IgnoreCase = True 
				Set linkRegExpMatches = linkRegExp.Execute(tweettext)

				For Each linkRegExpMatch In linkRegExpMatches
					tweettext = Replace(tweettext, linkRegExpMatch, "<a href='" & linkRegExpMatch & "'>" & linkRegExpMatch & "</a>")
				Next

				'insert links to tweeters
				Dim twitRegExp, twitRegExpMatch, twitRegExpMatches
				Set twitRegExp = New regexp 
				twitRegExp.Pattern = "@[^\s]+"
				twitRegExp.Global = True 
				twitRegExp.IgnoreCase = True 
				Set twitRegExpMatches = twitRegExp.Execute(tweettext)

				For Each twitRegExpMatch In twitRegExpMatches
					tweettext = Replace(tweettext, twitRegExpMatch, "<a href='http://twitter.com/" & Right(twitRegExpMatch, Len(twitRegExpMatch)-1) & "'>" & twitRegExpMatch & "</a>")
				Next

				'create title
				tweettitle = Left(tweettext, InStr(tweettext," "))

				%>
				<div class="tweet">
					<div class="tweetdatum">
						<%=Left(Right(Left(tweetdate,19),15),6)%>
						<%=" "%>
						<%=right(tweetdate,4)%>
						<%'=Right(Right(Left(tweetdate,19),15),8)%></div>
					<div class="tweetbericht"><%=tweettext%></div>
				</div>
				<%
			End if
		  next

		End if
		set xmlHttp = Nothing ' clear HTTP object
		Set twitterxmldom = Nothing ' clear XML

%>
	</div>
	<br>
<%End Sub%>

<%sub r_nieuwsbrief(nieuwsbrief_id)
	if nieuwsbrief_id = "" or isnull(nieuwsbrief_id) then
		set rs = adocon.execute("SELECT top 1 * FROM nieuwsbrief ORDER BY jaar DESC, maand DESC")
		nieuwsbrief_id = rs("nieuwsbrief_id")
	else
		set rs = adocon.execute("SELECT * FROM nieuwsbrief WHERE nieuwsbrief_id="&nieuwsbrief_id)
	end if
	%>
	<div class="rechtsdiv" style="background-color:#f90">
	<h3><%=maandvanjaar(rs("maand")) & " " & rs("jaar")%></h3>
	<%set rs = nothing
	set rs = adocon.execute("SELECT * FROM artikel WHERE nieuwsbrief_id="&nieuwsbrief_id&" ORDER BY volgorde ASC")
	if not rs.EOF then
	do until rs.EOF%>
		<a href="default.asp?a=nieuwsbrieven&id=<%=nieuwsbrief_id%>&artikel_id=<%=rs("artikel_id")%>"><%=rs("titel")%></a><br>
	<%rs.movenext : loop
	rs.close
	end if
	set rs = nothing%>
	</div>
	<br>
<%end sub%>

<%sub r_nieuwsbrieven(aantal)%>
	<div class="rechtsdiv" style="background-color:#fc0">
	<h3>Nieuwsbrieven Archief</h3>
	<%set rs = adocon.execute("SELECT top " & aantal & " * FROM nieuwsbrief ORDER BY jaar DESC, maand DESC")
	if not rs.EOF then
	jaar = rs("jaar")
	displayText("<strong>" & jaar & "</strong><div class=""agendaitems"">")
	do until rs.EOF
		if jaar <> rs("jaar") then
			jaar = rs("jaar")
			displayText("</div><strong>" & jaar & "</strong><div class=""agendaitems"">")
		end if
		%>
		<a href="default.asp?a=nieuwsbrieven&id=<%=rs("nieuwsbrief_id")%>"><%=maandvanjaar(rs("maand"))%></a><br>
	<%rs.movenext : loop
	rs.close
	end if
	set rs = nothing%>
	</div>
	<br>
<%end sub%>

<%sub r_agenda(aantal_geschiedenis)%>
	<div class="rechtsdiv" style="background-color:#3cf">
	<h3>Skateology Agenda</h3>
	<strong>Iedere week</strong>
	<div class="agendaitems">
	<a href="default.asp?a=agenda&id=10" class="agendaitems">Dinsdag - Avondskate</a><br>
	<a href="default.asp?a=agenda&id=11" class="agendaitems">Zondag - Middagskate</a>
	</div>
	<%set rs = adocon.execute("SELECT * FROM agenda WHERE datum>=FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & "','dd-mm-yyyy') ORDER BY datum ASC, agenda_id asc")
	if not rs.EOF then
	maand = month(rs("datum"))
	displayText("<strong>" & maandvanjaar(maand) & "</strong><div class=""agendaitems"">")
	do until rs.EOF
		if maand <> month(rs("datum")) then
			maand = month(rs("datum"))
			displayText("</div><strong>" & maandvanjaar(maand) & " ")
			displayText(year(rs("datum")) & "</strong><div class=""agendaitems"">")
		end if%>
		<a href="default.asp?a=agenda&id=<%=rs("agenda_id")%>" class="agendaitems"><%=tweegetallen(day(rs("datum"))) & " - " & rs("titel")%></a><br>
	<%rs.movenext : loop
	rs.close
	end if
	set rs = nothing%>
	</div></div>
	<br>
<%end sub%>

<%sub r_spelregels()%>
	<div class="rechtsdiv" style="background-color:#fc6">
	<h3>Gebruiksaanwijzing</h3>
	Op deze pagina zijn de algemene spelregels voor het skaten met Skateology te vinden.<br><br>
	Dit reglement is tevens downloadbaar als Adobe Acrobat bestand.<br>
	<a href="bestanden/algemeen_reglement_skateology.pdf">Algemeen Reglement.pdf</a> (58 Kb)<br><br>
	Vragen of opmerkingen kun je richten aan <a href="default.asp?a=vereniging">het bestuur</a>.
	</div>
	<br>
<%end sub%>

<%sub r_fotos()%>
	<div class="rechtsdiv" style="background-color:#999">
	<h3>Met dank aan...</h3>
	Tjerk<br>
	<a href="http://www.skateroutes.nl" target=_blank>www.skateroutes.nl</a>
	<br>
	<br>
	Lukas<br>
	<a href="http://www.3l.nl" target=_blank>www.3l.nl</a>
	<br>
	<br>
	Jan Jouke<br>
	<a href="http://picasaweb.google.nl/neohipper" target=_blank>JJ op Picasa</a>
	<br>
	<br>
	</div>
<%end sub

sub r_inschrijven()%>
	<div class="rechtsdiv" style="background-color:#99cc99;">
	<h3>Inschrijven voor Skateology</h3>
	<p>Om het inschrijfformulier zo toegankelijk mogelijk te maken is hij hier in twee formaten te vinden. De bovenste opent een nieuw venster met daarin een formulier dat u in kunt vullen en kunt printen. De onderste kunt u opslaan op de harde schijf en later uitprinten en invullen. </p>
  <p>Bij ondertekening van dit formulier geeft de deelnemer aan op de hoogte te zijn van het <a href="http://www.skateology.nl/bestanden/algemeen_reglement_skateology.pdf">Algemeen Reglement</a></p>
	<p><A onclick="window.open('popup_inschrijfformulier.asp','Inschrijfformulier','width=650, height=530,directories=no,location=no, menubar=no, scrollbars=yes,status=no,toolbar=no,resizable=yes');return false" href="popup_inschrijfformulier.asp">Inschrijfformulier invullen en printen</a><br>
		<em>HTML formulier</em></p>
	<!--<p><a href="bestanden/inschrijfformulier_acrobat.pdf" target="_blank">Inschrijfformulier downloaden</a><br>
		<em>Adobe Acrobat </em></p>--><%
end sub

sub r_3House()%>
	<div class="rechtsdiv" align="center">
	<img src="images/logo-3house.gif" width="150" height="114">
	</div>
<%end sub

sub r_skateweekend()
	%>
	
	<%
end sub
%>