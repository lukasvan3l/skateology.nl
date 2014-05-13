<%sub l_welkom()%>
	<h1>Skaten in Leiden en omgeving</h1>
	<%set rs = adocon.execute("SELECT * FROM tekst WHERE tekst_id=1")
	do until rs.EOF

		displayText(rs("tekst"))

	rs.movenext : loop
	rs.close : set rs = nothing%>
<!--
	<p>Ben je tussen de 18 en 35 jaar (of voel je je zo jong) en op zoek naar andere inlineskaters? Dan is inline skate vereniging Skateology misschien wel iets voor jou!</p>
	<p>Sinds 1997 rijdt Skateology elke <a href="default.asp?a=agenda&id=10">dinsdagavond</a> en <a href="default.asp?a=agenda&id=11">zondagmiddag</a> een skatetocht door Leiden en omstreken. Op dinsdag verzamelen we om 20.00 uur op de Beestenmarkt in Leiden. Skate gerust een keer mee! Ook in de winter gaan we gewoon door. Kijk voor andere Skateology activiteiten, zoals deelname aan night skates, zondag specials en het skate weekend  in de <a href="default.asp?a=agenda">agenda</a>!</p>
	<p>Ge&iuml;nteresseerd?</p>
	<p>Kom een keer langs op een dinsdagavond (lees wel eerst even de <a href="default.asp?a=spelregels">spelregels</a>) of neem contact op met iemand van <a href="default.asp?a=vereniging">het bestuur</a>. Lidmaatschap kost &euro;12,50 per jaar. <a href="default.asp?a=inschrijven">Klik hier voor een inschrijfformulier.</a></p>
	<p class="bijschrift"><strong>Skatelessen?</strong></p>
	<blockquote>
		<p class="bijschrift"> Skateology heeft als partner 3HOUSE. Dit bedrijf organiseert skatelessen en events&nbsp;voor alle leeftijden en niveaus. Voor meer informatie <a href="default.asp?a=3House"><font color="#999999">klik hier</font></a>. </p>
	</blockquote>
	<p class="bijschrift"><strong>Belangrijke mededeling van het bestuur over veiligheid</strong></p>
	<ul>
		<li class="bijschrift">Het dragen van een helm, pols-, knie- en elleboogbeschermers wordt door het bestuur dringend aangeraden. Een val is altijd mogelijk en kan ernstig letsel tot gevolg hebben. Deelname aan Skateology activiteiten is altijd op eigen risico. </li>
		<li class="bijschrift">Verlichting is bij schemering en donker verplicht: jouw zichtbaarheid is ook voor de andere deelnemers van belang.</li>
		<li class="bijschrift">Minimaal kunnen remmen, wendbaar en stabiel zijn. Voor eventuele skatelessen verwijzen we je door naar <a href="default.asp?a=3House"><font color="#999999">3House</font></a>. </li>
		<li class="bijschrift">Skateology rijdt alleen bij droog wegdek. </li>
	</ul>-->
<%end sub

sub l_nieuwsbrief()
	if request.querystring("artikel_id") = "" and request.querystring("id") = "" then
		set rs2 = adocon.execute("SELECT top 1 * FROM nieuwsbrief ORDER BY jaar DESC, maand DESC")
		set rs = adocon.execute("SELECT top 1 * FROM artikel WHERE nieuwsbrief_id="&rs2("nieuwsbrief_id")&" ORDER BY volgorde ASC")
		set rs2 = nothing
	elseif request.querystring("artikel_id") = "" then
		set rs = adocon.execute("SELECT top 1 * FROM artikel WHERE nieuwsbrief_id="&request.querystring("id")&" ORDER BY volgorde ASC")
	else
		set rs = adocon.execute("SELECT * FROM artikel WHERE nieuwsbrief_id="&request.querystring("id")&" AND artikel_id="&request.querystring("artikel_id"))
	end if
	
	if not rs.EOF then
		response.write("<h1>")
		if not request.querystring("artikel_id") = "" then
			response.write rs("titel")
		else
			set rs2 = adocon.execute("SELECT maand, jaar FROM nieuwsbrief WHERE nieuwsbrief_id="&rs("nieuwsbrief_id"))
				displayText("Skateology Nieuwsbrief " & maandvanjaar(rs2("maand")) & " " & rs2("jaar"))
			set rs2 = nothing
		end if
		%></h1>
		<p><%=replace(rs("tekst"), vbcrlf, "<br>")%></p>
		<%	
	end if

end sub

sub l_agenda()
	if request.querystring("id") = "" then
		set rs = adocon.execute("SELECT * FROM agenda LEFT JOIN lid ON agenda.contactpersoon=lid.lid_id WHERE agenda_id=10")
		'set rs = adocon.execute("SELECT TOP 1 * FROM agenda LEFT JOIN lid ON agenda.contactpersoon=lid.lid_id WHERE datum >= FORMAT('" & day(now) & "-" & month(now) & "-" & year(now) & "','dd-mm-yyyy') ORDER BY datum ASC, agenda_id asc")
	else
		set rs = adocon.execute("SELECT * FROM agenda LEFT JOIN lid ON agenda.contactpersoon=lid.lid_id WHERE agenda_id="&request.querystring("id"))
	end if
	if rs.EOF then
		displayText("<h1>Geen agendaitems</h1>")
	else%>
		<h1><%=rs("titel")%></h1>
		<p><%
		if rs("agenda_id") = 10 then
			displayText("Iedere dinsdagavond")
		elseif rs("agenda_id") = 11 then
			displayText("Iedere zondagmiddag")
		else
			displayText(dagvanweek(DatePart("w", rs("datum"))) & " " & day(rs("datum")) & " " & maandvanjaar(month(rs("datum"))) & " " & year(rs("datum")))
		end if%></p>
		<blockquote><%=replace(rs("bericht"),vbcrlf,"<br>")%></blockquote>
		<%
		if rs("prijs") <> 0 then
			displayText("<p><strong>Verbonden kosten:</strong><br>&euro; "&formatnumber(rs("prijs"))&"</p>")
		end if
		
		if rs("contactpersoon") <> "" then
			displayText("<p><strong>Contactpersoon:</strong><br>")
			set rs2 = adocon.execute("SELECT * FROM functie LEFT JOIN lid ON functie.lid_id=lid.lid_id WHERE jaartal_id="&skatejaar&" AND bestuur=true AND functie.lid_id="&rs("contactpersoon"))
			if not rs2.EOF then
				call naam(rs2)
				'displayText("<br>"&rs2("telefoonnummer"))
				displayText("<br><a href=""info@skateology.nl"">info@skateology.nl</a>")
				displayText("</p>")
			else
				call naam(rs)
				displayText("<br><a href=""info@skateology.nl"">info@skateology.nl</a>")
				displayText("</p>")
			end if
		end if
	end if
	set rs = nothing
end sub

sub l_fotos()%>
	<h1>Fotoalbums</h1>
	<br>
       <script type="text/javascript">username='skateologyleiden'; photosize='400'; columns='3';</script>
       <script type="text/javascript" src="/inc/pwa.js"></script>
<%end sub

sub l_routes()%>
	<h1>Skate routes</h1>
	<p><a href="http://www.skateroutes.nl" target=_blank>www.skateroutes.nl</a><br>
	Sinds jaar en dag houdt Tjerk zich liefdevol bezig met het routebeheer van Skateology. Op zijn website zijn veel routes te vinden, door heel Nederland en ver daarbuiten. Nagenoeg alle routes die Skateology rijdt en heeft gereden worden op deze website gepresenteerd met printbare routekaart.</p>
	<p><a href="http://www.inline-skate.nl/skateroutes" target=_blank>www.inline-skate.nl</a><br>
	Een portal voor in-line skaters waar ook veel skateroutes te vinden zijn door heel Nederland.</p>
	<p><a href="http://home.wanadoo.nl/skeelerroutes/" target=_blank>Skeelerroutes in Zeeland</a><br>
	Deze website bevat een aantal zeer gedetailleerde kaarten van de provincie Zeeland. Met kleurcodes is van bijna elke weg de kwaliteit aangegeven.</p>
<%end sub

sub l_spelregels()%>
	<h1>Algemeen Reglement Skateology</h1>
	<%set rs = adocon.execute("SELECT * FROM tekst WHERE tekst_id=2")
	do until rs.EOF

		displayText(rs("tekst"))

	rs.movenext : loop
	rs.close : set rs = nothing%>
	
	<!--<p>Onderstaand het Algemeen Reglement van Skateology zoals deze op 27 mei door de ALV is goedgekeurd. Dit reglement is een uitbreiding op de statuten van de vereniging. Om de volledige statuten in handen te krijgen kun je een verzoek sturen naar <a href="javascript:mailto('info')">Skateology</a>.</p>
	
	<p><strong>Artikel 1: Algemene bepalingen</strong></p>
	<ol>
		<li>De vereniging genaamd Skateology, hierna te noemen &quot;de vereniging&quot; is opgericht op 26 september 1997 en is gevestigd te Leiden.</li>
		<li>Het algemeen reglement is van toepassing in onverbrekelijke samenhang met de statuten van de vereniging, zoals deze laatstelijk zijn gewijzigd en opnieuw vastgesteld door de algemene ledenvergadering de dato 22 september 2002. </li>
	</ol>
	
	<p><strong>Artikel 2: Definities</strong></p>
	<ol>
		<li>(Aspirant) leden dienen 18 jaar of ouder te zijn.</li>
		<li>De dinsdagtocht is toegankelijk voor leden en aspirant leden, zondagtochten zijn alleen toegankelijk voor leden.</li>
		<li>Aspirant leden mogen &eacute;&eacute;nmaal vrijblijvend meeskaten, daarna dienen zij zich in te schrijven als lid.</li>
		<li>Het bestuur kan in voorkomende gevallen uitzonderingen maken op de in lid 2 genoemde zondagtochten en op de in lid 3 genoemde eenmalige introductie. </li>
		<li>Verenigingsactiviteiten zijn activiteiten die zijn aangekondigd in de verenigingsagenda. Skateroutes kunnen in gezamenlijk overleg worden aangepast door de deelnemers. </li>
	</ol>
	
	<p><strong>Artikel 3: Risico's</strong></p>
	<ol>
		<li>Het beoefenen van de skatesport draagt risico's met zich mee. In het bijzonder wordt hierbij gedoeld op in het algemeen aan de skatesport verbonden risico's in onverwachte situaties als gevolg van onder meer: 
			<ol type="a"><li>Algemene instabiliteit inherent aan het dragen van skates</li>
				<li>Obstakels en oneffenheden in het wegdek. </li>
				<li>Skaten op de openbare weg: skaters hebben de wettelijke status van voetganger, maar de snelheid van fietsers. Dit kan leiden tot onvoorspelbare reacties van andere verkeersdeelnemers.</li>
				<li>Skaten in het verkeer op een niet afgezet parcours. </li>
				<li>Skaten in groepen: afleiding door sociale interactie, beperkter zicht op mogelijke obstakels en nabijheid van andere skaters. </li>
				<li>Skaten op niet vooraf verkend terrein. </li>
				<li>Skaten in een wisselend tempo, met onderling verschillende snelheden. </li>
				<li> Skaten op nat wegdek met o.a. vermindering van de stabiliteit en een verlenging van de remweg tot gevolg, </li>
			</ol>
		</li>
		<li>Ongelukken bij het skaten kunnen ernstige gevolgen hebben. Daarbij wordt in het bijzonder gedoeld op (ernstige) verwondingen, botbreuken en hoofdletsel, met mogelijk de dood tot gevolg. </li>
		<li>Het bestuur heeft de plicht (aspirant) leden te wijzen op de risico&acute;s die het beoefenen van de skatesport met zich meebrengt. Zij voert daartoe een actief veiligheidsbeleid. </li>
		<li>Deelnemers aan activiteiten zijn zich bewust van de risico&acute;s en dragen bij aan de beperking van deze risico&acute;s. (Aspirant) leden nemen onder eigen risico en verantwoording deel aan Skateology activiteiten. </li>
	</ol>
	
	<p><strong>Artikel 4: Beschermingsmiddelen</strong></p>
	<ol>
		<li>Het dragen van beschermende middelen wordt door de vereniging dringend geadviseerd. Hierbij gaat het in het bijzonder om de volgende skate-beschermingsmiddelen:
			<ol type="a">
				<li>Een skate-, wielren- of mountainbikehelm die de boven-, voor-, zij- en achterkant van het hoofd goed beschermt, bij voorkeur een CE goedgekeurde helm. </li>
				<li>Polsbeschermers</li>
				<li>Kniebeschermers</li>
				<li>Elleboogbeschermers</li>
			</ol>
		</li>
		<li> Het niet of slechts gedeeltelijk dragen van beschermingsmiddelen geschiedt geheel voor eigen risico en verantwoording van de deelnemer. </li>
	</ol>
	
	<p><strong>Artikel 5: Verlichting</strong></p>
	<ol>
		<li>Het dragen van verlichting bij beoefening van de skatesport in verenigingsverband is vanaf de inval van de schemering verplicht. Hierbij dient de deelnemer minimaal werkende en goed zichtbare verlichting te voeren aan zowel de voor- als achterzijde. Bij voorkeur dient de voorzijde voorzien te zijn van een wit of geel licht, de achterzijde van een rood licht. </li>
	</ol>
	
	<p><strong>Artikel 6: Skatevaardigheid</strong></p>
	<ol>
		<li>Van mensen die (willen) deelnemen aan een Skateology activiteit wordt verwacht dat ze de skatetechniek in zoverre beheersen dat ze: 
			<ol type="a"><li>goed kunnen remmen</li>
				<li>goed wendbaar zijn en goed kunnen sturen </li>
				<li>goed kunnen omgaan met oneffenheden in het wegdek </li>
				<li>stabiel zijn op de skates, het skaten onder controle hebben </li>
				<li>goed verkeersinzicht hebben </li>
				<li> oplettend en anticiperend rijden </li>
			</ol>
		</li>
		<li>Deelnemers moeten kunnen meeskaten zonder anderen in gevaar te brengen. </li>
	</ol>
	
	<p><strong>Artikel 7: Signalen</strong></p>
	<ol>
		<li> Deelnemers gebruiken tijdens de tocht een aantal algemeen aanvaarde signalen om gevaren aan te duiden. </li>
		<li> De voorste (of achterste) deelnemers duiden een gevaar, waarna ze tot achter (of voor) in de groep worden doorgegeven. </li>
		<li> Deelnemers worden geacht elkaar aan te spreken op het juist en consequent gebruik van waarschuwingssignalen. Het gaat hierbij in het bijzonder om de volgende signalen: 
			<ol type="a">
				<li>Als de groep moet remmen steekt iedereen zijn handen in de lucht. </li>
				<li>Als de groep afslaat bij een bocht in het parcours wordt met de vinger in de juiste richting gewezen. </li>
				<li>Als sprake is van obstakels en / of tegenliggers wordt de locatie hiervan aangeduid door de arm en hand naar achteren te houden aan de kant van het probleem (het gaat hier om de hele hand in tegenstelling tot de vinger bij het afslaan). Dit dient ruim van tevoren te gebeuren. </li>
				<li>De situatie kan verbaal benoemd worden (bijvoorbeeld Stoppen! Rechts! Links! Paaltje! Let op! Tegenligger! Fietser voor! Auto achter!).</li>
			</ol>
		</li>
	</ol>
	
	<p><strong>Artikel 8: Nat wegdek</strong></p>
	<ol>
		<li>Skate activiteiten gaan alleen door als het wegdek bij aanvang van de tocht droog is. </li>
		<li>Wanneer het tijdens een activiteit begint met regenen wordt de veiligste route terug gereden of wordt de activiteit (tijdelijk) gestaakt. </li>
	</ol>
	
	<p><strong>Artikel 9: Algemene gedragsregels tijdens een verenigingsactiviteit</strong></p>
	<ol>
		<li>Deelnemers worden geacht zich te houden aan de algemeen geldende verkeersregels. </li>
		<li>Deelnemers worden geacht waar mogelijk goed rechts te houden zodat zowel inhalend als tegenliggend verkeer geen last ondervindt. </li>
		<li>Deelnemers worden geacht vooraf en tijdens de activiteiten geen alcohol of drugs te gebruiken. </li>
		<li>Deelnemers worden geacht geen walkmans of vergelijkbare apparatuur te dragen, noch mobiel te bellen tijdens het skaten. </li>
		<li>Veiligheid is een gezamenlijke verantwoordelijkheid voor en door deelnemers; deelnemers worden geacht elkaar aan te spreken op gedrag dat de eigen of andermans veiligheid in gevaar brengt. </li>
		<li>De bestuursleden, of in afwezigheid van deze de overige deelnemers van een activiteit waar aspirant leden aanwezig zijn, dragen zorg voor een introductie in de Algemene Reglementen van de vereniging. </li>
		<li>Het bestuur draagt er zorg voor dat de leden op de hoogte zijn van hetgeen minimaal in deze introductie zal moeten worden verteld. </li>
	</ol>
	
	<p><strong>Artikel 10: Verzekering deelnemers</strong></p>
	<ol>
		<li>Van alle deelnemers wordt verwacht dat zij beschikken over een eigen WA- en ziektekostenverzekering. Een eigen ongevallenverzekering wordt aangeraden. </li>
		<li>Leden van de vereniging kunnen gebruik maken van de via de Skatebond Nederland afgesloten Univ&eacute;-ongevallen-, rechtshulp- en aansprakelijkheidsdekking. De dekking is beperkt tot de bepalingen in de verzekeringspolis en uitsluitend voor zover aan de bepalingen in de verzekeringspolis is voldaan. De meest recente versie van deze polisvoorwaarden is in te zien op de website van Skatebond Nederland (<a href="http://www.skatebond.nl/" target="_blank">www.skatebond.nl</a>). 
			<ol type="a">
				<li>&Eacute;&eacute;n van de specifiek hier te noemen uitsluitende bepalingen in de polisvoorwaarden is het niet dragen van een helm (Clausule 199).</li>
			</ol>
		</li>
	</ol>
	
	<p><strong>Artikel 11: Richtlijnen bij ongevallen</strong></p>
	<ol>
		<li>In geval er een valpartij of ongeluk plaatsvindt dan wordt van de niet-betrokken deelnemers verwacht dat ze, indien nodig: 
			<ol type="a">
				<li>Hulp verlenen</li>
				<li>Zonodig professionele hulp inschakelen </li>
				<li>Er zorg voor dragen dat de gewonde veilig bij een arts, bij een ziekenhuis of thuis komt</li>
			</ol>
		</li>
	</ol>
	
	<p><strong>Artikel 12: Uitsluiting deelname</strong></p>
	<ol>
		<li>Het bestuur of leden van het bestuur kunnen hun recht uitoefenen, zoals vastgelegd in artikel 7 van de statuten, om deelnemers uit te sluiten van verdere deelname aan activiteiten, indien zij van mening zijn dat in het betreffende geval de veiligheid van het individu of de groep in het geding komt. Het niet voldoen aan het algemeen reglement kan aanleiding vormen tot uitsluiting. </li>
	</ol>
	
	<p><strong>Artikel 13: Slotbepalingen</strong></p>
	<ol>
		<li>Iedere deelnemer en verenigingsorgaan heeft zich tijdens verenigingsactiviteiten te houden aan de bepalingen van dit reglement. </li>
		<li>Alle afwijkingen van dit reglement zijn op eigen risico en vallen dus buiten de verantwoordelijkheid van de vereniging. </li>
		<li>Na vaststelling van het reglement wordt de inhoud van het reglement zo spoedig mogelijk bekend gemaakt aan de leden. </li>
		<li>In gevallen waarin de statuten of het reglement niet voorzien, beslist het bestuur. </li>
	</ol>
	<p><strong>Opgemaakt en goedgekeurd door het Bestuur en de Algemene Ledenvergadering van Skateology,</strong></p>
	<p><strong>Leiden, 27 mei 2004</strong></p>-->
<%end sub

sub l_vereniging()%>
	<h1>Verenigingsinformatie</h1>

	<%set rs = adocon.execute("SELECT * FROM tekst WHERE tekst_id=3")
	do until rs.EOF

		displayText(rs("tekst"))

	rs.movenext : loop
	rs.close : set rs = nothing%>

	<!--<p>Skateology is een vereniging voor inline skaters in Leiden en omgeving. De vereniging is opgericht in 1997 en heeft bijna 60 leden. Skateology skate in ieder geval twee keer in de week. Op de <a href="default.asp?a=agenda&id=10">dinsdagavond tochten</a> rijden er meestal iets meer mensen mee dan op de <a href="default.asp?a=agenda&id=11">zondagtochten</a>. In de zomer rijden we op de dinsdagavond gemiddeld met 10 tot 20 skaters. In de winter zijn dit er 5 tot 10. </p>
	<p>Het is altijd weer een verrassing wie er deze keer weer op de Beestenmarkt verschijnen: het motto van Skateology is namelijk &quot;Vrijheid, blijheid en veiligheid&quot;. Altijd meeskaten is dus zeker geen verplichting; dit maakt dat leden gewoon komen als ze zin hebben in een gezellige skatetocht. Aanmelden is voor de meeste activiteiten dus ook niet nodig. </p>
	<p>Skateology is een vereniging, geen “nightskate” waar je onbeperkt vrijblijvend mee kan rijden. Skateology activiteiten zijn dus in principe voor de leden, maar op dinsdagavond zijn aspirantleden altijd welkom om een keer mee te rijden. Lees voor ervaringen van aspirantleden <a href="default.asp?a=ervaringen&id=1">het verhaal van Hester</a> of <a href="default.asp?a=ervaringen&id=3">de sfeerimpressie uit de Haagse Courant</a>. </p>
	<p>Spreekt Skateology je aan? Kom dan een keer langs op een <a href="default.asp?a=agenda&id=10">dinsdagavond</a>. Lees dan wel eerst de <a href="default.asp?a=spelregels">spelregels</a>. Lid worden? Ga dan naar het <a href="default.asp?a=inschrijven">inschrijfformulier</a>. De statuten zijn op te vragen via <a href="javascript:mailto('info')">Skateology</a>.</p>
	<p>Als je nog vragen hebt over Skateology, mail dan naar <a href="javascript:mailto('info')" target="_blank">Skateology</a>.</p>-->

	<h1>Bestuur</h1>
	<table cellpadding=2 cellspacing=0 width="70%">
	<%set rs = adocon.execute("SELECT * FROM functie LEFT JOIN lid ON functie.lid_id=lid.lid_id WHERE jaartal_id="&skatejaar&" AND bestuur=true ORDER BY functie DESC")
	do until rs.EOF%>
		<tr> 
			<td nowrap class="tekst"><b><%=rs("functie")%></b></td>
			<td nowrap class="tekst">
        <!--<a href="mailto:<%=rs("functie.email")%>"><%=rs("functie.email")%></a>-->
        <a href="javascript:mailto('<%=rs("functie")%>')">Verstuur e-mail</a>
      </td>
		</tr>
		<tr> 
			<td nowrap class="tekst"><%call naam(rs)%></td>
			<td nowrap class="tekst"><%
			If rs("functie") = "Voorzitter" then
				response.write(rs("telefoonnummer"))
			End if
			%></td>
		</tr>
	<%rs.movenext : loop
	rs.close : set rs = nothing%>
	</table>
<%end sub

sub l_inschrijven()%>
	<h2>Inschrijven voor Skateology</h2>
	<p>Om het inschrijfformulier zo toegankelijk mogelijk te maken is hij hier in twee formaten te vinden. De bovenste opent een nieuw venster met daarin een formulier dat u in kunt vullen en kunt printen. De onderste kunt u opslaan op de harde schijf en later uitprinten en invullen. </p>
	<p><A onclick="window.open('popup_inschrijfformulier.asp','Inschrijfformulier','width=650, height=530,directories=no,location=no, menubar=no, scrollbars=yes,status=no,toolbar=no,resizable=yes');return false" href="popup_inschrijfformulier.asp">Inschrijfformulier invullen en printen</a><br>
		<em>HTML formulier</em></p>
	<p><a href="/bestanden/inschrijfformulier_acrobat.pdf" target="_blank">Inschrijfformulier downloaden</a><br>
		<em>Adobe Acrobat </em></p><%
end sub

sub l_ervaringen()
	set rs = adocon.execute("SELECT * FROM ervaring WHERE ervaring_id="&request.querystring("id"))%>
	<h1><%=rs("titel")%></h1>
	
	<em><%=date2datum(rs("datum"), false) & " door "& rs("auteur")%></em>
	<%displayText(rs("tekst"))
	set rs = nothing
end sub

sub l_3House()%>
	<h2>Skatelessen aangeboden door 3House</h2>
	<p>Heb je volledige controle over je skates of heb je het gevoel dat je nog wel wat beter zou willen leren remmen? Dan biedt 3HOUSE jou de mogelijkheid om je veiliger en meer relaxed op je skates te voelen, maar ook voor de gevorderde skater is er nog altijd wat te schaven aan de techniek. Dus geef je op voor de beginners of de gevorderden cursus!</p>
	
	<p><strong>Wat kan je van de lessen voor beginners verwachten?</strong><br>
	Onder leiding van een instructeur met een rijkserkend diploma van de Skatebond Nederland word je geleerd om op een veilige en vooral ontspannen manier te skaten in het dagelijkse verkeer.</p>
	<p><strong>Wat kan je van de lessen voor gevorderden verwachten?</strong><br>
	Mocht je de smaak te pakken hebben na de lessen voor beginners of al een flink aantal skatetechnieken beheersen dan kan je starten met de lessen voor gevorderden. Verfijning van eerder geleerde technieken, nieuwe remmethodes en achteruitrijden zullen allemaal in het programma voorkomen.</p>
	<p><em>Neem contact op met 3HOUSE voor gedetailleerde informatie.<br>
	(<a href="javascript:mailto('info','3house')">E-mail</a> of  06-18089556)</em></p>
	
	<p><em>Al bekend bij 3HOUSE? Stuur dan een email naar ons adres met de ontbrekende gegevens.</em></p>
	<p>Alle gegevens op een rijtje:</p>
	<table border="1" cellpadding="2" cellspacing="0">
	<tr class="tekst">
		<td colspan="2"><p align="center"><strong>Beginners </strong></p></td>
	</tr>
	<tr class="tekst">
		<td width="84" valign="top"><p>Tijd </p></td>
		<td width="229" valign="top"><p>18:30 – 19:45 </p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Locatie </p></td>
		<td valign="top"><p>Boshuizerkade 83 (achter de Vijf Mei hal) </p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Duur </p></td>
		<td valign="top"><p><strong>5 </strong> lessen van <strong>75 </strong> minuten </p></td>
	</tr>
	<!--
	<tr class="tekst">
		<td valign="top"><p>Huur skates </p></td>
		<td valign="top"><p>&euro; 5,- (skates + bescherming) per les </p></td>
	</tr>-->
	<tr class="tekst">
		<td valign="top"><p>Kosten </p></td>
		<td valign="top"><p>&euro; 50,- </p></td>
	</tr>
	</table>
	<br>
	<table border="1" cellpadding="2" cellspacing="0">
	<tr class="tekst">
		<td colspan="2"><p align="center"><strong>Gevorderden </strong></p></td>
	</tr>
	<tr class="tekst">
		<td width="84" valign="top"><p>Tijd </p></td>
		<td width="240" valign="top"><p>20:00 – 21:15 </p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Locatie </p></td>
		<td valign="top"><p>Boshuizerkade 83 (achter de Vijf Mei hal) </p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Duur </p></td>
		<td valign="top"><p><strong>5 </strong> lessen van <strong>75 </strong> minuten </p></td>
	</tr>
	<!--
	<tr class="tekst">
		<td valign="top"><p>Huur skates </p></td>
		<td valign="top"><p>&euro; 5,- (skates + bescherming) per les </p></td>
	</tr>
	-->
	<tr class="tekst">
		<td valign="top"><p>Kosten </p></td>
		<td valign="top"><p>&euro; 50,- </p></td>
	</tr>
	</table>
	<p><em>Tot ziens bij de skatelessen!</em></p>
	<hr>
	<p><strong>What could you expect from the basic course?</strong><br>
	Under the supervision of a certified instructor (by the &ldquo;Skatebond Nederland&rdquo;) you will be tought the basic techniques to take part in the day to day traffic. Safety and fun will be priority.</p>
	<p><strong>What could you expect from the advanced course?</strong><br>
	If you got the right feeling in the basic course or you&rsquo;re already a bit more familiar with your skates it is possible to start the advanced course. Refinement of techniques, new braking methodes and riding backwards will be all part of the programm.</p>
	<p><em>Please contact 3HOUSE for detailed information.<br>
	(<a href="javascript:mailto('info','3house')">Email</a> or  06-18089556)</em></p>
	<p><em>Have you already given your data to 3HOUSE? Then please send an email to our address with the missing data.</em></p>
	<p>All the details:</p>
	<table cellspacing="0" cellpadding="2" border="1">
	<tr class="tekst">
		<td colspan="2"><p align="center"><strong>Basic </strong></p></td>
	</tr>
	<tr class="tekst">
		<td width="143" valign="top"><p>Date </p></td>
		<td width="297" valign="top"><p>Unknown</p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Time </p></td>
		<td valign="top"><p>18:30 – 19:45 </p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Location </p></td>
		<td valign="top"><p>Boshuizerkade 83 (behind the “Vijf Mei hal”) </p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Duration </p></td>
		<td valign="top"><p><strong>5 </strong> lessons of <strong>75 </strong> minutes </p></td>
	</tr>
	<!--<tr class="tekst">
		<td valign="top"><p>Rental cost of skates </p></td>
		<td valign="top"><p>&euro; 5,- (skates + protection(including helmet) per lesson </p></td>
	</tr>-->
	<tr class="tekst">
		<td valign="top"><p>Cost of the lessons </p></td>
		<td valign="top"><p>&euro; 50,- </p></td>
	</tr>
	</table>
	<br>
	<table border="1" cellpadding="2" cellspacing="0">
	<tr class="tekst">
		<td colspan="2"><p align="center"><strong>Advanced </strong></p></td>
	</tr>
	<tr class="tekst">
		<td width="143" valign="top"><p>Date </p></td>
		<td width="303" valign="top"><p>Unknown</p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Time </p></td>
		<td valign="top"><p>20:00 – 21:15 </p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Location </p></td>
		<td valign="top"><p>Boshuizerkade 83 (behind the “Vijf Mei hal”) </p></td>
	</tr>
	<tr class="tekst">
		<td valign="top"><p>Duration </p></td>
		<td valign="top"><p><strong>5 </strong> lessons of <strong>75 </strong> minutes </p></td>
	</tr>
	<!--
	<tr class="tekst">
		<td valign="top"><p>Rental costs of skates </p></td>
		<td valign="top"><p>&euro; 5,- (skates + protection(including helmet) per lesson </p></td>
	</tr>-->
	<tr class="tekst">
		<td valign="top"><p>Costs of the lessons </p></td>
		<td valign="top"><p>&euro; 50,- </p></td>
	</tr>
	</table>          
	<p><em>See you at the skate lessons!</em></p>
	<hr>
	<p><strong>Belangrijke mededeling van 3HOUSE</strong></p>
	<ul>
		<li>Het dragen van een helm, pols-, knie- en elleboogbeschermers is verplicht bij alle lessen van 3HOUSE. Een deelnemer zonder volledige bescherming wordt geweigerd uit de lessen. Een val is altijd mogelijk en kan ernstig letsel tot gevolg hebben.</li>
		<li> Minimale vereisten voor deelname aan de lessen: in balans kunnen staan op de skates en enigszins kunnen voortbewegen <br></li>
	</ul><%
end sub

sub l_skateweekend2007()
%>

	<h1>Skateweekend 2007</h1>
	<p>20 t/m 22 Mei 2005, Koudekerke, Zeeland. Zin om mee te gaan? Schrijf je hieronder in! Let wel, inschrijven is meteen betalingsverplichting (€60 p.p.).<br />
	Meer info over het skateweekend vind je in de <a href="http://www.skateology.nl/default.asp?a=agenda&id=12">agenda</a>.</p>
	
	<form method="post" action="default.asp?a=skateweekend_save">
	<table width="100%">
	<tr>
		<td valign="top">Naam</td>
		<td><input type="text" name="naam" size="40" maxlength="50"></td>
	</tr>
	<tr>
		<td valign="top">Aankomst:</td>
		<td><select name="aankomst">
				<option>Vrijdag middag</option>
				<option selected>Vrijdag avond</option>
				<option>Zaterdag ochtend</option>
				<option>Anders (vul in bij opmerkingen)</option>
			</select></td>
	</tr>
	<tr>
		<td valign="top">Vertrek:</td>
		<td><select name="vertrek">
				<option>Zaterdag avond</option>
				<option>Zondag ochtend</option>
				<option>Zondag middag</option>
				<option selected>Zondag avond</option>
			</select></td>
	</tr>
	<tr>
		<td valign="top">Vervoer:</td>
		<td><select name="vervoer">
				<option>Auto / eigen vervoer</option>
				<option>Openbaar vervoer</option>
				<option>Hopelijk meerijden, anders OV</option>
			</select></td>
	</tr>
	<tr>
		<td valign="top">Aantal vrije plaatsen</td>
		<td valign="top"><select name="vrijeplaatsen">
			<%for i = 0 to 5%>
				<option value="<%=i%>"><%=i%></option>
			<%next%>
			</select> (alleen indien auto)</td>
	</tr>
	<tr>
		<td valign="top">Eetwensen /<br>
			Allergien /<br>
			Vegetarisch</td>
		<td valign="top"><textarea cols=30 rows=3 name="eetwensen"></textarea></td>
	</tr>
	<tr>
		<td valign="top">Contactpersoon thuis:</td>
		<td valign="top"><textarea cols=30 rows=2 name="contactthuis"></textarea></td>
	</tr>
	<tr>
		<td valign="top">Opmerkingen:</td>
		<td valign="top"><textarea cols=30 rows=3 name="overig"></textarea></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>

<%
end sub
sub l_skateweekend()
	%>
	<h1>Skateweekend 2005</h1>
	<p>20 t/m 22 Mei 2005, Koudekerke, Zeeland. Zin om mee te gaan? Schrijf je hieronder in! Let wel, inschrijven is meteen betalingsverplichting (€60 p.p.).<br />
	Meer info over het skateweekend vind je in de <a href="http://www.skateology.nl/default.asp?a=agenda&id=12">agenda</a>.</p>
	
	<form method="post" action="default.asp?a=skateweekend_save">
	<table width="100%">
	<tr>
		<td valign="top">Naam</td>
		<td><input type="text" name="naam" size="40" maxlength="50"></td>
	</tr>
	<tr>
		<td valign="top">Aankomst:</td>
		<td><select name="aankomst">
				<option>Vrijdag middag</option>
				<option selected>Vrijdag avond</option>
				<option>Zaterdag ochtend</option>
				<option>Anders (vul in bij opmerkingen)</option>
			</select></td>
	</tr>
	<tr>
		<td valign="top">Vertrek:</td>
		<td><select name="vertrek">
				<option>Zaterdag avond</option>
				<option>Zondag ochtend</option>
				<option>Zondag middag</option>
				<option selected>Zondag avond</option>
			</select></td>
	</tr>
	<tr>
		<td valign="top">Vervoer:</td>
		<td><select name="vervoer">
				<option>Auto / eigen vervoer</option>
				<option>Openbaar vervoer</option>
				<option>Hopelijk meerijden, anders OV</option>
			</select></td>
	</tr>
	<tr>
		<td valign="top">Aantal vrije plaatsen</td>
		<td valign="top"><select name="vrijeplaatsen">
			<%for i = 0 to 5%>
				<option value="<%=i%>"><%=i%></option>
			<%next%>
			</select> (alleen indien auto)</td>
	</tr>
	<tr>
		<td valign="top">Eetwensen /<br>
			Allergien /<br>
			Vegetarisch</td>
		<td valign="top"><textarea cols=30 rows=3 name="eetwensen"></textarea></td>
	</tr>
	<tr>
		<td valign="top">Contactpersoon thuis:</td>
		<td valign="top"><textarea cols=30 rows=2 name="contactthuis"></textarea></td>
	</tr>
	<tr>
		<td valign="top">Opmerkingen:</td>
		<td valign="top"><textarea cols=30 rows=3 name="overig"></textarea></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="Opslaan"></td>
	</tr>
	</table>
	</form>
	<%
end sub

sub l_skateweekend_save()
	sql = "INSERT INTO skateweekend(naam, aankomst, vertrek, vervoer, vrijeplek, eetwensen, contactpersoon, overigen, skatejaar) VALUES("
	sql = sql & "'" & replace(request.form("naam"),"'", "''") & "'"
	sql = sql & ",'" & replace(request.form("aankomst"),"'", "''") & "'"
	sql = sql & ",'" & replace(request.form("vertrek"),"'", "''") & "'"
	sql = sql & ",'" & replace(request.form("vervoer"),"'", "''") & "'"
	sql = sql & "," & request.form("vrijeplaatsen")
	sql = sql & ",'" & replace(request.form("eetwensen"),"'", "''") & "'"
	sql = sql & ",'" & replace(request.form("contactthuis"),"'", "''") & "'"
	sql = sql & ",'" & replace(request.form("overig"),"'", "''") & "'"
	sql = sql & "," & skatejaar & ")"
	adocon.execute(sql)
	%>
	<h1>Skateweekend 2005</h1>
	<p>Je staat met de volgende gegevens ingeschreven voor het Skateweekend 2005:</p>
	<table width="90%" align="center">
	<%
	for each i in request.form()
		displayText("<tr><td>" & i & "</td><td>" & request.form(i) & "</td></tr>")
	next
	%>
	</table>
	<p>Mocht er iets niet kloppen, <a href="javascript:mailto('penningmeester')">stuur een e-mail</a>.</p>
	<h2>Betalen</h2>
	<p>Kloppen de gegevens? Maak dan, indien je dit nog niet hebt gedaan, &euro; 60,- over naar rekening 9151332 t.n.v. Skateology onder vermelding van je naam en "skateweekend". Zodra je dit gedaan hebt, sta je volledig ingeschreven voor het Skateweekend.</p>
	<h2>Verdere gang van zaken</h2>
	<p>Een of twee weken voor het weekend krijgt iedereen een informatiepakketje toegestuurd. Hierin staat onderandere wat je mee dient te nemen, wie er vrije plekken in de auto hebben, en een wat gedetailleerdere planning voor het weekend.</p>
	<p>Tot skates!</p>
	<%
end sub
%>