<%@LANGUAGE = VBScript%>
<% 
Option Explicit
Response.Buffer = TRUE
Response.Expires = 0
session.lcid = 1043
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<!-- #include file="inc/dandate.inc" -->
<!-- #include file="inc/adovbs.inc" -->
<!-- #include file="inc/functions.asp" -->

<%checkUser()%>

<head>
	<title>Skateology Admin - Ingelogd als <%=session("username")%></title>
	<link rel="STYLESHEET" type="text/css" href="../css/admin.css" />
	<link rel="STYLESHEET" type="text/css" href="../css/admin_print.css" media="print" />
	<META NAME="ROBOTS" CONTENT="NOINDEX">
	<META NAME="ROBOTS" CONTENT="NOFOLLOW">
	<script src="../inc/clientsniff.js" type="text/javascript"></script>
	<script src="../inc/javascript.js" type="text/javascript"></script>
	<script language="JavaScript" src="include/yusasp_ace.js"></script>
	<script language="JavaScript" src="include/yusasp_color.js"></script>
	<script>
	function SubmitForm()
		{
		if(obj1.displayMode == "HTML")
			{
			alert("Zet a.u.b. de HTML codering uit voor het opslaan.")
			return ;
			}
		Form.txtContent.value = obj1.getContentBody()
		Form.submit()
		}
	function LoadContent()
		{ obj1.putContent(idTextarea.value) }
	</script>
	<%if request.querystring("alert") = "opgeslagen" then%>
		<script>alert('De bewerking is opgeslagen');</script>
	<%elseif request.querystring("alert") = "verwijderd" then%>
		<script>alert('Het item is verwijderd');</script>
	<%end if%>
</head>
<body onload="preloadImages(); <%
	if (request.querystring("m") = "tekst" AND isnumeric(request.querystring("a")) AND request.querystring("a") <> "") or (request.querystring("ace") = "true") then
		%>LoadContent();<%
	end if
	%>">

<!-- #include file="menu.inc" -->

<table border="0" cellpadding="0" cellspacing="0" width="805" align="center" height="98%">
<tr class="adminheader">
	<td background="images/top_fill.gif" align="right" colspan="1">
		<a href="login.asp?a=logout"><%=session("username")%> Uitloggen</a>
		</td>
	<td background="images/top_fill.gif" align="right" colspan="1">
		<img src="images/top_logo.gif" width="325" height="69" alt="Skateology Admin"></td>
</tr>

<tr>
	<td width="116" height="100%">
		<table width="100%" height="100%" class="adminnavigation">
		<tr>
			<td valign="top">
			<br>
			<A HREF="default.asp?m=&a="
					ONMOUSEOVER="clearTimeout(timeOn); 
						hideAllLayers(); 
						changeImages('menu_1','images/left_1home_over.gif');
						clearTimeout(timeOn);"
					ONMOUSEOUT="btnTimer();
						changeImages('menu_1','images/left_1home.gif');">
			<img name="menu_1" src="images/left_1home.gif" width="116" height="26" alt=""></a><br>
			<A HREF="default.asp?m=lid&a=nieuw"
					ONMOUSEOVER="clearTimeout(timeOn); 
						hideAllLayers(); 
						showLayer('menu_2layer'); 
						changeImages('menu_2','images/left_2nieuw_over.gif');
						clearTimeout(timeOn);"
					ONMOUSEOUT="btnTimer();
						changeImages('menu_2','images/left_2nieuw.gif');">
			<img name="menu_2" src="images/left_2nieuw.gif" width="116" height="23" alt=""></a><br>
			<A HREF="default.asp?m=ledenlijst&a="
					ONMOUSEOVER="clearTimeout(timeOn); 
						hideAllLayers(); 
						showLayer('menu_3layer'); 
						changeImages('menu_3','images/left_3overzicht_over.gif');
						clearTimeout(timeOn);"
					ONMOUSEOUT="btnTimer();
						changeImages('menu_3','images/left_3overzicht.gif');">
			<img name="menu_3" src="images/left_3overzicht.gif" width="116" height="23" alt=""></a><br>
			<A HREF="default.asp?m=ledenlijst&a="
					ONMOUSEOVER="clearTimeout(timeOn); 
						hideAllLayers(); 
						showLayer('menu_4layer'); 
						changeImages('menu_4','images/left_4ledenlijst_over.gif');
						clearTimeout(timeOn);"
					ONMOUSEOUT="btnTimer();
						changeImages('menu_4','images/left_4ledenlijst.gif');">
			<img name="menu_4" src="images/left_4ledenlijst.gif" width="116" height="25" alt=""></a><br>
			<A HREF="default.asp?m=tekst&a="
					ONMOUSEOVER="clearTimeout(timeOn); 
						hideAllLayers(); 
						showLayer('menu_5layer'); 
						changeImages('menu_5','images/left_5teksten_over.gif');
						clearTimeout(timeOn);"
					ONMOUSEOUT="btnTimer();
						changeImages('menu_5','images/left_5teksten.gif');">
			<img name="menu_5" src="images/left_5teksten.gif" width="116" height="24" alt=""></a><br>
			<A HREF="default.asp?m=email&a="
					ONMOUSEOVER="clearTimeout(timeOn); 
						hideAllLayers(); 
						showLayer('menu_6layer'); 
						changeImages('menu_6','images/left_6email_over.gif');
						clearTimeout(timeOn);"
					ONMOUSEOUT="btnTimer();
						changeImages('menu_6','images/left_6email.gif');">
			<img name="menu_6" src="images/left_6email.gif" width="116" height="23" alt=""></a><br>
			</td>
		</tr>
		<tr>
			<td valign="bottom">
			<img src="images/left_logo.gif" width="116" height="189" alt="">
			</td>
		</tr>
		</table>
		</td>
	<td align="left" valign="top" width="689">
	<br>
	<table border=0 cellspacing=0 cellpadding=0 width="100%">
	<tr>
		<td width="10">&nbsp;</td>
		<td>
		<%'CONTENT%>
		
		<%select case request.querystring("m")%>
		<%case "lid"%>
			<!-- #include file="mod/lid.inc" -->
		<%case "nieuwsbrief"%>
			<!-- #include file="mod/nieuwsbrief.inc" -->
		<%case "agenda"%>
			<!-- #include file="mod/agenda.inc" -->
		<%case "ervaring"%>
			<!-- #include file="mod/ervaring.inc" -->
		<%case "fotoalbum"%>
			<!-- #include file="mod/fotoalbum.inc" -->
		<%case "kopfoto"%>
			<!-- #include file="mod/kopfoto.inc" -->
		<%case "bestand"%>
			<!-- #include file="mod/bestand.inc" -->
		<%case "bestuursjaar"%>
			<!-- #include file="mod/bestuursjaar.inc" -->
		<%case "tekst"%>
			<!-- #include file="mod/tekst.inc" -->
		<%case "email"%>
			<!-- #include file="mod/email.inc" -->
		<%case "ledenlijst"%>
			<!-- #include file="mod/ledenlijst.inc" -->
		<%case "referrer"%>
			<!-- #include file="mod/referrer.inc" -->
		<%end select%>
				
		<%'CONTENT%>
		</td>
	</tr>
	</table>
	<br>
	</td>
</tr>

<tr class="adminfooter">
	<td background="images/footer.gif" height="15" colspan="2">&nbsp;</td>
</tr>
</table>

</body></html>
<!-- #include file="inc/functionsclose.asp" -->