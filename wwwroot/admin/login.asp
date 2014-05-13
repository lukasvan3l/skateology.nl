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

<%
if request.querystring("a") = "logout" then
	
	session("username") = ""
	session("niveau") = ""
	session.abandon
	response.redirect("login.asp")
	
elseif request.querystring("a") = "login" then
	
	set rs = adocon.execute("SELECT * FROM functie WHERE jaartal_id=" & skatejaar & " AND username='" & changequotes(request.form("username")) & "'")
	if rs.EOF then
		set rs2 = adocon.execute("SELECT * FROM functie WHERE username='" & changequotes(request.form("username")) & "' AND wachtwoord='" & trim(request.form("password")) & "'")
		if rs2.EOF = true then
			response.redirect("login.asp?a=f1")
		else
			response.redirect("login.asp?a=f3")
		end if
		set rs2 = nothing
	else
		if trim(request.form("password")) <> trim(rs("wachtwoord")) then
			response.redirect("login.asp?a=f2")
		end if
	end if
	
	Response.Cookies("skateology.nl")("admin") = rs("username")
	Response.Cookies("skateology.nl").Expires = Date + 365
	
	session("username") = rs("username")
	session("functie_id") = rs("functie_id")
	session("lid_id") = rs("lid_id")
	session.Timeout=60
	rs.close : set rs = nothing
	
	response.redirect("default.asp")
	
Else

'   set rs2 = adocon.execute("SELECT * FROM functie WHERE jaartal_id=" & skatejaar & " AND username='tjerk'")
'	If rs2.EOF = false then
'		response.write("wachtwoord is: " + rs2("wachtwoord"))
'	End if
'	rs2.close : Set rs2 = nothing

End if
%>

<head>
	<title>Skateology Admin</title>
	<link rel="STYLESHEET" type="text/css" href="../css/admin.css">
	<META NAME="ROBOTS" CONTENT="NOINDEX">
	<META NAME="ROBOTS" CONTENT="NOFOLLOW">
	<script src="../inc/clientsniff.js" type="text/javascript"></script>
	<script src="../inc/javascript.js" type="text/javascript"></script>
	<script language="JavaScript" src="include/yusasp_ace.js"></script>
	<script language="JavaScript" src="include/yusasp_color.js"></script>
	<script>
		col=255;
		function fadein() { 
			document.getElementById("fadeintext").style.color="rgb(" + col + "," + col + "," + col + ")"; 
			col-=5; 
			if(col>0) setTimeout('fadein()', 40);
		}
	</script>
</head>
<body onload="fadein();">

<p>&nbsp;</p>
<H1 align="center">Skateology - Admin</H1>
<H2 align="center">&nbsp;
<%if request.querystring("a") = "f1" then%>
	<span id="fadeintext">De combinatie naam en wachtwoord komt niet overeen.</span>
<%elseif request.querystring("a") = "f3" then%>
	<span id="fadeintext">Sorry, oud-bestuur heeft hier geen toegang meer.</span>
<%elseif request.querystring("a") = "f2" then%>
	<span id="fadeintext">Het wachtwoord klopt niet.</span>
<%else%>
	<span id="fadeintext">&nbsp;</span>
<%end if%>
</H2>
<form action="login.asp?a=login" method="post" name="loginform">
<table width="200" align="center">
	<tr>
		<td>Naam:</td>
		<td><input type="text" name="username" value="<%=request.Cookies("skateology.nl")("admin")%>" size="20" maxlength="20"></td>
	</tr>
	<tr>
		<td>Wachtwoord:</td>
		<td><input type="password" name="password" value="" size="20" maxlength="20"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="log in"></td>
	</tr>
</table>
</form>

<%if request.Cookies("skateology.nl")("admin") = "" then%>
	<!-- dit script moet onder het formulier zelf geplaats worden -->
	<script type="text/javascript" language="javascript">
		document.forms['loginform'].elements['username'].focus();
	</script>
<%else%>
	<!-- dit script moet onder het formulier zelf geplaats worden -->
	<script type="text/javascript" language="javascript">
		document.forms['loginform'].elements['password'].focus();
	</script>
<%end if%>


</body></html>
<!-- #include file="inc/functionsclose.asp" -->