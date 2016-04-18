<%@EnableSessionState=False%>
<% Response.Expires = -1 %>

<HTML>
<HEAD>
<TITLE>Enviando fichero</TITLE>
<style type='text/css'>td {font-family:verdana; font-size: 9pt }</style>
</HEAD>

<% If Request("b") = "IE" Then %> <!-- Internet Explorer -->
<BODY BGCOLOR="#C0C0B0">
<IFRAME src="arbol_bar.asp?PID=<%= Request.QueryString("PID") & "&to=" & Request.QueryString("to") %>" 
title="Progreso subida" scrolling=no frameborder=0 framespacing=10 width=369 height=115></IFRAME>
<TABLE BORDER="0" WIDTH="100%" cellpadding="2" cellspacing="0">
  <TR><TD ALIGN="center">
     Para detener el envío pulsa <B>STOP</B> en el navegador.
  </TD></TR>
</TABLE>
</BODY>

<%Else%> <!-- Netscape Navigator etc ... -->

<FRAMESET ROWS="65%, 35%" COLS="100%" border="0" framespacing="0" frameborder="NO">
<FRAME SRC="arbol_bar.asp?PID=<%= Request("PID") & "&to=" & Request("to") %>" noresize scrolling="NO" frameborder="NO" name="sp_body">
<FRAME SRC="note.htm" noresize scrolling="NO" frameborder="NO" name="sp_note">
</FRAMESET>

<%End If%>

</HTML>
