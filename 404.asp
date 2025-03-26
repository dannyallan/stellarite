<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html>

<head>
<title><% =Application("IDS_ErrorUnspecified") %></title>
<link href="<% =Application("av_CRMDir") %>common/css.asp" rel="stylesheet" type="text/css" />
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" scroll="no">

<table border=0 cellspacing=0 cellpadding=0 width="100%" height="90%">
  <tr><td align=center valign=middle>

	<table border=0 cellspacing=10 cellpadding=20 width=300 height=150>
	  <tr class="hRow">
		<td class="dFont">
		<span class="hFont"><% =Application("IDS_ErrorUnspecified") %></span><br /><br />
		<ul>
			<li><a href="Javascript:history.back();"><% =Application("IDS_BackOne") %></a></li><br />
			<li><a href="Javascript:history.go(-2);"><% =Application("IDS_BackTwo") %></a></li><br />
			<li><a href="Javascript:window.close();"><% =Application("IDS_CloseWindow") %></a></li><br />
		</ul>
		</td>
	  </tr>
	</table>

  </td></tr>
</table>

</body>
</html>