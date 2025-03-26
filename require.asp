<%
	Sub showProblem(fProb,fText)
		If Request.QueryString("prob") = fProb Then
			Response.Write(vbTab & "<li><b><font color=red>" & fText & "</font></b></li>" & vbCrLf)
		Else
			Response.Write(vbTab & "<li>" & fText & "</li>" & vbCrLf)
		End If
	End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="en">

<head>
<title>Requirements</title>
<link href="common/css.asp" rel="stylesheet" type="text/css" />
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table border=0 cellspacing=10 width="100%">
  <tr class="hRow">
	<th class="hFont"><a name=browser></a>Browser requirements</th>
  </tr>
  <tr>
	<td class="dFont">The Stellarite web interface has four requirements:
	<ul>
<%
	Call showProblem("br","Microsoft Internet Explorer 5.01+")
	Call showProblem("br","Netscape 6.0+")
	Call showProblem("sr","Minimum 800 x 600 resolution")
	Call showProblem("js","JavaScript Enabled")
	Call showProblem("ck","Cookies Enabled")
%>
	</ul>

	<p>To find out the number of the browser version you are currently running, please do the following:
	<ul>
	 <li><strong>Netscape Navigator</strong> - from the Help menu, select <strong>About Communicator</strong>.</li>
	 <li><strong>Microsoft Internet Explorer</strong> - from the Help menu, select <strong>About Internet Explorer</strong>.</li>
	</ul>
	If necessary, you might want to visit the <a href="http://www.browsers.com" target="_blank">CNET Browsers</a> website to download the
	appropriate version of one of these browsers.</p>
	</td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr class="hRow">
	<th class="hFont"><a name=enabling></a><strong>&nbsp;Enabling JavaScript and Cookies</strong></th>
  </tr>
  <tr>
	<td class="dFont">
	<strong>Microsoft Internet Explorer 5.0+</strong><br />
	<ol>
	 <li>From the <strong>Tools</strong> menu, select <strong>Internet Options</strong>.</li>
	 <li>Click on the <strong>Security</strong> tab.</li>
	 <li>Select <strong>Internet</strong> as the &quot;Web content zone&quot;.</li>
	 <li>If you have never customized your security settings for this zone (I.E., <strong>Custom</strong> is
	 not selected for the <strong>Security Level</strong>):
	  <ul type=square>
	   <li>Select (by moving the slider) at least <strong>Medium</strong> for the security level, in order
	   to fully utilize all of the features of Stellarite.</li>
	   <li>Click <strong>OK</strong> in the &quot;Internet Options&quot; window, and you're done!</li>
	  </ul>
	 <li>Otherwise:
	  <ul type=square>
	   <li>Click the <strong>Custom Level...</strong> button to display the &quot;Security Settings&quot;
	   window.</li>
	   <li>Select either the <strong>Enable</strong> or the <strong>Prompt</strong> button for the
	   <strong>Allow cookies that are stored on your computer</strong> option.</li>
	   <li>Select either the <strong>Enable</strong> or the <strong>Prompt</strong> button for the
	   <strong>Allow per-session cookies (not stored)</strong> option.</li>
	   <li>Select the <strong>Enable</strong> option for the <strong>Active Scripting</strong> option.</li>
	   <li>Select either the <strong>Enable</strong> or the <strong>Prompt</strong> button for the
	   <strong>Scripting of Java applets</strong> option.</li>
	   <li>Click <strong>OK</strong> in the &quot;Security Settings&quot; window to return to the
	   &quot;Internet Options&quot; window.</li>
	   <li>Click <strong>OK</strong> in the &quot;Internet Options&quot; window.</li>
	  </ul>
	 </li>
	</ol>
	<br />
	<strong>Netscape Navigator 6.0 and above</strong><br />
	<ol>
	 <li>From the <strong>Edit</strong> menu, select <strong>Preferences</strong>.</li>
	 <li>Click on the <strong>Advanced</strong> category on the left side of the <strong>Preferences</strong> window.</li>
	 <li>Check the <strong>Enable Java</strong> checkbox.</li>
	 <li>Check the <strong>Enable JavaScript</strong> checkbox.</li>
	 <li>In the <strong>Cookies</strong> section, select one of the following buttons:
	  <ul type=square><strong>
	   <li>Accept all cookies</li>
	   <li>Accept only cookies that get sent back to the originating server</li></strong>
	  </ul>
	 </li>
	 <li>Click <strong>OK</strong>.</li>
	</ol>
	</td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr class="hRow">
	<th class="hFont"><a name=loadProbs></a><strong>&nbsp;Problems loading Stellarite web pages</strong></th>
  </tr>
  <tr>
	<td class="dFont">
	<p>The cache is the location on your computer's hard disk where web pages and files (such as graphics) are
	stored as you view them. This speeds up the display of pages you frequently visit or have already seen,
	because your browser can open them from your hard disk instead of from the web.</p>

	<p>If you are having a problem getting a Stellarite page to load, it may be because your cache contains an
	old version of the page and is trying to load that instead of the newer version. By clearing the cache
	you are forcing your browser to go to the web and download a new copy of your page.</p>

	<p>To clear your disk cache, follow the instructions below:</p>

	<strong>Microsoft Internet Explorer 5.0+</strong><br />
	<ol>
	  <li>From the <strong>Tools</strong> menu, select <strong>Internet Options</strong>.</li>
	  <li>Click the <strong>General</strong> tab.</li>
	  <li>Click <strong>Delete Files...</strong> and click <strong>OK</strong> on the
	  &quot;Delete Files&quot; popup window.</li>
	  <li>Click <strong>OK</strong> on the &quot;Internet Options&quot; window.</li>
	</ol>

	<strong>Netscape Navigator 6.0 and above</strong><br />
	<ol>
	  <li>From the <strong>Edit</strong> menu, select <strong>Preferences</strong>.</li>
	  <li>Click the <strong>Advanced</strong> category on the left side of the window to expand it, and select <strong>Cache</strong> below
	  it.</li>
	  <li>Click <strong>Clear Disk Cache</strong>.</li>
	  <li>Click <strong>OK</strong>.</li>
	</ol>
	</td>
  </tr>
</table>

</body>
</html>