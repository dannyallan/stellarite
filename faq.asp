<!--#include file="_inc\functions.asp" -->
<%
	strTitle = getIDS("IDS_FAQs")
	strModName = strTitle

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

  <ul class="dFont">
	<li><a href="#1">Can I navigate the application using keystrokes only?</a></li>
	<li><a href="#2">Why do my drop down menus not expand?</a></li>
	<li><a href="#3">Why do my dates and times appear to be incorrect?</a></li>
	<li><a href="#4">What are the minimum browser requirements?</a></li>
	<li><a href="#5">Where do I go for more help?</a></li>
  </ul>

  <hr />

  <p class="dFont"><a name="1" class="hFont">Can I navigate the application using keystrokes only?</a>
  <br />The Stellarite CRM implements many features for increased accessability.  For example, all
  form elements should have an associated label, and all auditory and visual elements should have
  equivalent alternatives.  You may also use keystrokes for many common functions:

  <br /><br />
  <table border=0 cellspacing=0 width="50%">
	<tr class="dRow1">
	  <td class="dFont">Alt - (comma)</td>
	  <td class="dFont">Previous Record</td>
	</tr>
	<tr class="dRow2">
	  <td class="dFont">Alt - (period)</td>
	  <td class="dFont">Next Record</td>
	</tr>
	<tr class="dRow1">
	  <td class="dFont">Alt - N</td>
	  <td class="dFont">New Record</td>
	</tr>
	<tr class="dRow2">
	  <td class="dFont">Alt - O</td>
	  <td class="dFont">Open / Edit Record</td>
	</tr>
	<tr class="dRow1">
	  <td class="dFont">Alt - S</td>
	  <td class="dFont">Save Record</td>
	</tr>
	<tr class="dRow2">
	  <td class="dFont">Alt - D</td>
	  <td class="dFont">Delete Record</td>
	</tr>
	<tr class="dRow1">
	  <td class="dFont">Alt - F</td>
	  <td class="dFont">Find Record</td>
	</tr>
	<tr class="dRow2">
	  <td class="dFont">Alt - P</td>
	  <td class="dFont">Print Details</td>
	</tr>
	<tr class="dRow1">
	  <td class="dFont">Alt - X</td>
	  <td class="dFont">Cancel / Exit</td>
	</tr>
  </table>
  </p>


  <p class="dFont"><a name="2" class="hFont">Why do my drop down menus not expand?</a>
  <br />This is a known problem with Netscape browsers.  Although the Stellarite system attempts
  to be compatible with both Netscape and Internet Explorer, there are some instances when this
  is not possible.  Netscape does not allow expanding Select fields if they are enclosed in
  a scrollable DIV layer.  You are able to use the arrow key to select different options.</p>

  <p class="dFont"><a name="3" class="hFont">Why do my dates and times appear to be incorrect?</a>
  <br />The Stellarite CRM derives all of its dates and times based on the timezone of the PC
  accessing the system.  If the timezone on your PC is incorrect, all of the dates and times
  will appear to be off by the error factor.  We calculate the current time as approximately
  <% =showDate(1,Now()) %> on your PC and exactly <% =Now %> on the Stellarite CRM Server.</p>

  <p class="dFont"><a name="4" class="hFont">What are the minimum browser requirements?</a>
  <br />For more information on the minimum browser requirements, please visit the
  <a href="http://www.stellarite.com/require.asp">Stellarite Website.</p>

  <p class="dFont"><a name="5" class="hFont">Where do I go for more help?</a>
  <br />For more information, please email the
  <a href="mailto:<% =Application("av_EmailFrom") %>">Stellarite Administrator</a> or visit the
  <a href="http://www.stellarite.com/default.asp">Stellarite Website</a>.</p>

</div>
<%
	Call DisplayFooter(1)
%>