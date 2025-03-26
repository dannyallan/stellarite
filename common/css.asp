<%
	Response.ContentType = "text/css"
%>

body		{background-color: #FFFFFF;
			color: black;
			font-family: Verdana, Arial, Helvetica;
			font-size: 70%;
			margin: 0px;
			overflow: hidden;}

form		{margin-top: 0; margin-bottom: 0;}

p,ul,ol		{margin-top: 2; margin-bottom: 8px;}

a:link      {text-decoration: <% =Application("av_LinkStyle") %>;}
a:visited   {text-decoration: <% =Application("av_VisitedStyle") %>;}
a:hover     {text-decoration: <% =Application("av_HoverStyle") %>;}

.lIndent	{padding-left:10px;}
.rIndent	{padding-right:10px;}

div.hr      {display: block;
			height: 12px;}

.dFont,.wFont,.dText,.dNum,.dCurrency,.eTab,.dTab,.button,
.oText,.oEmail,.oPhone,.oRGB,.oDate,.oMemo,.oLink,.oNum,.oBool,.oByte,.oInt,.oLong,.oCurrency,
.mText,.mEmail,.mPhone,.mRGB,.mDate,.mMemo,.mLink,.mNum,.mBool,.mByte,.mInt,.mLong,.mCurrency
			{font-family: <% =Application("av_DefaultFont") %>;
			font-size: <% =Application("av_DefaultSize") %>;
			color: <% =Application("av_DefaultColor") %>;}

.tFont      {font-family: <% =Application("av_TitleFont") %>;
			font-size: <% =Application("av_TitleSize") %>;
			font-weight: <% =Application("av_TitleWeight") %>;
			color: <% =Application("av_TitleColor") %>;
			padding-left: 5px;
			padding-right: 5px;}

.hFont      {font-family: <% =Application("av_HeaderFont") %>;
			font-size: <% =Application("av_HeaderSize") %>;
			font-weight: <% =Application("av_HeaderWeight") %>;
			color: <% =Application("av_HeaderColor") %>;
			text-align: left;}

.bFont     {font-family: <% =Application("av_LabelFont") %>;
			font-size: <% =Application("av_LabelSize") %>;
			font-weight: <% =Application("av_LabelWeight") %>;
			color: <% =Application("av_LabelColor") %>;
			padding-right: 10px;}

.pFont      {font-family: <% =Application("av_PrintFont") %>;
			font-size: <% =Application("av_PrintSize") %>;}

.wFont      {font-weight: bold;
			color: red;}

.oText,.oEmail,.oPhone,.oRGB,.oDate,.oMemo,.oLink,.oNum,.oBool,.oByte,.oInt,.oLong,.oCurrency
			{border: thin groove;
			padding-left: 3px;
			background-color: <% =Application("av_FormEnabled") %>;}

.mText,.mEmail,.mPhone,.mRGB,.mDate,.mMemo,.mLink,.mNum,.mBool,.mByte,.mInt,.mLong,.mCurrency
			{border: thin groove;
			padding-left: 3px;
			background-color: <% =Application("av_FormMandatory") %>;}

.dText,.dNum,.dCurrency
			{border: thin groove;
			padding-left: 3px;
			background-color: <% =Application("av_FormDisabled") %>;}

.hRow       {background-color: <% =Application("av_MinorColor") %>;
			padding:2px;}

.hScr		{position: relative;
			top: expression(this.offsetParent.scrollTop-2);
			left: -1;}

.dRow1      {background-color: #FFFFFF;}
.dRow2      {background-color: <% =Application("av_DefaultAltBG") %>;}
.dRow3      {background-color: <% =Application("av_DefaultHiBG") %>;}

.eTab       {background-color: <% =Application("av_MajorColor") %>;
			color: white;
			cursor: default;}

.dTab       {background-color: <% =Application("av_MinorColor") %>;
			color: black;
			cursor: pointer;}

input.button     {margin-top: 5px;}

div.dvMod, div.dvBorder, div.dvNoBorder,
div.dvHeader, div.dvFooter, iframe.iBorder
			{-moz-box-sizing: border-box;
			box-sizing: border-box;}

div.dvBorder, div.dvNoBorder, div.dvHeader, div.dvFooter, div.dvMod
			{width: 100%;}

div.dvMod   {padding-left: 10px;
			padding-right: 10px;
			padding-top: 10px;}

div.dvNoBorder, div.dvHeader, div.dvFooter, div.dvRightMenu
				{overflow: auto;
			padding: 0px;}

div.dvBorder    {overflow: auto;
			padding: 10px;}

div.dvRightMenu	{margin-bottom: 20px;}

iframe.iBorder    {border: 2px solid <% =Application("av_MajorColor") %>;
			margin-left: 10px;
			margin-right: 10px;
			width:100%;}

div.dvHeader	{background-color: <% =Application("av_MinorColor") %>;
			margin: 10px;
			padding: 5px;}

div.dvFooter    {background-color: <% =Application("av_MinorColor") %>;
			margin: 10px;
			padding: 5px;
			text-align: right;}


div.menuBar, div.menuBar a.menuButton, div.menu, div.menu a.menuItem {
			font-family: "MS Sans Serif", Arial, sans-serif;
			font-size: 8pt;
			font-style: normal;
			font-weight: normal;
			color: #FFFFFF;
}

div.menuBar {
			background-color: <% =Application("av_MajorColor") %>;
			border-style: solid;
			border-width: 0px;
			border-top-color: <% =Application("av_MajorColorLight") %>;
			border-top-width: 1px;
			border-bottom-color: <% =Application("av_MajorColorDark") %>;
			border-bottom-width: 1px;
			padding: 5px 2px 5px 2px;
			text-align: left;
}

div.menuBar a.menuButton {
			background-color: transparent;
			border: 1px solid <% =Application("av_MajorColor") %>;
			color: #FFFFFF;
			cursor: default;
			left: 0px;
			margin: 1px;
			padding: 2px 6px 2px 6px;
			position: relative;
			text-decoration: none;
			top: 0px;
			z-index: 100;
}

div.menuBar a.menuButton:hover {
			background-color: transparent;
			border-color: 1px outset <% =Application("av_MajorColor") %>;
}

div.menuBar a.menuButtonActive, div.menuBar a.menuButtonActive:hover {
			background-color: <% =Application("av_MajorColorDark") %>;
			border-color: 1px inset <% =Application("av_MajorColor") %>;
			left: 1px;
			top: 1px;
}

div.menu {
			background-color: #C0C0C0;
			border: 2px solid;
			border-color: #F0F0F0 #909090 #909090 #F0F0F0;
			left: 0px;
			padding: 0px 1px 1px 0px;
			position: absolute;
			top: 0px;
			visibility: hidden;
			z-index: 101;
}

div.menu a.menuItem {
			color: #000000;
			cursor: default;
			display: block;
			padding: 3px 1em;
			text-decoration: none;
			white-space: nowrap;
}

div.menu a.menuItem:hover, div.menu a.menuItemHighlight {
			background-color: <% =Application("av_MajorColorDark") %>;
			color: #FFFFFF;
}

div.menu a.menuItem span.menuItemText {}

div.menu a.menuItem span.menuItemArrow {
			margin-right: -.75em;
}

div.menu div.menuItemSep {
			border-top: 1px solid #909090;
			border-bottom: 1px solid #F0F0F0;
			margin: 4px 2px;
}

.userName     {color: #FFFFFF;
			font-family: "MS Sans Serif", Arial, Tahoma,sans-serif;
			font-size: 8pt;
			font-style: normal;
			font-weight: normal;
			padding-right: 10px;
			text-align: right;
			width: 100%;
}

.breadcrumb {-moz-box-sizing: border-box;
			background-color: <% =Application("av_MajorColorLight") %>;
			box-sizing: border-box;
			padding-top: 3px;
			padding-left: 10px;
			padding-bottom: 5px;
			position: relative;
			left: 0px;
			top: 0px;
}

.breadcrumbtext {
			color: #FFFFFF;
			font-family: Verdana, Arial, Tahoma,sans-serif;
			font-size: 7pt;
}

.C0,.C1,.C2,.C3,.C4,.C5
		{font-family: Verdana, Arial, Helvetica;
		font-size: 8pt;}

.C0    {background-color: #FFFFFF;}
.C1    {background-color: #F2F2F2;}
.C2    {background-color: #DEDEDE;}
.C3    {background-color: #808080;}
.C4    {color: white; background-color: #333333;}
.C5    {color: white; background-color: #000000;}