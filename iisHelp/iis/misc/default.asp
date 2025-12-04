<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">

<% @Language = "VBScript" %>

<html dir=ltr><head><title>Microsoft インターネット インフォメーション サービス&nbsp;5.1 マニュアル</title>

<META NAME="ROBOTS" CONTENT="NOINDEX">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=shift_jis">

</head>

<% Set OBJbrowser = Server.CreateObject("MSWC.BrowserType")
	BrsType = Request.ServerVariables("HTTP_User-Agent")
	MachType=Request.ServerVariables("HTTP_UA-CPU")
	If InStr(BrsType, "MSIE") Then
		If InStr(BrsType, "Windows") Then
			File="contents.asp" 
			Size="30"
			Scroll="Auto"
		Else
			File="coflat.htm"
			Size="34"
			Scroll="Yes"
		End If
		If MachType="Alpha" Then
			File="contalph.asp"
			Size="30"
			Scroll="Auto" 
		End If
	Else

		File="coflat.htm"
		Size="34"
		Scroll="Yes"
	End If
%>

 
<%
        If Request.QueryString("jumpurl") <> "" Then
                strMainUrl = Server.HTMLEncode(Request.QueryString("jumpurl"))
		quote = chr(34)
		strMainUrl = replace(str, quote, quote & quote)
        Else
                strMainUrl = "../htm/core/iiwltop.htm" 
        End If
 %>


<!--frameset cols="275,*"-->


<frameset rows="<% =Size%>,*" FRAMEBORDER="0" FRAMESPACING="0">
	<frame src="navbar.asp" name="NavBar" scrolling="No" noresize marginheight="0" marginwidth="0" framespacing="0" frameborder="No">
	<frameset cols="284,*">
        	<frame src=<% =File%> name="contents"  scrolling=<% =Scroll%> FRAMEBORDER="0" FRAMESPACING="0">
        	<frame src="<% =strMainUrl%>" name="main" FRAMEBORDER="0" FRAMESPACING="0">
	</frameset>
</frameset>


<noframes>

<body bgcolor="#FFFFFF" text="#000000"><font face="ＭＳ Ｐゴシック">

<h1>Microsoft インターネット インフォメーション サービス&nbsp;5.1 マニュアル</h1>

<p>Microsoft インターネット インフォメーション サービス&nbsp;5.1 マニュアルを表示するには、フレームをサポートするブラウザを使用する必要があります。マニュアルを表示するには、次のアイコンをクリックして、Microsoft&reg; Internet Explorer の最新バージョンをダウンロードしてください。</p>

<p><A HREF="http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=inews" target="_top"><IMG SRC="../../common/bestwith.gif" ALT="ここをクリックして開始します" ALIGN="BOTTOM" BORDER="0" VSPACE="7" WIDTH="88" HEIGHT="31" HSPACE="5"></A></p>

<hr class="iis" size="1">
<p align="center"><em><a href="../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation. All rights reserved.</a></em></p>
</font>
</noframes>

</body>
</html>