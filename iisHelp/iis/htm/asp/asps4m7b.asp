<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>クライアント側スクリプト作成</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
	TempString = navigator.appVersion
	if (navigator.appName == "Microsoft Internet Explorer"){	
// Check to see if browser is Microsoft
		if (TempString.indexOf ("4.") >= 0){
// Check to see if it is IE 4
			document.writeln('<link rel="stylesheet" type="text/css" href="/iishelp/common/coua.css">');
		}
		else {
			document.writeln('<link rel="stylesheet" type="text/css" href="/iishelp/common/cocss.css">');
		}
	}
	else if (navigator.appName == "Netscape") {						
// Check to see if browser is Netscape
		document.writeln('<link rel="stylesheet" type="text/css" href="/iishelp/common/coua.css">');
	}
	else
		document.writeln('<link rel="stylesheet" type="text/css" href="/iishelp/common/cocss.css">');
//-->
</script> <SCRIPT LANGUAGE="VBScript">
<!--
Sub Window_OnLoad()
   Dim frmContents
   On Error Resume Next
   If Not Parent Is Nothing Then
      Set frmContents = Parent.Contents
      If Not frmContents Is Nothing Then
            frmContents.Window.TOCSynch_Click
      End If
   End If
End Sub
//--></SCRIPT><META NAME="DESCRIPTION" CONTENT="インターネット インフォメーション サービス リファレンス情報">
<META HTTP-EQUIV="PICS-Label" CONTENT='(PICS-1.1 "<http://www.rsac.org/ratingsv01.html>" l comment "RSACi North America Server" by "inet@microsoft.com <mailto:inet@microsoft.com>" r (n 0 s 0 v 0 l 0))'>
<META NAME="MS.LOCALE" CONTENT="JA">
<META NAME="MS-IT-LOC" Content="インターネット インフォメーション サービス"> 
</HEAD>

<BODY BGCOLOR="#FFFFFF" TEXT="#000000">

<H2><A NAME="_client_side_scripting"></A> <A NAME="_client_side_scripting"></A>クライアント側スクリプト作成</H2>

<H6>概要</H6>

<P>ASP は、サーバー側スクリプト作成です。クライアント側スクリプト作成は、ASP を適切に補い、ActiveX<sup>&reg;</sup> コントロールなどの多数の拡張機能を提供します。これにより、アプリケーションはさらに強力で、ユーザーに使いやすいものになります。</P>

<H6>コードの説明</H6>

<P>この例では、使用するスクリプト言語でクライアント側スクリプトをインクルードする方法を示します。このスクリプトは、DoIt というサブルーチンを &lt;SCRIPT&gt; タグの中で定義します。スクリプトが ASP のサブルーチンであることを示す、RUNAT=SERVER 属性が存在しない点に注意してください。このページには、ユーザーが使用するボタンが 1 つあり、これをクリックすると、クライアント ブラウザ上で DoIt サブルーチンが実行されます。</P>

<P>このサンプルは、ASP スクリプトをクライアント側スクリプトに結びつける利点を示しています。ASP が &lt;SCRIPT&gt; タグを検出したときに、ASP は単にこのブロック内のコードをすべて無視するわけではありません。ASP は、サーバーにとって意味のある、デリミタ (&lt;% ... %&gt;) で指定されたスクリプト要素の検出、解析、および実行を続けます。したがってこの例では、クライアント側スクリプトの中で、Session.SessionID メソッドから戻されたセッション情報が戻ります。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
