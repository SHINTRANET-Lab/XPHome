<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>HTTP サーバー変数</TITLE>
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

<H2><A NAME="_http_server_variables"></A> <A NAME="_http_server_variables"></A>HTTP サーバー変数</H2>

<H6>概要</H6>

<P>クライアント ブラウザとサーバー間で実行される各 HTTP 処理には、潜在的に有効な大量の情報の交換が含まれます。このサンプルでは、サーバー変数アクセスによる、この情報のアクセス方法の 1 つを示します。サーバー変数を使用すると、SERVER_NAME 変数でサーバーの名前が、または ALL_HTTP 変数で HTTP ヘッダーが取得できます。</P>

<H6>コードの説明</H6>

<P>サンプルの構造は簡単です。2 列のテーブルを作成し、テーブルの各行に <I>name, value</I> の変数ペアを入れていきます。個々のサーバー変数は、Request.ServerVariables ("ALL_HTTP") などのように、名前をキーとして使用して取得されます。さらに、ServerVariables はコレクションなので、VBScript の For Each ... Next ループまたは JScript の Enumerator オブジェクトを使用し、コレクションにあるすべての変数を反復して処理することができます。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
