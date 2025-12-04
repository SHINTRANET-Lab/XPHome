<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>Browser Capabilities</TITLE>
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

<H2><A NAME="_browser_capabilities"></A> <A NAME="_browser_capabilities"></A>Browser Capabilities</H2>

<H6>概要</H6>

<P>すべてのブラウザの機能が同じわけではありません。相違点を簡単に明確化するために、ASP には Browser Capabilities コンポーネントが備わっています。このコンポーネントは、BrowserType オブジェクトを使用して、スクリプトにクライアント ブラウザの機能の説明を提供します。</P>

<H6>コードの説明</H6>

<P>最初に、BrowserType オブジェクトのインスタンスを作成し、オブジェクト変数の <I>bc</I> に割り当てる必要があります。次に、<I>object.property</I> 構文を使用して、オブジェクトから各プロパティを順番に要求します。スクリプト例を実行して、ブラウザの動作を試してください。</P>

<P>Internet Explorer Version 5.0 以降を使用している場合は、表示される機能の一覧の中に、新しいクライアント機能の cookie を使用して判断された機能が含まれています。これについては、「<A HREF="/iishelp/iis/htm/asp/eadg6isz.htm">クライアント機能</A>」を参照してください。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
