<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>サーバー側インクルード</TITLE>
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

<H2><A NAME="_server_side_includes"></A> <A NAME="_server_side_includes"></A>サーバー側インクルード</H2>

<H6>概要</H6>

<P>ASP スクリプトの開発では、コードのモジュール化と再利用が可能であることがたいへん有効です。たとえば、HTML ページと ASP ページのページごとに、著作権情報を表示する場合、ASP ではサーバー側インクルードという機能を提供してこの問題を解決します。これは、テキスト ファイル、グラフィカル イメージ、または ASP 関数などの特定のファイルをインクルードするようにサーバーに指示するディレクティブです。著作権情報は、1 つのファイルになっていて、Web サイトのこのほかのファイルにインクルードできます。また、著作権情報に変更があった場合、1 つのファイルを変更するだけでよく、50 または 500 も変更しなくて済みます。</P>

<P>ファイルをインクルードする構文は、次のとおりです。</P>

<PRE><CODE>&lt;!-- #include PathType=Name --&gt;
</CODE></PRE>

<P>PathType パラメータは、キーワードの FILE または VIRTUAL からなります。このキーワードは、指定された <I>Name</I> 文字列が物理パスであるのか、または仮想パスであるのかを示します。 </P>

<H6>コードの説明</H6>

<P>この例では、#include ディレクティブを使用してファイル HeaderInfo.asp をインクルードします。このスクリプトが実行されると、ASP は #include ディレクティブを検出するまで、スクリプトを 1 行ずつ、1 文字ずつ読み込みます。#include ディレクティブを検出すると、指定されたファイルのコンテンツを 1 行ずつ読み込みます。次に、サンプル スクリプトの残りの部分が読み込まれます。これが終了すると、インクルードされたファイルも含み、スクリプトがすべて実行されます。 </P>

<P><span class=le><B>注&nbsp;&nbsp;&nbsp;</B></span>ASP スクリプトがインクルードしたファイルに大量の関数および変数が含まれているときに、ファイルをインクルードしたスクリプトがそれらを使用しない場合、これらの使用しない構造体が占める余分なリソースは、パフォーマンスに悪い影響を与え、最終的に Web アプリケーションの拡張性が損なわれる可能性があります。したがって、不要な情報も含む 1、2 の大きなインクルード ファイルをインクルードするのではなく、インクルード ファイルを複数の小さなファイルに分割し、これらの中から ASP スクリプトで必要なファイルだけをインクルードすることをお勧めします。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
