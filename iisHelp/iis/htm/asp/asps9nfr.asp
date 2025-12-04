<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>ループ処理</TITLE>
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

<H2><A NAME="_looping"></A> <A NAME="_looping"></A>ループ処理</H2>

<H6>概要</H6>

<P>ループ処理は、プログラム言語の最も重要なフロー制御方法の 1 つです。ループ処理構造は、タスクの繰り返し処理が必要なアプリケーションの基本構造となります。これらの繰り返しタスクには、変数に 1 を加える処理、テキストファイルの読み込み、または電子メール メッセージの処理と送信などがあります。 </P>

<H6>コードの説明</H6>

<P>VBScript と JScript には、いくつかのループ処理機構が備わっています。 このサンプルでは、3 つの最もよく使用されるループ処理ステートメントの、For ... Next、Do ... Loop、および While ... Wend を示します。これらの 3 つのステートメントは微妙に異なり、スクリプトの作成状況により、使用に最適なステートメントが決まります。ただし、このサンプルでは、各タイプのループ処理ステートメントを使用して同じタスクを実行します。タスクは、徐々にフォント サイズを大きくしながらあいさつメッセージを 5 回出力する処理です。各ループ処理ステートメントで、変数 <I>i</I> が初期化され、<I>i</I> が 5 より大きくならないというループの判定条件が定義されます。ループを繰り返すたびに、変数が 1 ずつインクリメントされます。 </P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
