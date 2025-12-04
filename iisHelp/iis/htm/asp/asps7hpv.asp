<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>フィールドに値を代入する</TITLE>
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

<H2><A NAME="_populating_fields"></A> <A NAME="_populating_fields"></A>フィールドに値を代入する</H2>

<H6>概要</H6>

<P>フォームを使用してユーザーからの入力を集めることができますが、フォームは情報の表示にも使用できます。たとえば、クライアント ブラウザがディレクトリ検索エンジンにアクセスした場合、フォームを使ってその検索結果を表示することができます。検索スクリプト (ASP に実装できます) は入力を受け取り、データベースにアクセスし、クエリ文字列の結果を表示フォームに送信します。このサンプルは、この表示フォームがどのように表示されるかを示す簡単な例です。</P>

<H6>コードの説明</H6>

<P>このサンプルでは、データはハードコードされていますが、情報は対話型フォーム、データベース、またはテキスト ファイルから取得できます。このサンプルは、変数を初期化することから始まります。次に、HTML &lt;FORM&gt; タグでフォームが作成され、4 つのテキスト ボックスが定義されます。スクリプトは、サーバー側スクリプト デリミタ &lt;%= ...<B> </B>%&gt; を使用して、初期化の中で設定された値をテキスト ボックスに代入します。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
