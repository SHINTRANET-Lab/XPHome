<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>cookie の使い方</TITLE>
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

<H2><A NAME="_using_cookies"></A> <A NAME="_using_cookies"></A>cookie の使い方</H2>

<H6>概要</H6>

<P>cookie を使用して、特定のクライアント、セッション、またはアプリケーションに関する情報を保存します。次にこの保存された情報を使用して、クライアント ブラウザのセッションをカスタマイズしたり簡素化したりします。</P>

<H6>コードの説明</H6>

<P>このサンプルは、アプリケーションで特定の cookie の値を問い合わせる方法を示します。IIS では、Request.Cookies コレクションを使用して、ASP スクリプトで cookie が使用できるようにします。この例では、最初に標準的なコレクションアクセス形式の <I>object.collection</I>(<I>keyname</I>) を使用して、cookie の CookieVBScript または CookieJScript を問い合わせます。次に、cookie を現在の日付と時刻にリセットします。</P>

<P>最初のクエリの結果が NULL 文字列である場合、これはクライアント ブラウザがこのページを以前に利用したことがないことを示し、初回の歓迎メッセージが表示されます。CookieVBScript または CookieJScript の最初のクエリで値が戻された場合は、クライアント ブラウザが以前に利用したことを示すだけではなく、最後にいつ利用されたかを示します。</P>

<P><span class=le><B>注&nbsp;&nbsp;&nbsp;</B></span>IIS は、HTML がクライアント ブラウザに送信される前に、所定の Web ページまたはスクリプトに必要なすべての HTTP ヘッダーを送信します。したがって、Response.Cookies コレクション メンバの設定を含む、応答の HTTP ヘッダーを変更するすべてのステートメントとメソッドが、スクリプトの &lt;HTML&gt; タグより前に配置されている必要があります。サーバーが HTML コンテンツをクライアント ブラウザに戻す送信を開始した後で、スクリプトが HTTP ヘッダーを変更しようとすると、エラーが発生します。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
