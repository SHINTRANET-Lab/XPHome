<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>有効期限の情報を設定する</TITLE>
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

<H2><A NAME="_setting_expiration_information"></A> <A NAME="_setting_expiration_information"></A>有効期限の情報を設定する</H2>

<H6>概要</H6>

<P>サーバー上の各ページまたはスクリプトに、有効期限の日時を設定することができます。有効期限は、2000 年 1 月 1 日などの絶対的な日付、またはクライアント ブラウザがそのページを最初にダウンロードした時刻から 600 分などの相対的日付のいずれかの形式で指定できます。ページの有効期限が来る前に、クライアント ブラウザが再度同じページを要求した場合、クライアント ブラウザはキャッシュされたページのコピーを使用します。</P>

<H6>コードの説明</H6>

<P>この例では、スクリプトでファイルに有効期限の日時を設定する方法を示します。Response.Expires プロパティを使用して、相対有効期限を設定します。設定単位は分です。したがって、たとえばこのプロパティが 10 に設定された場合、このページは 10 分後に有効期限が切れます。</P>

<P>Response.ExpiresAbsolute プロパティを使用して、絶対有効期限を指定します。この例では、1999 年 1 月 1 日午後 1 時 30 分 15 秒にページの有効期限が切れるように指定されます。有効期限を割り当てるときに、時刻または日付のいずれかを省略できます。</P>

<P><span class=le><B>注&nbsp;&nbsp;&nbsp;</B></span>IIS は、HTML がクライアント ブラウザに送信される前に、所定の Web ページまたはスクリプトに必要なすべての HTTP ヘッダーを送信します。したがって、Response.Expires および Response.ExpiresAbsolute プロパティの設定を含む、応答の HTTP ヘッダーを変更するすべてのステートメントとメソッドを、スクリプトの &lt;HTML&gt; タグより前に配置する必要があります。サーバーが HTML コンテンツをクライアント ブラウザに戻す送信を開始した後で、スクリプトが HTTP ヘッダーを変更しようとすると、エラーが発生します。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
