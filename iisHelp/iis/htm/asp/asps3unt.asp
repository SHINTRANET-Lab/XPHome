<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>簡単なクエリ</TITLE>
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

<H2><A NAME="_simple_query"></A> <A NAME="_simple_query"></A>簡単なクエリ</H2>

<H6>概要</H6>

<P>データベースはたいへん複雑なシステムであり、データにアクセスするツールは強力で応答性が高くなければなりませんが、同様に、簡単なデータベース アクセス タスクが容易に実行できることも重要です。このサンプルは、ADO でこのようなアクセス タスクを実行する簡単な方法を示します。</P>

<H6>コードの説明</H6>

<P>このサンプル アプリケーションの目的は、Microsoft<sup>&reg;</sup> Access データベースから小さなレコードセットを取得し、結果を出力することです。最初のステップでは、Server.CreateObject メソッドを使用して Connection オブジェクトのインスタンスを作成します。サンプルでは、Connection オブジェクトのインスタンスを使用して OLE DB データ プロバイダを開き、次に Connection インスタンスを再度使用して SQL SELECT コマンドを実行し、Authors テーブルからすべてのレコードを取得します。最後にスクリプトは、戻されたレコードセット コレクションを繰り返し処理して、結果を表示します。その後、レコードセットと OLE DB データ ソース接続の両方が閉じられます。</P>

<P><span class=le><B>重要&nbsp;&nbsp;&nbsp;</B></span>このサンプルを適切に実行するためには、サーバー上に OLE DB を適切に構成する必要があります。 </P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
