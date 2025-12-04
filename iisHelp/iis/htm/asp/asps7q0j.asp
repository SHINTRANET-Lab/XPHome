<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>レコードを追加および削除する</TITLE>
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

<H2><A NAME="_add_delete_records"></A> <A NAME="_add_delete_records"></A>レコードを追加および削除する</H2>

<H6>概要</H6>

<P>このサンプルでは、ASP と ADO を使用してデータベースのレコードを追加および削除する方法を示します。</P>

<H6>コードの説明</H6>

<P>最初に、CreateObject を使用して、Connection オブジェクトのインスタンスを作成します。次にこのオブジェクトを使用して、OLE DB データ プロバイダへの接続を開きます。CreateObject を再度使用し、今度は空の Recordset オブジェクトを作成します。ActiveConnection プロパティが、新しい Connection オブジェクトを参照するように設定されます。</P>

<P>ADO の Recordset オブジェクトは、新しいレコードをデータベースに追加するための AddNew メソッドを備えていますが、SQL INSERT コマンドをデータベース エンジンに直接送信するほうが、より高い拡張性を得ることができます。このサンプルでは、Recordset.Execute コマンドを適切な SQL コマンド文字列と共に使用して、新しい作成者用の情報を挿入します。</P>

<P>この時点で、別の Recordset オブジェクト インスタンスが作成され、別の SQL コマンドで開かれます。追加されたばかりのレコードが選択され、次に、SQL DELETE コマンドを直接データベース エンジンに渡すことにより、このレコードが削除されます。この後、スクリプトが終了します。</P>

<P><span class=le><B>重要&nbsp;&nbsp;&nbsp;</B></span>このサンプルを適切に実行するためには、サーバー上に OLE DB を適切に構成する必要があります。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
