<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>基本トランザクション</TITLE>
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

<H2><A NAME="_basic_transaction"></A> <A NAME="_basic_transaction"></A>基本トランザクション</H2>

<H6>概要</H6>

<P>ASP を使用すると、@TRANSACTION ディレクティブをスクリプトに入れるだけで、コンポーネント サービスに備わる信頼性を容易に活用できます。このディレクティブは、データベース処理またはメッセージ キュー サービスのメッセージの送信など、そのページで発生している変更をトランザクションと見なすようにコンポーネント サービスに指示します。トランザクション サービスが管理する変更は、コミットされて永続的なものとなるか、または中止されて、データベース内、またはキューとして変更の加えられる前の状態に戻るかのいずれかです。</P>

<H6>コードの説明</H6>

<P>このサンプルでは、@TRANSACTION ディレクティブを使用して、ページ全体がトランザクションであると宣言します。また、さらに完了タスク、またはクリーンアップ タスクを実行するために呼び出される、2 つの別のプロシージャ用のスクリプト コマンドもいくつか含まれます。OnTransactionCommit は、スクリプトの実行が成功したとき、または ObjectContext.SetComplete メソッドが呼び出されたときに呼び出されます。同様に、OnTransactionAbort は、スクリプトの実行で何らかの処理エラーが発生したとき、または ObjectContext.SetAbort メソッドが呼び出されたときに呼び出されます。</P>

<P>このサンプルは単にメッセージをいくつか出力して終了するだけのものであり、既定でコミットします。ディレクティブがスクリプトをトランザクションであると宣言しているため、実行が成功して終了すると、スクリプトの中で加えられた変更が自動的にコミットされ (この例にはありませんが)、OnTransactionCommit プロシージャが起動されてメッセージが出力されます。</P>

<P><span class=le><B>重要&nbsp;&nbsp;&nbsp;</B></span>@TRANSACTION ディレクティブは、.asp ファイルの先頭の行に配置する必要があります。先頭にない場合は、エラーが発生します。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
