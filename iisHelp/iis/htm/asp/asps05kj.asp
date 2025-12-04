<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>ストアド プロシージャを実行する</TITLE>
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

<H2><A NAME="_executing_stored_procedures"></A><A NAME="_executing_stored_procedures"></A>ストアド プロシージャを実行する</H2>

<H6>概要</H6>

<P>Microsoft<sup>&reg;</sup> SQL Server などが提供するストアド プロシージャは、大型の基幹データベース アプリケーションを円滑かつ効率的に運用するために欠かせないものです。この例では、ASP スクリプト内で ADO を使用して、この機能にアクセスする方法を示します。</P>

<H6>コードの説明</H6>

<P>このスクリプトは、最初に Connection オブジェクトのインスタンスを作成し、これを使用してサンプル データベース <I>pubs</I> への OLE DB 接続を開きます。次に、Command オブジェクトという特殊なオブジェクトを作成します。Command オブジェクトの CommandText プロパティは、発行するコマンドの文字列に設定されます。この例では、ストアド プロシージャの名前、ByRoyalty になっています。Command オブジェクトの Parameters プロパティには、Parameter オブジェクトのコレクションが入ります。このスクリプトでは、Append メソッドを使用して、新しいパラメータをこのコレクションに追加します。</P>

<P>CreateParameter を使用してパラメータ インスタンスの名前を指定して構成した後は、Command オブジェクト自体がコレクションであるかのように、指定されたパラメータ名を使用して直接パラメータの値をアクセスできます。たとえば、次のようになります。 </P>

<PRE><CODE>oCmd("@Percentage") = 75 </CODE></PRE>

<P>このコードでは、スクリプトが、<CODE>@Percentage</CODE> というラベルをつけたパラメータに値 <CODE>75</CODE> を割り当てます。Command オブジェクトの Execute メソッドが起動され、結果のレコードセットが、既にスクリプトで定義されているオブジェクト変数 <I>oRs </I>に割り当てられます。結果のレコードセットの先頭のレコードが表示されます。</P>

<P><span class=le><B>重要&nbsp;&nbsp;&nbsp;</B></span>このサンプルが正しく動作するためには、IIS を実行しているコンピュータと同じコンピュータ上に Microsoft SQL Server をインストールし、適切に構成する必要があります。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
