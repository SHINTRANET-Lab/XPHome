<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>配列</TITLE>
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

<H2><A NAME="_arrays"></A> <A NAME="_arrays"></A>配列</H2>

<H6>概要</H6>

<P>配列は、本質的には複数の値を保持できる変数です。VBScript と JScript は、どちらもほとんどのデータ型の配列をサポートします。このサンプルでは、ASP スクリプトで配列を作成して使用する方法を示します。</P>

<P>配列の各要素には、インデックスを使用してアクセスします。たとえば配列 <I>A</I> で、<I>A(0)</I> は先頭の要素を参照し、<I>A(1)</I> は 2 番目の要素を参照します。VBScript でも JScript でも、番号は 0 から始まる点に注意してください。</P>

<P>配列のサイズは、固定サイズにも動的に変化するサイズにもできます。VBScript も JScript も、どちらも新しい配列のために変数名を宣言し、格納スペースを割り当てる必要があります。これは、VBScript では Dim ステートメントで行い、JScript では new ステートメントで行います。特定のサイズが明示的に記述された場合は、配列は固定サイズになりますが、宣言で特定のサイズが省略されている場合は、配列のサイズは動的に変化すると見なされます。 </P>

<H6>コードの説明</H6>

<P>このサンプルでは、固定サイズの配列 <I>aFixed</I> と、動的変動サイズの配列 <I>aColors</I> の 2 つの配列が作成されます。次にスクリプトは、適切な要素にアクセスするためのインデックス表記である <I>ArrayName</I>(<I>Index</I>) を使用して、特定の文字列の値を配列の各要素に割り当てます。スクリプトは For ループを使用して各配列をループ処理し、テーブルに結果を表示します。</P>

<P>JScript では、動的にサイズが変化する配列は、割り当ての中でインデックス付けされた数値の最も高い要素が入るように、暗黙的に拡張される点に注意してください。これとは対称的に VBScript では、スクリプトの作成元が ReDim ステートメントで動的配列を明示的にリサイズする必要があります。 </P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
