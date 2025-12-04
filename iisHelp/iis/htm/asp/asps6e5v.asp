<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>変数</TITLE>
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

<H2><A NAME="_variables"></A> <A NAME="_variables"></A>変数</H2>

<H6>概要</H6>

<P>作成に使用されるプログラム言語が何であっても、あらゆるアプリケーションで何らかの変数が使用されています。ASP スクリプトも例外ではありません。VBScript でも JScript でも、変数の作成と管理は簡単です。 </P>

<P>言語により、変数宣言の方法は異なります。JScript と VBScript は、どちらも変数とその宣言についてはたいへん柔軟性があります。VBScript では、どの変数も最初に Dim ステートメントで宣言されたときに、自動的にバリアント型であると見なされます。各変数には、最終的に数値型や配列などの内部処理形式が割り当てられます。JScript も同様です。変数は最初に var ステートメントで宣言されます。一般に、どちらの言語も、型変換を含めたデータ型管理のほとんどを自動的に実行しようとします。実際、新しい変数を使用するために、Dim ステートメントまたは var ステートメントをまったく使用する必要がないほどです。これらのステートメントは、それぞれの言語で省略可能です。</P>

<H6>コードの説明</H6>

<P>このサンプルでは、いくつかの異なる種類の変数を宣言し、それらを使用して簡単な操作を実行してから、特別な &lt;% = ...%&gt; スクリプト デリミタを使用してクライアント ブラウザにそれらを表示します。整数が変数 <I>intVariable</I> に割り当てられ、その変数自体に加算されて、合計がクライアント ブラウザに送信されます。<I>StrVariable</I> がファースト ネームと等しくなるように設定され、Smith に追加されてクライアント ブラウザに送信されます。ブール値と日付が同様に宣言または作成され、初期化と処理が行われて表示されます。 </P>

<H6>解説</H6>

<P>この例で特に興味を引くのが、日付変数を示した部分の最後のステップです。VBScript では、変数に最初にリテラルの日付文字列が割り当てられて表示されます。次にこれがリセットされ、VBScript の Now 関数の戻り値が割り当てられます。この関数は、現在のシステム時間を返します。JScript の例では、JScript の Date 関数が使用され、これで両方の値を設定します。つまり、パラメータを関数に渡すことにより最初のリテラルを設定し、次に、関数にパラメータを何も渡さないことにより現在のシステム時間と等しくなるように変数を設定します。 </P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
