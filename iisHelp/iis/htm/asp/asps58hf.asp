<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>関数とプロシージャ</TITLE>
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

<H2><A NAME="_functions_and_procedures"></A> <A NAME="_functions_and_procedures"></A>関数とプロシージャ</H2>

<H6>概要</H6>

<P>関数とプロシージャを使用すると、特定のタスクを実行するたびに同じコードのブロックを再度作成する必要がなくなります。VBScript と JScript では、スクリプトの任意の場所から関数またはプロシージャを呼び出すことができます。この例では、ASP を用いてこれらのツールを作成し、使用する方法を示します。 </P>

<P>ASP ページに関数が何もない場合、クライアント ブラウザが要求するたびに、ASP エンジンは単にファイル全体を開始から終了まで処理します。ただし関数とプロシージャは、コードのそのほかの部分と一緒に順次処理されるのではなく、呼び出されたときにのみ実行されます。 </P>

<P>関数とプロシージャは、Function ステートメントを使用して VBScript または JScript で記述できます。さらに VBScript は、値を戻す関数と戻さない関数を区別します。値を戻す関数は Sub ステートメントで表記し、これがサブルーチンであることを示します。 </P>

<H6>コードの説明</H6>

<P>このサンプルでは、1 つの関数 PrintOutMsg を定義します。この関数はパラメータとして 1 つのメッセージ、およびこのメッセージが Response.Write メソッドでクライアント ブラウザに書き込まれる回数を指定する数値を使用します。このサンプルの目的に従って、関数はメッセージが出力された回数をクライアント ブラウザに戻します。 </P>

<H6>解説</H6>

<P>&lt;SCRIPT&gt; タグの RUNAT 属性に注意することが大切です。RUNAT 属性が含まれていない場合、ASP はスクリプトがクライアント側スクリプトであると想定し、ブラウザで処理するためにコードをブラウザに戻します。さらに ASP は PrintOutMsg 関数呼び出しを認識せず、エラーを返して中止します。 </P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
