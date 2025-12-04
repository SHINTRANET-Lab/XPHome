<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>Ad Rotator</TITLE>
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

<H2><A NAME="_ad_rotator"></A> <A NAME="_ad_rotator"></A>Ad Rotator</H2>

<H6>概要</H6>

<P>Ad Rotator コンポーネントは、AdRotator オブジェクトを作成します。このオブジェクトは、広告イメージを自動的に入れ替えて Web ページに表示します。このコンポーネントは、クライアント ブラウザが Web ページを開いたり再読み込みするたびに、新しい広告を表示するように設計されています。</P>

<H6>コードの説明</H6>

<P>このサンプルでは、Ad Rotator コンポーネントの使用方法を示します。スクリプト自体は簡単な構造であり、AdRotator オブジェクトのインスタンスを作成して、ページが読み込まれるたびに GetAdvertisement メソッドを呼び出します。GetAdvertisement は、単一の広告エントリを Ad Rotator スケジュール ファイル、この例では adrot.txt から返します。特別なスクリプト デリミタ &lt;% = ... %&gt; が、結果をクライアント ブラウザに表示します。</P>

<P>この例のスケジュールファイル adrot.txt は、どちらかといえば単純です。アスタリスク (*) の上の 4 行は、グローバル ファイル設定であり、ファイル内のすべてのスケジュール エントリに影響します。最も便利なグローバル設定は、リダイレクト指定です。このサンプルでは、現在表示されているエントリがどれであっても、ユーザーが広告をクリックすると、ユーザーは示された .asp ファイルに移動します。これらのスクリプト、つまり DLL は、通常は特定の広告にヒットした回数をカウントし、ユーザー情報を集め、次に要求から URL を取得してから、最初に要求された URL にクライアント ブラウザを再びリダイレクトします。</P>

<P>アスタリスクの下にある各エントリは 4 行で構成され、表示するイメージ、ハイパーリンク、代替テキスト、および Web ページが利用されたときにそのエントリが表示される相対的な確率をそれぞれ記述します。したがって、Microsoft<sup>&reg;</sup> インターネット インフォメーション サービスの広告イメージがこのページのヒットで表示される確率は 80 パーセントです。一方、Microsoft Internet Explorer のイメージが表示される確率は、ヒットの中の 20 パーセントです。</P>

<P><span class=le><B>注&nbsp;&nbsp;&nbsp;</B></span>スケジュール ファイルの広告エントリが対応する URL を持たない場合、このエントリのハイパーリンク行にはハイフン (-) を入れなければなりません。ハイフンが入っていない場合、Ad Rotator コンポーネントはエラーを返します。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
