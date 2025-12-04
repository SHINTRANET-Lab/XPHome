<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>POST によるユーザー フォーム インプット</TITLE>
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

<H2><A NAME="_user_form_input_with_post"></A> <A NAME="_user_form_input_with_post"></A>POST によるユーザー フォーム インプット</H2>

<H6>概要</H6>

<P>Web 上での対話の最も基本的なフォームは、おそらく HTML フォームです。ASP はフォームを置き換えるのではなく、むしろフォームを強化して、その実装と管理を容易にします。</P>

<P>HTML の &lt;FORM&gt; タグは、フォームが情報を処理スクリプトに渡すために使用するメソッドを指定します。POST メソッド属性は、フォームからの情報が個別の HTTP 接続を通して処理スクリプトまたはプログラムに渡されることを示します。処理スクリプトまたはプログラムは、情報を解析し、要求されたタスクを実行して、クライアント ブラウザに出力を戻します。</P>

<H6>コードの説明</H6>

<P>この例では、HTML POST メソッド属性を使用して簡単なフォームを実装する方法、および ASP でフォームを作成する際のたいへん有効な機能である、フォームと実際の処理コードを結びつけて同じファイルにする機能を示します。この例では、2 つのテキスト入力ボックスを備えた小さなフォームを作成します。1 つの入力ボックスにはユーザーの名前 (<I>fname</I>) を入力し、ほかの 1 つには名字 (<I>lname</I>) を入力します。Request.Forms コレクションがアクセスされ、要求から <I>fname</I> 変数と <I>lname</I> 変数の値を取得して、結果をページの下部に表示します。</P>

<P>初めてこのスクリプトを実行した場合、線の下にはテキストがなにも表示されません。これは、開始時にこのスクリプトに渡せるフォーム情報が何もないためです。また、ASP は存在しない情報を探す Request.Forms の検索を無視するためです。ただし、[Submit] を押した場合、ページが再度読み込まれ、テキスト ボックスに入力した情報がスクリプトで使用できるようになります。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
