<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>スクロールできるクエリ</TITLE>
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

<H2><A NAME="_scrollable_query"></A> <A NAME="_scrollable_query"></A>スクロールできるクエリ</H2>

<H6>概要</H6>

<P>インターネット環境のアプリケーションを設計するときに、データベース クエリがクライアント ブラウザに返す情報の量を制限することが望ましい場合がよくあります。この例では、スクリプトが ASP と ADO を使用して、クライアント ブラウザに 1 回に渡す行の数を制限しながら、ユーザーがクエリのすべての結果を閲覧できる方法を示します。</P>

<H6>コードの説明</H6>

<P>スクリプトはいくつかのコード セクションで構成され、これらのセクションが一緒になってこのタスクを実行します。最初に、通常どおりにデータベースがアクセスされ、Connection オブジェクトと Recordset オブジェクトが作成されます。Recordset オブジェクトの PageSize プロパティが 4 に設定され、レコードセットが開かれてデータベース テーブルの Authors からのクエリ結果の値が代入されます。4 つの結果レコードの最初の論理ページが表になって表示されます。ユーザーがほかのレコードセットのページを表示するために、2 つのボタン PgUp と PgDn が設けられています。</P>

<P>ユーザーがボタンをクリックすると、ページが再びアクセスされます。このときは、POST メソッドを使用して、変数のいくつかをこのページの次のコピーに渡します。ユーザーが現在表示しているページを保存するために変数 <I>PageNo</I> が使用され、<I>Mv</I> 変数で次のフォームへのスクロールする方向を渡します。たとえばユーザーが PgDn をクリックした場合、ページが再びアクセスされ、<I>Mv</I> は PgDn に設定され、<I>PageNo</I> は 1 に設定されます。スクリプトはこの情報を使用して現在のページ番号に 1 を追加し、AbsolutePage を使用して次の結果ページを表示します。</P>

<P><span class=le><B>重要&nbsp;&nbsp;&nbsp;</B></span>このサンプルを適切に実行するためには、サーバー上に OLE DB を適切に構成する必要があります。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
