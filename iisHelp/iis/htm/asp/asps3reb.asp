<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML DIR=LTR>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=shift_jis">
<TITLE>クエリの結果を制限する</TITLE>
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

<H2><A NAME="_limit_query_results"></A> <A NAME="_limit_query_results"></A>クエリの結果を制限する</H2>

<H6>概要</H6>

<P>インターネット環境では、データベース クエリがクライアント ブラウザに戻す情報の量を制限することにより、多くの利点があります。この例では、スクリプトが ASP と ADO を使用して、戻される行の数を制限する方法を示します。</P>

<H6>コードの説明</H6>

<P>サンプルでは、最初に Connection オブジェクトのインスタンスを作成し、このオブジェクトの Open メソッドで OLE DB 接続を開きます。CreateObject を再度使用して、空の Recordset オブジェクトのインスタンスを作成します。新しい Recordset オブジェクトの ActiveConnection プロパティが、開かれた OLE DB 接続を指すように設定され、SQL ソース文字列が割り当てられ、カーソル タイプが指定されます。戻される結果を制限するのは、Recordset オブジェクトの PageSize プロパティです。この例では、このプロパティの値は 10 に設定されます。つまり、ADO が最大 10 レコードを戻すことを示します。最後に Open メソッドが呼び出され、SQL 検索文字列を満たす最初の 10 レコードを ADO が検索します。</P>

<P>ADO が検索結果を戻して Recordset オブジェクトに配置すると、スクリプトはページをループ処理し、テーブル内の各レコードのすべてのフィールドを表示します。次にスクリプトは一般的な終了操作を行い、レコードセットと接続の両方を閉じます。</P>

<P>SQL クエリが 10 個より多いレコードを戻した場合、このスクリプトではそれらを表示しないことに注意してください。代わりに、余分なレコードは追加の論理ページに保存されます。PageCount プロパティは、データの論理ページが何ページ戻されたかを示します。</P>

<P><span class=le><B>重要&nbsp;&nbsp;&nbsp;</B></span>このサンプルを適切に実行するためには、サーバー上に OLE DB を適切に構成する必要があります。</P>
<hr class="iis" size="1">
<p align="center"><em><a href="../../../common/colegal.htm">&copy; 1997-2001 Microsoft Corporation.All rights reserved.</a></em></p>
</BODY>
</HTML>
