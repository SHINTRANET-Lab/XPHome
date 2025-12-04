<HTML DIR=LTR>
<HEAD>
<TITLE>ASP 検索フォーム</TITLE>

<META NAME="ROBOTS" CONTENT="NOINDEX">

<META HTTP-EQUIV="Content-Type" content="text/html; charset=shift_jis">

<style>
<!--
a:link	 	{color: white; text-decoration:none;}
a:visited 	{color: white; text-decoration:none;}
a:active 	{color: white; text-decoration:none;}
a:hover 	{color: white; text-decoration:underline;}
a		{font-size: 14px; font-family: ＭＳ Ｐゴシック}
-->
</style>

<Script Language="JavaScript">
<!--
function ChangeList(y,z) {

window.location.href="search.asp?Searchset="+(y)+"&SearchString="+(z);

}

//-->
</Script>

<SCRIPT LANGUAGE="VBScript">
<!--
Sub contents_onfocus
	deactivateAll
	contents.childNodes(0).src = "NoCont-active.gif"
End Sub

Sub contents_onblur
	contents.childNodes(0).src = "NoCont.gif"
End Sub

Sub contents_onmouseover
	contents.childNodes(0).src = "NoCont-active.gif"
End Sub

Sub contents_onmouseout
	contents.childNodes(0).src = "Nocont.gif"	
End Sub

Sub index_onfocus
	deactivateAll
	index.childNodes(0).src = "NoIndex-Active.gif"
End Sub

Sub index_onblur
	index.childNodes(0).src = "NoIndex.gif"
End Sub

Sub index_onmouseover
	index.childNodes(0).src = "NoIndex-Active.gif"
End Sub

Sub index_onmouseout
	index.childNodes(0).src = "NoIndex.gif"	
End Sub

sub deactivateAll()
	index.childNodes(0).src = "noindex.gif"
	contents.childNodes(0).src = "Nocont.gif"
end sub

-->
</SCRIPT>

</HEAD>

<BODY bgcolor="#000000" onload="Activate()">
<font face="ＭＳ Ｐゴシック">

<SPAN STYLE="LEFT:  0px; POSITION: relative; TOP: 4px">	<A id="contents" HREF="contents.asp" hidefocus>	<IMG SRC="NoCont.gif" border="0" alt="目次"></A></SPAN>
<SPAN STYLE="LEFT: -4px; POSITION: relative; TOP: 4px">	<A id="index" HREF="index.asp" hidefocus>			<IMG SRC="NoIndex.gif" border="0" alt="索引"></A></SPAN>
<SPAN STYLE="LEFT: -8px; POSITION: relative; TOP: 4px">															<IMG SRC="Search.gif" border="0" alt="検索"></SPAN>

<Script Language="JavaScript">
<!--
function Activate() {
      document.iissrch.SearchString.focus();
}

//-->
</Script>
<TABLE bgcolor="#ffffff" width="262" height="82%" border="0">
<% SearchString=Server.HTMLEncode(Request.QueryString("SearchString"))%> <% If SearchString="undefined" Then SearchString="" %>

<% SearchSet=Server.HTMLEncode(Request.QueryString("SearchSet"))%> <% if SearchSet="" then SearchSet=0%>
<FORM ACTION="Query.asp?SearchType=<%=SearchSet%>" name="iissrch" id="iissrch" target="main" METHOD="POST">
<TR border="0" bgcolor="#ffffff" valign="top"><TD>
<IMG SRC="white.gif"> <font size="-1">検索文字列<br>
<INPUT TYPE="TEXTarea" NAME="SearchString" SIZE="27" MAXLENGTH="100" Value="<% =SearchString%>">
<table>
<tr><td width=65%></td><td>
<INPUT NAME="Action" TYPE="SUBMIT" VALUE="検索"</td></tr><tr><td><font size="-1">検索タイプの選択</font></td></tr></table>


<%If SearchSet=0 Then%>
<SELECT NAME="SearchType" ONCHANGE=ChangeList(SearchType.selectedIndex,SearchString.value)>
<Option Selected=True Value="1">標準検索
<Option Value="2">完全一致
<Option Value="3">すべての語を含む
<Option Value="4">いずれかの語を含む
<Option Value="5">論理検索
</Select>
<%End If%>

<%If SearchSet=1 Then%>
<SELECT NAME="SearchType" ONCHANGE=ChangeList(SearchType.selectedIndex,SearchString.value)>
<Option Value="1">標準検索
<Option Selected=True Value="2">完全一致
<Option Value="3">すべての語を含む
<Option Value="4">いずれかの語を含む
<Option Value="5">論理検索
</Select>
<%End If%>

<%If SearchSet=2 Then%>
<SELECT NAME="SearchType" ONCHANGE=ChangeList(SearchType.selectedIndex,SearchString.value)>
<Option Value="1">標準検索
<Option Value="2">完全一致
<Option Selected=True Value="3">すべての語を含む
<Option Value="4">いずれかの語を含む
<Option Value="5">論理検索
</Select>
<%End If%>

<%If SearchSet=3 Then%>
<SELECT NAME="SearchType" ONCHANGE=ChangeList(SearchType.selectedIndex,SearchString.value)>
<Option Value="1">標準検索
<Option Value="2">完全一致
<Option Value="3">すべての語を含む
<Option Selected=True Value="4">いずれかの語を含む
<Option Value="5">論理検索
</Select>
<%End If%>

<%If SearchSet=4 Then%>
<SELECT NAME="SearchType" ONCHANGE=ChangeList(SearchType.selectedIndex,SearchString.value)>
<Option Value="1">標準検索
<Option Value="2">完全一致
<Option Value="3">すべての語を含む
<Option Value="4">いずれかの語を含む
<Option Selected=True Value="5">論理検索
</Select>
<%End If%>




<%If SearchSet=0 Then%>
<div style="margin-left: -.25in">
<font size="-1">
<ul>
<li>
語句または質問を入力します。
<li>
その単語のすべての語形を含みます。
<li>通常は多数の結果が返ります。
</div>

<br><b>使用例</b>
<div style="margin-left: .17in">
複数のサイトをホストする<br> ディレクトリの権限を設定する<br> IIS の変更点
</div>
</font> <%End If%>

<%If SearchSet=1 Then%>
<div style="margin-left: -.25in">
<font size="-1">
<ul>
<li>
語句をそのままの形で検索します。
<li>
大文字小文字を区別しません。
</div>
<br><b>使用例</b>
<div style="margin-left: .17in">
認証<br> SSL<br> データベース アクセス<br> 接続プーリング
</div>
</font> <%End If%>

<%If SearchSet=2 Then%>
<div style="margin-left: -.25in">
<font size="-1">
<ul>
<li>
語句の順番を問いません。
<li>
通常は少数の結果が返ります。
</div>
<br><b>使用例</b>
<div style="margin-left: .17in">
名前 パスワード アカウント<br> リモート 管理 インターネット<br> レジストリ メタベース 構成<br>
</div>
</font> <%End If%>

<%If SearchSet=3 Then%>
<div style="margin-left: -.25in">
<font size="-1">
<ul>
<li>
語句の出現回数の多いトピックから表示されます。
<li>
通常は多数の結果が返ります。
</div>
<br><b>使用例</b>
<div style="margin-left: .17in">
セキュリティ ハッカー ファイアウォール<br> Web アプリケーション スクリプト ASP<br> ユーザー 権限 アクセス権 拒否<br>
</div>
</font>

<%End If%>

<%If SearchSet=4 Then%>
<div style="margin-left: -.25in">
<font size="-1">
<ul>
<li>
AND、OR、NEAR、および AND NOT の演算子がサポートされます。
<li>
リテラル文字列は、二重引用符 (&quot;) で囲みます。
<li>
語句をグループ化するには、かっこを使用します。
</div>
<br><b>使用例</b>
<div style="margin-left: .17in">
証明書 NEAR インストール<br> "IIS スナップイン" AND 管理<br> (v-root OR virtual) AND (プログラム OR アプリケーション)<br>
</div>
</font>

<%End If%>


<p>


<INPUT TYPE="hidden" NAME="CiResultsSize" value= "on"><br> <BR>

</TD></TR>
</FORM>
</TABLE>

<div align="right" ><A target="main" href="/iishelp/iis/htm/core/NavigationHelp.htm">ナビゲーション ヘルプ</A></div>
</BODY>
</HTML>

