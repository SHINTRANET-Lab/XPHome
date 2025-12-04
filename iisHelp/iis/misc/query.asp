<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML 3.0//EN" "html.dtd">
<HTML DIR=LTR>
<HEAD>

<META HTTP-EQUIV="Content-Type" content="text/html; charset=shift_jis">

<META NAME="ROBOTS" CONTENT="NOINDEX">

<%
' ********** INSTRUCTIONS FOR QUICK CUSTOMIZATION **********
'
' This form is set up for easy customization. It allows you to modify the
' page logo, the page background, the page title and simple query
' parameters by modifying a few files and form variables. The procedures
' to do this are explained below.
'
'
' *** Modifying the Form Logo:

' The logo for the form is named is2logo.gif. To change the page logo, simply
' name your logo is2logo.gif and place in the same directory as this form. If
' your logo is not a GIF file, or you don't want to copy it, change the following
' line so that the logo variable contains the URL to your logo.

        FormLogo = "is2logo.gif"

'
' *** Modifying the Form's background pattern.

' You can use either a background pattern or a background color for your
' form. If you want to use a background pattern, store the file with the name
' is2bkgnd.gif in the same directory as this file and remove the remark character
' the single quote character) from the line below. Then put the remark character on
' the second line below.
'
' If you want to use a different background color than white, simply edit the
' bgcolor line below, replacing white with your color choice.

'       FormBG = "background = " & chr(34) & "is2bkgnd.gif" & chr(34)
        FormBG = "bgcolor = " & chr(34) & "#FFFFFF" & chr(34)


' *** Modifying the Form's Title Text.

' The Form's title text is set on the following line.
%>

    <TITLE>検索結果</TITLE>

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
</script>

<%
'
' *** Modifying the Form's Search Scope.
'
' The form will search from the root of your web server's namespace and below
' (deep from "/" ). To search a subset of your server, for example, maybe just
' a PressReleases directory, modify the scope variable below to list the virtual path to
' search. The search will start at the directory you specify and include all sub-
' directories.

        FormScope = "/iishelp/iis"

'
' *** Modifying the Number of Returned Query Results.
'
' You can set the number of query results returned on a single page
' using the variable below.
				
				'was 10
        PageSize = 10

'
' *** Setting the Locale.
'
' The following line sets the locale used for queries. In most cases, this
' should match the locale of the server. You can set the locale below.

        SiteLocale = "JA"

' ********** END QUICK CUSTOMIZATION SECTIONS ***********
noise=",about,after,all,also,an,another,any,and,are,as,at,be,because,been,before,being,between,both,but,by,came,can,come,could,did,do,each,for,from,get,got,has,had,he,have,her,here,him,himself,his,how,if,in,into,is,it,like,make,many,me,might,more,most,much,must,my,never,near,now,of,on,only,or,other,our,out,over,said,same,see,should,since,some,still,such,take,than,that,the,their,them,then,there,these,they,this,those,through,to,too,under,up,very,was,way,we,well,were,what,where,which,while,who,with,would,you,your,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,$,1,2,3,4,5,6,7,8,9,0,_,!,&,~,|,?,I,"

punc2="$,1234567890_!&~|? #+()"
punc="$,1234567890_!&~|?#@%^+"
%>

<%
' Set Initial Conditions
    NewQuery = FALSE
    UseSavedQuery = FALSE
    rSearchString = ""

' Did the user press a SUBMIT button to execute the form? If so get the form variables.
    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        rSearchString = Request.Form("SearchString")
	SearchType=Request.QueryString("SearchType")
	if SearchString<>"" then
        Session("SearchStringDisplay")=Server.HTMLEncode(rSearchString)
	end if
        rFreeText = Request.Form("FreeText")
	QueryForm = "Query.Asp"
	rCiResultsSize = Request.Form("CiResultsSize")
 	CiLimits = Request.Form("CiLimits")
        ' NOTE: this will be true only if the button is actually pushed.
        ' if Request.Form("Action") = "Search" then
            NewQuery = TRUE
	    if CiLimits = "on" then
			RankBase=50
	    else
			RankBase=1000
	    end if
        ' end if
    end if
    if Request.ServerVariables("REQUEST_METHOD") = "GET" then
        rSearchString = Request.QueryString("SearchString")
	SearchType=Request.QueryString("SearchType")
	QueryForm = "Query.Asp"
	rCiResultsSize = Request.QueryString("CiResultsSize")
                rFreeText = Request.QueryString("FreeText")
                FormScope = Server.HTMLEncode(Request.QueryString("sc"))
				RankBase = CInt(Request.QueryString("RankBase"))
        if Request.QueryString("pg") <> "" then
            NextPageNumber = CInt(Request.QueryString("pg"))
            NewQuery = FALSE
            UseSavedQuery = TRUE
        else
            NewQuery = SearchString <> ""
        end if
    end if

    ' remove HTML special characters, they are ignored in search
    SearchString = replace(replace(replace(rSearchString, "<", ""), ">", ""), """", "")

    ' encode these texts to avoid cross site scripting
    CiResultsSize = Server.HTMLEncode(rCiResultsSize)
    FreeText = Server.HTMLEncode(rFreeText)

%>

</HEAD>

<BODY <%=FormBG%>>

<%
  if NewQuery then
    set Session("Query") = nothing
    set Session("Recordset") = nothing
    NextRecordNumber = 1

	'Strip punctuation from search term
	for x = 1 to len(SearchString)
	   testpunc= mid(SearchString,x,1)
	   if instr(punc,testpunc) then
		SearchStringErr= SearchStringErr
	   else
		SearchStringErr = SearchStringErr + testpunc
	   end if
	next
	SearchString = SearchStringErr


  if SearchType=0 Then
	'Strip noise words from search term
	SearchStringComp=SearchString+" "
	for x = 1 to len(SearchStringComp)
	if mid(SearchStringComp,x,1)=" " Then
		ncompare2 = ","+ncompare+","
		if instr(noise,ncompare2) = 0 then
			NewCompare=NewCompare+" "+ncompare
		End If
		ncompare=""
	else
		ncompare=ncompare+mid(SearchString,x,1)
	end if
	next    	
	x = len(NewCompare)
	if left(NewCompare,1) = " " Then
	   NewCompare = right(NewCompare,(x-1))
	end if
	SearchString=NewCompare
        CompSearch = "$CONTENTS " + SearchString
  end if
 
  if SearchType=1 Then
    CompSearch = chr(34) + SearchString + chr(34)
  end if

  if SearchType=2 Then
	'Strip noise words from search term
	SearchStringComp=SearchString+" "
	for x = 1 to len(SearchStringComp)
	if mid(SearchStringComp,x,1)=" " Then
		ncompare2 = ","+ncompare+","
		if instr(noise,ncompare2) = 0 then
			NewCompare=NewCompare+" "+ncompare
		End If
		ncompare=""
	else
		ncompare=ncompare+mid(SearchString,x,1)
	end if
	next    	
	x = len(NewCompare)
	if left(NewCompare,1) = " " Then
	   NewCompare = right(NewCompare,(x-1))
	end if
	SearchString=NewCompare
	slen=len(SearchString)
	for k = 1 to slen
	slet = Mid(SearchString,k,1)
	  if slet <> " " then
            ss1=ss1+slet
	  else
	    ss1=ss1+ " AND "
	  end if
	Next
        CompSearch=ss1
	If Right(CompSearch,5) = " AND " Then CompSearch = Left(CompSearch,Len(CompSearch)-5)
  end if

 if SearchType=3 Then
	'Strip noise words from search term
	SearchStringComp=SearchString+" "
	for x = 1 to len(SearchStringComp)
	if mid(SearchStringComp,x,1)=" " Then
		ncompare2 = ","+ncompare+","
		if instr(noise,ncompare2) = 0 then
			NewCompare=NewCompare+" "+ncompare
		End If
		ncompare=""
	else
		ncompare=ncompare+mid(SearchString,x,1)
	end if
	next    	
	x = len(NewCompare)
	if left(NewCompare,1) = " " Then
	   NewCompare = right(NewCompare,(x-1))
	end if
	SearchString=NewCompare
	slen=len(SearchString)
	for k = 1 to slen
	slet = Mid(SearchString,k,1)
	  if slet <> " " then
            ss1=ss1+slet
	  else
	    ss1=ss1+ " OR "
	  end if
	Next
        CompSearch=ss1
	If Right(CompSearch,4) = " OR " Then CompSearch = Left(CompSearch,Len(CompSearch)-4)
  end if

  if SearchType=4 Then
	'Strip noise words from search term
	NCompare=""
	NewCompare=""
	SearchStringComp=SearchString+" "
	for x = 1 to len(SearchStringComp)
	if mid(SearchStringComp,x,1)=" " Then
		ncompare2 = ","+ncompare+","
		if instr(noise,ncompare2) = 0 then
			NewCompare=NewCompare+" "+ncompare
		End If
		ncompare=""
	else
		ncompare=ncompare+mid(SearchString,x,1)
	end if
	next    	
	x = len(NewCompare)
	if left(NewCompare,1) = " " Then
	   NewCompare = right(NewCompare,(x-1))
	end if
	SearchString=NewCompare
    CompSearch = SearchString
  end if


    set Q = Server.CreateObject("ixsso.Query")
        set util = Server.CreateObject("ixsso.Util")

    Q.Query = CompSearch
    Q.Catalog = "Web" 
    Q.SortBy = "rank[d]"
    Q.Columns = "DocTitle, vpath, filename, size, write, characterization, rank"
	Q.MaxRecords = RankBase 

        if FormScope <> "/" then
                util.AddScopeToQuery Q, FormScope, "deep"
        end if

        if SiteLocale<>"" then
                Q.LocaleID = util.ISOToLocaleID(SiteLocale)
        end if
    On Error Resume Next
    set RS = Q.CreateRecordSet("nonsequential")

    RS.PageSize = PageSize
    Test = RS.PageSize
    ActiveQuery = TRUE



  elseif UseSavedQuery then
    if IsObject( Session("Query") ) And IsObject( Session("RecordSet") ) then
      set Q = Session("Query")
      set RS = Session("RecordSet")


      if RS.RecordCount <> -1 and NextPageNumber <> -1 then
        RS.AbsolutePage = NextPageNumber
        NextRecordNumber = RS.AbsolutePosition
      end if

      ActiveQuery = TRUE
    else
      Response.Write "エラー : 保存されているクエリーがありません。"
    end if
  end if


If Err<>424 Then

  if ActiveQuery then
    if not RS.EOF then
 %>

<p>
<HR WIDTH=80% ALIGN=center SIZE=3>
<%LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
KLastRecordOnPage=LastRecordOnPage
If KLastRecordOnPage>RS.RecordCount Then KLastRecordOnPage=RS.RecordCount%>


<b>検索文字列 <%=Session("SearchStringDisplay")%></b><br><br> <b><i><font size="3"> <%=RS.RecordCount%> 件の検索結果のうち <%=NextRecordNumber%> から <%=KLastRecordOnPage%> までを表示しています。</b></i></font><br>
<p>

<%
        LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
        CurrentPage = RS.AbsolutePage
        if RS.RecordCount <> -1 AND RS.RecordCount < LastRecordOnPage then
            LastRecordOnPage = RS.RecordCount
        end if

 %>

<%

%>

<%'if Not RS.EOF and NextRecordNumber <= LastRecordOnPage then
	 
	if Not RS.EOF and NextRecordNumber <= LastRecordOnPage then%>
		<table border=0>
<% end if %>

<%

Do While Not RS.EOF and NextRecordNumber <= LastRecordOnPage

        ' This is the detail portion for Title, Description, URL, Size, and
    ' Modification Date.



TmpExt = Server.HTMLEncode( RS("filename") )
FullExt = Right(TmpExt, 3)

If FullExt <> "cnt" and FullExt <> "hhc" and FullExt <> "hpj" and FullExt <> "hlp" and FullExt <> "rtf" and FullExt <> "asf" and FullExt <> "gid" and FullExt <> "fts" and FullExt <> "wmp" and FullExt <> "hhk" and FullExt <> "txt" and FullExt <> "ass" and FullExt <> "idq" and FullExt <> "ncr" and FullExt <> "ncl" and FullExt <> "url" and FullExt <> "css" and FullExt <> "prp" and FullExt <> "htx" and FullExt <> "htw" and FullExt <> "tmp" and FullExt <> "mdb" and FullExt <> "xls" and FullExt <> "chm" Then


    ' If there is a title, display it, otherwise display the filename.
%>
    <p>
	<tr class="RecordTitle">
		
    	        <td><b><%=NextRecordNumber%>.</b></td>
		<b class="RecordTitle"> <td><b><%if VarType(RS("DocTitle")) = 1 or RS("DocTitle") = "" then%> <a href="<%=RS("vpath")%>" class="RecordTitle"><%= Server.HTMLEncode( RS("filename") )%></a> <%else%> <a href="<%=RS("vpath")%>" class="RecordTitle"><%= Server.HTMLEncode(RS("DocTitle"))%></a> <%end if%></b></b><br>
		
			<%if VarType(RS("characterization")) = 8 and RS("characterization") <> "" then%> <%= RS("characterization")%>
		
		<%end if%> <%if CiResultsSize = "on" then%> <%end if%>
		</td>
	</tr>
	<tr>
	</tr>

<%
else
   NextRecordNumber = NextRecordNumber-1
end if%>

<%
          RS.MoveNext
          NextRecordNumber = NextRecordNumber+1
      Loop
 %>

</table>

<P><BR>

<%
  else   ' NOT RS.EOF
      if NextRecordNumber <> 1 then
          Response.Write "クエリーはこれで終わりです。<P>"
      end if

  end if ' NOT RS.EOF%>

<%
  if Q.QueryIncomplete then
'    If the query was not executed because it needed to enumerate to
'    resolve the query instead of using the index, but AllowEnumeration
'    was FALSE, let the user know %>

    <P>
    <I><B>クエリーが完了できませんでした。クエリーを再度送信してください。<BR> 技術的説明: このクエリーを完了するためには、AllowEnumeration を TRUE に設定する必要があります。</B></I><BR> <%end if


  if Q.QueryTimedOut then
'    If the query took too long to execute (for example, if too much work
'    was required to resolve the query), let the user know %>
    <P>
    <I><B>クエリーの実行時間が長すぎました。</B></I><BR> <%end if%>


<TABLE>

<%
'    This is the "previous" button.
'    This retrieves the previous page of documents for the query.
%>

<%SaveQuery = FALSE%> <%if CurrentPage > 1 and RS.RecordCount <> -1 then %>
    <td align=left>
        <form action="<%=QueryForm%>" method="get">
            <INPUT TYPE="HIDDEN" NAME="SearchString" VALUE="<%=SearchString%>">
                        <INPUT TYPE="HIDDEN" NAME="FreeText" VALUE="<%=FreeText%>">
	    <INPUT TYPE="HIDDEN" NAME="CiResultsSize" VALUE="<%=CiResultsSize%>">
            <INPUT TYPE="HIDDEN" NAME="sc" VALUE="<%=FormScope%>">
            <INPUT TYPE="HIDDEN" name="pg" VALUE="<%=CurrentPage-1%>" >
			<INPUT TYPE="HIDDEN" NAME = "RankBase" VALUE="<%=RankBase%>">
            <input type="submit" value="<< 前ページ">
        </form>
    </td>
        <%SaveQuery = TRUE%> <%end if%>

<%
'    This is the "next" button for unsorted queries.
'    This retrieves the next page of documents for the query.

  if Not RS.EOF or NextRecordNumber = 9 then%>
    <td align=right>
        <form action="<%=QueryForm%>" method="get">
            <INPUT TYPE="HIDDEN" NAME="SearchString" VALUE="<%=SearchString%>">
                        <INPUT TYPE="HIDDEN" NAME="FreeText" VALUE="<%=FreeText%>">
	    <INPUT TYPE="HIDDEN" NAME="CiResultsSize" VALUE="<%=CiResultsSize%>">
            <INPUT TYPE="HIDDEN" NAME="sc" VALUE="<%=FormScope%>">
            <INPUT TYPE="HIDDEN" name="pg" VALUE="<%=CurrentPage+1%>" >
			<INPUT TYPE="HIDDEN" NAME = "RankBase" VALUE="<%=RankBase%>">

                <% NextString = "次ページ >>"%>
            <input type="submit" value="<%=NextString%>">
        </form>
    </td>
    <%SaveQuery = TRUE%> <%end if%>

</TABLE>

<% ' Display the page number %> <%if RS.RecordCount = 0 then%>クエリー <%=SearchString%> に一致するドキュメントがありませんでした。<br><br>

	対処方法
	<UL><LI>このエラーについてメールでお知らせください。<a href="mailto:iisdocs@microsoft.com?subject=<%=SearchString%>-search%20term%20not%20matched&body=The%20term%20'<%=SearchString%>'%20produced%20no%20matches.">(mailto:iisdocs@microsoft.com)</a> ただし、英語のメールのみとさせていただきます。
	<LI>関連する用語についてインデックスをチェックしてください。
	<LI>スペルと構文の両方をチェックしてください。
	<LI>ほかの検索オプションを試してください (標準検索、完全一致、いずれかの語を含む、すべての語を含む、および論理検索の各オプションが使用できます)。
	<LI>後で再度クエリーを実行してください。インデックス サービス開始直後の場合、IIS マニュアルのカタログの作成に数分かかります。
	</UL>
    <%else%>

<%=CurrentPage%> <%if RS.PageCount <> -1 then
     Response.Write " / " & RS.PageCount
  end if %> ページ<%end if%>

<%
    ' If either of the previous or back buttons were displayed, save the query
    ' and the recordset in session variables.
    if SaveQuery then
        set Session("Query") = Q
        set Session("RecordSet") = RS
    else
        RS.close
        Set RS = Nothing
        Set Q = Nothing
        set Session("Query") = Nothing
        set Session("RecordSet") = Nothing
    end if
 %> <% end if %>


</BODY>
</HTML>
<%else%> <%


	'Strip noise words from search term
	NCompare=""
	NewCompare=""
	SearchStringComp=SearchString+" "
	for x = 1 to len(SearchStringComp)
	if mid(SearchStringComp,x,1)=" " Then
		ncompare2 = ","+ncompare+","
		if instr(noise,ncompare2) = 0 then
			NewCompare=NewCompare+" "+ncompare
		End If
		ncompare=""
	else
		ncompare=ncompare+mid(SearchString,x,1)
	end if
	next    	
	x = len(NewCompare)
	if left(NewCompare,1) = " " Then
	   NewCompare = right(NewCompare,(x-1))
	end if
	SearchString=NewCompare

	'Strip punctuation from search term
	SearchStringErr = ""
	for x = 1 to len(SearchString)
	   testpunc= mid(SearchString,x,1)
	   if instr(punc2,testpunc) then
		SearchStringErr= SearchStringErr
	   else
		SearchStringErr = SearchStringErr + testpunc
	   end if
	next
	SearchString = SearchStringErr
	CompSearch=SearchString

%>


<%if SearchString = "" or instr(SearchString,"*") or instr(CompSearch,")") or instr(CompSearch,"(") or right(CompSearch,3)="OR " or right(CompSearch,4)="AND " then%> <b>インデックス サービスでクエリーを処理できませんでした。<p></b><br>

クエリーを変えて再実行してください。一般的な単語 ("get"、"for"、および "many" など) には、インデックスのないものがあります。また、句読点 (カンマ、ピリオドなど) はクエリーの中では使用しないでください。

<%else%>

インデックス サービスが開始されていません。<p>
*<%=CompSearch%>* <br>


IIS マニュアルで検索クエリーを実行するには、最初にインデックス サービスを起動する必要があります。<br>
&nbsp;<p>
インデックス サービスを起動するには、以下の操作を実行します。
<ol>
<li>IIS を実行しているコンピュータ上で、[マイ コンピュータ] アイコンを右クリックし、[管理] をクリックします。<p>
<li>MMC の [サービスとアプリケーション] ノードを展開します。<p>
<li>[インデックス サービス] をクリックします。</p>
<li>MMC の [操作] メニューの [開始] をクリックします。<p>
</ol>&nbsp;<p>
注: インデックス サービスで IIS マニュアルの一覧を作成するには、数分かかる場合があります。<p>
マニュアルをリモートで表示しながら検索機能を使用するには、マニュアルがインストールされているコンピュータ上でインデックス サービスを実行する必要があります。インデックス サービスを起動できない場合は、Web サイト管理者に連絡してください。
<%end if%> <% end if %>

