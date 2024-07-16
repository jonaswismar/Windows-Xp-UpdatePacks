﻿<%@ CODEPAGE=65001 'UTF-8%>
<%' certrqxt.asp - (CERT)srv web - (R)e(Q)uest, e(XT)ernally created
  ' Copyright (C) Microsoft Corporation, 1998 - 1999 %>
<!-- #include FILE=certsbrt.inc -->
<!-- #include FILE=certdat.inc -->

<!-- Windows Security Update, KB2518295 has replaced some of the CA Web Enrollment ASP files -->
<!-- Please see http://www.support.microsoft.com/kb/2518295 for the back-up location of the previous ASP files -->

<%
	Dim sBrowserDependentLineBreak
	If "Text"<>sBrowser Then
		sBrowserDependentLineBreak="<BR>"
	Else
		sBrowserDependentLineBreak=""
	End If
%>
<HTML>
<Head>
	<Meta HTTP-Equiv="Content-Type" Content="text/html; charset=UTF-8">
	<Title>Microsoft 証明書サービス</Title>
</Head>
<Body BgColor=#FFFFFF Link=#0000FF VLink=#0000FF ALink=#0000FF OnLoad="postLoad();"><Font ID=locPageFont Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">

<Table Border=0 CellSpacing=0 CellPadding=4 Width=100% BgColor=#008080>
<TR>
	<TD><Font Color=#FFFFFF><LocID ID=locMSCertSrv><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B><I>Microsoft</I></B> 証明書サービス &nbsp;--&nbsp; <%=sServerDisplayName%> &nbsp;</Font></LocID></Font></TD>
	<TD ID=locHomeAlign Align=Right><A Href="/certsrv"><Font Color=#FFFFFF><LocID ID=locHomeLink><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B>ホーム</B></Font></LocID></Font></A></TD>
</TR>
</Table>

<Form Name=UIForm OnSubmit="goNext();return false;" Action="certlynx.asp" Method=Post>
<Input Type=Hidden Name=SourcePage Value="certrqxt">

<P><LocID ID=locPageTitle> <B> 証明書の要求または更新要求の送信 </B></LocID>
<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>

<%If "IE"=sBrowser Then%>
<Span ID=spnFixTxt Style="display:none">
	<Table Border=0 CellSpacing=0 CellPadding=4 Style="Color:#FF0000"><TR><TD><LocID ID=locBlankError>
		<I>要求フィールドに何か入力してください。</I>フィールドに要求を貼り付けてから再実行してください。
	</TD></LocID></TR></Table>
</Span>
<%End If%>

<P><LocID ID=locInstructions>
CA に保存された要求を送信するには、外部ソース (Web サーバーなど) によって生成された Base 64 エンコード CMD または PKCS #10 証明書の要求または PKCS #7 の更新の要求を、保存されている要求ボックスに貼り付けます。</LocID>
</P>

<Table Border=0 CellSpacing=0 CellPadding=0>
	<TR> <!-- establish column widths. -->
		<TD><Img Src="certspc.gif" Alt="" Height=1 Width=<%=L_LabelColWidth_Number%>></TD> <!-- label column, top border -->
		<TD RowSpan=59><Img Src="certspc.gif" Alt="" Height=1 Width=4></TD>                <!-- label spacing column -->
		<TD></TD>                                                                          <!-- field column -->
	</TR>
	
	<TR>
		<TD ColSpan=3><Font Face="Arial" Size=-1><Label For=locTaRequest><LocID ID=locSavedReqHead><B>保存された要求:</B></LocID></Label></Font></TD>
	</TR><TR><TD ColSpan=3 BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD>
	</TR><TR><TD ColSpan=3><Img Src="certspc.gif" Alt="" Height=3 Width=1></TD></TR>
		
	<TR>
		<TD Align=Left><Span ID=spnPasteLabel><LocID ID=SavedReqLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>Base 64 エンコード 
		<%=sBrowserDependentLineBreak%> 証明書要求
		<%=sBrowserDependentLineBreak%> (CMC または
		<%=sBrowserDependentLineBreak%>  PKCS&nbsp;#10 または
		<%=sBrowserDependentLineBreak%>  PKCS&nbsp;#7):</Font></LocID></Span></TD>
		<TD><TextArea ID=locTaRequest Rows=6 Cols=40 Name=taRequest Wrap=Off></TextArea></TD>
	</TR><TR><TD ColSpan=3 Height=3></TD>
	</TR><TR><TD></TD>
		<TD><%If "IE"=sBrowser Then%>
		<LocID ID=locBrowse><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><Span tabindex=0 Style="cursor:hand; color:#0000FF; text-decoration:underline;"
			OnContextMenu="return false;"
			OnKeyDown="if (13==event.keyCode) {BeginRead();blur();return false;} else if (9==event.keyCode) {return true;};return false;"
			OnClick="BeginRead();blur();return false;"
			OnMouseOver="window.status=L_BrowseLink_Message;return true;" 
			OnMouseOut="window.status='';return true;">挿入するファイルを参照してください</Span>。
		</Font></LocID>
		<Span ID=spRead Style="display:none">
		<Table Border=0 CellSpacing=0 CellPadding=0>
		<TR><TD Height=5></TD>
		<TR>
			<TD Width=6></TD>
			<TD Width=3 BgColor=#008080></TD>
			<TD Width=4></TD>
			<TD>
				<LocID ID=locFileNameLabel>完全なパス名:</LocID> <Input ID=locFlRequest Type=File Size=40 Name=flRequest><BR>
				<Input ID=locBtnRead Type=Button Value="読み込み" onClick="FinishRead();blur();" Style="font-weight:bold">
				<Input ID=locBtnCancel Type=Button Value="キャンセル" onClick="spRead.style.display='none';blur();">

			</TD>
		</TR>
		</Table>
		</Span>
		<%End If%></TD>
	</TR>	 

	<%If "Enterprise"=sServerType Then%>
	<TR>
		<TD ColSpan=3><LocID ID=locCertTmplFont><Font Face="Arial" Size=-1><%If "Text"=sBrowser Then%><P><%Else%><BR><%End If%><Label For=lbCertTemplateID><LocID ID=locTemplateHead><B>証明書テンプレート:</B></LocID></Label></Font></LocID></TD>
	</TR><TR><TD ColSpan=3 BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD>
	</TR><TR><TD ColSpan=3><Img Src="certspc.gif" Alt="" Height=3 Width=1></TD>
	</TR><TR><TD></TD>
		<TD><Select Name=lbCertTemplate ID=lbCertTemplateID>
<%
	Dim nWriteTemplateResult
	nWriteTemplateResult=WriteTemplateList() 
%>
		</Select></TD>
	</TR>	 
	<%End If%>

	<TR>
		<TD ColSpan=3><LocID ID=locAttrFont><Font Face="Arial" Size=-1><%If "Text"=sBrowser Then%><P><%Else%><BR><%End If%><Label For=locTaAttrib><LocID ID=locAttribHead><B>追加属性:</B></LocID></Label></Font></LocID></TD>
	</TR><TR><TD ColSpan=3 BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD>
	</TR><TR><TD ColSpan=3><Img Src="certspc.gif" Alt="" Height=6 Width=1></TD>
	</TR>

	<TR>
		<TD Align=Right><LocID ID=locAttribLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>属性:</Font></LocID></TD>
		<TD><TextArea ID=locTaAttrib Name=taAttrib Wrap=Off Rows=2 Cols=30></TextArea></TD>
	</TR>

<%If "StandAlone"<>sServerType And 0<>nWriteTemplateResult Then%>
<!-- submit button removed if there was an error getting the templates -->
<%Else%>
	<TR><TD ColSpan=3><Font Size=-1><BR></Font></TD></TR>
	<TR><TD ColSpan=3 BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR>
	<TR><TD ColSpan=3><Img Src="certspc.gif" Alt="" Height=3 Width=1></TD></TR>

	<TR><TD><TD Align=Right><LocID ID=locSubmitAlign>
		<Input Type=Submit ID=btnSubmit Value="送信 &gt;" <%If "IE"=sBrowser Then%> Style="width:.75in"<%End If%>>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</LocID></TD></TR>
	<TR><TD ColSpan=3 Height=20></TD></TR>
<%End If%>

</Table>
<P>


<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>
<!-- White HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#FFFFFF><Img Src="certspc.gif" Alt="" Height=5 Width=1></TD></TR></Table>

</Form>
</Font>
<!-- ############################################################ -->
<!-- End of standard text. Scripts follow  -->
	
<%bIncludeXEnroll=False%>
<%bIncludeGetCspList=False%>
<%bIncludeTemplateCode=True%>
<!-- #include FILE=certsgcl.inc -->

<!-- This form we fill in and submit 'by hand'-->
<Form Name=SubmittedData Action="certfnsh.asp" Method=Post>
	<Input Type=Hidden Name=Mode>             <!-- used in request ('newreq'|'chkpnd') -->
	<Input Type=Hidden Name=CertRequest>      <!-- used in request -->
	<Input Type=Hidden Name=CertAttrib>       <!-- used in request -->
	<Input Type=Hidden Name=FriendlyType>     <!-- used on pending -->
	<Input Type=Hidden Name=ThumbPrint>       <!-- used on pending -->
	<Input Type=Hidden Name=TargetStoreFlags> <!-- used on install ('0'|CSSLM)-->
	<Input Type=Hidden Name=SaveCert>         <!-- used on install ('no'|'yes')-->
</FORM>

<Script Language="JavaScript">

	//----------------------------------------------------------------
	// Strings to be localized
	<%If "IE"=sBrowser Then%>
	;
	var L_BrowseLink_Message="挿入するファイルの参照";
	var L_ReadProhibited_ErrorMessage="Web ブラウザのセキュリティ設定は、このページがディスクにアクセスするのを制限しています。\n手動でデータを貼り付けるか、ブラウザの信頼しているサイトにこのページを追加してください。";
	var L_Unexpected_ErrorMessage="\"ファイルを読み取ろうとしているときに予期しないエラーが発生しました。\\n\\nエラー: \"+nResult";
	var L_FileNotFound_ErrorMessage="指定したファイルが見つからないか、指定したドライブの準備ができていません。\n有効なファイル名を入力してください。";
	var L_NoName_ErrorMessage="ファイル名を入力してください。";
	<%End If%>
	<%If "StandAlone"<>sServerType Then%>
	;
	var L_TemplateLoadErrNoneFound_ErrorMessage="証明書のテンプレートは見つかりませんでした。証明書をこの CA から要求するアクセス許可がありません、または Active Directory にアクセス中にエラーが発生しました。";
	var L_TemplateLoadErrUnexpected_ErrorMessage="\"証明書のテンプレートの一覧を取得中に予期しないエラー (\"+sErrorNumber+\") が発生しました。\"";
	<%End If%>
	;
	var L_NoBlank_ErrorMessage="要求フィールドを空白のままにしないでください。\nフィールドに要求を貼り付けてから再実行してください。";
	var L_SavedReqCert_Text="保存された要求証明書";

	//================================================================
	// INITIALIZATION ROUTINES

	//----------------------------------------------------------------
	// This contains the functions we want executed immediately after load completes
	function postLoad() {
		<%If "StandAlone"<>sServerType And 0<>nWriteTemplateResult Then%>
		handleLoadError(<%=nWriteTemplateResult%>, L_TemplateLoadErrNoneFound_ErrorMessage, L_TemplateLoadErrUnexpected_ErrorMessage);
		<%End If%>
	}

	<%If "StandAlone"<>sServerType Then%>
	//----------------------------------------------------------------
	// handle errors from GetTemplateList()
	function handleLoadError(nResult, sNoneFound, sUnexpected) {
		if (-1==nResult) {
			alert(sNoneFound);
		} else {
			var sErrorNumber="0x"+toHex(nResult);
			alert(eval(sUnexpected));
		}
		//document.UIForm.btnSubmit.disabled=true;
	}
	<%End If%>


	<%If "IE"=sBrowser Then%>
	//================================================================
	// FILE READ ROUTINES

	//----------------------------------------------------------------
	// IE SPECIFIC:
	// make sure that we have permision to do a read, then show 
	// the file name box
	function BeginRead() {
		if (true==TestRead()) {
			spRead.style.display='';
			document.UIForm.flRequest.focus()
		} else {
			alert(L_ReadProhibited_ErrorMessage);
		}
	}

	//----------------------------------------------------------------
	// IE SPECIFIC:
	function FinishRead() {
		spnFixTxt.style.display='none';
		if (""==document.UIForm.flRequest.value) {
			handleReadError(5);
			return;
		}
		var nResult=GetFileData(); // use VBScript to read the file, since it can handle errors
		if (0!=nResult) {
			handleReadError(nResult);
			return;
		}
		spRead.style.display='none';
		document.UIForm.btnSubmit.focus()
	}

	//----------------------------------------------------------------
	// IE SPECIFIC:
	function handleReadError(nResult) {
		var sMessage;
		var elemFocusMe=null;
		if (429==nResult) {
			sMessage=L_ReadProhibited_ErrorMessage;
			elemFocusMe=document.UIForm.flRequest;
		} else if (53==nResult || 76==nResult || 71==nResult) {
			sMessage=L_FileNotFound_ErrorMessage;
			elemFocusMe=document.UIForm.flRequest;
		} else if (5==nResult) {
			sMessage=L_NoName_ErrorMessage;
			elemFocusMe=document.UIForm.flRequest;
		} else {
			sMessage=eval(L_Unexpected_ErrorMessage);
		}
		
		// Show the error message
		alert(sMessage);

		// place focus on offending control
		if (null!=elemFocusMe) {
			elemFocusMe.focus();
		}
	}
	<%End If%>

	//================================================================
	// SUBMIT ROUTINES

	//----------------------------------------------------------------
	// determine what to do when the submit button is pressed
	function goNext() {
		SubmitRequest();
	}

	//----------------------------------------------------------------
	// set a label to normal style
	function markLabelNormal(spn) {
		<%If "IE"=sBrowser Then%>
		spn.style.color="#000000";
		spn.style.fontWeight='normal';
		<%End If%>
	}

	//----------------------------------------------------------------
	// set a label to error state
	function markLabelError(spn) {
		<%If "IE"=sBrowser Then%>
		spn.style.color='#FF0000';
		spn.style.fontWeight='bold';
		<%End If%>
	}

	//----------------------------------------------------------------
	function validateRequest() {
		<%If "IE"<>sBrowser Then%>
		// work around for NN: label marking does nothing
		var spnPasteLabel;
		<%End If%>

		markLabelNormal(spnPasteLabel);
				
		// Check for an empty request
		if (""==document.UIForm.taRequest.value) {
			bOK=false;
			markLabelError(spnPasteLabel);
			<%If "IE"=sBrowser Then%>
			spnFixTxt.style.display='';
			window.scrollTo(0,0);
			<%Else%>
			alert(L_NoBlank_ErrorMessage); 
			<%End If%>
			document.UIForm.taRequest.focus();
			return false;
		}

		// everything is OK
		return true;
	}

	//----------------------------------------------------------------
	function SubmitRequest() {

		<%If "IE"=sBrowser Then%>
		spnFixTxt.style.display='none';
		<%End If%>

		// check that the form is filled in
		if (false==validateRequest()) {
			return;
		}

		// set request
		document.SubmittedData.CertRequest.value=document.UIForm.taRequest.value;

		// set defaults for values we need on install
		document.SubmittedData.TargetStoreFlags.value=0; // 0=Use default (=user store), but ignored when saving cert.
		document.SubmittedData.SaveCert.value="yes";
		document.SubmittedData.Mode.value="newreq";
		document.SubmittedData.FriendlyType.value=L_SavedReqCert_Text;
		// append the local date to the type
		document.SubmittedData.FriendlyType.value+=" ("+(new Date()).toLocaleString()+")";
		//not created by xenroll, not supported
		document.SubmittedData.ThumbPrint.value="";

		// make sure the arributes end cr/lf
		var sAttrib=document.UIForm.taAttrib.value;
		if (sAttrib.lastIndexOf("\r\n")!=sAttrib.length-2 && sAttrib.length>0) {
			sAttrib=sAttrib+"\r\n";
		}

		<%If "Enterprise"=sServerType Then%>
		// add an attribute for the cert type

		// get the selected template
		var sRealName = getTemplateStringInfo(CTINFO_INDEX_REALNAME, null);

		// set the cert template
		sAttrib+="CertificateTemplate:"+sRealName+"\r\n";

		<%End If%>

		// for interop debug purposes
		sAttrib+="UserAgent:<%=Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))%>\r\n";

		// set the attributes
		document.SubmittedData.CertAttrib.value=sAttrib;

		// Submit the cert request and move forward in the wizard
		document.SubmittedData.submit();
	}

</Script>

<%If "IE"=sBrowser Then%>
<Script Language="VBSCRIPT">

	'=================================================================
	' FILE READ ROUTINES

	'-----------------------------------------------------------------
	' IE SPECIFIC:
	' See if we have permision to access the file system
	Function TestRead()
		Dim filesystem
		On Error Resume Next
		
		' See if we're allowed to create the FileSystem object
		Set filesystem=CreateObject("Scripting.FileSystemObject")
		' Security may not allow this
		If Err.Number<>0 Then
			TestRead=False
		Else
			TestRead=True
		End If
	End Function

	'-----------------------------------------------------------------
	' IE SPECIFIC:
	' read the given file into the text-area
	Function GetFileData()
		Dim filesystem, file
		On Error Resume Next
		
		' First, create the FileSystem object
		Set filesystem=CreateObject("Scripting.FileSystemObject")
		' Security may not allow this
		If Err.Number<>0 Then
			GetFileData=Err.Number
			Exit Function
		End If
	
		' open the specified file	
		Set file=filesystem.OpenTextFile(document.UIForm.flRequest.value, 1 , false) '1->ForReading, false->don't create
		' file may not exist
		If Err.Number<>0 Then
			GetFileData=Err.Number
			Exit Function
		End If
		
		' read the data and stash it into the form
		document.UIForm.taRequest.value=file.ReadAll
		' catch any read errors
		If Err.Number<>0 Then
			GetFileData=Err.Number
			Exit Function
		End If
		
		' clean up
		file.Close
		Set file=Nothing
		Set filesystem=Nothing
		GetFileData=0
	End Function

</Script> 
<%End If '"IE"=sBrowser%>

</Body>
</HTML>