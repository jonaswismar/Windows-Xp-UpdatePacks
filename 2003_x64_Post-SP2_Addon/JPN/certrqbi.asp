﻿<%@ CODEPAGE=65001 'UTF-8%>
<%' certrqbi.asp - (CERT)srv web - (R)e(Q)uest, (B)asic (I)nformation
  ' Copyright (C) Microsoft Corporation, 1998 - 1999 %>
<!-- #include FILE=certsbrt.inc -->
<!-- #include FILE=certdat.inc -->
<!-- #include FILE=certrqtp.inc -->

<!-- Windows Security Update, KB2518295 has replaced some of the CA Web Enrollment ASP files -->
<!-- Please see http://www.support.microsoft.com/kb/2518295 for the back-up location of the previous ASP files -->

<%
	' Strings To Be Localized
	Const L_MoreOptions_Message="クリックすると詳細オプションが表示されます。"
%>
<HTML>
<Head>
	<Meta HTTP-Equiv="Content-Type" Content="text/html; charset=UTF-8">
	<Title>Microsoft 証明書サービス</Title>
</Head>
<Body BgColor=#FFFFFF Link=#0000FF VLink=#0000FF ALink=#0000FF <%If "IE"=sBrowser Then%> OnLoad="postLoad();" <%End If%>><Font ID=locPageFont Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">

<Table Border=0 CellSpacing=0 CellPadding=4 Width=100% BgColor=#008080>
<TR>
	<TD><Font Color=#FFFFFF><LocID ID=locMSCertSrv><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B><I>Microsoft</I></B> 証明書サービス &nbsp;--&nbsp; <%=sServerDisplayName%> &nbsp;</Font></LocID></Font></TD>
	<TD ID=locHomeAlign Align=Right><A Href="/certsrv"><Font Color=#FFFFFF><LocID ID=locHomeLink><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B>ホーム</B></Font></LocID></Font></A></TD>
</TR>
</Table>

<Form Name=UIForm OnSubmit="goNext();return false;" Action="certlynx.asp" Method=Post>
<Input Type=Hidden Name=SourcePage Value="certrqbi">

<P ID=locPageTitle> <B> <%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_FRIENDLYNAME)%> - 識別情報 </B>
<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>

<%If "IE"=sBrowser Then%>
<Span ID=spnFixTxt Style="display:none">
	<Table Border=0 CellSpacing=0 CellPadding=4 Style="Color:#FF0000"><TR><TD><LocID ID=locBadCharError>
		<I><B>赤く</B>表示されているフィールドを修正してください。</I>
		名前フィールドを空白にすることはできません。
		電子メール アドレスは文字 A-Z、a-z、0-9、およびいくつかの共通シンボルを使えますが、拡張文字は使用できません。
		国/地域フィールドには 2 桁の ISO 3166 国/地域コードを使ってください。
	</LocID></TD></TR></Table>
</Span>
<Span ID=spnErrorTxt Style="display:none">
	<Table Border=0 CellSpacing=0 CellPadding=4 Style="Color:#FF0000">
	<TR><TD><LocID ID=locErrMsgBasic>
		証明書の要求を作成しているときに<B>エラーが発生しました</B>。
		正しい CSP を選択したか確認するか、
		または管理者に問い合わせてください。
	</LocID></TD></TR><TR><TD><Span ID=spnErrorDetailsBtn>
		<Table Border=0 CellSpacing=0 CellPadding=0>
		<TR> <TD Width=20></TD><TD>
			<Input ID=locBtnDetails Type=Button Value="詳細 &gt;&gt;" OnClick="showErrorDetails();blur();">
		</TD></TR>
		</Table>
	</Span></TD></TR><TR><TD><Span ID=spnErrorDetails1 Style="display:none">
		<LocID ID=locErrorCause><B>原因:</B></LocID><BR>
		<Span ID=spnErrorMsg></Span>
	</Span></TD></TR><TR>
		<TD><Span ID=spnErrorDetails2 Style="display:none"><LocID ID=locErrorNumber><Font Size=-2>エラー: <Span ID=spnErrorNum></Span></Font></LocID></Span></TD>
	</TR>
	</Table>
</Span>
<%End If%>

<P>
<Table Border=0 CellSpacing=0 CellPadding=0>
	<TR> <!-- establish column widths. -->
		<TD Height=4 Width=<%=L_LabelColWidth_Number%>></TD> <!-- label column, top border -->
		<TD RowSpan=50 Width=4></TD>                         <!-- label spacing column -->
		<TD></TD>                                            <!-- field column -->
	</TR>
	<!-- <TR><TD ColSpan=3 Height=15></TD></TR>-->

<%If "StandAlone"=sServerType Then%>
	<TR>
		<TD ColSpan=3><LocID ID=locInstructions><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial"> 
			次のボックスに要求された情報を入力して、証明書を完了してください。</Font></LocID></TD>
	</TR>
	<TR><TD ColSpan=3 Height=4></TD></TR>
	<TR>
		<TD ID=locNameAlign Align=Right><Span ID=spnNameLabel><LocID ID=locNameLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>名前:</Font></LocID></Span></TD>
		<TD><Input ID=locTbCommonName Type=Text MaxLength=64 Size=42 Name=tbCommonName></TD>
	</TR><TR>
		<TD ID=locEmailAlign Align=Right><Span ID=spnEmailLabel><LocID ID=locEmailLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>電子メール:</Font></LocID></Span></TD>
		<TD><Input ID=locTbEmail Type=Text MaxLength=128 Size=42 Name=tbEmail></TD>
	</TR><TR>
		<TD Height=8></TD> <TD></TD>
	</TR><TR>
		<TD ID=locCompanyAlign Align=Right><Span ID=spnCompanyLabel><LocID ID=locOrgLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>会社:</Font></LocID></Span></TD>
		<TD><Input ID=locTbOrg Type=Text MaxLength=64 Size=42 Name=tbOrg Value="<%=sDefaultCompany%>"></TD>
	</TR><TR>
		<TD ID=locDepartmentAlign Align=Right><Span ID=spnDepartmentLabel><LocID ID=locOrgUnitLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>部署:</Font></LocID></Span></TD>
		<TD><Input ID=locTbOrgUnit Type=Text MaxLength=64 Size=42 Name=tbOrgUnit Value="<%=sDefaultOrgUnit%>"></TD>
	</TR><TR>
		<TD Height=8></TD> <TD></TD>
	</TR><TR>
		<TD ID=locCityAlign Align=Right><Span ID=spnCityLabel><LocID ID=locLocalityLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>市区町村:</Font></LocID></Span></TD>
		<TD><Input ID=locTbLocality Type=Text MaxLength=128 Size=42 Name=tbLocality Value="<%=sDefaultLocality%>"></TD>
	</TR><TR>
		<TD ID=locStateAlign Align=Right><Span ID=spnStateLabel><LocID ID=locStateLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>都道府県:</Font></LocID></Span></TD>
		<TD><Input ID=locTbState Type=Text MaxLength=128 Size=42 Name=tbState Value="<%=sDefaultState%>"></TD>
	</TR><TR>
		<TD ID=locCountryAlign Align=Right><Span ID=spnCountryLabel><LocID ID=locCountryLabel><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>国/地域コード:</Font></LocID></Span></TD>
		<TD><Input ID=locTbCountry Type=Text MaxLength=2 Size=2 Name=tbCountry Value="<%=sDefaultCountry%>"></TD>
	</TR>

<%Else%>
	<TR>
		<TD ID=locReadyToGo ColSpan=3><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
			これ以上の識別情報は必要ありません。
			<%If "IE"=sBrowser Then%><LocID ID=locReadyToGo2>証明書を完了するには [送信] をクリックしてください。</LocID><%End If%></Font></TD>
	</TR>
<%End If%>

<%If "IE"=sBrowser Then%>
	<TR ID=trMoreOptHide><TD Height=12></TD><TD></TD></TR>
	<TR ID=trMoreOptHide>
		<TD><Font Size=-1><Span ID=spnShowMoreOptions tabindex=0 Style="cursor:hand; color:#0000FF; text-decoration:underline;"
				OnContextMenu="return false;"
				OnMouseOver="window.status='<%=L_MoreOptions_Message%>'; return true;" 
				OnMouseOut="window.status=''; return true;" 
				OnKeyDown="if (13==event.keyCode) {showMoreOptions();return false;} else if (9==event.keyCode) {return true;};return false;"
				OnClick="showMoreOptions();return false;">
			<LocID ID=locMoreOpt>詳細オプション &gt;&gt;</LocID></Span></Font>
		</TD>
		<TD></TD>
	</TR>

  <!-- More options -->
	<TR ID=trMoreOptShow Style="display:none">
		<TD ID=locMoreOptHead ColSpan=3><Font Size=-1><BR><B>詳細オプション:</B></Font></TD>
	</TR>
	<TR ID=trMoreOptShow Style="display:none"><TD ColSpan=3 Height=2 BgColor=#008080></TD></TR>
	<TR ID=trMoreOptShow Style="display:none"><TD ColSpan=3 Height=3></TD></TR>

	<TR ID=trMoreOptShow Style="display:none">
		<TD ColSpan=3><Font Face="Arial"><Label For=lbCSPID><LocID ID=locCSPInstr>
			暗号化サービス プロバイダを選択してください:</LocID><Label></Font></TD>
	</TR>

	<TR ID=trMoreOptShow Style="display:none"><TD Height=4></TD> <TD></TD></TR>
	<TR ID=trMoreOptShow Style="display:none">
		<TD ID=locCSPLabel Align=Right><Font Size=-1>CSP:</Font></TD>
		<TD><Select Name=lbCSP ID=lbCSPID>
			<Option ID=locLoading>読み込んでいます...</Option>
			</Select>
		</TD>
	</TR>

	<TR ID=trMoreOptShow Style="display:none"><TD Height=8></TD> <TD></TD></TR>
	<TR ID=trMoreOptShow Style="display:none">
		<TD></TD>
		<TD>
			<Table Border=0 CellSpacing=0 CellPadding=0><TR>
				<TD><Input Type=Checkbox ID=cbStrongKey Name=cbStrongKey></TD>
				<TD><Font Size=-1><Label For=cbStrongKey ID=locStrongKeyLabel>秘密キーの保護を強力にする</Label></Font></TD>
			</TR></Table>
		</TD>
	</TR>

	<TR ID=trMoreOptShow Style="display:none"><TD Height=8></TD> <TD></TD></TR>
	<TR ID=trMoreOptShow Style="display:none">
		<TD ID=locRequestFormatLabel Align=Right><LocID ID=locRequestFormat><Font Size=-1>要求の形式:</Font></LocID></TD>
		<TD>
			<Input Type=Radio ID=rbFormatPKCS10 Name=rbRequestFormat Value="0" Checked><Label For=rbFormatPKCS10 ID=locFormatPKCS10Label>CMC</Label>
			<LocID ID=locSpc5>&nbsp;&nbsp;&nbsp;<LocID>
			<Input Type=Radio ID=rbFormatCMC Name=rbRequestFormat Value="1"><Label For=rbFormatCMC ID=locFormatCMCLabel>PKCS10</Label>
		</TD>
	</TR>

	<TR ID=trMoreOptShow Style="display:none">
		<TD ColSpan=3><LocID ID=locAdvancedLink><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><BR>
			ここにはない詳細オプションが必要な場合は、
			<A Href="certrqma.asp">証明書の要求の詳細設定フォーム</A> を使用してください。</Font></LocID></TD>
	</TR>
  <!-- end More options -->


<%Else '"NN"=sBrowser%>
</Form>
<Form Name=SubmittedData Action="certfnsh.asp" OnSubmit="return goNext();" Method=Post>
	<Input Type=Hidden Name=Mode>             <!-- used in request ('newreq'|'chkpnd') -->
<!--<Input Type=Hidden Name=CertRequest>-->   <!-- used in request -->
	<Input Type=Hidden Name=CertAttrib>       <!-- used in request -->
	<Input Type=Hidden Name=FriendlyType>     <!-- used on pending -->
	<Input Type=Hidden Name=ThumbPrint>       <!-- used on pending -->
	<Input Type=Hidden Name=TargetStoreFlags> <!-- used on install ('0'|CSSLM)-->
	<Input Type=Hidden Name=SaveCert>         <!-- used on install ('no'|'yes')-->


	<TR><TD ColSpan=3 Height=18></TD></TR>
	<TR>
		<TD ID=locStrengthInst ColSpan=3><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
			キーの強度を選択してください:</Font></TD>
	</TR>
	<TR><TD ColSpan=3 Height=3></TD></TR>
	<TR>
		<TD ID=locStrengthLabel Align=Right><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1>キーの強度:</Font></TD>
		<TD><KeyGen Name=CertRequest Challenge="provePequalsNP"></TD>
	</TR>

<%End If%>


	<TR><TD ColSpan=3><Font Size=-1><BR></Font></TD></TR>
	<TR><TD ColSpan=3 Height=2 BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR>
	<TR><TD ColSpan=3 Height=3></TD></TR>
	<TR><TD></TD>
		<TD ID=locSubmitAlign Align=Right>
		<Input ID=locBtnSubmit Type=Submit Name=btnSubmit Value="送信 &gt;" <%If "IE"=sBrowser Then%> Style="width:.75in"<%End If%>>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</TD></TR>
	<TR><TD ColSpan=3 Height=40></TD></TR>

</Table>
<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>
<!-- White HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#FFFFFF><Img Src="certspc.gif" Alt="" Height=5 Width=1></TD></TR></Table>

</Form>
</Font>
<!-- ############################################################ -->
<!-- End of standard text. Scripts follow  -->

<%bIncludeXEnroll=True%>
<%bIncludeGetCspList=True%>
<%bIncludeTemplateCode=True%>
<%bIncludeCheckClientCode=True%>
<!-- #include FILE=certsgcl.inc -->

<%If "IE"=sBrowser Then%>
<!-- IE SPECIFIC: This form we fill in and submit 'by hand'. NN does it differently. -->
<Form Name=SubmittedData Action="certfnsh.asp" Method=Post>
	<Input Type=Hidden Name=Mode>             <!-- used in request ('newreq'|'chkpnd') -->
	<Input Type=Hidden Name=CertRequest>      <!-- used in request -->
	<Input Type=Hidden Name=CertAttrib>       <!-- used in request -->
	<Input Type=Hidden Name=FriendlyType>     <!-- used on pending -->
	<Input Type=Hidden Name=ThumbPrint>       <!-- used on pending -->
	<Input Type=Hidden Name=TargetStoreFlags> <!-- used on install ('0'|CSSLM)-->
	<Input Type=Hidden Name=SaveCert>         <!-- used on install ('no'|'yes')-->
</FORM>
<%End If%>

<Script Language="JavaScript">

	//================================================================
	// PAGE GLOBAL VARIABLES

	//----------------------------------------------------------------
	// Strings to be localized
	var L_StillLoading_ErrorMessage="このページの読み込みがまだ終了していません。数秒待ってから、再実行してください。";
	var L_Generating_Message="要求を生成しています...";
	<%If "IE"=sBrowser Then%>
	;
	var L_CspLoadErrNoneFound_ErrorMessage="CSP の一覧を取得中に予期しないエラーが発生しました:\nCSP が見つかりませんでした。";
	var L_CspLoadErrUnexpected_ErrorMessage="\"CSP の一覧を取得中に予期しないエラー (\"+sErrorNumber+\") が発生しました。\"";
	var L_Waiting_Message="サーバーの応答を待っています...";
	var L_ErrNameUnknown_ErrorMessage="(不明)";
	var L_SugCauseNone_ErrorMessage="なし";
	var L_SugCauseBadCSP_ErrorMessage="選択した CSP は要求を処理できませんでした。別の CSP で実行してください。";
	var L_SugCauseKeysetFull_ErrorMessage="セキュリティ トークンには、コンテナを追加するために利用できる記憶域がありません。";
	var L_SugCauseBadSetting_ErrorMessage="選択した CSP は適用した 1 つ以上の設定をサポートしていません。別の設定または別の CSP を使用して実行してください。";
	var L_SugCauseBadChar_ErrorMessage="無効な文字を入力しました。これは検証したときに確認されていなければならないので、バグをレポートしてください。";
	var L_SugCauseNoProfile_ErrorMessage="ユーザーのプロファイルは一時プロファイルです。";
	var L_SugCauseCancelled_ErrorMessage="ユーザーによって操作が取り消されました。";
	<%Else%>
	;
	var L_BadChars_ErrorMessage="名前フィールドは空白ではなりません。電子メール アドレスは文字 A-Z、a-z、0-9、およびいくつかの共通シンボルを使えますが、拡張文字は使用できません。国/地域フィールドには 2 桁の ISO 3166 国/地域コードを使ってください。";
	<%End If%>


	<%If "IE"=sBrowser Then%>
	// IE is not ready until XEnroll has been loaded
	var g_bOkToSubmit=false;
	<%Else%>
	//  We start with this variable true since it doesn't do anything
	//  for Netscape anyway.
	var g_bOkToSubmit=true;
	<%End If%>
	var g_bSubmitPending=false;

	<%If "IE"=sBrowser Then%>
	//================================================================
	// INITIALIZATION ROUTINES

	//----------------------------------------------------------------
	// IE SPECIFIC: 
	// This contains the functions we want executed immediately after load completes
	function postLoad() {
		// Load an XEnroll object into the page
		loadXEnroll("postLoadPhase2()");
		handleCMCFormat();
	}
	function postLoadPhase2() {
		// continued from above
		var nResult;
		var sCSPList ="<%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_CSPLIST)%>";
		<%If "Enterprise"=sServerType Then%>
			var sUserAgent=navigator.userAgent;
			if (-1 == sUserAgent.indexOf("Windows NT 5.1"))
			{
				var sCSPList ="<%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_CSPLIST2)%>";
			}
		<%End If%>

                if ("" != sCSPList)
                {
                        // get csp from template
			updateCSPListFromStrings(sCSPList);
			nResult = 0;
                }
                else
                {
		        // get the CSP list from local xenroll
		        nResult=GetCSPList();
                }
		if (0!=nResult) {
			handleLoadError(nResult, L_CspLoadErrNoneFound_ErrorMessage, L_CspLoadErrUnexpected_ErrorMessage);
			return;
		}

		// Now we're ready to go
		g_bOkToSubmit=true;
	}

	//----------------------------------------------------------------
	// IE SPECIFIC: handle errors from GetCSPList() and GetTemplateList()
	function handleLoadError(nResult, sNoneFound, sUnexpected) {
		if (-1==nResult) {
			alert(sNoneFound);
		} else {
			var sErrorNumber="0x"+toHex(nResult);
			alert(eval(sUnexpected));
		}
		disableAllControls();
	}

	//================================================================
	// PAGE MANAGEMENT ROUTINES

	//----------------------------------------------------------------
	// IE SPECIFIC: morph method for the error details drop-down
	function showErrorDetails() {
		spnErrorDetailsBtn.style.display='none';
		spnErrorDetails1.style.display='';
		spnErrorDetails2.style.display='';
	}

	//----------------------------------------------------------------
	// IE SPECIFIC: morph method for the "more options" drop down
	function showMoreOptions() {
		var nIndex;
		for (nIndex=0; nIndex<trMoreOptHide.length; nIndex++) { //>
			trMoreOptHide[nIndex].style.display='none';
		}
		for (nIndex=0; nIndex<trMoreOptShow.length; nIndex++) { //>
			trMoreOptShow[nIndex].style.display='';
		}
	}

	//----------------------------------------------------------------
	// handle CMC Format
	function handleCMCFormat() {
		if (!isClientAbleToCreateCMC())
		{
			//no cmc, disable it, only pkcs10
			document.UIForm.rbRequestFormat[0].disabled=true;
			document.UIForm.rbRequestFormat[1].disabled=true;
			document.UIForm.rbRequestFormat[1].checked=true;
		}
	}

	<%End If%>
		
	//================================================================
	// SUBMIT ROUTINES

	//----------------------------------------------------------------
	// determine what to do when the submit button is pressed
	function goNext() {
		if (false==g_bOkToSubmit) {
			alert(L_StillLoading_ErrorMessage);
			return false;
		} else if (true==g_bSubmitPending) {
			// ignore this, as there is UI already.
			return false;
		} else {
			return SubmitRequest();
		}
	}

	<%If "StandAlone"=sServerType Then%>
	//----------------------------------------------------------------
	// check for invalid characters
	var gc_IA5Chars=" !\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKMLNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
	function isValidIA5String(sSource) {
		var nIndex;
		for (nIndex=sSource.length-1; nIndex>=0; nIndex--) {
			//if (sSource.charCodeAt(nIndex)>127) {  // NOTE: this is better, but not compatible with old browsers.
			if (-1==gc_IA5Chars.indexOf(sSource.charAt(nIndex))) {
				return false;
			}
		};
		return true;
	}

	//----------------------------------------------------------------
	// check for invalid characters
	function isValidCountryField(tbCountry) {
		tbCountry.value=tbCountry.value.toUpperCase();
		var sSource=tbCountry.value;
		var nIndex, ch;
		if (0!=sSource.length && 2!=sSource.length) {
			return false;
		}
		for (nIndex=sSource.length-1; nIndex>=0; nIndex--) {
			ch=sSource.charAt(nIndex)
			if (ch<"A" || ch>"Z") {
				return false;
			}
		};
		return true;
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
	// check that the form has data in it
	function validateRequest() {
		<%If "IE"<>sBrowser Then%>
		// work around for NN: label marking does nothing
		var spnNameLabel, spnEmailLabel, spnCompanyLabel, spnDepartmentLabel, spnCityLabel, spnStateLabel, spnCountryLabel;
		<%End If%>

		markLabelNormal(spnNameLabel);
		markLabelNormal(spnEmailLabel);
		markLabelNormal(spnCompanyLabel);
		markLabelNormal(spnDepartmentLabel);
		markLabelNormal(spnCityLabel);
		markLabelNormal(spnStateLabel);
		markLabelNormal(spnCountryLabel);
		
		var bOK=true;
		var fldFocusMe=null;
		// check in 'reverse' order so that focus gets set to last item
		// don't set focus immediately because we'd get funny scrolling effects.
		if (false==isValidCountryField(document.UIForm.tbCountry)) {
			bOK=false;
			fldFocusMe=document.UIForm.tbCountry;
			markLabelError(spnCountryLabel);
		}
		// document.UIForm.tbState.value OK
		// document.UIForm.tbLocality.value OK
		// document.UIForm.tbOrgUnit.value OK
		// document.UIForm.tbOrg.value OK
		if (false==isValidIA5String(document.UIForm.tbEmail.value)
			<%If "1.3.6.1.5.5.7.3.4"=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_OID) Then 'e-mail Protection%>
				|| ""==document.UIForm.tbEmail.value
			<%End If%>
			) {
			bOK=false;
			fldFocusMe=document.UIForm.tbEmail;
			markLabelError(spnEmailLabel);
		}
		if (""==document.UIForm.tbCommonName.value) {
			bOK=false;
			fldFocusMe=document.UIForm.tbCommonName;
			markLabelError(spnNameLabel);
		}

		if (false==bOK) {
			<%If "IE"=sBrowser Then%>
			spnFixTxt.style.display='';
			window.scrollTo(0,0);
			<%Else%>
			alert (L_BadChars_ErrorMessage);
			<%End If%>
			fldFocusMe.focus();
		}

		return bOK;
	}
	<%End If '"StandAlone"=sServerType%>



	<%If "IE"=sBrowser Then%>
	//----------------------------------------------------------------
	// IE SPECIFIC:
	function SubmitRequest() {

		g_bSubmitPending=true;

		spnErrorTxt.style.display='none';
		spnFixTxt.style.display='none';

		<%If "StandAlone"=sServerType Then%>
		// check that the form is filled in
		if (false==validateRequest()) {
			g_bSubmitPending=false;
			return;
		}
		<%End If%>

		// show a nice message since request creation can take a while
		ShowTransientMessage(L_Generating_Message);
	
		// Make the message show up on the screen,
		// then continue with 'SubmitRequest':
		// Pause 1 mS before executing phase 2,
		// so screen will have time to repaint.
		setTimeout("SubmitRequestPhase2();", 10);
	}
	function SubmitRequestPhase2() {
		// continued from above

		// some constants defined in wincrypt.h: (line ~234)
		var CRYPT_EXPORTABLE=1;
		var CRYPT_USER_PROTECTED=2;
		var AT_KEYEXCHANGE=1;
		var AT_SIGNATURE=2;
		var PROV_DSS=3;
		var PROV_DSS_DH=13;
		var XECR_PKCS10_V2_0=1;
		var XECR_CMC=3;

		<%If "StandAlone"=sServerType Then%>
		// set the identifying info
		var sDistinguishedName=""
		if (""!=document.UIForm.tbCountry.value) {
			sDistinguishedName+="C=\""+document.UIForm.tbCountry.value.replace(/"/g, "\"\"")   +"\";";
		}
		if (""!=document.UIForm.tbState.value) {
			sDistinguishedName+="S=\""+document.UIForm.tbState.value.replace(/"/g, "\"\"")     +"\";";
		}
		if (""!=document.UIForm.tbLocality.value) {
			sDistinguishedName+="L=\""+document.UIForm.tbLocality.value.replace(/"/g, "\"\"")  +"\";";
		}
		if (""!=document.UIForm.tbOrg.value) {
			sDistinguishedName+="O=\""+document.UIForm.tbOrg.value.replace(/"/g, "\"\"")       +"\";";
		}
		if (""!=document.UIForm.tbOrgUnit.value) {
			sDistinguishedName+="OU=\""+document.UIForm.tbOrgUnit.value.replace(/"/g, "\"\"")   +"\";";
		}
		if (""!=document.UIForm.tbEmail.value) {
			sDistinguishedName+="E=\""+document.UIForm.tbEmail.value.replace(/"/g, "\"\"")     +"\";";
		}
		if (""!=document.UIForm.tbCommonName.value) {
			sDistinguishedName+="CN=\""+document.UIForm.tbCommonName.value.replace(/"/g, "\"\"")+"\";";
		}
		<%Else%>
		// the distinguished name is not used for enterprise CAs
		var sDistinguishedName="";
		<%End If%>

		// set defaults for values we need on install
		document.SubmittedData.CertAttrib.value="UserAgent:<%=Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))%>\r\n";
		document.SubmittedData.TargetStoreFlags.value=0; // 0=Use default (=user store)
		document.SubmittedData.SaveCert.value="no";
		document.SubmittedData.Mode.value="newreq";
		document.SubmittedData.FriendlyType.value="<%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_FRIENDLYNAME)%>";
		// append the local date to the type
		document.SubmittedData.FriendlyType.value+=" ("+(new Date()).toLocaleString()+")";

		<%If "StandAlone"=sServerType Then%>

		// set the cert type information
		var sCertUsage="<%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_OID)%>";

		<%Else%>

		// set the cert template, we know this is v1 template
		var XECT_EXTENSION_V1=1;
		XEnroll.addCertTypeToRequestEx(XECT_EXTENSION_V1, "<%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_TEMPLATE)%>", 0, false, 0);

		var sCertUsage=""; // ignored

		<%End If%>
		
		// set the CSP
		var nCSPIndex=document.UIForm.lbCSP.selectedIndex;
		XEnroll.ProviderName=document.UIForm.lbCSP.options[nCSPIndex].text;
		var nProvType=document.UIForm.lbCSP.options[nCSPIndex].value
		XEnroll.ProviderType=nProvType;

		// default to exchange keys, unless we're doing DSS which only does sig.
		if (PROV_DSS==nProvType || PROV_DSS_DH==nProvType) {
			XEnroll.KeySpec=AT_SIGNATURE;
		} else {
			XEnroll.KeySpec=AT_KEYEXCHANGE;
		}

		// set 'Strong private key protection'
		if (document.UIForm.cbStrongKey.checked) {
			XEnroll.GenKeyFlags|=CRYPT_USER_PROTECTED;
		}
		<% If "Enterprise"=sServerType Then%>
			if ("True"=="<%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_EXPORTABLE)%>")
			{
				XEnroll.GenKeyFlags|=CRYPT_EXPORTABLE;
			}
		<%End If%>

		// set request format
		lRequestFlag=XECR_CMC;
		if (document.UIForm.rbRequestFormat[1].checked) {
			lRequestFlag=XECR_PKCS10_V2_0;
		}

		// build the certificate request
		var nResult=CreateRequest(lRequestFlag, sDistinguishedName, sCertUsage); // ask VB to do it, since it can handle errors

		if (0 == nResult)
		{
			//always get thumbprint in case of pending
			document.SubmittedData.ThumbPrint.value=XEnroll.ThumbPrint;
		}

		// hide the message box
		HideTransientMessage();

		//see if it was cancelled
		if (document.UIForm.cbStrongKey.checked && (0==(0x8010006e^nResult)))
		{
			//ERROR_CANCELLED, likely from dialog, out
			g_bSubmitPending=false;
			XEnroll.reset();
			return;
		}

		// deal with an error if there was one
		if (0!=nResult) {
			handleError(nResult);
			g_bSubmitPending=false;
			return;
		}

		// put up a new wait message
		ShowTransientMessage(L_Waiting_Message);

		// Submit the cert request and move forward in the wizard
		document.SubmittedData.submit();
	}

	//----------------------------------------------------------------
	// IE SPECIFIC:
	function handleError(nResult) {
		var sSugCause=L_SugCauseNone_ErrorMessage;
		var sErrorName=L_ErrNameUnknown_ErrorMessage;
		// analyze the error - funny use of XOR ('^') because obvious choice '==' doesn't work
		if (0==(0x80090008^nResult)) {
			sErrorName="NTE_BAD_ALGID";
			sSugCause=L_SugCauseBadCSP_ErrorMessage;
		} else if (0==(0x80090016^nResult)) {
			sErrorName="NTE_BAD_KEYSET";
			sSugCause=L_SugCauseBadCSP_ErrorMessage;
		} else if (0==(0x80090019^nResult)) {
			sErrorName="NTE_KEYSET_NOT_DEF";
			sSugCause=L_SugCauseBadCSP_ErrorMessage;
		} else if (0==(0x80090020^nResult)) {
			sErrorName="NTE_FAIL";
			sSugCause=L_SugCauseBadCSP_ErrorMessage;
		} else if (0==(0x80090023^nResult)) {
			sErrorName="NTE_TOKEN_KEYSET_STORAGE_FULL";
			sSugCause=L_SugCauseKeysetFull_ErrorMessage;
		} else if (0==(0x80090009^nResult)) {
			sErrorName="NTE_BAD_FLAGS";
			sSugCause=L_SugCauseBadSetting_ErrorMessage;
		} else if (0==(0x80092002^nResult)) {
			sErrorName="CRYPT_E_BAD_ENCODE";
			//sSugCause="";
		} else if (0==(0x80092022^nResult)) {
			sErrorName="CRYPT_E_INVALID_IA5_STRING";
			sSugCause=L_SugCauseBadChar_ErrorMessage;
		} else if (0==(0x80092023^nResult)) {
			sErrorName="CRYPT_E_INVALID_X500_STRING";
			sSugCause=L_SugCauseBadChar_ErrorMessage;
		} else if (0==(0x80090024^nResult)) {
			sErrorName = "NTE_TEMPORARY_PROFILE";
			sSugCause = L_SugCauseNoProfile_ErrorMessage;
		} else if (0==(0x800704C7^nResult)) { 
			sErrorName = "ERROR_CANCELLED";
			sSugCause = L_SugCauseCancelled_ErrorMessage;
		} else if (0==(0x8000FFFF^nResult)) {
			sErrorName="E_UNEXPECTED";
		}
		
		var sErrorNum="0x"+toHex(nResult)+" - "+sErrorName;

		// modify the document text and appearance to show the error message
		spnErrorNum.innerText=sErrorNum;
		spnErrorMsg.innerText=sSugCause;
		spnFixTxt.style.display='none';
		spnErrorTxt.style.display='';

		// back to the top so the messages show
		window.scrollTo(0,0);

		// reset XEnroll so the user can select a different CSP, etc.
		XEnroll.reset();
	}

	<%Else '"NN"=sBrowser%>

	//----------------------------------------------------------------
	// NN SPECIFIC:
	function SubmitRequest() {

		<%If "StandAlone"=sServerType Then%>
		// check that the form is filled in
		if (false==validateRequest()) {
			return false;
		}
		<%End If%>

		ShowTransientMessage(L_Generating_Message);
	
		// set defaults for values we need on install
		var sAttrib="challenge: provePequalsNP\r\n";
		<%If "StandAlone"=sServerType Then%>
		if (""!=document.UIForm.tbCountry.value) {
			sAttrib+=   "country: "+document.UIForm.tbCountry.value   +"\r\n";
		}
		if (""!=document.UIForm.tbState.value) {
			sAttrib+=     "state: "+document.UIForm.tbState.value     +"\r\n";
		}
		if (""!=document.UIForm.tbLocality.value) {
			sAttrib+=  "locality: "+document.UIForm.tbLocality.value  +"\r\n";
		}
		if (""!=document.UIForm.tbOrg.value) {
			sAttrib+=       "org: "+document.UIForm.tbOrg.value       +"\r\n";
		}
		if (""!=document.UIForm.tbOrgUnit.value) {
			sAttrib+=   "orgunit: "+document.UIForm.tbOrgUnit.value   +"\r\n";
		}
		if (""!=document.UIForm.tbEmail.value) {
			sAttrib+=     "email: "+document.UIForm.tbEmail.value     +"\r\n";
		}
		if (""!=document.UIForm.tbCommonName.value) {
			sAttrib+="commonname: "+document.UIForm.tbCommonName.value+"\r\n";
		}
		<%End If%>
		<%If "StandAlone"=sServerType Then%>
		sAttrib+="CertificateUsage:	<%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_OID)%>\r\n";
		<%Else%>
		sAttrib+="CertificateTemplate: <%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_TEMPLATE)%>\r\n";
			<%End If%>
		sAttrib+="UserAgentString: <%=Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))%>\r\n";

		document.SubmittedData.CertAttrib.value=sAttrib;

		document.SubmittedData.TargetStoreFlags.value=0; // 0=Use default (=user store), but ignored by Netscape
		document.SubmittedData.SaveCert.value="no";
		document.SubmittedData.Mode.value="newreq NN";
		document.SubmittedData.FriendlyType.value="<%=rgAvailReqTypes(CInt(Request.QueryString("type")), FIELD_FRIENDLYNAME)%>";
		// append the local date to the type
		document.SubmittedData.FriendlyType.value+=" ("+(new Date()).toLocaleString()+")";

		// keygen and submit
		return true;
	}

	<%End If%>

</Script> 

<%If "IE"=sBrowser Then%>
<Script Language="VBSCRIPT">
	'-----------------------------------------------------------------
	' IE SPECIFIC:
	' call XEnroll to create a request, since javascript has no error handling
	Function CreateRequest(lFlags, sDistinguishedName, sCertUsage)
		On Error Resume Next
		XEnroll.ReuseHardwareKeyIfUnableToGenNew=False
		document.SubmittedData.CertRequest.value= _
			XEnroll.CreateRequest(lFlags, sDistinguishedName, sCertUsage)
		CreateRequest=Err.Number
	End Function
</Script> 
<%End If%>

</Body>
</HTML>
