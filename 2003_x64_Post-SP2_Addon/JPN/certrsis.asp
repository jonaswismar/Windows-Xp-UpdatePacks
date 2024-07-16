<%@ CODEPAGE=65001 'UTF-8%>
<%' certrsis.asp - (CERT)srv web - (R)e(S)ult: (IS)sued
  ' Copyright (C) Microsoft Corporation, 1998 - 1999 %>
<!-- #include FILE=certsbrt.inc -->
<!-- #include FILE=certdat.inc -->
<!-- #include FILE=certsrck.inc -->
<!-- #include FILE=certcert.inc -->

<!-- Windows Security Update, KB2518295 has replaced some of the CA Web Enrollment ASP files -->
<!-- Please see http://www.support.microsoft.com/kb/2518295 for the back-up location of the previous ASP files -->

<%  ' came from certfnsh.asp

    Const CRYPT_STRING_ANY      = &H00000007
    Const FR_PROP_FULLRESPONSE  = &H00000001
    	
	Set ICertRequest=Session("ICertRequest")
	sMode=Request.Form("Mode")

	If "IE"=sBrowser And "no"=Request.Form("SaveCert") Then
		' get the cert chain and save in on this page so the client can install it
		Public sPKCS7
		Dim sCertificate
		sCertificate=ICertRequest.GetFullResponseProperty(FR_PROP_FULLRESPONSE, 0, PROPTYPE_BINARY, CR_OUT_BASE64)
		sPKCS7=FormatBigString(sCertificate, "	sPKCS7=sPKCS7 & ")
	End If

	'-----------------------------------------------------------------
	' Format the big string as a concatenated VB string, breaking at the embedded newlines
	Function FormatBigString(sSource, sLinePrefix)
		Dim sResult, bCharsLeft, nStartChar, nStopChar, chQuote
		sResult=""
		chQuote=chr(34)
		bCharsLeft=True
		nStopChar=1

		While (bCharsLeft)
			nStartChar=nStopChar
			nStopChar=InStr(nStopChar, sSource, vbNewLine)

			If (nStopChar>0) Then
				sResult=sResult & sLinePrefix & chQuote & Mid(sSource, nStartChar, nStopChar-nStartChar) & chQuote & " & vbNewLine"

				If (nStopChar>=Len(sSource)-Len(vbNewLine)) Then
					bCharsLeft=False
				End If

			Else
				bCharsLeft=False
			End if
			sResult=sResult & vbNewLine
			nStopChar=nStopChar+Len(vbNewLine)
		Wend
		FormatBigString=sResult
	End Function
%>
<HTML>
<Head>
	<Meta HTTP-Equiv="Content-Type" Content="text/html; charset=UTF-8">
	<Title>Microsoft Active Directory 証明書サービス</Title>
</Head>
<Body BgColor=#FFFFFF Link=#0000FF VLink=#0000FF ALink=#0000FF <%If "IE"=sBrowser Then%> OnLoad="postLoad();" <%End If%>><Font ID=locPageFont Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">

<Table Border=0 CellSpacing=0 CellPadding=4 Width=100% BgColor=#008080>
<TR>
	<TD><Font Color=#FFFFFF><LocID ID=locMSCertSrv><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B><I>Microsoft</I></B> Active Directory 証明書サービス &nbsp;--&nbsp; <%=sServerDisplayName%> &nbsp;</Font></LocID></Font></TD>
	<TD ID=locHomeAlign Align=Right><A Href="/certsrv"><Font Color=#FFFFFF><LocID ID=locHomeLink><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B>ホーム</B></Font></LocID></Font></A></TD>
</TR>
</Table>

<P ID=locPageTitle> <B> 証明書は発行されました </B>
<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>

<P ID=locInfo> 要求した証明書は要求者に発行されました</P>

<P><Form Name=UIForm>

<Table Border=0 CellSpacing=0 CellPadding=0>
<%If "no"<>Request.Form("SaveCert") Then%>
<TR><TD></TD>
	<TD></TD>
	<TD></TD>
	<TD><Font ID=locEncFont Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
		<Input Type=Radio ID=rbDerEnc Name=rbEncoding Checked><Label For=rbDerEnc ID=locDerEnc0>DER エンコード</Label>
		<LocID ID=locSep0>&nbsp;または&nbsp;</LocID>
		<Input Type=Radio ID=rbB64Enc Name=rbEncoding><Label For=rbB64Enc ID=locB64Enc0>Base 64 エンコード</Label>
		</Font>
	</TD>
</TR>
<TR><TD ColSpan=4 Height=4></TD></TR>
<%End If%>
<TR>
	<TD><Img Src="certspc.gif" Alt="" Height=1 Width=40></TD>
	<TD><Img Src="certcert.gif" Alt="" Width=32 Height=24></TD>
	<TD><Img Src="certspc.gif" Alt="" Height=1 Width=5></TD>
	<%If "IE"=sBrowser And "no"=Request.Form("SaveCert") Then%>
	<TD><Font ID=locInstCert1Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial"><Span tabindex=0 ID=spnInstall
		Style="cursor:hand; color:#0000FF; text-decoration:underline;"
		OnContextMenu="return false;"
		OnMouseOver="window.status='<%=L_InstallCert_Message%>';return true;" 
		OnMouseOut="window.status='';return true;" 
		OnKeyDown="if (13==event.keyCode) {Install();return false;} else if (9==event.keyCode) {return true;};return false;"
		OnClick="Install();return false;"><LocID ID=locInstallCert1>この証明書のインストール</LocID></Span>
		<Span ID=spnAlreadyInstalled Style="display:none"><LocID ID=locAlreadyInstalledCert>(この証明書は既にインストールされています)</LocID></Span></Font></TD>
		
	<%ElseIf "NN"=sBrowser And "no"=Request.Form("SaveCert") Then%>
	<TD><Font ID=locInstCert2Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial"><A Href="certnew.cer?ReqID=<%=ICertRequest.GetRequestId()%>&amp;Mode=inst&amp;Enc=b64" OnMouseOver="window.status='<%=L_InstallCert_Message%>';return true;" OnMouseOut="window.status='';return true;"><LocID ID=locInstallCert2>この証明書のインストール</LocID></A></Font></TD>
	
	<%ElseIf "UnknownClient"=sBrowser Then%>
	<TD><Font ID=locDwnld1Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
		<LocID ID=locDownloadCert1>証明書のダウンロード: </LocID><A Href="certnew.cer?ReqID=<%=ICertRequest.GetRequestId()%>&amp;Enc=bin"><LocID ID=locDerEnc1>DER エンコード</LocID></A><LocID ID=locSep1> または </LocID><A Href="certnew.cer?ReqID=<%=ICertRequest.GetRequestId()%>&amp;Enc=b64"><LocID ID=locB64Enc1>Base 64 エンコード</LocID></A><BR>
		<LocID ID=locDownloadCertChain1>証明書チェーンのダウンロード: </LocID><A Href="certnew.p7b?ReqID=<%=ICertRequest.GetRequestId()%>&amp;Enc=bin"><LocID ID=locDerEnc2>DER エンコード</LocID></A><LocID ID=locSep2> または </LocID><A Href="certnew.p7b?ReqID=<%=ICertRequest.GetRequestId()%>&amp;Enc=b64"><LocID ID=locB64Enc2>Base 64 エンコード</LocID></A>
		</Font>
		</TD>
	<%ElseIf "IE"=sBrowser Then%>
	<TD><Font ID=locDwnld2Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
			<Span tabindex=0 Style="cursor:hand; color:#0000FF; text-decoration:underline;"
				OnContextMenu="return false;"
				OnMouseOver="window.status='<%=L_DownloadCert_Message%>'; return true;" 
				OnMouseOut="window.status=''; return true;" 
				OnKeyDown="if (13==event.keyCode) {handleGetCert();return false;} else if (9==event.keyCode) {return true;};return false;"
				OnClick="handleGetCert();return false;">
			<LocID ID=locDownloadCert2>証明書のダウンロード</LocID></Span>
		<BR>
			<Span tabindex=0 Style="cursor:hand; color:#0000FF; text-decoration:underline;"
				OnContextMenu="return false;"
				OnMouseOver="window.status='<%=L_DownloadChain_Message%>'; return true;" 
				OnMouseOut="window.status=''; return true;" 
				OnKeyDown="if (13==event.keyCode) {handleGetChain();return false;} else if (9==event.keyCode) {return true;};return false;"
				OnClick="handleGetChain();return false;">
			<LocID ID=locDownloadCertChain2>証明書チェーンのダウンロード</LocID></Span>
		</Font></TD>
	<%Else%>
	<TD><Font ID=locDwnld3Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
			<A Href="#"
				OnMouseOver="window.status='<%=L_DownloadCert_Message%>'; return true;" 
				OnMouseOut="window.status=''; return true;" 
				OnClick="handleGetCert();return false;">
			<LocID ID=locDownloadCert3>証明書のダウンロード</LocID></A>
		<BR>
			<A Href="#"
				OnMouseOver="window.status='<%=L_DownloadChain_Message%>'; return true;" 
				OnMouseOut="window.status=''; return true;" 
				OnClick="handleGetChain();return false;">
			<LocID ID=locDownloadCertChain3>証明書チェーンのダウンロード</LocID></A>
		</Font></TD>
	<%End If%>
</TR>
</Table>
</P>

<P>
    <Span ID=spnNotTrusted Style="display:none"><LocID ID=locNotTrusted>
	<Font Color=#FF0000>この CA は信頼されていません。</Font>この証明機関から発行された証明書を信頼するには、
	<A Href="certnew.cer?ReqID=CACert&amp;Renewal=<%=nRenewals%>&amp;Mode=inst&amp;Enc=b64"
		OnMouseOver="window.status='<%=L_InstallThisCACert_Message%>';return true;"
		OnMouseOut="window.status='';return true;"
	>この CA 証明書をインストール</A>してください。
	</LocID></Span>
</P>


<%If "no"=Request.Form("SaveCert") Then%>

<p>
<table border=0 cellspacing=0 cellpadding=4 ID="Table1">
    
	<TR>
		<TD><Input Type=Checkbox Name=cbSaveResponse ID=cbSaveResponse OnClick="handleSaveResponse();"><Label For=cbSaveResponse ID=locSaveResponseLabel>応答の保存</Label></TD>
	</TR>

</table>
</p>

<%End If%>
<p>
<Font Size=-1 face="Courier New">
<SPAN ID=spnSavedResponse tabIndex=-1 title=spnSavedResponseLabel Style="display:none">
</SPAN>
</font>

</p>
</Form>

<%If "IE"=sBrowser And "no"=Request.Form("SaveCert") And "chkpnd"=Request.Form("Mode")Then%>
<!-- This option is shown if install fails -->
<Span ID=spnRmpn Style="display:none">
<Form Action="certrmpn.asp" Method=Post>
<Input Type=Hidden Name=Action Value="rmpn">
<Input Type=Hidden Name=ReqID Value="<%=ICertRequest.GetRequestId()%>">
<P><Input ID=locBtnRemove Type=Submit Value="削除"> - 保留中の要求一覧からこの要求を削除します。
</Form>
</Span>
<%End If%>

<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>
<!-- White HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#FFFFFF><Img Src="certspc.gif" Alt="" Height=5 Width=1></TD></TR></Table>

</Font>
<!-- ############################################################ -->
<!-- End of standard text. Scripts follow  -->
	
<%bIncludeXEnroll=True%>
<%bIncludeGetCspList=False%>
<%bNeedLHInstallObject=True%>
<!-- #include FILE=certcnst.inc -->
<!-- #include FILE=xenrprxy.inc -->
<!-- #include FILE=certsgcl.inc -->

<%If "IE"=sBrowser And "no"=Request.Form("SaveCert") Then%>
<!-- This form passes data to certrmpn.asp -->
<Span Style="display:none">
<Form Name=SubmittedData Action="certrmpn.asp" Method=Post>
<Input Type=Hidden Name=Action Value="inst">
<Input Type=Hidden Name=ReqID Value="<%=ICertRequest.GetRequestId()%>">
</Form>
<!-- This form is used to try to prevent installing twice -->
<Form Name=State><Input Type=Hidden Name=AlreadyInstalled Value=""></Form>
</Span>

<Script Language="JavaScript">
	//================================================================
	// PAGE GLOBAL VARIABLES

	//----------------------------------------------------------------
	// Strings To Be Localized
	var L_StillLoading_ErrorMessage="このページの読み込みがまだ終了していません。数秒待ってから、再実行してください。";
	var L_NoPrivKey_ErrorMessage="証明書が既にインストールされている可能性があります。\n\n(証明書をインストールできませんでした:\nこの証明書に対応する秘密キーが見つかりませんでした。)";
	var L_CannotAddRoot_ErrorMessage="証明書がインストールされました。\nしかし、ローカル ストアにルート証明書を追加することはできません。\nドメイン グループ ポリシーが原因である可能性があります。\n\nCA 管理者に連絡して、証明書にチェーン確認エラーがないかどうかを\n確認してください。";
	var L_RootIsNotAdded_ErrorMessage="証明書がインストールされました。\nしかしルート証明書はインストールされていません。\nホームページに行くと CA ルート証明書をダウンロードすることができます。";
	var L_UnknownInstallFailure_ErrorMessage="\"証明書をインストールできません:\\n\\エラー: \"+sErrorNumber";
	var L_Cancelled_ErrorMessage="ユーザーによって操作が取り消されました。";

	// IE is not ready until XEnroll has been loaded
	var g_bOkToInstall=false;

	//================================================================
	// INITIALIZATION ROUTINES

	//----------------------------------------------------------------
	// This contains the functions we want executed immediately after load completes
	function postLoad() {
		<%If False=bLH Then%>		
		    // Load an XEnroll object into the page
		    loadXEnroll("postLoadPhase2()"); 
        <%Else%>        
            // No need to load xenroll.  Call the postLoad directly:
            postLoadPhase2();            
        <%End If%>
	}
	function postLoadPhase2() {
		// continued from above

		// Now we're ready to go
		g_bOkToInstall=true;

		// if the user installed the cert then pressed back,
		// we'll try to keep them from getting confused by preventing
		// them from trying to install again and getting an error.
		if (""!=document.State.AlreadyInstalled.value) {
			spnAlreadyInstalled.style.display="";
			spnInstall.style.display="none";
		}	
	}

	//----------------------------------------------------------------
	function markInstalled() {
		document.State.AlreadyInstalled.value="Y";
	}

	//================================================================
	// INSTALL ROUTINES

	//----------------------------------------------------------------
	// perform substitution on the error string, because VBScript cannot
	function evalErrorMessage(sErrorNumber) {
		return eval(L_UnknownInstallFailure_ErrorMessage);
	}
</Script>

<Script Language="VBScript">
	Public sPKCS7
	sPKCS7=""
<%=sPKCS7%>
</Script>

<Script Language="JavaScript">

	function handleSaveResponse() {
		if (document.UIForm.cbSaveResponse.checked) {
			spnSavedResponse.innerText = sPKCS7;
			spnSavedResponse.style.display = '';
		} else {
			spnSavedResponse.style.display = 'none';
		}
	}

</Script>

<Script Language="VBScript">
	    
	'-----------------------------------------------------------------
	' Install the certificate
	Sub Install()
		On Error Resume Next

		Dim     bMachine
		Dim     sMessage
		
		If False=g_bOkToInstall Then
			Alert L_StillLoading_ErrorMessage
			Exit Sub
		End If
		
		bMachine = False
		
		Call XE_reset()
		
		<%If IsNumeric(Request.Form("TargetStoreFlags")) Then%>
			<%If Request.Form("TargetStoreFlags") > 0 Then%>
			
				Dim dwStoreFlags
		    
				dwStoreFlags = <%=Request.Form("TargetStoreFlags")%>
				dwStoreFlags = dwStoreFlags And CERT_SYSTEM_STORE_LOCAL_MACHINE
				if 0 = dwStoreFlags Then
					bMachine = False
				Else
					bMachine = True
				End If
			
			<%End If%>
		<%End If%>

        Call XE_Enroll_AcceptResponse(g_objEnroll, sPKCS7, bMachine)


		If &H800B0109=Err.number Then
			
			' this error is likely coming from CertEnroll, indicating that the CA is not trusted.
			' Give the user the option to install the CA cert, and fail
			spnNotTrusted.style.display=""
			
		Else 
			If 0=Err.Number Or &H80095001=Err.Number Then
				If &H80095001=Err.Number Then
					'inform user about root cert install failure
					Alert L_CannotAddRoot_ErrorMessage
				End If
				' Certificate has been successfully installed. Go to 'success' page
				markInstalled
				document.SubmittedData.submit
			Else
				If Err.Number=&H80092004 Then 'CRYPT_E_NOT_FOUND
					' the private key was not found - most likely this is an attempt to reinstall
					sMessage=L_NoPrivKey_ErrorMessage
				ElseIf Err.Number=&H800704C7 Then 'ERROR_CANCELLED
									' The operation was canceled by the user.
									sMessage=L_Cancelled_ErrorMessage
				Else
					' unknown error
					sMessage=evalErrorMessage("0x" & Hex(Err.Number))
				End If
				'don't popup if cancel, downlevel OS such as NT4 returns H80070002 in case of dlg cancel
				If &H800704c7<>Err.Number And &H8010006e<>Err.Number And &H80070002<>Err.Number Then
					Alert sMessage
				End If
				<%If "chkpnd"=Request.Form("Mode") Then%>
				' give the user the option to remove this broken cert from their list
				spnRmpn.style.display=""
				<%End If%>
			End If
		End If
				
	End Sub
</Script>

<%ElseIf "UnknownClient"<>sBrowser Then%>
<Script Language="JavaScript">
	//================================================================
	// INITIALIZATION ROUTINES

	//----------------------------------------------------------------
	// This contains the functions we want executed immediately after load completes
	function postLoad() {
		// do nothing
	}

	//================================================================
	// LINK HANDLERS

	//----------------------------------------------------------------
	// Get the requested cert
	function handleGetCert() {
		location="certnew.cer?ReqID=<%=ICertRequest.GetRequestId()%>&"+getEncoding();
	}
	//----------------------------------------------------------------
	// Get the requested certificate chain
	function handleGetChain() {
		location="certnew.p7b?ReqID=<%=ICertRequest.GetRequestId()%>&"+getEncoding();
	}

	//----------------------------------------------------------------
	// return the ecoding parameter based upon the radio button
	function getEncoding() {
		if (true==document.UIForm.rbEncoding[0].checked) {
			return "Enc=bin";
		} else {
			return "Enc=b64";
		}
	}

</Script>
<%Else '"IE"<>sBrowser%>

<!-- No Scripts -->

<%End If%>

</Body>
</HTML>
<%Session.Abandon()%>