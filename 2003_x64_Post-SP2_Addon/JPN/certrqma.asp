<%@ CODEPAGE=65001 'UTF-8%>
<%' certrqma.asp - (CERT)srv web - (R)e(Q)uest, (M)ore (A)dvanced
  ' Copyright (C) Microsoft Corporation, 1998 - 1999 %>
<!-- #include FILE=certsbrt.inc -->
<!-- #include FILE=certdat.inc -->

<!-- Windows Security Update, KB2518295 has replaced some of the CA Web Enrollment ASP files -->
<!-- Please see http://www.support.microsoft.com/kb/2518295 for the back-up location of the previous ASP files -->

<% 
	On Error Resume Next

	' Exporting keys to a pvk file is only used by old code signing tools.
	' (This is different from exporting both cert and keys in a pfx file.)
	' Set this flag to true if you really need this functionality
	bEnableExportKeyToFile = True

	Dim bFailed, nError
	bFailed=False
	nError = 0

	If "Enterprise"=sServerType And "IE"=sBrowser Then

	        ' get CA exchange cert and save into this page

		Const CR_OUT_BASE64      =&H00000001
		Const CR_PROP_CAEXCHGCERT=15
		Const PROPTYPE_BINARY=3

		Set ICertRequest2=Server.CreateObject("CertificateAuthority.Request")
		Public sCAExchangeCert
		Dim sCertificate
		sCertificate=ICertRequest2.GetCAProperty(sServerConfig, CR_PROP_CAEXCHGCERT, 0, PROPTYPE_BINARY, CR_OUT_BASE64)
		sCAExchangeCert=FormatBigString(sCertificate, "	sCAExchange=sCAExchange & ")

		'&H800B0113, CERT_E_INVALID_POLICY, treat it as OK
		If Err.Number<>0 And Err.Number <> &H800B0113 Then
			' CA may be down.
			bFailed=True
			nError=Err.Number
		End If
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
	<Title>Microsoft 証明書サービス</Title>
</Head>
<%If True=bFailed Then %>
<Body BgColor=#FFFFFF Link=#0000FF VLink=#0000FF ALink=#0000FF><Font ID=locPageFont Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
<%Else%>
<Body BgColor=#FFFFFF Link=#0000FF VLink=#0000FF ALink=#0000FF OnLoad="postLoad();"><Font ID=locPageFont Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
<%End If%>

<Table Border=0 CellSpacing=0 CellPadding=4 Width=100% BgColor=#008080>
<TR>
	<TD><Font Color=#FFFFFF><LocID ID=locMSCertSrv><Font Face="MS UI Gothic" Size=-1><B><I>Microsoft</I></B> 証明書サービス &nbsp;--&nbsp; <%=sServerDisplayName%> &nbsp;</Font></LocID></Font></TD>
	<TD ID=locHomeAlign Align=Right><A Href="/certsrv"><Font Color=#FFFFFF><LocID ID=locHomeLink><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B>ホーム</B></Font></LocID></Font></A></TD>
</TR>
</Table>

<%If True=bFailed Then %>
<P ID=locPageTitle1><Font Color=#FF0000><B>エラー</B></Font>
<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>

<P ID=locErrorMsg> 予期しないエラーが発生しました。<BR><BR>
<%If nError=&H800706BA Or nError=&H80070005 Then%>
	<LocID ID=locSvcNotStarted>証明機関サービスは開始されていません。</LocID>
<%ElseIf nError=&H800b0101 Then%> 
	<LocID ID=locSvcNotValidDate>現在のシステム クロックまたは署名されたファイルのタイムスタンプと検証したところ、必要な証明書は有効期間内にはありません。</LocID>
<%Else%>
	<%=ICertRequest.GetErrorMessageText(nError, CR_GEMT_HRESULT_STRING)%>
<%End If%>

<%Else 'True<>bFailed%>

<Form Name=UIForm OnSubmit="goNext();return false;" Action="certlynx.asp" Method=Post>
<Input Type=Hidden Name=SourcePage Value="certrqma">

<P ID=locPageTitle> <B> 証明書の要求の詳細設定 </B>
<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=></TD></TR></Table>

<Span ID=spnFixTxt Style="display:none">
	<Table Border=0 CellSpacing=0 CellPadding=4 Style="Color:#FF0000"><TR><TD ID=locBadCharError>
		<I><B>赤く</B>表示されているフィールドを修正してください。</I>
		電子メール アドレスは文字 A-Z、a-z、0-9、およびいくつかの共通シンボルを使えますが、拡張文字は使用できません。
		国/地域フィールドには 2 桁の ISO 3166 国/地域コードを使ってください。
	</TD></TR></Table>
</Span>
<Span ID=spnErrorTxt Style="display:none">
	<Table Border=0 CellSpacing=0 CellPadding=4 Style="Color:#FF0000">
	<TR><TD ID=locErrMsgBasic>
		証明書の要求を作成しているときに<B>エラーが発生しました</B>。
		CSP がすべての設定をサポートしているか、
		また入力したデータが有効であるか確認してください。
	</TD></TR><TR><TD>
		<LocID ID=locErrorCause><B>原因:</B></LocID><BR>
		<Span ID=spnErrorMsg></Span>
	</TD></TR><TR>
		<TD ID=locErrorNumber><Font Size=-2>エラー: <Span ID=spnErrorNum></Span></Font></TD>
	</TR>
	</Table>
</Span>


<Table Border=0 CellSpacing=0 CellPadding=0>
	<TR> <!-- establish column widths. -->
		<TD Width=<%=L_LabelColWidth_Number%>></TD> <!-- label column, top border -->
		<TD RowSpan=59 Width=4></TD>                <!-- label spacing column -->
		<TD></TD>                                   <!-- field column -->
	</TR>
	
<%If "Enterprise"=sServerType Then%>	<!-- Enterprise Options -->

	<TR>
		<TD ID=locTemplateHead ColSpan=3><Font Size=-1><BR><Label For=lbCertTemplateID><locID ID=locTemplateHead><B>証明書テンプレート:</B></locID></Label></Font></TD>
	</TR><TR><TD ColSpan=3 Height=2 BgColor=#008080></TD>
	</TR><TR><TD ColSpan=3 Height=6></TD>
	</TR><TR><TD></TD>
		<TD><Select Name=lbCertTemplate ID=lbCertTemplateID OnChange="handleTemplateChange();">
<%
	Dim nWriteTemplateResult
	nWriteTemplateResult=WriteTemplateList() 
%>
		</Select></TD>
	</TR>

<%End If '"Enterprise"=sServerType%>
	<TR><TD ColSpan=3>
<%If "Enterprise"=sServerType Then%>	<!-- Enterprise Options -->
<Span ID=spnIDInfo Style="display:none">
<%End If '"Enterprise"=sServerType%>
<Table Border=0 CellSpacing=0 CellPadding=0>
	<TR> <!-- establish column widths. -->
		<TD Width=<%=L_LabelColWidth_Number%>></TD> <!-- label column, top border -->
		<TD RowSpan=59 Width=4></TD>                <!-- label spacing column -->
		<TD></TD>                                   <!-- field column -->
	</TR>


	<TR>
<%If "StandAlone"=sServerType Then%>
		<TD ID=locIdentHeadStandAlone ColSpan=3><Font Size=-1><BR><B>識別情報:</B></Font></TD>
	</TR><TR><TD ColSpan=3 Height=2 BgColor=#008080></TD>
<%Else%>
		<TD ID=locIdentHeadEnterprise ColSpan=3><Font Size=-1><BR><B>オフライン テンプレート用の識別情報:</B></Font></TD>
	</TR><!--<TR><TD ColSpan=3 Height=2 BgColor=#008080></TD>-->
<%End If%>
	</TR><TR><TD ColSpan=3 Height=6></TD>
	</TR><TR>
		<TD ID=locNameAlign Align=Right><Span ID=spnNameLabel><LocID ID=locNameLabel><Font Size=-1>名前:</Font></LocID></Span></TD>
		<TD><Input ID=locTbCommonName Type=Text MaxLength=64 Size=42 Name=tbCommonName></TD>
	</TR><TR>
		<TD ID=locEmailAlign Align=Right><Span ID=spnEmailLabel><LocID ID=locEmailLabel><Font Size=-1>電子メール:</Font></LocID></Span></TD>
		<TD><Input ID=locTbEmail Type=Text MaxLength=128 Size=42 Name=tbEmail></TD>
	</TR><TR>
		<TD Height=8></TD> <TD></TD>
	</TR><TR>
		<TD ID=locCompanyAlign Align=Right><Span ID=spnCompanyLabel><LocID ID=locOrgLabel><Font Size=-1>会社:</Font></LocID></Span></TD>
		<TD><Input ID=locTbOrg Type=Text MaxLength=64 Size=42 Name=tbOrg Value="<%=sDefaultCompany%>"></TD>
	</TR><TR>
		<TD ID=locDepartmentAlign Align=Right><Span ID=spnDepartmentLabel><LocID ID=locOrgUnitLabel><Font Size=-1>部署:</Font></LocID></Span></TD>
		<TD><Input ID=locTbOrgUnit Type=Text MaxLength=64 Size=42 Name=tbOrgUnit Value="<%=sDefaultOrgUnit%>"></TD>
	</TR><TR>
		<TD Height=8></TD> <TD></TD>
	</TR><TR>
		<TD ID=locCityAlign Align=Right><Span ID=spnCityLabel><LocID ID=locLocalityLabel><Font Size=-1>市区町村:</Font></LocID></Span></TD>
		<TD><Input ID=locTbLocality Type=Text MaxLength=128 Size=42 Name=tbLocality Value="<%=sDefaultLocality%>"></TD>
	</TR><TR>
		<TD ID=locStateAlign Align=Right><Span ID=spnStateLabel><LocID ID=locStateLabel><Font Size=-1>都道府県:</Font></LocID></Span></TD>
		<TD><Input ID=locTbState Type=Text MaxLength=128 Size=42 Name=tbState Value="<%=sDefaultState%>"></TD>
	</TR><TR>
		<TD ID=locCountryAlign Align=Right><Span ID=spnCountryLabel><LocID ID=locCountryLabel><Font Size=-1>国/地域コード:</Font></LocID></Span></TD>
		<TD><Input ID=locTbCountry Type=Text MaxLength=2 Size=2 Name=tbCountry Value="<%=sDefaultCountry%>"></TD> 
	</TR>

</Table>
<%If "Enterprise"=sServerType Then%>
</Span>
<%End If '"Enterprise"=sServerType%>
	</TD></TR>

<%If "StandAlone"=sServerType Then%> <!-- Stand-Alone Options -->

	<TR>
		<TD ID=locEKUHead ColSpan=3><Font Size=-1><BR><B>必要な証明書の種類:</B></Font></TD>
	</TR><TR><TD ColSpan=3 Height=2 BgColor=#008080></TD>
	</TR><TR><TD ColSpan=3 Height=6></TD>
	</TR><TR><TD></TD>
		<TD><Select Name=lbUsageOID OnChange="handleUsageOID(true);">
			<Option ID=locCliAuthCert Selected Value="1.3.6.1.5.5.7.3.2"> クライアント認証証明書
			<Option ID=locEmailCert   Value="1.3.6.1.5.5.7.3.4"> 電子メール保護証明書
			<Option ID=locSrvAuthCert Value="1.3.6.1.5.5.7.3.1"> サーバー認証証明書
			<Option ID=locCodeSgnCert Value="1.3.6.1.5.5.7.3.3"> コード署名証明書
			<Option ID=locTimStmpCert Value="1.3.6.1.5.5.7.3.8"> タイムスタンプの署名証明書
			<Option ID=locIPSecCert   Value="1.3.6.1.5.5.8.2.2"> IPSec 証明書
			<Option ID=locUserEKUCert Value="**"> その他...
		</Select></TD>
	</TR>

	<TR><TD ID=locEkuAlign Align=Right><Span ID=spnEKUOther1 Style="display:none"><LocID ID=locUserEKULabel><Font Size=-1>OID:</Font></LocID></Span></TD>
		<TD><Span ID=spnEKUOther2 Style="display:none"><Input ID=locTbEKUOther Type=Text Name=tbEKUOther Value="1.3.6.1.5.5.7.3."></Span></TD>
	</TR>

<%End If%> <!-- common -->

	<TR>
		<TD ID=locKeyOptHead ColSpan=3><Font Size=-1><BR><B>キーのオプション:</B></Font></TD>
	</TR><TR><TD ColSpan=3 Height=2 BgColor=#008080></TD>
	</TR><TR><TD ColSpan=3 Height=6></TD>
	</TR>

	<TR>
		<TD></TD>
		<TD><Font Size=-1>
			<Input Type=Radio ID=rbKG1 Name=rbKeyGen Value="0" OnClick="handleKeyGen();" Checked><Label For=rbKG1 ID=locNewKeyLabel>新しいキー セットを作成する</Label>
			<LocID ID=locSpc3>&nbsp;&nbsp;&nbsp;<LocID>
			<Input Type=Radio ID=rbKG2 Name=rbKeyGen Value="1" OnClick="handleKeyGen();"><Label For=rbKG2 ID=locExistKeyLabel>既存のキー セットを使用する</Label>
		</Font></TD>
	</TR>

	<TR><TD ColSpan=3 Height=4></TD></TR>
	<TR>
		<TD ID=locCSPLabel Align=Right><Font Size=-1><Label For=lbCSPID><locID ID=locCSPLabel>CSP:</locID></Label></Font></TD>
		<TD><Select Name=lbCSP ID=lbCSPID OnChange="handleCSPChange();">
			<Option ID=locLoading>読み込んでいます...</Option>
		</Select></TD>
	</TR>
	<TR ID=trBadCSPForKeySpec Style="display:none">
		<TD></TD>
		<TD BgColor=#FFFFE0><LocID ID=locBadCSPForKeySpec><Font Size=-1><Span ID=spnBadCSPForKeySpecMsg></Span></Font></LocID></TD>
	</TR>

	<TR><TD ColSpan=3 Height=4></TD></TR>
	<TR>
		<TD ID=locKeyUsageLabel Align=Right><Font Size=-1>キー使用法:</Font></TD>
		<TD><Font Size=-1>
			<Span ID=spnKeyUsageKeyExchange><Input Type=Radio ID=rbKU1 Name=rbKeyUsage Value="0" Checked OnClick="handleKeyUsageChange(false);"><Label For=rbKU1 ID=locKUExch>交換</Label><LocID ID=locSpc1>&nbsp;&nbsp;&nbsp;<LocID></Span>
			<Span ID=spnKeyUsageSignature><Input Type=Radio ID=rbKU2 Name=rbKeyUsage Value="1" OnClick="handleKeyUsageChange(false);"><Label For=rbKU2 ID=locKUSig>署名</Label><LocID ID=locSpc2>&nbsp;&nbsp;&nbsp;<LocID></Span>
			<Span ID=spnKeyUsageBoth><Input Type=Radio ID=rbKU3 Name=rbKeyUsage Value="2" OnClick="handleKeyUsageChange(false);"><Label For=rbKU3 ID=locKUBoth>両方</Label></Span></Font></TD>
	</TR>

	<TR><TD ColSpan=3 Height=4></TD></TR>
	<TR>
		<TD ID=locKeySizeLabel Align=Right ><Font Size=-1><Label For=locTbKeySize><locID ID=locKeySizeLabel>キーのサイズ:</locID></Label></Font></TD>
		<TD><Table Border=0 CellPadding=0 CellSpacing=0>
			<TR>
				<TD RowSpan=2><Input ID=locTbKeySize Type=Text Name=tbKeySize Value="0" MaxLength=5 Size=4 OnPropertyChange="handleKeySizeChange();">&nbsp;</TD>
				<TD ID=locKeySizeMinLabel Align=Right><Font Size=-2>最小:</Font></TD>
				<TD ID=locKeySizeMin Align=Right><Font Size=-2><Span ID=spnKeySizeMin></Span></Font></TD>
				<TD ID=locKeySizeCommon RowSpan=2><Font Size=-2>&nbsp;&nbsp;(キーのサイズ: <Span ID=spnKeySizeCommon></Span>)</Font></TD>
			</TR><TR>
				<TD ID=locKeySizeMaxLabel Align=Right><Font Size=-2>最大:</Font></TD>
				<TD ID=locKeySizeMax Align=Right><Font Size=-2><Span ID=spnKeySizeMax></Span></Font></TD>
			</TR>
		</Table></TD>
	</TR>
	<TR ID=trKeySizeBad Style="display:none">
		<TD></TD>
		<TD BgColor=#FFFFE0><LocID ID=locKeySizeBad><Font Size=-1><Span ID=spnKeySizeBadMsg></Span></Font></LocID></TD>
	</TR>
	<TR ID=trKeySizeBadSpc Style="display:none"><TD ColSpan=3 Height=4></TD></TR>
	<TR ID=trKeySizeWarn Style="display:none">
		<TD></TD>
		<TD BgColor=#FFFFE0><LocID ID=locKeySizeWarning><Font Size=-1><I>警告: 大きなキー サイズは作成するのに時間がかかる可能性があります。</I></Font></LocID></TD>
	</TR>
	<TR ID=trKeyGenWarn Style="display:none">
		<TD></TD>
		<TD><LocID ID=locKeyGenWarning><Font Size=-1><I>このサイズのキーは、指定されたコンテナに
		指定された使用法のキーが既に存在しない場合に</I>のみ<I>
		生成されます。</I></Font></LocID></TD>
	</TR>

	<TR ID=trGenContNameSpc><TD ColSpan=3 Height=4></TD></TR>
	<TR ID=trGenContName>
		<TD></TD>
		<TD><Font Size=-1>
			<Input Type=Radio ID=rbGCN1 Name=rbGenContName Value="0" OnClick="handleGenContName();" Checked><Label For=rbGCN1 ID=locAutoContNameLabel>自動キー コンテナ名</Label>
			<LocID ID=locSpc4>&nbsp;&nbsp;&nbsp;<LocID>
			<Input Type=Radio ID=rbGCN2 Name=rbGenContName Value="1" OnClick="handleGenContName();"><Label For=rbGCN2 ID=locUserContNameLabel>ユーザーが指定したキー コンテナ名</Label>
		</Font></TD>
	</TR>

	<TR ID=trContNameSpc><TD ColSpan=3 Height=4></TD></TR>
	<TR ID=trContName Style="display:none">
		<TD ID=locContainerNameLabel Align=Right><Font Size=-1>コンテナ名:</Font></TD>
		<TD><Font Size=-1><Input ID=locTbContainerName Type=Text Name=tbContainerName Size=20></Font></TD>
	</TR>

	<TR ID=trMarkExportSpc><TD ColSpan=3 Height=4></TD></TR>
	<TR ID=trMarkExport><TD></TD>
		<TD><Font Size=-1><Input Type=Checkbox Name=cbMarkKeyExportable ID=cbMarkKeyExportable OnClick="handleMarkExport(false);"><Label For=cbMarkKeyExportable ID=locMarkExportLabel>エクスポート可能なキーとしてマークする</Label>
		<%If bEnableExportKeyToFile Then%>
		<Span ID=spnMarkKeyExportable Style="display:none">
			<BR><Img Src="certspc.gif" Alt="" Height=1 Width=25><Input Type=Checkbox Name=cbExportKeys ID=cbExportKeys OnClick="handleExportKeys();"><Label For=cbExportKeys ID=locExportToFileLabel>キーをファイルへエクスポートする</Label>
			<Span ID=spnExportKeys Style="display:none">
				<BR><Img Src="certspc.gif" Alt="" Height=1 Width=25><LocID ID=locExpFileNameLabel>完全なパス名:</LocID> <Input ID=locTbExportKeyFile Type=Text Name=tbExportKeyFile Size=20 Value="*.pvk">
			</Span>
		</Span>
		<%End If%>
		</Font></TD>
	</TR>

	<TR ID=trStrongKeySpc><TD ColSpan=3 Height=4></TD></TR>
	<TR ID=trStrongKey>
		<TD></TD>
		<TD><Font Size=-1><Input Type=Checkbox ID=cbStrongKey Name=cbStrongKey OnClick="handleStrongKeyAndLMStore(false);"><Label For=cbStrongKey ID=locStrongKeyLabel>秘密キーの保護を強力にする</Label></Font></TD>
	</TR>

	<TR ID=trLMStoreSpc><TD ColSpan=3 Height=4></TD></TR>
	<TR ID=trLMStore><TD></TD>
		<TD><Font Size=-1><Input Type=Checkbox Name=cbLocalMachineStore ID=cbLocalMachineStore OnClick="handleStrongKeyAndLMStore(false);"><Label For=cbLocalMachineStore ID=locLMStoreLabel>ローカル コンピュータの証明書ストアに証明書を格納する</Label><BR>
		<LocID ID=locAdminWarning><Img Src="certspc.gif" Alt="" Height=1 Width=25><I>ユーザーの証明書ストアの代わりにローカル コンピュータ<BR>
                  <Img Src="certspc.gif" Alt="" Height=1 Width=25>ストアに証明書を格納します。ルート CA の<BR>
                  <Img Src="certspc.gif" Alt="" Height=1 Width=25>証明書はインストールしません。ローカル<BR>
                  <Img Src="certspc.gif" Alt="" Height=1 Width=25>コンピュータ ストアの中にキーを生成する、または<BR>
                  <Img Src="certspc.gif" Alt="" Height=1 Width=25>キーを使用するには管理者でなければなりません。</I></Font></LocID></TD>
	</TR>

	<TR>
		<TD ID=locAddOptHead ColSpan=3><Font Size=-1><BR><B>追加オプション:</B></Font></TD>
	</TR><TR><TD ColSpan=3 Height=2 BgColor=#008080></TD>
	</TR><TR><TD ColSpan=3 Height=3></TD>
	</TR>

	<TR><TD ColSpan=3 Height=6></TD></TR>
	<TR>
		<TD ID=locRequestFormatLabel Align=Right><Font Size=-1>要求の形式:</Font></TD>
		<TD>
		<Input Type=Radio ID=rbFormatCMC Name=rbRequestFormat Value="0" Checked><Label For=rbFormatCMC ID=locFormatCMCLabel>CMC</Label>
		<LocID ID=locSpc5>&nbsp;&nbsp;&nbsp;<LocID>
		<Input Type=Radio ID=rbFormatPKCS10 Name=rbRequestFormat Value="1"><Label For=rbFormatPKCS10 ID=locFormatPKCS10Label>PKCS10</Label>
		</TD>
	</TR>
	<TR><TD ColSpan=3 Height=4></TD></TR>

	<TR><TD ColSpan=3 Height=4></TD></TR>
	<TR>
		<TD ID=locHashAlgLabel Align=Right><Font Size=-1><Label For=lbHashAlgorithmID><locID ID=locHashAlgLabel>ハッシュ アルゴリズム:</locID></Label></Font></TD>
		<TD><Select Name=lbHashAlgorithm ID=lbHashAlgorithmID></Select></TD>
	</TR>
	<TR><TD></TD><TD ID=locHashAlgWarning><Font Size=-1><I>要求を署名するためだけに使用されます。</I></Font></TD></TR>

	<TR><TD ColSpan=3 Height=8></TD></TR>
	<TR><TD></TD>
		<TD><Font Size=-1><Input Type=Checkbox Name=cbSaveRequest ID=cbSaveRequest OnClick="handleSaveReq();"><Label For=cbSaveRequest ID=locSaveReqLabel>要求をファイルへ保存します。</Label>
		<Span ID=spnSaveRequest Style="display:none">
			<BR><Img Src="certspc.gif" Alt="" Height=1 Width=25><LocID ID=locReqFileNameLabel>完全なパス名:</LocID> <Input ID=locTbSaveReqFile Type=Text Name=tbSaveReqFile Size=20>
			<BR><Img Src="certspc.gif" Alt="" Height=1 Width=25><LocID ID=locSaveReqWarning><B>この要求は保存されますが、送信されません。</B></LocID>
		</Span>
		</Font></TD>
	</TR>

	<TR><TD ColSpan=3 Height=6></TD>
	</TR><TR>
		<TD ID=locAttribLabel Align=Right><Font Size=-1><Span ID=spnSubmitAttrLable><Label For=locTaAttrib><locID ID=locAttribLabel>属性:</locID></Label></Span></Font></TD>
		<TD><Span ID=spnSubmitAttrBox><TextArea ID=locTaAttrib Name=taAttrib Wrap=Off Rows=2 Cols=30></TextArea></SPan></TD>
	</TR>

	<TR><TD ColSpan=3 Height=6></TD>
	</TR><TR>
		<TD ID=locFriendlyNameLabel Align=Right><Font Size=-1><Label For=locTbFriendlyName><locID ID=locFriendlyNameLabel>フレンドリ名:</locID></Label></Font></TD>
		<TD><Font Size=-1><Input ID=locTbFriendlyName Type=Text Name=tbFriendlyName Size=20></Font></TD>
	</TR>

	<TR><TD ColSpan=3><Font Size=-1><BR></Font></TD></TR>
	<TR><TD ColSpan=3 Height=2 BgColor=#008080></TD></TR>
	<TR><TD ColSpan=3 Height=3></TD></TR>
	<TR>
		<TD></TD>
		<TD ID=locSubmitAlign Align=Right>
			<Input ID=locBtnSubmit Type=Submit Name=btnSubmit Value="送信 &gt;" Style="width:.75in">
			<Input ID=locBtnSave Type=Submit Name=btnSave Value="保存" Style="width:.75in; display:none">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</TD>
	</TR>
	<TR><TD ColSpan=3 Height=20></TD></TR>


</Table>
</P>


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

<!-- This form we fill in and submit 'by hand'-->
<Span Style="display:none">
<Form Name=SubmittedData Action="certfnsh.asp" Method=Post>
	<Input Type=Hidden Name=Mode>             <!-- used in request ('newreq'|'chkpnd') -->
	<Input Type=Hidden Name=CertRequest>      <!-- used in request -->
	<Input Type=Hidden Name=CertAttrib>       <!-- used in request -->
	<Input Type=Hidden Name=FriendlyType>     <!-- used on pending -->
	<Input Type=Hidden Name=ThumbPrint>       <!-- used on pending -->
	<Input Type=Hidden Name=TargetStoreFlags> <!-- used on install ('0'|CSSLM)-->
	<Input Type=Hidden Name=SaveCert>         <!-- used on install ('no'|'yes')-->
</FORM>
</Span>

<Script Language="JavaScript">

	//================================================================
	// PAGE GLOBAL VARIABLES

	//----------------------------------------------------------------
	// Strings to be localized
	var L_CspLoadErrNoneFound_ErrorMessage="CSP の一覧を取得中に予期しないエラーが発生しました:\nCSP が見つかりませんでした。";
	var L_CspLoadErrUnexpected_ErrorMessage="\"CSP の一覧を取得中に予期しないエラー (\"+sErrorNumber+\") が発生しました。\"";
	var L_SetKeySize_Message="\"キーのサイズを次の値に設定してください \"+nKeySize";
	var L_WarningTemplateKeySize_Message="\"キー サイズの最小値が \" + nKeySize + \" である証明書の種類を選択しました。これは選択された CSP の最高値よりも大きな数値です。\\n別の CSP を選択してください。\"";
	var L_RecommendOneKeySize_Message="\"\"+nKeySize+\" は無効なキー サイズです。一番近い有効なキー サイズは \"+sCloseBelow+\" です。\"";
	var L_RecommendTwoKeySizes_Message="\"\"+nKeySize+\" は無効なキー サイズです。一番近い有効なキー サイズは \"+sCloseBelow+\" と \"+sCloseAbove+\" です。\"";
	var L_StillLoading_ErrorMessage="このページの読み込みがまだ終了していません。数秒待ってから、再実行してください。";
	var L_KeySizeNotNumber_ErrorMessage="キーのサイズを入力してください。";
	var L_KeySizeBadNumber_ErrorMessage="\"有効なキー サイズを入力してください。キーのサイズは\\n\"+g_nCurKeySizeMin+\" から \"+g_nCurKeySizeMax+\" の間で、\"+g_nCurKeySizeInc+\"の倍数である必要があります。\"";
	var L_CSPNotSupportTemplateKeySpec_Message="\"テンプレートで定義されている種類のキーをサポートしない CSP を選択した可能性があります。テンプレート内のキーを修正するか、別の CSP または証明書テンプレートを選択してください。\"";
	var L_TemplateKeySizeTooBig_ErrorMessage = "\"キー サイズの最小値が \" + g_nCurTemplateKeySizeMin + \"である証明書の種類を選択しました。\\nこれは \" + g_nCurKeySizeMax + \"よりも大きな数値です。\\n数値を変更するか、別の CSP を選択してください。\"";
	var L_NoCntnrName_ErrorMessage="キー コンテナ名を入力してください。";
	var L_BadOid_ErrorMessage="有効な OID を入力するか、定義済みの証明書の種類を選択してください。\n複数の OID を入力するときは、カンマで区切ってください。";
	var L_NoExportFileName_ErrorMessage="キーのエクスポートで使用するファイル名を入力してください。";
	var L_NoSaveReqFileName_ErrorMessage="要求を保存するファイル名を入力してください。";
	var L_Generating_Message="要求を生成しています...";
	var L_UserEKUCert_Text="\"User-EKU (\"+sCertUsage+\") 証明書\"";
	var L_RequestSaved_Message="要求はファイルに保存されました。";
	var L_Waiting_Message="サーバーの応答を待っています...";
	var L_ErrNameUnknown_ErrorMessage="(不明)";
	var L_SugCauseNone_ErrorMessage="なし";
	var L_SugCauseBadCSP_ErrorMessage="選択した CSP は要求を処理できませんでした。別の CSP で実行してください。";
	var L_SugCauseKeysetFull_ErrorMessage="セキュリティ トークンには、コンテナを追加するために利用できる記憶域がありません。";
	var L_SugCauseBadSetting2_ErrorMessage="選択した CSP は、キーのサイズ、キーの指定、ハッシュ アルゴリズムなど 1 つ以上の設定をサポートしていません。別の設定または別の CSP を使用して再実行してください。";
	var L_SugCauseBadKeyContainer_ErrorMessage="指定したキー コンテナが存在しないか、または選択した CSP は要求を処理できませんでした。既存のキー コンテナの名前を入力するか、'新しいキーセットの作成' を選択するか、または別の CSP を使用して再実行してください。";
	var L_SugCauseExistKeyContainer_ErrorMessage="コンテナの名前は既に使用されています。新しいキーを作成するときは、新しいコンテナ名を使用してください。";
	var L_SugCauseBadChar_ErrorMessage="無効な文字を入力しました。これは検証したときに確認されていなければならないので、バグをレポートしてください。";
	var L_SugCauseBadHash_ErrorMessage="選択したハッシュ アルゴリズムは署名するために使用できません。別のハッシュ アルゴリズムを選択してください。";
	var L_SugCauseNoFileName_ErrorMessage="ファイル名が入力されていません。";
	var L_SugCauseCryptArchivableNotSupp_ErrorMessage="選択した CSP では、アーカイブされるがエクスポートされないキーの作成はサポートされていません。";
	var L_ErrNameNoFileName_ErrorMessage="(無名)";
	var L_SugCauseNotAdmin_ErrorMessage="ローカル コンピュータのストアでキーを生成するには、管理者でなければなりません。";
	var L_ErrNamePermissionDenied_ErrorMessage="アクセス許可は拒否されました";
	var L_SugCausePermissionToWrite_ErrorMessage = "ファイルをこのパスに保存するための書き込みのアクセス許可がありません";
	var L_SugCauseBadFileName_ErrorMessage="指定したファイルは有効なファイル名ではありません。別のファイル名を指定してください。";
	var L_SugCauseBadDrive_ErrorMessage="指定されたドライブの準備ができていません。ドライブにディスクを挿入するか、別の名前を入力してください。";
	var L_SugCauseNoProfile_ErrorMessage="ユーザーのプロファイルは一時プロファイルです。";
	var L_SugCauseCAExSignerNotFound_ErrorMessage="証明書チェーンは信頼されたルート機関には構築できません。";
	var L_SugCauseCAExNotTrusted_ErrorMessage="証明書チェーンは処理されましたが、ルート証明書で停止されました。証明書は信頼プロバイダから信頼されていません。";
	var L_DownLevelClients_ErrorMessage="このエラーは新しい秘密キー用にキーのアーカイブを要求したときに起こる可能性があります。このプラットフォームではこれはサポートされていません。";
	var L_SugCauseCancelled_ErrorMessage="ユーザーによって操作が取り消されました。";
	var L_SCARD_E_NOMEMORYMSG_ErrorMessage="メモリ不足のため、このコマンドを完了できません。";
	var L_SCARD_F_WAITEDTOOLONG_ErrorMessage="内部整合性タイマが切れました。";
	var L_SCARD_E_INSUFFICIENTBUFFER_ErrorMessage="返されるデータを受け取るバッファは、小さすぎます。";
	var L_SCARD_E_UNKNOWNREADER_ErrorMessage="指定された読み取り装置名は認識されません。";
	var L_SCARD_E_NOSMARTCARD_ErrorMessage="この操作にはスマート カードが必要ですが、デバイスに現在、スマート カードがありません。";
	var L_SCARD_E_UNKNOWNCARD_ErrorMessage="指定したスマート カード名が認識されません。";
	var L_SCARD_E_NOTREADY_ErrorMessage="読み取り装置またはスマート カードは、コマンドを受け取る準備ができていません。";
	var L_SCARD_F_COMMERROR_ErrorMessage="内部通信エラーが検出されました。";
	var L_SCARD_E_NOSERVICE_ErrorMessage="スマート カード リソース マネージャが実行されていません。";
	var L_SCARD_E_SERVICESTOPPED_ErrorMessage="スマート カード リソース マネージャをシャットダウンしました。";
	var L_SCARD_E_NOREADERSAVAILABLE_ErrorMessage="スマート カード リーダーが見つかりません。";
	var L_SCARD_E_COMMDATALOST_ErrorMessage="スマート カードとの通信エラーが検出されました。操作をもう一度やり直してください。";
	var L_SCARD_E_NOKEYCONTAINER_ErrorMessage="要求されたキー コンテナはスマート カードには存在しません。";
	var L_SCARD_W_UNPOWEREDCARD_ErrorMessage="これ以上通信できないようにするため、スマート カードの電源を切りました。";
	var L_SCARD_W_REMOVEDCARD_ErrorMessage="スマート カードが取り外されたため、これ以上通信することはできません。";
	var L_SCARD_W_WRONGCHV_ErrorMessage="与えられた暗証番号 (PIN) が間違っているため、カードにアクセスできません。";
	var L_SCARD_W_CHVBLOCKED_ErrorMessage="暗証番号 (PIN) エントリの最大数に達したため、カードにアクセスできません。";
	var L_SCARD_W_EOF_ErrorMessage="スマート カード ファイルの終わりに到達しました。";
	var L_SCARD_W_CANCELLEDBYUSER_ErrorMessage="ユーザーによって操作は取り消されました。";
	var L_SCARD_W_CARDNOTAUTHENTICATED_ErrorMessage="PIN がスマート カードに提示されました。";

	<%If "Enterprise"=sServerType Then%>
	;
	var L_TemplateLoadErrNoneFound_ErrorMessage="証明書のテンプレートは見つかりませんでした。証明書をこの CA から要求するアクセス許可がありません、または Active Directory にアクセス中にエラーが発生しました。";
	var L_TemplateLoadErrUnexpected_ErrorMessage="\"証明書のテンプレートの一覧を取得中に予期しないエラー (\"+sErrorNumber+\") が発生しました。\"";
	var L_TemplateCert_Text= "sFriendlyName+\" 証明書\"";
	<%End If%>

	// IE is not ready until XEnroll has been loaded
	var g_bOkToSubmit=false;
	var g_bSubmitPending=false;

	// some constants defined in wincrypt.h:
	var CRYPT_EXPORTABLE=1;
	var CRYPT_USER_PROTECTED=2;
	var CRYPT_MACHINE_KEYSET=0x20;
	var AT_KEYEXCHANGE=1;
	var AT_SIGNATURE=2;
	var CERT_SYSTEM_STORE_LOCATION_SHIFT=16;
	var CERT_SYSTEM_STORE_LOCAL_MACHINE_ID=2;
	var CERT_SYSTEM_STORE_LOCAL_MACHINE=CERT_SYSTEM_STORE_LOCAL_MACHINE_ID << CERT_SYSTEM_STORE_LOCATION_SHIFT;
	var ALG_CLASS_ANY=0
	var ALG_CLASS_SIGNATURE=1<<13;
	var ALG_CLASS_HASH=4<<13;
	var PROV_DSS=3;
	var PROV_DSS_DH=13;
	var PROV_DH_SCHANNEL=18;


	// convenience constants, for readability
	var KEY_USAGE_EXCH=0;
	var KEY_USAGE_SIG=1;
	var KEY_USAGE_BOTH=2;

	var XEKL_KEYSIZE_MIN=1;
	var XEKL_KEYSIZE_MAX=2;
	var XEKL_KEYSIZE_INC=3;
	var XEKL_KEYSIZE_DEFAULT=4;
	var XEKL_KEYSPEC_KEYX=1;
	var XEKL_KEYSPEC_SIG=2;

	// defaults
	var KEY_LEN_MIN_DEFAULT=384;
	var KEY_LEN_MAX_DEFAULT=16384;
	var KEY_LEN_MY_DEFAULT=1024;
	var KEY_LEN_INC_DEFAULT=8;

	// for key size
	var g_nCurKeySizeMax;
	var g_nCurKeySizeMin;
	var g_nCurKeySizeDefault;
	var g_nCurKeySizeInc;
	var g_bCSPUpdate;

	var g_nCurTemplateKeySizeMin = 0; //init to 0

	var XECR_PKCS10_V2_0=1;
	var XECR_PKCS7=2;
	var XECR_CMC=3;

	var XECT_EXTENSION_V1=1;
	var XECT_EXTENSION_V2=2;

	//================================================================
	// INITIALIZATION ROUTINES

	function removeV2KATemplate()
	{
		var CT_FLAG_ALLOW_PRIVATE_KEY_ARCHIVAL=0x00000001;
		//downlevel machines, no V2 templates with KA
		var nTemplateCount = document.UIForm.lbCertTemplate.length;
		var n, sTemplate, sCTEOID;
		for (n = nTemplateCount - 1; n > -1 ; --n)
		{
			sTemplate = document.UIForm.lbCertTemplate.options[n].value;
			sCTEOID = getTemplateStringInfo(CTINFO_INDEX_EXTOID, sTemplate);
			var lFlags=getTemplateValueInfo(CTINFO_INDEX_PRIVATEKEYFLAG, sTemplate);
			if ("" != sCTEOID && 0x0 != (lFlags & CT_FLAG_ALLOW_PRIVATE_KEY_ARCHIVAL))
			{
				//v2 template with KA
				document.UIForm.lbCertTemplate.options.remove(n);				
			}
		}
	}

	//----------------------------------------------------------------
	// This contains the functions we want executed immediately after load completes
	function postLoad() {
		// Load an XEnroll object into the page
		loadXEnroll("postLoadPhase2()"); 
		handleSaveReq();
		handleCMCFormat();
		<%If "Enterprise"=sServerType Then%>
			if (!isClientAbleToCreateCMC())
			{
				//downlevel machines
				removeV2KATemplate();
			}
		<%End If%>
	}
	function postLoadPhase2() {
		// continued from above
		var nResult;

		// get the CSP list
		nResult=GetCSPList();
		if (0!=nResult) {
			handleLoadError(nResult, L_CspLoadErrNoneFound_ErrorMessage, L_CspLoadErrUnexpected_ErrorMessage);
			return;
		}

		<%If "StandAlone"<>sServerType And 0<>nWriteTemplateResult Then%>
		handleLoadError(<%=nWriteTemplateResult%>, L_TemplateLoadErrNoneFound_ErrorMessage, L_TemplateLoadErrUnexpected_ErrorMessage);
		return;
		<%End If%>

		// Now we're ready to go
		g_bOkToSubmit=true;

		<%If "Enterprise"=sServerType Then%>
			handleTemplateChange();
		<%Else%>
			handleCSPChange();
		<%End If%>
		// dynamic styles are not preserved so
		// make sure dynamic UI is updated after 'back'
		handleKeyGen();
		handleMarkExport(false);
		handleExportKeys();
		<%If "StandAlone"=sServerType Then%>
		handleUsageOID(false);
		<%End If%>
	}

	//----------------------------------------------------------------
	// handle errors from GetCSPList() and GetTemplateList()
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

	<%If "StandAlone"=sServerType Then%>
	//----------------------------------------------------------------
	// handle the appearance of the text box when 'other...' is selected
	function handleUsageOID(bFocus) {
		if ("**"==document.UIForm.lbUsageOID.options[document.UIForm.lbUsageOID.selectedIndex].value) {
			spnEKUOther1.style.display='';
			spnEKUOther2.style.display='';
			if (bFocus) {
				document.UIForm.lbUsageOID.blur();
				document.UIForm.tbEKUOther.select();
				document.UIForm.tbEKUOther.focus();
			}
		} else {
			spnEKUOther1.style.display='none';
			spnEKUOther2.style.display='none';
		}
	}
	<%End If%>

	<%If "Enterprise"=sServerType Then%>
	//----------------------------------------------------------------


        function getTemplateValueInfo(nIndex, sTemplate)
	{
		var sValue=getTemplateStringInfo(nIndex, sTemplate);
		return parseInt(sValue);
	}

	// handle a change in the current template
	function isDNNeeded() {
		var sValue=getTemplateStringInfo(CTINFO_INDEX_OFFLINE, null);
		if ("O"==sValue)
		{
			//offline template needs DN
			return true;
		}

		//check template subject flag
		var lSubjectFlag = getTemplateValueInfo(CTINFO_INDEX_SUBJECTFLAG, null);
		var CT_FLAG_ENROLLEE_SUPPLIES_SUBJECT=0x00000001;
		return (0x0 != (lSubjectFlag & CT_FLAG_ENROLLEE_SUPPLIES_SUBJECT));
	}

        function isTemplateKeyArchival()
        {
		var CT_FLAG_ALLOW_PRIVATE_KEY_ARCHIVAL=0x00000001;
		var lFlags=getTemplateValueInfo(CTINFO_INDEX_PRIVATEKEYFLAG, null);
		return (0x0 != (lFlags & CT_FLAG_ALLOW_PRIVATE_KEY_ARCHIVAL));
	}

        function isSMimeCapabilities()
        {
		var CT_FLAG_INCLUDE_SYMMETRIC_ALGORITHMS=0x00000001;
		var lFlags=getTemplateValueInfo(CTINFO_INDEX_ENROLLFLAG, null);
		return (0x0 != (lFlags & CT_FLAG_INCLUDE_SYMMETRIC_ALGORITHMS));
	}

	function getTemplateMinKeySize()
	{
		var lKeyFlags = getTemplateValueInfo(CTINFO_INDEX_KEYFLAG, null);
		return (lKeyFlags & 0xFFFF0000) >> 16;
	}

	function updateCSPList()
	{
                //get csp list separated from template data
                var sCSPList = getTemplateStringInfo(CTINFO_INDEX_CSPLIST, null);
		if ("" != sCSPList)
		{
			updateCSPListFromStrings(sCSPList);
		}
		else
		{
			//remove current csps from list
			//strange reasons this remove code can't be in GetCSPList
			var n;
			var nCSP = document.UIForm.lbCSP.length;
			for (n = 0; n < nCSP-1; ++n)
			{
				document.UIForm.lbCSP.remove(0);
			}
			GetCSPList();
		}
	}

	//----------------------------------------------------------------
	// handle a change in the current template
	function handleTemplateChange()
	{
		if (false==isDNNeeded()) {
			spnIDInfo.style.display="none";
		} else {
			spnIDInfo.style.display="";
		}

                //update csp list from the template
                updateCSPList();
		handleCSPChange();

		//handle key spec
                var lKeySpec = getTemplateValueInfo(CTINFO_INDEX_KEYSPEC, null);
                var fDisabled = true;
                if ((0x0 != (AT_KEYEXCHANGE & lKeySpec)) &&
                    (0x0 != (AT_SIGNATURE & lKeySpec)) )
                {
			document.UIForm.rbKeyUsage[KEY_USAGE_BOTH].checked=true;
                }
                else if (0x0 != (AT_KEYEXCHANGE & lKeySpec))
                {
			document.UIForm.rbKeyUsage[KEY_USAGE_EXCH].checked=true;
                }
                else if (0x0 != (AT_SIGNATURE & lKeySpec))
                {
			document.UIForm.rbKeyUsage[KEY_USAGE_SIG].checked=true;
                }
                else
                {
			document.UIForm.rbKeyUsage[KEY_USAGE_BOTH].checked=true;
                        fDisabled = false;
                }
                document.UIForm.rbKeyUsage[KEY_USAGE_BOTH].disabled=fDisabled;
                document.UIForm.rbKeyUsage[KEY_USAGE_EXCH].disabled=fDisabled;
                document.UIForm.rbKeyUsage[KEY_USAGE_SIG].disabled=fDisabled;

		//update exportable control
		var lPrivateKeyFlags = getTemplateValueInfo(CTINFO_INDEX_PRIVATEKEYFLAG, null);
		var CT_FLAG_EXPORTABLE_KEY = 0x10;
		document.UIForm.cbMarkKeyExportable.checked = (0x0 != (lPrivateKeyFlags & CT_FLAG_EXPORTABLE_KEY));
		handleMarkExport(true);

		//update strong key protection control
                var CT_FLAG_STRONG_KEY_PROTECTION_REQUIRED = 0x20; 
                document.UIForm.cbStrongKey.checked = (0x0 != (lPrivateKeyFlags & CT_FLAG_STRONG_KEY_PROTECTION_REQUIRED));
                handleStrongKeyAndLMStore(true);

		//update template min key size
		g_nCurTemplateKeySizeMin = getTemplateMinKeySize();

                //update key size
		handleKeyUsageChange(false);

		//update CMC related
		handleCMCFormat();

		var lRASignatures = getTemplateValueInfo(CTINFO_INDEX_RASIGNATURE, null);
		var fSave = 0 < lRASignatures;
		//enforce save to file, can't submit if signing
		document.UIForm.cbSaveRequest.checked = fSave;
		document.UIForm.cbSaveRequest.disabled = fSave;
		handleSaveReq();

	}
	<%End If%>

	//----------------------------------------------------------------
	// handle a change in the current CSP
	function handleCSPChange() {

		if (0 == document.UIForm.lbCSP.length)
		{
			//no csp, disable submit button
			document.UIForm.btnSubmit.disabled = true;
			return;
		}
		else
		{
			document.UIForm.btnSubmit.disabled = false;
		}
		var nCSPIndex=document.UIForm.lbCSP.selectedIndex;
		XEnroll.ProviderName=document.UIForm.lbCSP.options[nCSPIndex].text;
		var nProvType=document.UIForm.lbCSP.options[nCSPIndex].value;
		XEnroll.ProviderType=nProvType;
		<%If "Enterprise"=sServerType Then%>
                var nTemplateKeySpec = getTemplateValueInfo(CTINFO_INDEX_KEYSPEC, null);
		<%End If%>

		// update the key spec options. If we support both, default to key exchange
		var nSupportedKeyUsages=XEnroll.GetSupportedKeySpec();
		if (0==nSupportedKeyUsages) {
			nSupportedKeyUsages=AT_SIGNATURE | AT_KEYEXCHANGE;
		}

		<%If "Enterprise"=sServerType Then%>
		if (0==nTemplateKeySpec) {
			nTemplateKeySpec=AT_SIGNATURE | AT_KEYEXCHANGE;
		}
		nSupportedKeyUsages = nTemplateKeySpec & nSupportedKeyUsages;
		<%End If%>

		if (PROV_DSS==nProvType || PROV_DSS_DH==nProvType || PROV_DH_SCHANNEL==nProvType) {
			nSupportedKeyUsages=AT_SIGNATURE;
		}

		if (0 == nSupportedKeyUsages)
		{
			spnBadCSPForKeySpecMsg.innerHTML=eval(L_CSPNotSupportTemplateKeySpec_Message);
			trBadCSPForKeySpec.style.display="";
		} else {
			trBadCSPForKeySpec.style.display="none";
		}

		if (nSupportedKeyUsages&AT_SIGNATURE) {
			spnKeyUsageSignature.style.display="";
			document.UIForm.rbKeyUsage[KEY_USAGE_SIG].checked=true;
		} else {
			spnKeyUsageSignature.style.display="none";
		}

		if (nSupportedKeyUsages&AT_KEYEXCHANGE) {
			spnKeyUsageKeyExchange.style.display="";
			document.UIForm.rbKeyUsage[KEY_USAGE_EXCH].checked=true;
		} else {
			spnKeyUsageKeyExchange.style.display="none";
		}

		if ((AT_SIGNATURE|AT_KEYEXCHANGE)==nSupportedKeyUsages) {
			spnKeyUsageBoth.style.display="";
			document.UIForm.rbKeyUsage[KEY_USAGE_BOTH].checked=true;
		} else {
			spnKeyUsageBoth.style.display="none";
		}

		handleKeyUsageChange(true);
		UpdateHashAlgList(nProvType);
	}

	//----------------------------------------------------------------
	// two cases invoke handleKeyUsageChange:
        // 1) csp selection change
        // 2) exchange vs. signature change
	function handleKeyUsageChange(bCSPChange) {
		// get the min, max, and default length from the CSP
		var bExchange=document.UIForm.rbKeyUsage[KEY_USAGE_EXCH].checked || document.UIForm.rbKeyUsage[KEY_USAGE_BOTH].checked ;

		g_nCurKeySizeMax=MyGetKeyLen(XEKL_KEYSIZE_MAX, bExchange);
		g_nCurKeySizeMin=MyGetKeyLen(XEKL_KEYSIZE_MIN, bExchange);
		
		<%If "Enterprise"=sServerType Then%>
		if (0 != g_nCurTemplateKeySizeMin)
		{
			g_nCurKeySizeMin=Math.max(g_nCurKeySizeMin, g_nCurTemplateKeySizeMin);
		}
		<%End If%>
		g_nCurKeySizeDefault=MyGetKeyLen(XEKL_KEYSIZE_DEFAULT, bExchange);
		g_nCurKeySizeInc=MyGetKeyLen(XEKL_KEYSIZE_INC, bExchange);

                // set to default lenth
                if ("0"==document.UIForm.tbKeySize.value || true == bCSPChange)
                {
                    //"0" likely init load or typed in, not bad go default
                    // or csp changed, set to default length
                    document.UIForm.tbKeySize.value = g_nCurKeySizeDefault;
                }

		// show the min and max
		spnKeySizeMin.innerText=g_nCurKeySizeMin;
		spnKeySizeMax.innerText=g_nCurKeySizeMax;

		// keep the key size in bounds
		var nKeySize=parseInt(document.UIForm.tbKeySize.value);
		if (isNaN(nKeySize) || nKeySize>g_nCurKeySizeMax) {
			document.UIForm.tbKeySize.value=g_nCurKeySizeMax;
		} else if (nKeySize<g_nCurKeySizeMin) { //>
			document.UIForm.tbKeySize.value=g_nCurKeySizeMin;
		}

		// update list of valid common key sizes
		var nPowerSize=128;
		var sCommonKeys="";
		while (nPowerSize<g_nCurKeySizeMin) { //>
			nPowerSize*=2;
		}
		while (nPowerSize<=g_nCurKeySizeMax) {
			sCommonKeys+=getKeySizeLinkHtmlString(nPowerSize)+" ";
			nPowerSize*=2;
		}
		spnKeySizeCommon.innerHTML=sCommonKeys;
		handleKeySizeChange();
	}

	//----------------------------------------------------------------
	function getKeySizeLinkHtmlString(nKeySize) {
		return "<Span tabindex=0 Style=\"cursor:hand; color:#0000FF; text-decoration:underline;\""
			+" OnContextMenu=\"return false;\""
			+" OnMouseOver=\"window.status='"+eval(L_SetKeySize_Message)+"';return true;\""
			+" OnMouseOut=\"window.status='';return true;\""
			+" OnMouseUp=\"window.status='"+eval(L_SetKeySize_Message)+"';return true;\""
			+" OnKeyDown=\"if (13==event.keyCode) {document.UIForm.tbKeySize.value='"+nKeySize+"';blur();return false;} else if (9==event.keyCode) {return true;};return false;\""
			+" OnClick=\"document.UIForm.tbKeySize.value='"+nKeySize+"';blur();return false;\">"
			+nKeySize+"</Span>";
	}

	//----------------------------------------------------------------
	// morphing routine
	function handleSaveReq() {
		if (document.UIForm.cbSaveRequest.checked) {
			spnSaveRequest.style.display='';
			document.UIForm.btnSubmit.style.display='none';
			document.UIForm.btnSave.style.display='';
			spnSubmitAttrLable.style.display='none';
			spnSubmitAttrBox.style.display='none';
		} else {
			spnSaveRequest.style.display='none';
			document.UIForm.btnSubmit.style.display='';
			document.UIForm.btnSave.style.display='none';
			spnSubmitAttrLable.style.display='';
			spnSubmitAttrBox.style.display='';
		}
	}

	//----------------------------------------------------------------
	// morphing routine
	function handleMarkExport(fFromTemplate) {
		<%If bEnableExportKeyToFile Then%>
		if (document.UIForm.cbMarkKeyExportable.checked) {
			spnMarkKeyExportable.style.display='';
			document.UIForm.cbMarkKeyExportable.disabled = false;
		} else {
			spnMarkKeyExportable.style.display='none';
			if (fFromTemplate)
			{
				//disable it to enforce template non-exportable
				document.UIForm.cbMarkKeyExportable.disabled = true;
			}
			else
			{
				//enable
				document.UIForm.cbMarkKeyExportable.disabled = false;
			}
		}
		<%End If%>
	}

	//----------------------------------------------------------------
	// morphing routine
	function handleExportKeys() {
		<%If bEnableExportKeyToFile Then%>
		if (document.UIForm.cbExportKeys.checked) {
			spnExportKeys.style.display='';
		} else {
			spnExportKeys.style.display='none';
		}
		<%End If%>
	}

	//----------------------------------------------------------------
	// morphing routine
	function handleKeyGen() {
		if (document.UIForm.rbKeyGen[0].checked) {
			// create new keyset
			trGenContName.style.display='';
			trGenContNameSpc.style.display='';
			trKeyGenWarn.style.display='none';

			handleGenContName();
			<%If "Enterprise"=sServerType Then%>
				handleTemplateChange();
			<%Else%>
				handleStrongKeyAndLMStore(false);
			<%End If%>

			trMarkExport.style.display='';
			trMarkExportSpc.style.display='';
		} else {
			// Use existing key set
			trGenContName.style.display='none';
			trGenContNameSpc.style.display='none';
			trKeyGenWarn.style.display='';

			handleGenContName();
			handleStrongKeyAndLMStore(false);

			document.UIForm.cbMarkKeyExportable.checked=false;
			trMarkExport.style.display='none';
			trMarkExportSpc.style.display='none';
 		}
	}

	//----------------------------------------------------------------
	// morphing routine
	function handleGenContName() {
		if (document.UIForm.rbGenContName[0].checked && document.UIForm.rbKeyGen[0].checked) {
			trContName.style.display='none';
			trContNameSpc.style.display='none';
		} else {
			trContName.style.display='';
			trContNameSpc.style.display='';
		}
	}

	//----------------------------------------------------------------
	// morphing routine
	function handleSetContainer() {
		if (document.UIForm.cbSetContainer.checked) {
			spnNewContainer.style.display='';
		} else {
			spnNewContainer.style.display='none';
		}
	}

	//----------------------------------------------------------------
	// morphing routine
	function handleKeySizeChange() {
		var sKeySize = document.UIForm.tbKeySize.value;
		if (0 == sKeySize.indexOf("0"))
		{
			//first digit is 0, wipe it out
			document.UIForm.tbKeySize.value = "";
			return;
		}
		var nKeySize=parseInt(sKeySize);
		if (isNaN(nKeySize)) {
			nKeySize=0;
		}
		if (nKeySize>2048) {
			trKeySizeWarn.style.display='';
		} else {
			trKeySizeWarn.style.display='none';
		}
		if (nKeySize<g_nCurKeySizeMin || nKeySize>g_nCurKeySizeMax || 0!=nKeySize%g_nCurKeySizeInc) {
			// clamp the current key size to be within the range
			var nCloseBelow=nKeySize;
			if (nCloseBelow<g_nCurKeySizeMin) { //>
				nCloseBelow=g_nCurKeySizeMin;
			} else if (nCloseBelow>g_nCurKeySizeMax) { 
				nCloseBelow=g_nCurKeySizeMax;
			}
			var nCloseAbove=nCloseBelow;
			// find closest values above and below
			nCloseBelow-=nCloseBelow%g_nCurKeySizeInc;
			nCloseAbove+=(g_nCurKeySizeInc-nCloseAbove%g_nCurKeySizeInc)%g_nCurKeySizeInc;
			var sCloseAbove=getKeySizeLinkHtmlString(nCloseAbove);
			var sCloseBelow=getKeySizeLinkHtmlString(nCloseBelow);
			if (g_nCurKeySizeMax < g_nCurTemplateKeySizeMin) {
				spnKeySizeBadMsg.innerHTML=eval(L_WarningTemplateKeySize_Message);
			} else if (nCloseAbove==nCloseBelow) {
				spnKeySizeBadMsg.innerHTML=eval(L_RecommendOneKeySize_Message);
			} else {
				spnKeySizeBadMsg.innerHTML=eval(L_RecommendTwoKeySizes_Message);
			}
			trKeySizeBad.style.display="";
			trKeySizeBadSpc.style.display="";
		} else {
			trKeySizeBad.style.display="none";
			trKeySizeBadSpc.style.display="none";
		}
	}

	//----------------------------------------------------------------
	// morphing routine
	function handleStrongKeyAndLMStore(fFromTemplate) {
                // If we took the value from the template, and the checkbox is checked, the disable it
                if (document.UIForm.cbStrongKey.checked && fFromTemplate) { 
                    //disable it to enforce template non-exportable
                    document.UIForm.cbStrongKey.disabled = true;
                } else { 
                    document.UIForm.cbStrongKey.disabled = false	;
                }

		if (document.UIForm.cbStrongKey.checked && document.UIForm.rbKeyGen[0].checked) {
			trLMStoreSpc.style.display='none';
			trLMStore.style.display='none';
			document.UIForm.cbLocalMachineStore.checked=false;
		} else {
			trLMStoreSpc.style.display='';
			trLMStore.style.display='';
		}

		if (document.UIForm.cbLocalMachineStore.checked || !document.UIForm.rbKeyGen[0].checked) {
			trStrongKeySpc.style.display='none';
			trStrongKey.style.display='none';
			document.UIForm.cbStrongKey.checked=false;
		} else {
			trStrongKeySpc.style.display='';
			trStrongKey.style.display='';
		}
	}
	//----------------------------------------------------------------
	// handle CMC Format
	function handleCMCFormat() {
		if (isClientAbleToCreateCMC())
		{
			<%If "Enterprise"=sServerType Then%>
			//change request format controls
			if (isTemplateKeyArchival())
			{
				//enforce CMC
				document.UIForm.rbRequestFormat[0].disabled=true;
				document.UIForm.rbRequestFormat[0].checked=true;
				document.UIForm.rbRequestFormat[1].disabled=true;
			}
			else
			{
				document.UIForm.rbRequestFormat[0].disabled=false;
				document.UIForm.rbRequestFormat[1].disabled=false;
			}
			<%End If%>
		}
		else
		{
			//no cmc, disable it, only pkcs10
			document.UIForm.rbRequestFormat[0].disabled=true;
			document.UIForm.rbRequestFormat[1].disabled=true;
			document.UIForm.rbRequestFormat[1].checked=true;
		}
	}

	//================================================================
	// SUBMIT ROUTINES

	//----------------------------------------------------------------
	// determine what to do when the submit button is pressed
	function goNext() {
		if (false==g_bOkToSubmit) {
			alert(L_StillLoading_ErrorMessage);
		} else if (true==g_bSubmitPending) {
			// ignore, because we are already prcessing a request.
		} else {
			SubmitRequest();
		}
	}
		
	//----------------------------------------------------------------
	// check for invalid characters and empty strings
	function isValidIA5String(sSource) {
		var nIndex;
		for (nIndex=sSource.length-1; nIndex>=0; nIndex--) {
			if (sSource.charCodeAt(nIndex)>127) {  // NOTE: this is better, but not compatible with old browsers.
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
	// check for invalid characters in an OID
	function isValidOid(sSource) {
		var nIndex, ch;
		if (0==sSource.length) {
			return true;
		}
		for (nIndex=sSource.length-1; nIndex>=0; nIndex--) {
			ch=sSource.charAt(nIndex)
			if (ch!="." && ch!="," && (ch<"0" || ch>"9")) {
				return false;
			}
		}
		return true;
	}

	//----------------------------------------------------------------
	// set a label to normal style
	function markLabelNormal(spn) {
		spn.style.color="#000000";
		spn.style.fontWeight='normal';
	}

	//----------------------------------------------------------------
	// set a label to error state
	function markLabelError(spn) {
		spn.style.color='#FF0000';
		spn.style.fontWeight='bold';
	}

	//----------------------------------------------------------------
	// check that the form has data in it
	function validateRequest() {
		markLabelNormal(spnNameLabel);
		markLabelNormal(spnEmailLabel);
		markLabelNormal(spnCompanyLabel);
		markLabelNormal(spnDepartmentLabel);
		markLabelNormal(spnCityLabel);
		markLabelNormal(spnStateLabel);
		markLabelNormal(spnCountryLabel);
		
		var bOK=true;

		<%If "Enterprise"=sServerType Then%>	
		if (true==isDNNeeded()) {
		<%End If%>
			var fldFocusMe=null;
			if (false==isValidCountryField(document.UIForm.tbCountry)) {
				bOK=false;
				fldFocusMe=document.UIForm.tbCountry;
				markLabelError(spnCountryLabel);
			}
			// document.UIForm.tbState.value OK
			// document.UIForm.tbLocality.value OK
			// document.UIForm.tbOrgUnit.value OK
			// document.UIForm.tbOrg.value OK
			if (false==isValidIA5String(document.UIForm.tbEmail.value))
			{
				bOK=false;
				fldFocusMe=document.UIForm.tbEmail;
				markLabelError(spnEmailLabel);
			}
			<%If "StandAlone"=sServerType Then%>
			if ("1.3.6.1.5.5.7.3.4"==document.UIForm.lbUsageOID.value && ""==document.UIForm.tbEmail.value)
			{
				bOK=false;
				fldFocusMe=document.UIForm.tbEmail;
				markLabelError(spnEmailLabel);
			}
			if (""==document.UIForm.tbCommonName.value)
			{
				bOK=false;
				fldFocusMe=document.UIForm.tbCommonName;
				markLabelError(spnNameLabel);
			}
			<%End If%>

			if (false==bOK) {
				spnFixTxt.style.display='';
				window.scrollTo(0,0);
				fldFocusMe.focus();
			}
		<%If "Enterprise"=sServerType Then%>	
		} // <- End if offline template
		<%End If%>

		<%If "StandAlone"=sServerType Then%>
		// Check the OID field
		if (true==bOK) {
			if ("**"==document.UIForm.lbUsageOID.options[document.UIForm.lbUsageOID.selectedIndex].value
				&& false==isValidOid(document.UIForm.tbEKUOther.value)) {
				alert(L_BadOid_ErrorMessage);
				document.UIForm.tbEKUOther.focus();
				bOK=false;
			}
		}
		<%End If%>

		// Check the keysize field
		if (true==bOK) {
			var nKeySize=parseInt(document.UIForm.tbKeySize.value);
			var sMessage;
			if (isNaN(nKeySize)) {
				sMessage=L_KeySizeNotNumber_ErrorMessage;
				bOK=false;
			} else if (g_nCurTemplateKeySizeMin > g_nCurKeySizeMax) {
				sMessage=eval(L_TemplateKeySizeTooBig_ErrorMessage);
				bOK = false;
			} else if (nKeySize < g_nCurKeySizeMin || nKeySize > g_nCurKeySizeMax || 0!=nKeySize%g_nCurKeySizeInc) {
				sMessage=eval(L_KeySizeBadNumber_ErrorMessage);
				bOK=false;
			}
			if (false==bOK) {
				alert (sMessage);
				document.UIForm.tbKeySize.focus();
			}
		}

		// Check the container name
		if (true==bOK) {
			if (document.UIForm.rbKeyGen[1].checked
				|| (document.UIForm.rbKeyGen[0].checked && document.UIForm.rbGenContName[1].checked)) {
				if (""==document.UIForm.tbContainerName.value) {
					bOK=false;
					alert(L_NoCntnrName_ErrorMessage);
					document.UIForm.tbContainerName.focus();
				}
			}
		}

		<%If bEnableExportKeyToFile Then%>
		// Check the exported private key file name
		if (true==bOK) {
			if (document.UIForm.rbKeyGen[0].checked 
				&& document.UIForm.cbMarkKeyExportable.checked 
				&& document.UIForm.cbExportKeys.checked) {
				if (""==document.UIForm.tbExportKeyFile.value) {
					bOK=false;
					alert(L_NoExportFileName_ErrorMessage);
					document.UIForm.tbExportKeyFile.focus();
				}
			}
		}
		<%End If%>

		// Check the saved-request file name
		if (true==bOK) {
			if (document.UIForm.cbSaveRequest.checked) {
				if (""==document.UIForm.tbSaveReqFile.value) {
					bOK=false;
					alert(L_NoSaveReqFileName_ErrorMessage);
					document.UIForm.tbSaveReqFile.focus();
				}
			}
		}

		return bOK;
	}

	//----------------------------------------------------------------
	function SubmitRequest() {
		g_bSubmitPending=true;

		// check that the form is filled in
		spnErrorTxt.style.display='none';
		spnFixTxt.style.display='none';
		if (false==validateRequest()) {
			g_bSubmitPending=false;
			return;
		}

		// show a nice message since request creation can take a while
		ShowTransientMessage(L_Generating_Message);

		// Make the message show up on the screen, 
		// then continue with 'SubmitRequest':
		// Pause 10 mS before executing phase 2, 
		// so screen will have time to repaint.
		setTimeout("SubmitRequestPhase2();", 10); 
	}
	function SubmitRequestPhase2() {
		// continued from above

		<%If "StandAlone"=sServerType Then%> 
		//
		// Stand-Alone Options
		//

		// set the extended key usage and certificate request 'friendly type'
		var nUsageIndex=document.UIForm.lbUsageOID.selectedIndex;
		var sCertUsage;
		if ("**"==document.UIForm.lbUsageOID.options[nUsageIndex].value) {
			sCertUsage=document.UIForm.tbEKUOther.value;
			document.SubmittedData.FriendlyType.value=eval(L_UserEKUCert_Text);
		} else {
			sCertUsage=document.UIForm.lbUsageOID.options[nUsageIndex].value;
			document.SubmittedData.FriendlyType.value=document.UIForm.lbUsageOID.options[nUsageIndex].text;
		}

		<%Else 'Enterprise%>
		//
		// Enterprise Options
		//

		// get cert template info
		var lCTEVer = XECT_EXTENSION_V1;
		var lCTEMajor = 0;
		var bCTEfMinor = false;
		var lCTEMinor = 0;
		var sRealName = getTemplateStringInfo(CTINFO_INDEX_REALNAME, null);
		var sFriendlyName = getTemplateStringInfo(CTINFO_INDEX_FRIENDLYNAME, null); 
		var sCTEOID = getTemplateStringInfo(CTINFO_INDEX_EXTOID, null);
		if ("" == sCTEOID) {
			//must v1 template, get template name
			sCTEOID = sRealName;			
		} else {
			// v2 template
			lCTEVer = XECT_EXTENSION_V2;
			lCTEMajor = getTemplateValueInfo(CTINFO_INDEX_EXTMAJ, null);
			bCTEfMinor = getTemplateValueInfo(CTINFO_INDEX_EXTFMIN, null);
			lCTEMinor = getTemplateValueInfo(CTINFO_INDEX_EXTMIN, null);
		}
		// set the cert template
		vbAddCertTypeToRequestEx(lCTEVer, sCTEOID, lCTEMajor, bCTEfMinor, lCTEMinor);
		document.SubmittedData.FriendlyType.value=eval(L_TemplateCert_Text);

		var sCertUsage=""; // ignored

		<%End If 'StandAlone or Enterprise%> 
		//
		// Common
		//

		// set the identifying info
		var sDistinguishedName="";
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
		<%If "Enterprise"=sServerType Then%> 
		if (false==isDNNeeded()) {
			sDistinguishedName="";
		}
		<%End If%>

		// append the local date to the type
		document.SubmittedData.FriendlyType.value+=" ("+(new Date()).toLocaleString()+")";

		//
		// Key Options subheading:
		//

		// set the 'SaveCert' flag to install the cert instead of saving
		document.SubmittedData.SaveCert.value="no";
		     
		// set the CSP
		var nCSPIndex=document.UIForm.lbCSP.selectedIndex;
		XEnroll.ProviderName=document.UIForm.lbCSP.options[nCSPIndex].text;
		XEnroll.ProviderType=document.UIForm.lbCSP.options[nCSPIndex].value;

		// set the key size (the upper 16 bits of GenKeyFlags)
		//  note: this value has already been validated
		var nKeySize=parseInt(document.UIForm.tbKeySize.value);
		XEnroll.GenKeyFlags=nKeySize<<16;

		// set the KeyUsage
		if (document.UIForm.rbKeyUsage[KEY_USAGE_EXCH].checked) {
			XEnroll.KeySpec=AT_KEYEXCHANGE;
			XEnroll.LimitExchangeKeyToEncipherment=true;
		} else if (document.UIForm.rbKeyUsage[KEY_USAGE_SIG].checked) {
			XEnroll.KeySpec=AT_SIGNATURE;
			XEnroll.LimitExchangeKeyToEncipherment=false;
		} else { // KEY_USAGE_BOTH
			XEnroll.KeySpec=AT_KEYEXCHANGE;
			XEnroll.LimitExchangeKeyToEncipherment=false;
		}

		// set the 'use existing key set' flag
		if (document.UIForm.rbKeyGen[0].checked) {
			XEnroll.UseExistingKeySet=false;
			if (document.UIForm.rbGenContName[1].checked) {
				XEnroll.ContainerName=document.UIForm.tbContainerName.value;
			}

			// set 'Strong private key protection'
			//   note: upper 16 bits already set as key size
			if (document.UIForm.cbStrongKey.checked) {
				XEnroll.GenKeyFlags|=CRYPT_USER_PROTECTED;
			}

			// mark the keys as exportable
			if (document.UIForm.cbMarkKeyExportable.checked) {
				XEnroll.GenKeyFlags|=CRYPT_EXPORTABLE;

				<%If bEnableExportKeyToFile Then%>
				// set the key export file (.pvk) and save the cert instead of installing
				if (document.UIForm.cbExportKeys.checked) {
					XEnroll.PVKFileName=document.UIForm.tbExportKeyFile.value;
					document.SubmittedData.SaveCert.value="yes";
				} 
				<%End If%>
			}

		} else {
			// set the 'use existing key set' flag
			XEnroll.UseExistingKeySet=true;
			XEnroll.ContainerName=document.UIForm.tbContainerName.value;
		}


		// place the keys in the local machine store
		if (document.UIForm.cbLocalMachineStore.checked) {

			// the keys attached to the dummy request cert go in the local machine store
			XEnroll.RequestStoreFlags=CERT_SYSTEM_STORE_LOCAL_MACHINE;

			// used in CryptAcquireContext
			XEnroll.ProviderFlags=CRYPT_MACHINE_KEYSET;

			// the keys attached to the final cert also go in the local machine store
			document.SubmittedData.TargetStoreFlags.value=CERT_SYSTEM_STORE_LOCAL_MACHINE;
		} else {

			// the keys attached to the final cert also go in the user store
			document.SubmittedData.TargetStoreFlags.value=0; // 0=Use default (=user store)
		}

		var dwCreateRequestFlag = XECR_CMC;
		if (document.UIForm.rbRequestFormat[1].checked)
		{
			dwCreateRequestFlag = XECR_PKCS10_V2_0;
		}

		<%If "Enterprise"=sServerType Then%>
		//SMIME capabilities
		XEnroll.EnableSMIMECapabilities = isSMimeCapabilities();

                //Key archival
		if (isTemplateKeyArchival())
		{
			var nResult = SetPrivateKeyArchiveCertificate(); //call VB
			if (0 != nResult)
			{
				handleError(nResult);
				return;
			}
		}
		<%End If%>

		if ("" != document.UIForm.tbFriendlyName.value)
		{
			//set friendly name property
			var CERT_FRIENDLY_NAME_PROP_ID=11;
			var XECP_STRING_PROPERTY=1;
			XEnroll.addBlobPropertyToCertificate(CERT_FRIENDLY_NAME_PROP_ID, XECP_STRING_PROPERTY, document.UIForm.tbFriendlyName.value);
		}

		//
		// Additional Options subheading:
		//

		// set the hash algorithm     
		var nHashIndex=document.UIForm.lbHashAlgorithm.selectedIndex;
		XEnroll.HashAlgID=document.UIForm.lbHashAlgorithm.options[nHashIndex].value;

		// set any extra attributes
		var sAttrib=document.UIForm.taAttrib.value;
		if (sAttrib.lastIndexOf("\r\n")!=sAttrib.length-2 && sAttrib.length>0) {
			sAttrib=sAttrib+"\r\n";
		}

		// for interop debug purposes
		sAttrib+="UserAgent:<%=Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))%>\r\n";

		document.SubmittedData.CertAttrib.value=sAttrib;

		// we are submitting a new request
		document.SubmittedData.Mode.value='newreq';

		// 
		// Create the request
		//

		var nResult;
		var HRESULT_ERROR_CANCELLED=0x800704c7;
		var SCARD_W_CANCELLED_BY_USER=0x8010006e;
		var PVK_HELPER_PASSWORD_CANCEL=0x80097004;

		if (document.UIForm.cbSaveRequest.checked) {

			// build and save the certificate request
			var sSaveReqFile=document.UIForm.tbSaveReqFile.value;
			nResult=CreateAndSaveRequest(dwCreateRequestFlag, sDistinguishedName, sCertUsage, sSaveReqFile); // ask VB to do it, since it can handle errors

		} else {
			// build the certificate request
			nResult=CreateRequest(dwCreateRequestFlag, sDistinguishedName, sCertUsage); // ask VB to do it, since it can handle errors
		}
		if (0 == nResult)
		{
			//always get thumbprint in case of pending
			document.SubmittedData.ThumbPrint.value=XEnroll.ThumbPrint;
		}

		// hide the message box
		HideTransientMessage();

		// reset XEnroll so the user can select a different CSP, etc.
		XEnroll.reset();
		// however, make sure it still matches the UI.
		XEnroll.ProviderName=document.UIForm.lbCSP.options[nCSPIndex].text;
		XEnroll.ProviderType=document.UIForm.lbCSP.options[nCSPIndex].value;

		// deal with an error if there was one
		if (0!=nResult) {
			g_bSubmitPending=false;
			if (0==(SCARD_W_CANCELLED_BY_USER^nResult) ||
			    0==(PVK_HELPER_PASSWORD_CANCEL^nResult))
			{
				//cancelled
				nResult=0;
				return;
			}


			<%If "Enterprise"=sServerType Then%>
			if (isTemplateKeyArchival() && !document.UIForm.cbMarkKeyExportable.checked) 
			{
				//they've tried to create an archivable key
				handleError2(nResult, L_SugCauseCryptArchivableNotSupp_ErrorMessage); 					
			} else {
				// use the regular error handling
				handleError(nResult);
			}
			<%Else%>
			// just use the regular error handling in the standalone case
			handleError(nResult);		
			<%End If%>

			return;
		}

		// check for special "no submit" case
		if (document.UIForm.cbSaveRequest.checked) {

			// just inform the user that it went OK, but don't submit
			alert(L_RequestSaved_Message);
			g_bSubmitPending=false;

		} else {

			// put up a new wait message
			ShowTransientMessage(L_Waiting_Message);

			// Submit the cert request and move forward in the wizard
			document.SubmittedData.submit();
		}
	}

	
	//----------------------------------------------------------------
	function handleError(nResult) {
		handleError2(nResult, 0);
	}

	//----------------------------------------------------------------
	function handleError2(nResult, sSugCauseIN) { 
		var sSugCause=L_SugCauseNone_ErrorMessage;
		var sErrorName=L_ErrNameUnknown_ErrorMessage;
		// analyze the error - funny use of XOR ('^') because obvious choice '==' doesn't work
		if (0==(0x80090008^nResult)) {
			sErrorName="NTE_BAD_ALGID";
			sSugCause=L_SugCauseBadSetting2_ErrorMessage;
		} else if (0==(0x80090016^nResult)) {
			sErrorName="NTE_BAD_KEYSET";
			if (document.UIForm.rbKeyGen[0].checked) {
				sSugCause=L_SugCauseBadCSP_ErrorMessage;
			} else {
				sSugCause=L_SugCauseBadKeyContainer_ErrorMessage;
			}
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
			sSugCause=L_SugCauseBadSetting2_ErrorMessage;
		} else if (0==(0x8009000F^nResult)) {
			sErrorName="NTE_EXISTS";
			sSugCause=L_SugCauseExistKeyContainer_ErrorMessage;
		} else if (0==(0x80092002^nResult)) {
			sErrorName="CRYPT_E_BAD_ENCODE";
			//sSugCause="";
		} else if (0==(0x80092022^nResult)) {
			sErrorName="CRYPT_E_INVALID_IA5_STRING";
			sSugCause=L_SugCauseBadChar_ErrorMessage;
		} else if (0==(0x80092023^nResult)) {
			sErrorName="CRYPT_E_INVALID_X500_STRING";
			sSugCause=L_SugCauseBadChar_ErrorMessage;
		} else if (0==(0x80070003^nResult)) {
			sErrorName="ERROR_PATH_NOT_FOUND";
			sSugCause=L_SugCauseBadFileName_ErrorMessage;
		} else if (0==(0x80070103^nResult)) {
			sErrorName="ERROR_NO_MORE_ITEMS";
			sSugCause=L_SugCauseBadHash_ErrorMessage;
		} else if (0==(0x8007007B^nResult)) {
			sErrorName="ERROR_INVALID_NAME";
			sSugCause=L_SugCauseBadFileName_ErrorMessage;
		} else if (0==(0x80070015^nResult)) {
			sErrorName="ERROR_NOT_READY";
			sSugCause=L_SugCauseBadDrive_ErrorMessage;
		} else if (0==(0x8007007F^nResult)) {
			sErrorName="ERROR_PROC_NOT_FOUND";
			sSugCause=L_DownLevelClients_ErrorMessage;
		} else if (0==(0x800704C7^nResult)) { 
			sErrorNamge="ERROR_CANCELLED";
			sSugCause=L_SugCauseCancelled_ErrorMessage;
		} else if (0==(0x80100006^nResult)) {
			sErrorName = "SCARD_E_NO_MEMORY";
			sSugCause = L_SCARD_E_NOMEMORYMSG_ErrorMessage;
		} else if (0==(0x80100007^nResult)) {
			sErrorName = "SCARD_F_WAITED_TOO_LONG";
			sSugCause = L_SCARD_F_WAITEDTOOLONG_ErrorMessage;
		} else if (0==(0x80100008^nResult)) {
			sErrorName = "SCARD_E_INSUFFICIENT_BUFFER";
			sSugCause = L_SCARD_E_INSUFFICIENTBUFFER_ErrorMessage;
		} else if (0==(0x80100009^nResult)) {
			sErrorName = "SCARD_E_UNKNOWN_READER";
			sSugCause = L_SCARD_E_UNKNOWNREADER_ErrorMessage;
		} else if (0==(0x8010000C^nResult)) {
			sErrorName = "SCARD_E_NO_SMARTCARD";
			sSugCause = L_SCARD_E_NOSMARTCARD_ErrorMessage;
		} else if (0==(0x8010000D^nResult)) {
			sErrorName = "SCARD_E_UNKNOWN_CARD";
			sSugCause = L_SCARD_E_UNKNOWNCARD_ErrorMessage;
		} else if (0==(0x80100010^nResult)) {
			sErrorName = "SCARD_E_NOT_READY";
			sSugCause = L_SCARD_E_NOTREADY_ErrorMessage;
		} else if (0==(0x80100013^nResult)) {
			sErrorName = "SCARD_F_COMM_ERROR";
			sSugCause = L_SCARD_F_COMMERROR_ErrorMessage;
		} else if (0==(0x8010001D^nResult)) {
			sErrorName = "SCARD_E_NO_SERVICE";
			sSugCause = L_SCARD_E_NOSERVICE_ErrorMessage;
		} else if (0==(0x8010001E^nResult)) {
			sErrorName = "SCARD_E_SERVICE_STOPPED";
			sSugCause = L_SCARD_E_SERVICESTOPPED_ErrorMessage;
		} else if (0==(0x8010002E^nResult)) {
			sErrorName = "SCARD_E_NO_READERS_AVAILABLE";
			sSugCause = L_SCARD_E_NOREADERSAVAILABLE_ErrorMessage;
		} else if (0==(0x8010002F^nResult)) {
			sErrorName = "SCARD_E_COMM_DATA_LOST";
			sSugCause = L_SCARD_E_COMMDATALOST_ErrorMessage;
		} else if (0==(0x80100030^nResult)) {
			sErrorName = "SCARD_E_NO_KEY_CONTAINER";
			sSugCause = L_SCARD_E_NOKEYCONTAINER_ErrorMessage;
		} else if (0==(0x80100067^nResult)) {
			sErrorName = "SCARD_W_UNPOWERED_CARD";
			sSugCause = L_SCARD_W_UNPOWEREDCARD_ErrorMessage;
		} else if (0==(0x80100069^nResult)) {
			sErrorName = "SCARD_W_REMOVED_CARD";
			sSugCause = L_SCARD_W_REMOVEDCARD_ErrorMessage;
		} else if (0==(0x8010006B^nResult)) {
			sErrorName = "SCARD_W_WRONG_CHV";
			sSugCause = L_SCARD_W_WRONGCHV_ErrorMessage;
		} else if (0==(0x8010006C^nResult)) {
			sErrorName = "SCARD_W_CHV_BLOCKED";
			sSugCause = L_SCARD_W_CHVBLOCKED_ErrorMessage;
		} else if (0==(0x8010006D^nResult)) {
			sErrorName = "SCARD_W_EOF";
			sSugCause = L_SCARD_W_EOF_ErrorMessage;
		} else if (0==(0x8010006E^nResult)) {
			sErrorName = "SCARD_W_CANCELLED_BY_USER";
			sSugCause = L_SCARD_W_CANCELLEDBYUSER_ErrorMessage;
		} else if (0==(0x8010006F^nResult)) {
			sErrorName = "SCARD_W_CARD_NOT_AUTHENTICATED";
			sSugCause = L_SCARD_W_CARDNOTAUTHENTICATED_ErrorMessage;
		} else if (0==(0x80090024^nResult)) {
			sErrorName = "NTE_TEMPORARY_PROFILE";
			sSugCause = L_SugCauseNoProfile_ErrorMessage;
		} else if (0==(0xFFFFFFFF^nResult)) {
			sErrorName=L_ErrNameNoFileName_ErrorMessage;
			sSugCause=L_SugCauseNoFileName_ErrorMessage;
		} else if (0==(0x800B010A^nResult)) {
			sErrorName = "CERT_E_CHAINING";
			sSugCause=L_SugCauseCAExSignerNotFound_ErrorMessage;
		} else if (0==(0x800B0109^nResult)) {
			sErrorName = "CERT_E_UNTRUSTEDROOT";
			sSugCause=L_SugCauseCAExNotTrusted_ErrorMessage;
		} else if (0==(0x8000FFFF^nResult)) {
			sErrorName="E_UNEXPECTED";
		} else if (0==(0x00000046^nResult)) {
			sErrorName=L_ErrNamePermissionDenied_ErrorMessage;
			if (document.UIForm.cbSaveRequest.checked) {
				sSugCause=L_SugCausePermissionToWrite_ErrorMessage;
			}
			else {
				sSugCause=L_SugCausePermissionToWrite_ErrorMessage;
			}
		}
		
		// modify the document text and appearance to show the error message
		spnErrorNum.innerText="0x"+toHex(nResult)+" - "+sErrorName;
	 	if (0 == sSugCauseIN) { 
			spnErrorMsg.innerText=sSugCause;
		} else { 
			spnErrorMsg.innerText=sSugCauseIN;
		}
		spnFixTxt.style.display='none';
		spnErrorTxt.style.display='';

		// back to the top so the messages show
		window.scrollTo(0,0);
	}


</Script> 
<Script Language="VBScript">
	' The current CA exchange certificate
	Public sCAExchangeCert
	sCAExchange=""
<%=sCAExchangeCert%>

	'-----------------------------------------------------------------
	' call XEnroll to create a request, since javascript has no error handling
	Function CreateRequest(dwCreateRequestFlag, sDistinguishedName, sCertUsage)
		On Error Resume Next
		XEnroll.ReuseHardwareKeyIfUnableToGenNew=False
		document.SubmittedData.CertRequest.value= _
			XEnroll.CreateRequest(dwCreateRequestFlag, sDistinguishedName, sCertUsage)
		CreateRequest=Err.Number
	End Function

	'-----------------------------------------------------------------
	' call XEnroll to create and save a request, since javascript has no error handling
	Function CreateAndSaveRequest(dwCreateRequestFlag, sDistinguishedName, sCertUsage, sSaveReqFile)
		On Error Resume Next
		XEnroll.ReuseHardwareKeyIfUnableToGenNew=False
		XEnroll.createFileRequest dwCreateRequestFlag, sDistinguishedName, sCertUsage, sSaveReqFile
		CreateAndSaveRequest=Err.Number
	End Function

	'----------------------------------------------------------------
	' handle a change in the current CSP, since javascript has no error handling
	Sub UpdateHashAlgList(nProvType)
		On Error Resume Next
		Dim nIndex, nAlgID, oElem, bList, lCSPType
		Const CALG_SSL3_SHAMD5=32776
		Const CALG_MAC=32773
		Const CALG_HMAC=32777
		Const CALG_MD5=32771

		'really strange, I can't use nProvType in following If compare
		'so I have to fetch from xenroll which is the same as nProvType:(
		lCSPType=XEnroll.ProviderType

		' clear the list
		While document.UIForm.lbHashAlgorithm.length>0
			document.UIForm.lbHashAlgorithm.options.remove(0)
		Wend

		' retrieve the list from XEnroll
		nIndex=0
		Do 
			' get the next AlgID
			nAlgID=XEnroll.EnumAlgs(nIndex, ALG_CLASS_HASH)
			If 0<>Err.Number Then 
				' no more algs
				Err.Clear
				Exit Do
			End If

			bList = True

			'GetAlgName is not cheap, try to reduce the call, check ID to filter out some unwanted hash

			'can't use the following hash
			If CALG_SSL3_SHAMD5=nAlgID Or CALG_MAC=nAlgID Or CALG_HMAC=nAlgID Then
				bList = False
			End If

			'DSS or DH won't work with MD5
			If CALG_MD5=nAlgID And PROV_DSS=lCSPType Or CALG_MD5=nAlgID And PROV_DSS_DH=lCSPType or CALG_MD5=nAlgID And PROV_DH_SCHANNEL=lCSPType Then
				bList = False
			End If

			If True=bList Then
				' get the corresponding name and create an option in the list box
				sName=XEnroll.GetAlgName(nAlgID)
				Set oElem=document.createElement("Option")
				oElem.text=sName
				oElem.value=nAlgID
				document.UIForm.lbHashAlgorithm.options.add(oElem)
			End If
			nIndex=nIndex+1

		Loop ' <- End alg enumeration loop

		' make sure the first one is selectd
		document.UIForm.lbHashAlgorithm.selectedIndex=0

	End Sub

	'----------------------------------------------------------------
	' call XEnroll to get the key length, since javascript has no error handling
	Function MyGetKeyLen(nSizeSpec, bExchange)
		On Error Resume Next
		Dim nKeySpec
		If True=bExchange Then
			nKeySpec=XEKL_KEYSPEC_KEYX
		Else
			nKeySpec=XEKL_KEYSPEC_SIG
		End If
		MyGetKeyLen=XEnroll.GetKeyLenEx(nSizeSpec, nKeySpec)
		If 0<>Err.Number Then
			If XEKL_KEYSIZE_MIN=nSizeSpec Then
				MyGetKeyLen=KEY_LEN_MIN_DEFAULT
			ElseIf XEKL_KEYSIZE_MAX=nSizeSpec Then
				MyGetKeyLen=KEY_LEN_MAX_DEFAULT
			ElseIf XEKL_KEYSIZE_DEFAULT=nSizeSpec Then
				MyGetKeyLen=KEY_LEN_MY_DEFAULT  'try 1024
			Else 'assume XEKL_KEYSIZE_INC=nSizeSpec
				MyGetKeyLen=KEY_LEN_INC_DEFAULT
			End If
		End If
		If XEKL_KEYSIZE_INC=nSizeSpec And 0=MyGetKeyLen Then
			MyGetKeyLen=KEY_LEN_INC_DEFAULT
		End If
	End Function

    '----------------------------------------------------
    ' set a certificate for key archive
	Function SetPrivateKeyArchiveCertificate()
		On Error Resume Next

        XEnroll.PrivateKeyArchiveCertificate=sCAExchange
	SetPrivateKeyArchiveCertificate = Err.Number
    End Function

    '----------------------------------------------------
    ' set request template extension
	Function vbAddCertTypeToRequestEx(lCTEVer, sCTEOID, lCTEMajor, bCTEfMinor, lCTEMinor)
		On Error Resume Next

		XEnroll.addCertTypeToRequestEx lCTEVer, sCTEOID, lCTEMajor, bCTEfMinor, lCTEMinor
		If 0 <> Err.Number Then
			'possible on downlevel not supporting v2 encoding, change to v1
			XEnroll.addCertTypeToRequestEx XECT_EXTENSION_V1, sCTEOID, lCTEMajor, bCTEfMinor, lCTEMinor
		End If
		vbAddCertTypeToRequestEx=Err.Number
	End Function

</Script> 

<%End If 'bFailed%>

</Body>
</HTML>
