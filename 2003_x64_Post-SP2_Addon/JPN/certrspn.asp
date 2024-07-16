<%@ CODEPAGE=65001 'UTF-8%>
<%' certrspn.asp - (CERT)srv web - (R)e(S)ult: (P)e(N)ding
  ' Copyright (C) Microsoft Corporation, 1998 - 1999 %>
<!-- #include FILE=certsbrt.inc -->
<!-- #include FILE=certdat.inc -->
<!-- #include FILE=certsrck.inc -->

<!-- Windows Security Update, KB2518295 has replaced some of the CA Web Enrollment ASP files -->
<!-- Please see http://www.support.microsoft.com/kb/2518295 for the back-up location of the previous ASP files -->

<%  ' came from certfnsh.asp

	Set ICertRequest=Session("ICertRequest")
	sMode=Request.Form("Mode")

	'Stop

	' If this is a new request, add it to the user's cookie
	If 0<>InStr(sMode,"newreq") Then
		AddRequest
	End If
%>
<HTML>
<Head>
	<Meta HTTP-Equiv="Content-Type" Content="text/html; charset=UTF-8">
	<Title>Microsoft Active Directory 証明書サービス</Title>
</Head>
<%If "IE"=sBrowser Then %>
<Body BgColor=#FFFFFF Link=#0000FF VLink=#0000FF ALink=#0000FF OnLoad="postLoad();"><Font ID=locPageFont Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
<%Else 'browsers other than IE may not be able to load postLoad script so we skip it %>
<Body BgColor=#FFFFFF Link=#0000FF VLink=#0000FF ALink=#0000FF"><Font ID=locPageFont Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial">
<%End If%>

<Table Border=0 CellSpacing=0 CellPadding=4 Width=100% BgColor=#008080>
<TR>
	<TD><Font Color=#FFFFFF><LocID ID=locMSCertSrv><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B><I>Microsoft</I></B> Active Directory 証明書サービス &nbsp;--&nbsp; <%=sServerDisplayName%> &nbsp;</Font></LocID></Font></TD>
	<TD ID=locHomeAlign Align=Right><A Href="/certsrv"><Font Color=#FFFFFF><LocID ID=locHomeLink><Font Face="MS UI Gothic, ＭＳ Ｐゴシック, Arial" Size=-1><B>ホーム</B></Font></LocID></Font></A></TD>
</TR>
</Table>

<P ID=locPageTitle> <B> 保留中の証明書 </B>
<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>

<%If 0<>InStr(sMode,"newreq") Then%>
<P ID=locInfoNewReq> 証明書の要求を受信しました。要求した証明書を管理者が発行するのを
お待ちください。 </P>

<P ID=locInfoReqID>
要求 ID は <%=nReqID%> です。
</P>

<%ElseIf "chkpnd"=sMode Then%>
<P ID=locInfoChkPnd> 証明書の要求はまだ保留になっています。要求した証明書を
管理者が発行するのをお待ちください。 </P>
<%End If%>

<P ID=locInstructions> 証明書を取得するには、1 日から 2 日後にこの Web サイトを再度参照してください。</P>
<P ID=locTimeoutWarning><Font Size=-1><B>注意:</B> 証明書を取得するには、<%=nPendingTimeoutDays%> 日以内に<B>この</B> Web ブラウザで
再度参照してください</Font></P>

<%If "chkpnd"=sMode Then%>
<Form Action="certrmpn.asp" Method=Post>
<Input Type=Hidden Name=Action Value="rmpn">
<Input Type=Hidden Name=ReqID Value="<%=Server.HTMLEncode(Request.Form("ReqID"))%>">
<P><Input ID=locBtnRemove Type=Submit Value="削除"> - 保留中の要求一覧からこの要求を削除します。
</Form>
<%End If%>


<!-- Green HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#008080><Img Src="certspc.gif" Alt="" Height=2 Width=1></TD></TR></Table>
<!-- White HR --><Table Border=0 CellSpacing=0 CellPadding=0 Width=100%><TR><TD BgColor=#FFFFFF><Img Src="certspc.gif" Alt="" Height=5 Width=1></TD></TR></Table>

</Font>
<!-- ############################################################ -->
<!-- End of standard text. Scripts follow  -->

<%If "IE"=sBrowser Then %>
<Script Language="VBScript">
	
	Sub postLoad
		On Error Resume Next

        <%If 0<>InStr(sMode,"newreq") Then%>
	
            '--------------------------------------------------------
            ' NOTE:
            '
            ' ENTERPRISE PENDING/AUTOENRL SUPPORT IS DISABLED BY DEFAULT.
            ' UN-COMMENT THE BELOW LINES TO RE-ENABLE
            '
            '--------------------------------------------------------
            
            <%If True=bLH Then%>
                
                '
                ' Dim n
		        ' Dim sServerConfig
		        ' Dim nReqId
		        ' Dim sCADNS
		        ' Dim sCAName
		        ' Dim sFriendlyName
                ' Dim objXEnroll
		        ' Dim sThumbPrint
                '
		        ' sThumbPrint="<%= escape(Request.Form("ThumbPrint")) %>"
		        ' If ""=sThumbPrint Then
		        '     Exit Sub
		        ' End If
		        ' Set objXEnroll = CreateObject("CEnroll.CEnroll.1")
		        '
		        '
		        ' sServerConfig="<%=sServerConfig%>"
		        ' nReqId = <%=nReqId%>
		        ' n = InStr(sServerConfig, "\")
		        ' sCADNS=Left(sServerConfig, n-1)
		        ' sCAName=Mid(sServerConfig, n+1)
		        ' sFriendlyName=""
		        'sFriendlyName="testName"
		        'Alert "requestId=" & nReqId & " CADNS=" & sCADNS & " CAName=" & sCAName & " ThumbPrint=" & sThumbPrint & " FriendlyName=" & sFriendlyName
		        ' objXEnroll.ThumbPrint=sThumbPrint
		        ' objXEnroll.setPendingRequestInfo nReqId, sCADNS, sCAName, sFriendlyName
                '
		        ' Set objXEnroll=Nothing
        		
		    <%End If%>
    	
        <%End If%>
	
	End Sub

</Script>
<%End If 'only support IE because other browsers may not load above script%>
</Body>
</HTML>
<%Session.Abandon()%>