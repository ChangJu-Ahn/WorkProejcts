<!-- #Include file="../inc/IncServer.asp" -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim LngRow
Dim GroupCount
Dim StrNextKey		' ���� �� 

On Error Resume Next

Call HideStatusWnd


If Request("txtMode") <> "" Then
	Dim objPopUp
	
	Set objPopUp = Server.CreateObject("B29016.B29016ListAcctDeptWithDate")
	
	objpopup.ImportBCalendarCalendarDt = UNIConvDate(Request("txtDate"))
    objPopUp.ImportBAcctDeptDeptCd = Request("txtCode")			'New query or Continuous query
    objPopUp.ImportBAcctDeptDeptNm = Request("txtName")		'set table name
    objPopup.ImportBAcctDeptInternalCd = Request("txtInternal")
    objPopUp.ServerLocation = ggServerIP
   
    objPopUp.ComCfg = gConnectionString
    objPopUp.Execute

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '��:
	   Set objPopUp = Nothing																	'��: ComProxy UnLoad
	   Response.End																				'��: Process End
	End If
    
    'Call SvrMsgBox(objPopUp.OperationStatusMessage, vbInformation, I_MKSCRIPT)
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If objPopUp.OperationStatusMessage <> "990000" Then
	   Call DisplayMsgBox(objPopUp.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	   Set objPopUp = Nothing																	'��: ComProxy UnLoad
	   Response.End																				'��: Process End
	End If

	GroupCount = objPopUp.ExportGroupCount
	
	If objPopUp.ExportItemBAcctDeptDeptCd(GroupCount) = objPopUp.ExportNextBAcctDeptDeptCd Then
		StrNextKey = ""
	Else
		StrNextKey = objPopUp.ExportNextBAcctDeptDeptCd
    End If
%>		    
<Script Language="vbscript">   
	Dim StrData
	
	With parent
<%
	For LngRow = 1 To GroupCount		
%>
		strData = strData & Chr(11) & "<%=ConvSPChars(objPopUp.ExportItemBAcctDeptDeptCd(LngRow))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(objPopUp.ExportItemBAcctDeptDeptNm(LngRow))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(objPopUp.ExportItemBAcctDeptInternalCd(LngRow))%>"
		strData = strData & Chr(11) & Chr(12)		
<%
    Next
%>
		'Call SvrMsgBox(strData, vbInformation, I_INSCRIPT)
		.lgCode = "<%=ConvSPChars(objPopUp.ExportNextBAcctDeptDeptCd)%>"
		.lgName = "<%=ConvSPChars(objPopUp.ExportNextBAcctDeptDeptNm)%>"
		.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
		.lgIntFlgMode = OPMD_UMODE
	
	    .ggoSpread.Source = parent.vspdData
		.ggoSpread.SSShowData strData
		.vspdData.focus

		If .vspdData.MaxRows = 0 Then
			parent.UNIMsgBox "�˻��� Data�� �����ϴ�", 48, parent.top.document.title
		End If

	End With

</Script>
<%
    Set objPopup = nothing
End If
%>