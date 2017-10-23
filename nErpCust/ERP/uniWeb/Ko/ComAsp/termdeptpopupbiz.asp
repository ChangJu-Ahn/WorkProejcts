<%
'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Comon Popup Business Part													*
'*  3. Program ID           : TermDeptBiz.asp															*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 기간별 부서팝업															*
'*  7. Modified date(First) : 2000/08/30																*
'*  8. Modified date(Last)  : 2000/08/30																*
'*  9. Modifier (First)     : Hwang Jeong Won															*
'* 10. Modifier (Last)      : Hwang Jeong Won															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :																			*
'*                            2000/08/30 : Coding Start													*
'********************************************************************************************************

%>
<!-- #Include file="../inc/IncServer.asp" -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngRow
Dim GroupCount
Dim StrNextKey		' 다음 값 

On Error Resume Next

Call HideStatusWnd


If Request("txtMode") <> "" Then
	Dim objPopUp
	
	Set objPopUp = Server.CreateObject("B29017.B29017ListAcctDeptWithTerm")
	
	objpopup.ImportFromBCalendarCalendarDt = UNIConvDate(Request("txtFromDate"))
	objpopup.ImportToBCalendarCalendarDt = UNIConvDate(Request("txtToDate"))
	objPopUp.ImportBAcctDeptOrgChangeId = Request("txtOrgId")
    objPopUp.ImportBAcctDeptDeptCd = Request("txtCode")			'New query or Continuous query
    objPopUp.ImportBAcctDeptOrgChangeDt = UNIConvDate(Request("txtChangeDt"))
    objPopUp.ImportBAcctDeptDeptFullNm = Request("txtName")		'set table name
    objPopUp.ImportZUsrMastRecUsrId = Request("txtUser")
    objPopup.ImportBAcctDeptInternalCd = Request("txtInternal")    
    objPopUp.ServerLocation = ggServerIP
    
    objPopUp.ComCfg = gConnectionString
    objPopUp.Execute

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------	
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '⊙:
	   Set objPopUp = Nothing																	'☜: ComProxy UnLoad
	'   Response.End																				'☜: Process End
	End If
    
    'Call SvrMsgBox(objPopUp.OperationStatusMessage, vbInformation, I_MKSCRIPT)
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If objPopUp.OperationStatusMessage <> "990000" Then
	   Call DisplayMsgBox(objPopUp.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	   Set objPopUp = Nothing																	'☜: ComProxy UnLoad
	   Response.End																				'☜: Process End
	End If

	GroupCount = objPopUp.ExportGroupCount
	
	If (objPopUp.ExportItemBAcctDeptOrgChangeId(GroupCount) = objPopUp.ExportNextBAcctDeptOrgChangeId) And _
	   (objPopUp.ExportItemBAcctDeptDeptCd(GroupCount) = objPopUp.ExportNextBAcctDeptDeptCd) Then
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
		strData = strData & Chr(11) & "<%=ConvSPChars(objPopUp.ExportItemBAcctDeptOrgChangeId(LngRow))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(objPopUp.ExportItemBAcctDeptOrgChangeDt(LngRow))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(objPopUp.ExportItemBAcctDeptDeptCd(LngRow))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(objPopUp.ExportItemBAcctDeptDeptFullNm(LngRow))%>"		
		strData = strData & Chr(11) & "<%=ConvSPChars(objPopUp.ExportItemBAcctDeptInternalCd(LngRow))%>"
		strData = strData & Chr(11) & Chr(12)		
<%
    Next
    
%>		
		.lgOrgId = "<%=ConvSPChars(objPopUp.ExportNextBAcctDeptOrgChangeId)%>"
		.lgCode = "<%=ConvSPChars(objPopUp.ExportNextBAcctDeptDeptCd)%>"
		.lgChangeDt = "<%=ConvSPChars(objPopUp.ExportNextBAcctDeptOrgChangeDt)%>"
		.lgName = "<%=ConvSPChars(objPopUp.ExportNextBAcctDeptDeptFullNm)%>"
		
		.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
		.lgIntFlgMode = OPMD_UMODE
				
	    .ggoSpread.Source = parent.vspdData
		.ggoSpread.SSShowData strData
		.vspdData.focus
							
		If .vspdData.MaxRows = 0 Then
			parent.UNIMsgBox "검색된 Data가 없습니다", 48, parent.top.document.title
		End If

	End With

</Script>
<%
    Set objPopup = nothing
End If
%>