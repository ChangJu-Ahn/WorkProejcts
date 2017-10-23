<%'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : (-)채권/출금등록 
'*  3. Program ID           : a4116ma1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : AP0041R
'*  7. Modified date(First) : 2001/02/10
'*  8. Modified date(Last)  : 2001/02/10
'*  9. Modifier (First)     : CHANG SUNG HEE
'* 10. Modifier (Last)      : CHANG SUNG HEE
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜: 

Dim pAR0139											'조회용 ComProxy Dll 사용 변수 

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount
Dim strItemSeq          

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strItemSeq = Request("txtItemSeq")

On Error Resume Next

Select Case strMode

Case CStr(UID_M0001)								'☜: 현재 조회/Prev/Next 요청을 받음 

	lgStrPrevKey = Request("lgStrPrevKey")
	
    Set pAR0139 = Server.CreateObject("Ar0139.ALookupRcptDcDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAR0139 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
'@ImportView
    pAR0139.ImportAAllcRcptAllcNo = Trim(Request("txtAllcNo"))
    pAR0139.ImportARcptDcSeq = strItemSeq
    pAR0139.CommandSent = "lookup"
    
    'Call SvrMsgBox("Condition ->" & Request("txtArNo") & " : " & Request("txtItemSeq") , vbInformation, I_MKSCRIPT)
    
    pAR0139.ServerLocation = ggServerIP

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
	pAR0139.ComCfg = gConnectionString
    pAR0139.Execute

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pAR0139 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAR0139.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAR0139.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAR0139.ExportErrEabSqlCodeSqlcode, _
						    pAR0139.ExportErrEabSqlCodeSeverity, _
						    pAR0139.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAR0139.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAR0139 = Nothing
		Response.End
	End If   
	
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
    GroupCount = pAR0139.ExportGroupCount

	' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
	' 문자/숫자 일 경우, 문맥에 맞게 처리함 
'	If pAR0139.ExportPIndReqIndReqmtNo(GroupCount) = pAR0139.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pAR0139.ExportNextPMPSRequirementIndReqmtNo
'    End If
%>

<Script Language=vbscript>
    Dim lngMaxRows       
    Dim strData
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		
		lngMaxRows = .frm1.vspdData3.MaxRows
		.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=GroupCount%>)		
<%      
	For LngRow = 1 To GroupCount
%>
	    strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportARcptDcDtlDtlSeq(LngRow))%>"        
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemCtrlCd(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemCtrlNm(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportARcptDcDtlCtrlVal(LngRow))%>"        
        strData = strData & Chr(11) & ""        
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportEabACtrlValRtnCtrlValC(LngRow))%>"  
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportARcptDcSeq(LngRow))%>"                    
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemTblId(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemDataColmId(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemDataColmNm(LngRow))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemColmDataType(LngRow))%>"
        strData = strData & Chr(11) & "<%=pAR0139.ExportItemACtrlItemDataLen(LngRow)%>"        
        strData = strData & Chr(11) & "<%=pAR0139.ExportItemAAssignAcctHqFg(LngRow)%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(pAR0139.ExportItemACtrlItemMajorCd(LngRow))%>"
        strData = strData & Chr(11) & "<%=LngRow%>"
        strData = strData & Chr(11) & Chr(12)        
        '
        .frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportARcptDcSeq(LngRow))%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportARcptDcDtlDtlSeq(LngRow))%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportARcptDcDtlCtrlVal(LngRow))%>"
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportEabACtrlValRtnCtrlValC(LngRow))%>"  
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportARcptDcSeq(LngRow))%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemTblId(LngRow))%>"
        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pAR0139.ExportItemACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=pAR0139.ExportItemAAssignAcctHqFg(LngRow)%>"
        .frm1.vspdData3.Col = 15
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAR0139.ExportItemACtrlItemMajorCd(LngRow))%>"

<%      
    Next
%>    
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData strData
		
'		.lgStrPrevKey = "<%=StrNextKey%>"

'		If .frm1.vspdData2.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> "" Then	<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
'			.DbQuery
'		Else
'			.frm1.hPlantCd.value = "<%=Request("txtPlantCd")%>"
'			.frm1.hReqStatus.value = "<%=Request("cboReqStatus")%>"			
'			.frm1.hFromReqrdDt.value = "<%=Request("txtFromReqrdDt")%>"
'			.frm1.hToReqrdDt.value = "<%=Request("txtToReqrdDt")%>"
'			.frm1.hItemCd.value = "<%=Request("txtItemCd")%>"
			
			.DbQueryOk2
'		End If
		
	End With
</Script>	
<% 
   
    Set pAR0139 = Nothing
End Select
%>
</Script>
