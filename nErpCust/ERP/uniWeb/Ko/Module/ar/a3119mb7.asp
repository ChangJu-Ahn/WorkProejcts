<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : f6102mb7
'*  4. Program Name         :
'*  5. Program Desc         : 계정별 관리항목 데이타 조회 
'*  6. Comproxy 리스트     : Ar0119
'*  6. Modified date(First) : 2000/10/7
'*  7. Modified date(Last)  : 송봉훈 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜: 


Dim pAr0119											'조회용 ComProxy Dll 사용 변수 

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          
Dim strItemSeq
Dim AcctNm

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 

On Error Resume Next

'Call SvrMsgBox("Condition ->" & Request("txtAdjustNo") & " : " & Request("txtSttlNo") , vbInformation, I_MKSCRIPT)

Select Case strMode

Case CStr(UID_M0001)								'☜: 현재 조회/Prev/Next 요청을 받음 

	lgStrPrevKey = Request("lgStrPrevKey")
	strItemSeq = Request("txtSttlNo")
	
    Set pAr0119 = Server.CreateObject("Ar0119.ALookupRcptAdjustDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------    
    If Err.Number <> 0 Then
		Set pAr0119 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
'@ImportView  
    
    pAr0119.ImportARcptAdjustAdjustNo = Trim(Request("txtAdjustNo"))
    pAr0119.ServerLocation = ggServerIP
	
    If Err.Number <> 0 Then
		Set pAr0119 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If
	
    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
    'pAr0119.ComCfg = "TCP Letitbe 2050"
    pAr0119.ComCfg = gConnectionString
    pAr0119.Execute
    
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------    
    If Err.Number <> 0 Then
		Set pAr0119 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAr0119.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAr0119.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAr0119.ExportErrEabSqlCodeSqlcode, _
						    pAr0119.ExportErrEabSqlCodeSeverity, _
						    pAr0119.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAr0119.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAr0119 = Nothing
		Response.End
	End If      
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
   	GroupCount = pAr0119.ExportGroupCount
'Call ServerMesgBox(pAr0119.ExportGroupCount, vbCritical, I_MKSCRIPT)
	' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
	' 문자/숫자 일 경우, 문맥에 맞게 처리함 
'	If pAr0119.ExportPIndReqIndReqmtNo(GroupCount) = pAr0119.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pAr0119.ExportNextPMPSRequirementIndReqmtNo
'    End If
%>

<Script Language=vbscript>
    Dim lngMaxRows       
    Dim strData
    Dim lRows
    Dim tmpDrCrFg	
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		
	lngMaxRows = .frm1.vspdData3.MaxRows
	.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=GroupCount%>)
<%      
	For LngRow = 1 To GroupCount
%>
<%'@ExportView - 고민중 %>

	    strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportARcptAdjustDtlDtlSeq(LngRow))%>"     '1   
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemCtrlCd(LngRow))%>"          '2 
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemCtrlNm(LngRow))%>"          '3
        If "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(pAr0119.ExportARcptAdjustDtlCtrlVal(LngRow))%>"    '4  
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportARcptAdjustDtlCtrlVal(LngRow))%>"    '4  
		ENd IF	
        strData = strData & Chr(11) & ""        												'5
        
		If "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
	        strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"  								'6
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportEabACtrlValRtnCtrlValC(LngRow))%>" '6
		End If          							'6
        strData = strData & Chr(11) & "<%=strItemSeq%>"											'7	
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemTblId(LngRow))%>" 			'8
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemDataColmId(LngRow))%>"		'9
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemDataColmNm(LngRow))%>"		'10
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>"    '11
        strData = strData & Chr(11) & "<%=pAr0119.ExportACtrlItemDataLen(LngRow)%>"        	'12
        strData = strData & Chr(11) & "<%=pAr0119.ExportAAssignAcctHqFg(LngRow)%>"		'13
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0119.ExportACtrlItemMajorCd(LngRow))%>"			'13
        strData = strData & Chr(11) & <%=LngRow%>												'14			
        strData = strData & Chr(11) & Chr(12)

		.frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportARcptAdjustDtlDtlSeq(LngRow))%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        If "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>" = "D" Then
			.frm1.vspdData3.Text = "<%=UNIDateClientFormat(pAr0119.ExportARcptAdjustDtlCtrlVal(LngRow))%>"   
        ELSE
			.frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportARcptAdjustDtlCtrlVal(LngRow))%>"   
        END IF
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
        if "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>" = "D" then		
        	.frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"  	
        ELSE	
			.frm1.vspdData3.Text =  "<%=ConvSPChars(pAr0119.ExportItemEabACtrlValRtnCtrlValC(LngRow))%>"
		end if
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemTblId(LngRow))%>"
        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pAr0119.ExportACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=pAr0119.ExportAAssignAcctHqFg(LngRow)%>"
		.frm1.vspdData3.Col = 15
        .frm1.vspdData3.Text = "<%=ConvSPChars(pAr0119.ExportACtrlItemMajorCd(LngRow))%>"
<%      
    Next
%>    
         
    .frm1.vspdData2.MaxRows = GroupCount
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowData strData
			
	.DbQueryOk2
		
	End With
</Script>	
<% 
   
    Set pAr0119 = Nothing
End Select
%>
</Script>
