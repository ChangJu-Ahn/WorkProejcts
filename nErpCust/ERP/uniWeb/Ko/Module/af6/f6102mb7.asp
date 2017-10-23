<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : f6102mb7
'*  4. Program Name         :
'*  5. Program Desc         : 계정별 관리항목 데이타 조회 
'*  6. Comproxy 리스트     : fp0038
'*  6. Modified date(First) : 2000/10/7
'*  7. Modified date(Last)  : 송봉훈 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜: 

Dim pCOM											'조회용 ComProxy Dll 사용 변수 

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

Select Case strMode

Case CStr(UID_M0001)								'☜: 현재 조회/Prev/Next 요청을 받음 

	lgStrPrevKey = Request("lgStrPrevKey")
	strItemSeq = Request("txtSttlNo")
	
    Set pCOM = Server.CreateObject("FP0038.Fp0038ListPpSttlDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pCOM = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
'@ImportView

    pCOM.ImportFPrpaymSttlSttlNo = Trim(Request("txtSttlNo"))
    pCOM.ImportFPrpaymPrpaymNo = Trim(Request("txtPrpaymNo"))

    pCOM.ServerLocation = ggServerIP

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
    pCOM.ComCfg = gConnectionString
    pCOM.Execute
    
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pCOM = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

	'-----------------------------------------
	'Com action result check area(DB,internal)
	'-----------------------------------------
	If Not (pCOM.OperationStatusMessage = MSG_OK_STR) Then
		'Call DisplayMsgBox(pCOM.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		Set pCOM = Nothing												'☜: ComProxy Unload
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If    
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
   	GroupCount = pCOM.ExportGroupCount

	' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
	' 문자/숫자 일 경우, 문맥에 맞게 처리함 
'	If pCOM.ExportPIndReqIndReqmtNo(GroupCount) = pCOM.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pCOM.ExportNextPMPSRequirementIndReqmtNo
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

		strData = strData & Chr(11) & "<%=strItemSeq%>"											'1
	    strData = strData & Chr(11) & "<%=pCOM.ExportItemFPrpaymSttlDtlDtlSeq(LngRow)%>"		'2
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemACtrlItemCtrlCd(LngRow))%>"			'3
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemACtrlItemCtrlNm(LngRow))%>"			'4
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemFPrpaymSttlDtlCtrlVal(LngRow))%>"	'5
        strData = strData & Chr(11) & ""														'6
		If "<%=ConvSPChars(pCOM.ExportItemACtrlItemColmDataType(LngRow))%>" = "D" Then					'7
	        strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemEabACtrlValRtnCtrlValC(LngRow))%>"
		End If
        strData = strData & Chr(11) & "<%=strItemSeq%>"											'8
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemACtrlItemTblId(LngRow))%>"			'9
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemACtrlItemDataColmId(LngRow))%>"		'10
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemACtrlItemDataColmNm(LngRow))%>"		'11
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemACtrlItemColmDataType(LngRow))%>"	'12
        strData = strData & Chr(11) & "<%=pCOM.ExportItemACtrlItemDataLen(LngRow)%>"			'13
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportItemAAssignAcctInputFg(LngRow))%>"		'14
        strData = strData & Chr(11) & <%=LngRow%>
        strData = strData & Chr(11) & Chr(12)

		.frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=pCOM.ExportItemFPrpaymSttlDtlDtlSeq(LngRow)%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemFPrpaymSttlDtlCtrlVal(LngRow))%>"
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
		If "<%=ConvSPChars(pCOM.ExportItemACtrlItemColmDataType(LngRow))%>" = "D" Then
        	.frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"
        Else
			.frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemEabACtrlValRtnCtrlValC(LngRow))%>"
		End If
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemACtrlItemTblId(LngRow))%>"
        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pCOM.ExportItemACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportItemAAssignAcctInputFg(LngRow))%>"

<%      
    Next
%>    
         
    .frm1.vspdData2.MaxRows = 0
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowData strData
			
	Call .DbQueryOk2
		
	End With
</Script>	
<% 
   
    Set pCOM = Nothing
End Select
%>
</Script>
