<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : f7102mb7
'*  4. Program Name         :
'*  5. Program Desc         : F_PrRcpt_Sttl_Dtl Query
'*  6. Comproxy 리스트     : fr0028
'*  6. Modified date(First) : 2000/10/7
'*  7. Modified date(Last)  : 2001/01/09
'*  8. Modifier (First)     : 송봉훈 
'*  9. Modifier (Last)      : 김희정 
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

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 

'Call SvrMsgBox("Condition ->" & Request("txtAcctCd") & " : " & Request("txtItemSeq") , vbInformation, I_MKSCRIPT)

Select Case strMode

Case CStr(UID_M0001)								'☜: 현재 조회/Prev/Next 요청을 받음 

	lgStrPrevKey = Request("lgStrPrevKey")
	strItemSeq   = Request("txtSttlNo")
	
    Set pCOM  = Server.CreateObject("FR0038.Fr0038ListPrSttlDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.number  <> 0 Then
		Set pCOM = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.Description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
	'@ImportView
    pCOM.ImportFPrrcptPrrcptNo			= UCase(Trim(Request("txtPrrcptNo")))
    pCOM.ImportFPrrcptSttlSttlNo		= Trim(Request("txtSttlNo"))
    pCOM.ImportNextFPrrcptSttlDtlDtlSeq = UNIConvNum(Request("lgStrPrevKey"),0)
    
    pCOM.ServerLocation = ggServerIP

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
    pCOM.ComCfg = gConnectionString 
    'pCOM.ComCfg = "TCP letitbe 2056"    
    pCOM.Execute
    
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.number  <> 0 Then
		Call ServerMesgBox(Err.number & Err.Description , vbCritical, I_MKSCRIPT)
		Set pCOM = Nothing												'☜: ComProxy Unload
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

	'-----------------------------------------
	'Com action result check area(DB,internal)
	'-----------------------------------------
	If Not (pCOM.OperationStatusMessage = MSG_OK_STR) Then
		'Call DisplayMsgBox(pCOM.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		Set pCOM = Nothing												'☜: ComProxy Unload
		Response.End 											'☜: 비지니스 로직 처리를 종료함 
	End If    
    
	LngMaxRow  = Request("txtMaxRows")										'Save previous Maxrow                                                
    GroupCount = pCOM.ExportDtlGrpCount

%>

<Script Language=vbscript>
    Dim lngMaxRows       
    Dim strData
    Dim lRows
    Dim tmpDrCrFg	
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		
	lngMaxRows	= .frm1.vspdData3.MaxRows
	.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=GroupCount%>)
<%      
	For LngRow  = 1 To GroupCount
%>
<%'@ExportView - 고민중 %>	    

		strData = strData & Chr(11) & "<%=strItemSeq%>"											'1
        strData = strData & Chr(11) & "<%=pCOM.ExportDtlItemFPrrcptSttlDtlDtlSeq(LngRow)%>"     '2
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemCtrlCd(LngRow))%>"			'3
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemCtrlNm(LngRow))%>"			'4
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemFPrrcptSttlDtlCtrlVal(LngRow))%>"	'5
        strData = strData & Chr(11) & ""														'6
		If "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemColmDataType(LngRow))%>" = "D" Then
	        strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"								'7
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemEabACtrlValRtnCtrlValC(LngRow))%>"
		End If
        strData = strData & Chr(11) & "<%=strItemSeq%>"											'8
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemTblId(LngRow))%>"			'9
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemDataColmId(LngRow))%>"		'10
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemDataColmNm(LngRow))%>"		'11
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemColmDataType(LngRow))%>"	'12
        strData = strData & Chr(11) & "<%=pCOM.ExportDtlItemACtrlItemDataLen(LngRow)%>"			'13
        strData = strData & Chr(11) & "<%=ConvSPChars(pCOM.ExportDtlItemAAssignAcctHqFg(LngRow))%>"		'14
        strData = strData & Chr(11) & <%=LngRow%>
        strData = strData & Chr(11) & Chr(12)

		.frm1.vspdData3.Row	 = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col  = 1
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col  = 2
        .frm1.vspdData3.Text = "<%=pCOM.ExportDtlItemFPrrcptSttlDtlDtlSeq(LngRow)%>"
        .frm1.vspdData3.Col  = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col  = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col  = 5
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemFPrrcptSttlDtlCtrlVal(LngRow))%>"
        .frm1.vspdData3.Col  = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col  = 7
		If "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemColmDataType(LngRow))%>" = "D" Then
        	.frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"
        Else
			.frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemEabACtrlValRtnCtrlValC(LngRow))%>"
		End If
        .frm1.vspdData3.Col  = 8
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col  = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemTblId(LngRow))%>"	
        .frm1.vspdData3.Col  = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col  = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col  = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col  = 13
        .frm1.vspdData3.Text = "<%=pCOM.ExportDtlItemACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col  = 14
        .frm1.vspdData3.Text = "<%=ConvSPChars(pCOM.ExportDtlItemAAssignAcctHqFg(LngRow))%>"       
        
<%      
    Next
%>             
    .frm1.vspdData2.MaxRows = 0
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowData strData
			
	.DbQueryOk2
		
	End With
</Script>	

<%  
    Set pCOM = Nothing
    
End Select
%>

</Script>
