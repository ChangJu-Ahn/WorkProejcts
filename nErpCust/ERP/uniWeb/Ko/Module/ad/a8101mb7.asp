<%
'=======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : a53017
'*  4. Program Name         :
'*  5. Program Desc         : 계정별 관리항목 데이타 조회 
'*  6. Modified date(First) : 2000/10/9
'*  7. Modified date(Last)  : 
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

Dim pA53017											'조회용 ComProxy Dll 사용 변수 

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

'Call SvrMsgBox("Condition ->" & Request("txtAcctCd") & " : " & Request("txtItemSeq") , vbInformation, I_MKSCRIPT)

Select Case strMode

	Case CStr(UID_M0001)								'☜: 현재 조회/Prev/Next 요청을 받음 

	lgStrPrevKey = Request("lgStrPrevKey")
	strItemSeq   = Request("txtItemSeq")
	
    Set pA53017 = Server.CreateObject("A53017.ALookupTempGlDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pA53017 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
    pA53017.InTempGlATempGlTempGlNo    = Request("txtTempGlNo")
    pA53017.InTemSeqATempGlItemItemSeq = Request("txtItemSeq")
    pA53017.ServerLocation             = ggServerIP

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
	pA53017.ComCfg = gConnectionString
    pA53017.Execute

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pA53017 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

	'-----------------------------------------
	'Com action result check area(DB,internal)
	'-----------------------------------------
	If Not (pA53017.OperationStatusMessage = MSG_OK_STR) Then
		Call DisplayMsgBox(pA53017.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		Set pA53017 = Nothing												'☜: ComProxy Unload
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If    
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
   	GroupCount = pA53017.OutGrpTempGlDtlCount

	' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
	' 문자/숫자 일 경우, 문맥에 맞게 처리함 
'	If pA53017.ExportPIndReqIndReqmtNo(GroupCount) = pA53017.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pA53017.ExportNextPMPSRequirementIndReqmtNo
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
        strData = strData & Chr(11) & "<%=pA53017.OutGrpATempGlDtlDtlSeq(LngRow)%>"				'1   
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemCtrlCd(LngRow))%>"              '2 
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemCtrlNm(LngRow))%>"              '3
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpATempGlDtlCtrlVal(LngRow))%>"                                                        '4  
        strData = strData & Chr(11) & ""        						'5
        
		If "<%=ConvSPChars(pA53017.OutGrpACtrlItemTblId(LngRow))%>" = "" And "<%=ConvSPChars(pA53017.OutGrpACtrlItemColmDataType(LngRow))%>" = "D" then
	        strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"                               '6
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpEabACtrlValRtnCtrlValC(LngRow))%>" '6
		End If

        strData = strData & Chr(11) & "<%=strItemSeq%>"											'7	
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemTblId(LngRow))%>" 				'8
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataColmId(LngRow))%>"			'9
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataColmNm(LngRow))%>"			'10
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemColmDataType(LngRow))%>"        '11
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataLen(LngRow))%>"        		'12
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpAAssignAcctInputFg(LngRow))%>"			'13
        strData = strData & Chr(11) & <%=LngRow%>												'14
        strData = strData & Chr(11) & Chr(12)

		.frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=pA53017.OutGrpATempGlDtlDtlSeq(LngRow)%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpATempGlDtlCtrlVal(LngRow))%>"
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpEabACtrlValRtnCtrlValC(LngRow))%>"
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemTblId(LngRow))%>"
	
        If "<%=ConvSPChars(pA53017.OutGrpACtrlItemTblId(LngRow))%>" = "" And "<%=ConvSPChars(pA53017.OutGrpACtrlItemColmDataType(LngRow))%>" = "D" then
			.frm1.vspdData3.Col = 7
        	.frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"
		End If

        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pA53017.OutGrpACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpAAssignAcctInputFg(LngRow))%>"
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
    Set pA53017 = Nothing

End Select
%>
</Script>
