<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2214QA2
'*  4. Program Name         : 판매계획대실적조회(품목그룹)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/16
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Yong Sik
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	
Const BIZ_PGM_ID 		= "S2214QB2.asp"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 85					                          '☆: SpreadSheet의 키의 갯수 

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                          
Dim IsOpenPop  

Dim lgCookValue 

Dim lgSaveRow 

Dim lgStrColorFlag

Dim lgIsInitSpreadSheetBeforeFncQuery	' FncQuery 실행전 InitSpreadSheet 실행여부 
Dim lgArrColHidden()

<% 
   BaseDate     = GetSvrDate                                                                  'Get DB Server Date
%>   
Dim FromDateOfDB
Dim ToDateOfDB
FromDateOfDB = UNIConvDateAToB(UniDateAdd("m", 0, "<%=BaseDate%>",parent.gServerDateFormat),parent.gServerDateFormat,parent.gDateFormat)
ToDateOfDB   = UNIConvDateAToB(UniDateAdd("m", 0, "<%=BaseDate%>",parent.gServerDateFormat),parent.gServerDateFormat,parent.gDateFormat)

'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

End Sub

'========================================================================================================
Sub SetDefaultVal()
	Frm1.txtConFromDt.Text	= cstr(FromDateOfDB)
	Frm1.txtConToDt.Text	    = cstr(ToDateOfDB)
	frm1.cboDisplayType.value = "H"
	frm1.cboConBaseCur.value = "L"
	Call cboConBaseCur_onChange
	frm1.cboConSpType.focus
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet(ByVal pvPsdNo)
	If pvPsdNo = "A" then
	    Call SetZAdoSpreadSheet("S2214QA201","S","A", "V20030326", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
	    Call SetSpreadLock("A")
	Else
		Call SetZAdoSpreadSheet("S2214QA202","S","B", "V20030326", parent.C_SORT_DBAGENT,frm1.vspdData2,C_MaxKey, "X", "X")
        Call SetSpreadLock("B")
	End If
	
	Redim lgArrColHidden( frm1.vspdData.MaxCols )
	lgIsInitSpreadSheetBeforeFncQuery = true
End Sub

'========================================================================================================
Sub SetSpreadLock(Byval iOpt)
    If iOpt = "A" Then
       With frm1
          .vspdData.ReDraw = False
          ggoSpread.Source = .vspdData 
            ggoSpread.SpreadLock 1, -1
          .vspdData.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2
            ggoSpread.SpreadLock 1, -1
            .vspdData2.ReDraw = True
       End With
    End If   
End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables
	Call InitComboBox	
	Call SetDefaultVal
	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")
    Call SetToolBar("1100000000011111")
    Call ggoOper.FormatDate(frm1.txtConFromDt, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtConToDt, Parent.gDateFormat, 2)

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
Function FncQuery() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False                                                              '⊙: Processing is NG
    
    Call ggoOper.ClearField(Document, "2")									      '⊙: Clear Contents  Field
    
    if frm1.cboDisplayType.value = "H" then
		ggoSpread.Source = frm1.vspdData
	else
		ggoSpread.Source = frm1.vspdData2
	end if
    Call ggoSpread.ClearSpreadData()
    
    Call InitVariables 														      '⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then								              '⊙: This function check indispensable field
       Exit Function
    End If
    
    If Trim(frm1.cboDisplayType.value) = "H" and frm1.txtConFromDt.year <> frm1.txtConToDt.year then
		'결과출력형식이 가로이면 계획년월의 시작년과 종료년이 같아야 함 
		Call DisplayMsgBox("202416","X","X","X")
		Exit Function
	End If
	
	'계획년월의 종료월은 시작월보다 크거나 같아야 함 
	If ValidDateCheck(frm1.txtConFromDt, frm1.txtConToDt) = False Then Exit Function
    
    If DbQuery = False Then 
       Exit Function
    End If   

    If Err.number = 0 Then
       FncQuery = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncExport(parent.C_MULTI)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncExcel = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncFind(parent.C_MULTI, True)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncFind = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================================
Function FncExit()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncExit = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
Function DbQuery() 
	Dim strVal, iSheetNo

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
	Call LayerShowHide(1)
	
	Call ChangeDisplayType
        
	if frm1.cboDisplayType.value = "H" then
		iSheetNo = "A"
	else
		iSheetNo = "B"
	end if
	        
	frm1.txtSelectList.value = EnCoding(GetSQLSelectList(iSheetNo))
	frm1.txtSelectListDT.value = GetSQLSelectListDataType(iSheetNo)
	frm1.txtTailList.value = MakeSQLGroupOrderByList(iSheetNo)
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then
       DbQuery = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
Function DbQueryOk()												

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												 '⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1

    Set gActiveElement = document.ActiveElement 
    Call SetQuerySpreadColor
    
    If frm1.cboDisplayType.value = "H" then
		If frm1.vspdData.MaxRows > 0 Then
	    	frm1.vspdData.Focus
		Else
			frm1.txtConFromDt.focus
		End If
	Else
		If frm1.vspdData2.MaxRows > 0 Then
	    	frm1.vspdData2.Focus
		Else
			frm1.txtConFromDt.focus
		End If
	End If

	Call FormatSpreadCellByCurrency	
	
    lgIsInitSpreadSheetBeforeFncQuery = false
    
End Function


'========================================================================================================
Sub InitComboBox()

	' 판매계획유형 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0023", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConSpType,lgF0,lgF1,parent.gColSep)

	'거래구분 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S4225", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConLocExpFlag,lgF0,lgF1,Chr(11))
	
	'품목그룹레벨 
	Call CommonQueryRs(" DISTINCT ITEM_GROUP_LEVEL, CAST(ITEM_GROUP_LEVEL AS VARCHAR(2)) + ' LEVEL' ", " B_ITEM_GROUP ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	'Call CommonQueryRs(" ITEM_GROUP_CD, ITEM_GROUP_NM  ", " B_ITEM_GROUP ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConGrpLvl,lgF0,lgF1,Chr(11))
	
	'결과출력형식 
	lgF0 = "H" & Chr(11) & "V" & Chr(11)
	lgF1 = "가로" & Chr(11) & "세로" & Chr(11)
	Call SetCombo2(frm1.cboDisplayType,lgF0,lgF1,Chr(11))
	
	'화폐기준 
	lgF0 = "D" & Chr(11) & "L" & Chr(11)
	lgF1 = "외화금액" & Chr(11) & "자국금액" & Chr(11)
	Call SetCombo2(frm1.cboConBaseCur,lgF0,lgF1,Chr(11))

End Sub

'========================================================================================================
'	Description : 품목그룹레벨 onChange 이벤트 처리 
'========================================================================================================
Sub cboConGrpLvl_onChange()
	lgBlnFlgChgValue = true	
	frm1.txtConItemGroupCd.value = ""
	frm1.txtConItemGroupNm.value = ""
End Sub

'========================================================================================================
'	Description : 화폐 onChange 이벤트 처리 
'========================================================================================================
Sub cboConBaseCur_onChange()
	lgBlnFlgChgValue = true	
	if frm1.cboConBaseCur.value = "L" then
		frm1.txtConCur.value = Parent.gCurrency
		Call ggoOper.SetReqAttr(frm1.txtConCur,"Q")
		frm1.btnCur.Disabled = true
	else
		frm1.txtConCur.value = ""
		Call ggoOper.SetReqAttr(frm1.txtConCur,"D")
		frm1.btnCur.Disabled = false
	end if
End Sub

'========================================================================================================
'	Description : 결과출력형식에 따라 spread 변경 
'========================================================================================================
Sub ChangeDisplayType()
	lgBlnFlgChgValue = true	
	if frm1.cboDisplayType.value = "H" then
		frm1.vspdData.style.display = "inline"
		frm1.vspdData2.style.display = "none"
	    Call SetColHiddenByMonth
		ggoSpread.Source = frm1.vspdData
	else
		frm1.vspdData.style.display = "none"
		frm1.vspdData2.style.display = "inline"
		ggoSpread.Source = frm1.vspdData2
	end if
End Sub

'========================================================================================================
'	Description : PopUp
'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(6), iArrField(6), iArrHeader(6)
	
	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case pvIntWhere
		Case 1 '품목그룹 
			iArrParam(0) = "품목그룹"
			iArrParam(1) = "B_ITEM_GROUP"						'TABLE 명 
			iArrParam(2) = Trim(frm1.txtConItemGroupCd.value)	' Code Condition
			iArrParam(3) = ""									' Name Cindition
			iArrParam(4) = "DEL_FLG=" & FilterVar("N", "''", "S") & "  AND ITEM_GROUP_LEVEL= " & FilterVar(frm1.cboConGrpLvl.value, "''", "S") & ""	' Where Condition
			iArrParam(5) = "품목그룹"						' TextBox 명칭 
					
			iArrField(0) = "ED15" & Parent.gColSep & "ITEM_GROUP_CD"
			iArrField(1) = "ED30" & Parent.gColSep & "ITEM_GROUP_NM"
				    
			iArrHeader(0) = "품목그룹"
			iArrHeader(1) = "품목그룹명"
			frm1.txtConItemGroupCd.focus
			
		Case 2 '화폐 
			iArrParam(0) = "화폐"
			iArrParam(1) = "B_CURRENCY"

			iArrParam(2) = Trim(frm1.txtConCur.value)
			 
			iArrParam(3) = ""          ' Name Cindition
			iArrParam(4) = ""          ' Where Condition
			iArrParam(5) = "화폐"  'TextBox 명칭 
			 
			iArrField(0) = "ED15" & Parent.gColSep & "CURRENCY"
			iArrField(1) = "ED30" & Parent.gColSep & "CURRENCY_DESC"
			    
			iArrHeader(0) = "화폐"		' Header명(0)
			iArrHeader(1) = "화폐명"	' Header명(1)
			frm1.txtConCur.focus
			
	End Select

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(iArrRet,pvIntWhere)
		OpenConPopup = True	
	End If	

End Function

'========================================================================================================
'	Description : OpenConPopup에서 Return되는 값 setting
'========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case 1 '품목그룹 
			.txtConItemGroupCd.Value = pvArrRet(0)
			.txtConItemGroupNm.Value = pvArrRet(1)	
		Case 2 '화폐 
			.txtConCur.Value = pvArrRet(0)
		End Select
	End With

	SetConPopup = True

End Function

'========================================================================================================
'	Description : 스프레트시트의 특정 컬럼의 배경색상을 변경 
'========================================================================================================
Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	Dim Spread
	
	if frm1.cboDisplayType.value = "H" then
		Set Spread = frm1.vspdData
	else
		Set Spread = frm1.vspdData2
	end if
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)
	
	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)

		Spread.Col = -1
		Spread.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				Spread.BackColor = RGB(204,255,153) '연두 
			Case "2"
				Spread.BackColor = RGB(176,234,244) '하늘색 
			Case "3"
				Spread.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				Spread.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				Spread.BackColor = RGB(255,255,153) '연노랑 
		End Select
	Next

End Sub

'========================================================================================================
'	Description : 스프레트시트의 특정 컬럼을 Hidden처리 
'========================================================================================================
Sub SetColHiddenByMonth()
	Dim iIntStartKey
	Dim iIntStartMonth, iIntEndMonth
	Dim iCnt
	
	iIntStartKey = 12 'Z_ADO_FIELD_INF 테이블에서 1월계획량에 해당하는 SEQ_NO
	iIntStartMonth = Cint(frm1.txtConFromDt.month)
	iIntEndMonth = Cint(frm1.txtConToDt.month)
	
	frm1.vspdData.ReDraw = false
	' ------------------------------------------------------------------------------------------------------------------
	' lgIsInitSpreadSheetBeforeFncQuery 전역변수를 사용하여 처리한 이유 :
	' InitSpreadSheet 함수를 실행한 직후 다시말해 그리드환경설정 정보를 배열에 담아주기 위해 
	' reload시에는 xml 폴더아래 파일로 저장된 그리드환경정보가 InitSpreadSheet 함수가 실행되면서 그리드에 적용되고 
	' 그 때의 그리드환경정보가 배열에 담겨진다.
	' 그리드환경설정 수정시에도 InitSpreadSheet 함수가 실행되고 수정된 환경이 그리드에 적용됨으로 
	' 수정된 그리드환경정보가 배열에 담겨진다.
	
	If lgIsInitSpreadSheetBeforeFncQuery then
		For iCnt = iIntStartKey to iIntStartKey+(6*12)-1
			frm1.vspdData.Col = iCnt
			If frm1.vspdData.ColHidden then
				lgArrColHidden(iCnt-iIntStartKey) = true
			Else
				lgArrColHidden(iCnt-iIntStartKey) = false
			End if
		Next
	End if
	' ------------------------------------------------------------------------------------------------------------------
		
	Call ggoSpread.SSSetColHidden(iIntStartKey, iIntStartKey+(12*6)-1, false)
	
	if iIntStartMonth = 1 then
		Call ggoSpread.SSSetColHidden(iIntStartKey+iIntEndMonth*6, iIntStartKey+12*6-1, True)
	elseif iIntEndMonth = 12 then
		Call ggoSpread.SSSetColHidden(iIntStartKey, iIntStartKey+(iIntStartMonth-1)*6-1, True)
	else
		Call ggoSpread.SSSetColHidden(iIntStartKey, iIntStartKey+(iIntStartMonth-1)*6-1, True)
		Call ggoSpread.SSSetColHidden(iIntStartKey+iIntEndMonth*6, iIntStartKey+12*6-1, True)
	end if
	
	For iCnt = iIntStartKey+((iIntStartMonth-1)*6) to iIntStartKey+(iIntEndMonth*6)-1
		Call ggoSpread.SSSetColHidden(iCnt, iCnt, lgArrColHidden(iCnt-iIntStartKey))
	Next

	frm1.vspdData.ReDraw = true

End Sub

'==================================================================================
Sub PopZAdoConfigGrid()
	
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			OpenOrderBy("A")
		Case "VSPDDATA2"			
			OpenOrderBy("B")
	End Select	
 
End Sub

'========================================================================================================
Sub OpenOrderBy(ByVal pvPsdNo)
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then         ' Means that nothing is happened!!!
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet(pvPsdNo)       
   End If
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData2
    
    If Frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call SetSpreadColumnValue("B",frm1.vspdData2,Col,Row)		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

'========================================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub 

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub

'========================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData2,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub

'========================================================================================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConFromDt.Focus
	End If
End Sub

Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConToDt.Focus
	End If
End Sub

'========================================================================================================
Sub txtConFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

Sub txtConToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
' 화폐별로 Cell Formating을 재설정한다.
Sub FormatSpreadCellByCurrency()
	on error resume next
	Dim iLngRow
	
	With frm1
		If .vspdData.style.display = "inline" Then
			For	iLngRow = 0 To 12
				Call ReFormatSpreadCellByCellByCurrency(.vspdData,-1, -1,GetKeyPos("A",5),GetKeyPos("A",9 + iLngRow * 6),"A" ,"Q","X","X")
				Call ReFormatSpreadCellByCellByCurrency(.vspdData,-1, -1,GetKeyPos("A",5),GetKeyPos("A",10 + iLngRow * 6),"A" ,"Q","X","X")
			Next
		End If
		
		If .vspdData2.style.display = "inline" Then
			Call ReFormatSpreadCellByCellByCurrency(.vspdData2, -1, -1, GetKeyPos("B",6), GetKeyPos("B",10), "A" , "Q", "X", "X")
			Call ReFormatSpreadCellByCellByCurrency(.vspdData2, -1, -1, GetKeyPos("B",6), GetKeyPos("B",11), "A" , "Q", "X", "X")
		End If
		
	End With
		
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
 <TR >
  <TD <%=HEIGHT_TYPE_00%>></TD>
 </TR>
 <TR HEIGHT=23>
  <TD WIDTH=100%>
   <TABLE <%=LR_SPACE_TYPE_10%>>
    <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <!--TD CLASS="CLSMTABP"-->
     <TD width="240">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTABP"><font color=white>판매계획대실적조회(품목그룹)</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right>&nbsp;</TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR HEIGHT=*>
  <TD WIDTH=100% CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>>
    <TR>
     <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD HEIGHT=20 WIDTH=100%>
      <FIELDSET CLASS="CLSFLD">
       <TABLE <%=LR_SPACE_TYPE_40%>>
	   <TR>
		<TD CLASS="TD5" NOWRAP>판매계획유형</TD>
		<TD CLASS="TD6"><SELECT Name="cboConSpType" ALT="판매계획유형" tag="12XXXU"></SELECT></TD>
        <TD CLASS="TD5" NOWRAP>결과출력형식</TD>
		<TD CLASS="TD6"><SELECT Name="cboDisplayType" ALT="결과출력형식" STYLE="WIDTH: 150px" tag="12"></SELECT></TD> 
	   </TR>
       <TR>       
         <TD CLASS="TD5" NOWRAP>계획년월</TD>
         <TD CLASS="TD6" NOWRAP>
          <TABLE CELLSPACING=0 CELLPADDING=0>
           <TR>
            <TD>
             <script language =javascript src='./js/s2214qa2_fpDateTime2_txtConFromDt.js'></script>
            </TD>
            <TD>
             &nbsp;~&nbsp;
            </TD>
            <TD>
             <script language =javascript src='./js/s2214qa2_fpDateTime2_txtConToDt.js'></script>
            </TD>
           </TR>
          </TABLE>
         </TD>
         <TD CLASS="TD5" NOWRAP>거래구분</TD>
		 <TD CLASS="TD6"><SELECT Name="cboConLocExpFlag" ALT="거래구분" STYLE="WIDTH: 150px" tag="1X"><OPTION Value=""></OPTION></SELECT></TD> 
        </TR> 
        <TR>
         <TD CLASS="TD5" NOWRAP>품목그룹레벨</TD>
		 <TD CLASS="TD6"><SELECT Name="cboConGrpLvl" ALT="품목그룹레벨" STYLE="WIDTH: 150px" tag="12"></SELECT></TD> 
         <TD CLASS="TD5" NOWRAP>품목그룹</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConItemGroupCd" ALT="품목그룹" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConItemGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup 1">&nbsp;<INPUT NAME="txtConItemGroupNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>화폐기준</TD>
		 <TD CLASS="TD6"><SELECT Name="cboConBaseCur" ALT="외화금액" STYLE="WIDTH: 150px" tag="12"></SELECT></TD> 
         <TD CLASS="TD5" NOWRAP>화폐</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConCur" ALT="화폐" TYPE="Text" MAXLENGTH=3 SiZE=10 tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCur" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup 2"></TD>
        </TR>
       </TABLE>
      </FIELDSET>
     </TD>
    </TR>
    <TR>
     <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <script language =javascript src='./js/s2214qa2_vspdData_vspdData.js'></script>
		 <script language =javascript src='./js/s2214qa2_vspdData2_vspdData2.js'></script>
        </TD>
       </TR>
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
 <TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
 </TR>
</TABLE>
<TEXTAREA class=hidden name=txtSelectList tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSelectListDT tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtTailList tag="24" TABINDEX="-1"></TEXTAREA>
</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
