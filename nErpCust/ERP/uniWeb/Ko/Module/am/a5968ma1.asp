<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        :
'*  3. Program ID           : A5968MA1
'*  4. Program Name         : 상여 지급기준 등록 
'*  5. Program Desc         : 상여 지급기준 등록 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/02/15
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : 권기수 
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================
Const BIZ_PGM_ID = "A5968MB1.asp"                                      'Biz Logic ASP

'========================================================================================================
Dim C_PAY_TYPE_CD
Dim C_PAY_TYPE_BT
Dim C_PAY_TYPE_NM 
Dim C_PAY_TYPE_CD_H 
Dim C_PAY_MM  
Dim C_FROM_MM 
Dim C_TO_MM   


Const COOKIE_SPLIT      = 4877	                                      '☆: Cookie Split String

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop
Dim lsConcd

'========================================================================================================
Sub initSpreadPosVariables()
	 C_PAY_TYPE_CD     = 1
	 C_PAY_TYPE_BT     = 2
	 C_PAY_TYPE_NM     = 3
	 C_PAY_TYPE_CD_H   = 4
	 C_PAY_MM          = 5
	 C_FROM_MM         = 6
	 C_TO_MM           = 7
End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE
	lgBlnFlgChgValue  = False
	lgIntGrpCount     = 0
    lgStrPrevKey      = ""
    lgStrPrevKeyIndex = ""
    lgSortKey         = 1
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================

Sub SetDefaultVal()
	Dim strYear, strMonth, strDay
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)
	
	frm1.fpdtWk_yyyy.Year	= strYear
	frm1.fpdtWk_yyyy.Month	= strMonth
	frm1.fpdtWk_yyyy.Day	= strDay
	
	Call ggoOper.FormatDate(frm1.fpdtWk_yyyy, gDateFormat, 3)
	frm1.fpdtWk_yyyy.focus
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value
'========================================================================================================
Sub CookiePage(Kubun)
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : Make key stream of query or delete condition data
'========================================================================================================
Sub MakeKeyStream(pOpt)

    Dim strYYYY
    Dim strYear,strMonth,strDay


    Call ExtractDateFrom(frm1.fpdtWk_yyyy.text,frm1.fpdtWk_yyyy.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

    strYYYY = Trim(frm1.fpdtWk_yyyy.text)
    lgKeyStream = strYYYY & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(frm1.txtBonusCd.value) & parent.gColSep                      '상여종류 
   '------ Developer Coding part (End   ) --------------------------------------------------------------

End Sub


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
    ggoSpread.SetCombo "01" & vbTab& "02" & vbTab& "03" & vbTab& "04" & vbTab& "05" & vbTab& "06" & vbTab& "07" & vbTab& "08" & vbTab& "09" & vbTab& "10" & vbTab& "11" & vbTab& "12", C_PAY_MM
    ggoSpread.SetCombo "01" & vbTab& "02" & vbTab& "03" & vbTab& "04" & vbTab& "05" & vbTab& "06" & vbTab& "07" & vbTab& "08" & vbTab& "09" & vbTab& "10" & vbTab& "11" & vbTab& "12", C_FROM_MM
    ggoSpread.SetCombo "01" & vbTab& "02" & vbTab& "03" & vbTab& "04" & vbTab& "05" & vbTab& "06" & vbTab& "07" & vbTab& "08" & vbTab& "09" & vbTab& "10" & vbTab& "11" & vbTab& "12", C_TO_MM
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  
		.ReDraw = false

       .MaxCols = C_TO_MM + 1
	   .Col = .MaxCols
       .ColHidden = True

'	   .Col = C_PAY_TYPE_CD_H
'       .ColHidden = True

		ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit    C_PAY_TYPE_CD     , "상여종류"   ,16,,, 01,2
		ggoSpread.SSSetButton  C_PAY_TYPE_BT
		ggoSpread.SSSetEdit    C_PAY_TYPE_NM     , "상여종류명" ,35,,, 35,2
		ggoSpread.SSSetEdit    C_PAY_TYPE_CD_H   , "상여종류"   ,01,,, 01,2
		ggoSpread.SSSetCombo   C_PAY_MM          , "실지급월"   ,19
		ggoSpread.SSSetCombo   C_FROM_MM         , "적용시작월"   ,19
		ggoSpread.SSSetCombo   C_TO_MM           , "적용종료월"   ,19
       
		Call ggoSpread.MakePairsColumn(C_PAY_TYPE_CD,C_PAY_TYPE_BT,"1")
		Call ggoSpread.SSSetColHidden(C_PAY_TYPE_CD_H,C_PAY_TYPE_CD_H,True)
		
	   .ReDraw = true

       Call SetSpreadLock
    End With
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired    C_PAY_TYPE_CD , -1 , C_PAY_TYPE_CD
        ggoSpread.SpreadLock       C_PAY_TYPE_NM , -1 , C_PAY_TYPE_NM
        ggoSpread.SSSetRequired    C_PAY_MM , -1, C_PAY_MM
        ggoSpread.SSSetRequired    C_FROM_MM , -1, C_FROM_MM
        ggoSpread.SSSetRequired    C_TO_MM , -1, C_TO_MM
        ggoSpread.SpreadLock	.vspdData.MaxCols, -1,.vspdData.MaxCols
        .vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
            ggoSpread.SSSetRequired    C_PAY_TYPE_CD , pvStarRow, pvEndRow
            ggoSpread.SSSetProtected   C_PAY_TYPE_NM , pvStarRow, pvEndRow
            ggoSpread.SSSetRequired    C_PAY_MM , pvStarRow, pvEndRow
            ggoSpread.SSSetRequired    C_FROM_MM , pvStarRow, pvEndRow
            ggoSpread.SSSetRequired    C_TO_MM , pvStarRow, pvEndRow
        .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_PAY_TYPE_CD     = iCurColumnPos(1)
		C_PAY_TYPE_BT     = iCurColumnPos(2)
		C_PAY_TYPE_NM     = iCurColumnPos(3)
		C_PAY_TYPE_CD_H   = iCurColumnPos(4)
		C_PAY_MM          = iCurColumnPos(5)
		C_FROM_MM         = iCurColumnPos(6)
		C_TO_MM           = iCurColumnPos(7)
	End Select
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If
       Next
    End If
End Sub

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'======================================================================================================
Function OpenBonus()
	Dim arrRet
	Dim arrParam(6), arrField(5), arrHeader(5)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "상여종류 팝업"		    	    <%' 팝업 명칭 %>
	arrParam(1) = "b_minor" 	<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtBonusCd.value		<%' Code Cindition%>
	arrParam(3) = ""								<%' Name Condition%>
	arrParam(4) = "major_cd = " & FilterVar("H0040", "''", "S") & "  AND MINOR_CD >= " & FilterVar("2", "''", "S") & "  AND MINOR_CD <= " & FilterVar("9", "''", "S") & " "
	arrParam(5) = "상여 코드"

    arrField(0) = "minor_cd"					<%' Field명(0)%>
    arrField(1) = "minor_nm"	     			<%' Field명(1)%>

    arrHeader(0) = "상여코드"				<%' Header명(0)%>
    arrHeader(1) = "상여명"				<%' Header명(1)%>


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBonusCd.focus
		Exit Function
	Else
		Call SetBonus(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetBonus()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetBonus(Byval arrRet)
	With frm1
		.txtBonusCd.focus
		.txtBonusCd.value = arrRet(0)
		.txtBonus.value	   = arrRet(1)
	End With
End Function

'======================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"		    				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"								' TABLE 명칭 
	arrParam(2) = frm1.txtBizAreaCd.value					' Code Condition
	arrParam(3) = "" 		            					' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "사업장"

    arrField(0) = "BIZ_AREA_CD"	     						' Field명(1)
    arrField(1) = "BIZ_AREA_NM"								' Field명(0)


    arrHeader(0) = "사업장코드"			    			' Header명(0)
    arrHeader(1) = "사업장명"								' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetBizArea(arrRet)
	End If

End Function

'======================================================================================================
Function SetBizArea(Byval arrRet)
	With frm1
		.txtBizAreaCd.focus
		.txtBizAreaCd.value = arrRet(0)
		.txtBizArea.value	   = arrRet(1)
	End With
End Function

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Select Case Col				'추가부분을 위해..select로..
	    Case C_PAY_TYPE_BT        'Cost center
	        frm1.vspdData.Col = C_PAY_TYPE_CD
            Call OpenBonus2(frm1.vspdData.Text, Row)
	End Select
	Call SetActiveCell(frm1.vspdData,Col - 1,frm1.vspdData.ActiveRow ,"M","X","X")

End Sub

'======================================================================================================
Function OpenBonus2(var_cd, Row)
	Dim arrRet
	Dim arrParam(6), arrField(5), arrHeader(5)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "상여종류 팝업"		    	    <%' 팝업 명칭 %>
	arrParam(1) = " b_minor" 	<%' TABLE 명칭 %>
	frm1.vspdData.Col = C_PAY_TYPE_CD
	Frm1.vspdData.Row = Row
	arrParam(2) = frm1.vspdData.text		<%' Code Cindition%>
	arrParam(3) = ""								<%' Name Condition%>
	arrParam(4) = "major_cd = " & FilterVar("H0040", "''", "S") & "  AND MINOR_CD >= " & FilterVar("2", "''", "S") & "  AND MINOR_CD <= " & FilterVar("9", "''", "S") & " "
	arrParam(5) = "배부기준 코드"

    arrField(0) = "minor_cd"					<%' Field명(0)%>
    arrField(1) = "minor_nm"	     			<%' Field명(1)%>

    arrHeader(0) = "상여코드"				<%' Header명(0)%>
    arrHeader(1) = "상여명"				<%' Header명(1)%>


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBonus2(arrRet, Row)
	End If

End Function
'=======================================================================================================
'======================================================================================================
'	Name : SetBonus2()
'	Description : Item Popup에서 Return되는 값 setting
'======================================================================================================
Function SetBonus2(Byval arrRet, Row)
	With frm1
        .vspdData.Col = C_PAY_TYPE_CD
		.vspdData.text = arrRet(0)
        .vspdData.Col = C_PAY_TYPE_NM
		.vspdData.text = arrRet(1)
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
	End With
End Function

'======================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")

    Call InitSpreadSheet

	Call InitVariables
    Call SetDefaultVal
    Call InitComboBox

	Call SetToolbar("1100110100101111")
    Call BtnDisabled(1)
	Call CookiePage (0)
End Sub

'======================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'======================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   --------------------------------------------------------------

    Call BtnDisabled(1)
    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncQuery = True

End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
End Function

'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD

    FncSave = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
    ggoSpread.Source = frm1.vspdData

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call MakeKeyStream("X")
	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncSave = True
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False
    Err.Clear

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow

            .ReDraw = True
		    .Focus
		 End If
	End With

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	' Clear key field
	'----------------------------------------------------------------------------------------------------
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.col = C_PAY_TYPE_CD 
			frm1.vspdData.text = ""
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.col = C_PAY_TYPE_NM 
			frm1.vspdData.text = ""
			
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCopy = True
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel()
    FncCancel = False
    Err.Clear

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCancel = False
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(Byval pvRowCnt)
    Err.Clear
	Dim imRow
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If
	If Not chkField(Document, "1") Then
       Exit Function
    End If


	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow

        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With

    Set gActiveElement = document.ActiveElement
    FncInsertRow = True
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False
    Err.Clear

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if

    With Frm1.vspdData
    	.focus
    	ggoSpread.Source = frm1.vspdData
    	lDelRows = ggoSpread.DeleteRow
    End With

    Set gActiveElement = document.ActiveElement
    FncDeleteRow = True
                                                 '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False
    Err.Clear
	Call Parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================================
Function FncPrev()
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext()
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel()
    FncExcel = False
    Err.Clear
	Call Parent.FncExport(parent.C_MULTI)
    FncExcel = True
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
    FncFind = False
    Err.Clear

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'======================================================================================================
Function DbQuery()
	Dim strVal

    Err.Clear
    DbQuery = False

    if LayerShowHide(1) = false then
		exit function
	end if

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
    End With

    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
End Function
'======================================================================================================
Function DbSave()
    Dim pP21011
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
	Dim strVal
	Dim strDel
	Dim strStDt, strEdDt,strExDt '시작년월,종료월,실지급월 

    Err.Clear

    DbSave = False
   if LayerShowHide(1) = false then
		exit function
	end if
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
  	With frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    lGrpCnt = 1

  	With frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0
           Select Case .vspdData.Text

            Case ggoSpread.InsertFlag                                      '☜: Update
										              strVal = strVal & "C" & parent.gColSep
                                                      strVal = strVal & lRow & parent.gColSep
                                                      strval = strval & Trim(.txtBonusCd.value) & parent.gColSep
                .vspdData.Col = C_PAY_TYPE_CD       : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_PAY_MM            : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_FROM_MM           : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_TO_MM             : strVal = strVal & Trim(.vspdData.text) & parent.gRowSep
                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U" & parent.gColSep
                                                      strVal = strVal & lRow & parent.gColSep
                .vspdData.Col = C_PAY_TYPE_CD       : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_PAY_MM            : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_FROM_MM           : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_TO_MM             : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_PAY_TYPE_CD_H     : strVal = strVal & Trim(.vspdData.text) & parent.gRowSep
                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                      strDel = strDel & "D" & parent.gColSep
                                                      strDel = strDel & lRow & parent.gColSep
                .vspdData.Col = C_PAY_TYPE_CD_H      : strDel = strDel & Trim(.vspdData.text) & parent.gRowSep
                lGrpCnt = lGrpCnt + 1
           End Select
       Next

       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
  	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      =  strDel & strVal

	End With


	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear
    DbDelete = False
    if LayerShowHide(1) = false then
		exit function
	end if

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003

    DbDelete = True
    Call RunMyBizASP(MyBizASP, strVal)
End Function

'======================================================================================================
Sub DbQueryOk()


    lgIntFlgMode = parent.OPMD_UMODE    
	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
    end if
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables
	'------ Developer Coding part (Start)  --------------------------------------------------------------
    Call MakeKeyStream("X")
     Call ggoOper.ClearField(Document, "2")									     '⊙: Clear Contents  Field
     ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	'------ Developer Coding part (End )   --------------------------------------------------------------

    Call DisableToolBar(parent.TBC_QUERY)
    If DBQuery = false Then
        Call RestoreToolBar()
        Exit Sub
    End If

	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
    end if
   Set gActiveElement = document.ActiveElement
End Sub

'======================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	Call InitVariables()
	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
    end if
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD,Input_alloc,  EFlag
    Dim lPay_Type, To_MM, From_MM, PAY_MM

	EFlag = False

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Select Case Col
		Case C_PAY_TYPE_CD
			Frm1.vspdData.Col = C_PAY_TYPE_CD
			lPay_Type = Frm1.vspdData.Text

			IF (lPay_Type = "" OR lPay_Type = NULL) THEN
			    Frm1.vspdData.Col = C_PAY_TYPE_NM
			    Frm1.vspdData.Text = ""
			Else
			    IntRetCD = CommonQueryRs( " minor_nm ", " b_minor " , " major_cd = " & FilterVar("H0040", "''", "S") & "  and minor_cd =  " & FilterVar(lPay_Type, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			    If IntRetCD = False Then
				    Call DisplayMsgBox("110100","X","X","X")
				    Frm1.vspdData.Col = C_PAY_TYPE_CD
				    Frm1.vspdData.Text = ""
				    Frm1.vspdData.Col = C_PAY_TYPE_NM
				    Frm1.vspdData.Text = ""
				    frm1.vspdData.Col = Col
				    Frm1.vspdData.Action = 0
				    Set gActiveElement = document.activeElement
				    EFlag = True
			    Else
				    Frm1.vspdData.Col = C_PAY_TYPE_NM
				    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
			    End If
			End IF
		Case C_PAY_MM
			PAY_MM = Frm1.vspdData.Text
			Frm1.vspdData.Col = C_FROM_MM
			From_MM = Frm1.vspdData.Text
			If Trim(From_MM) <> "" Then
				If Trim(From_MM) > Trim(PAY_MM) Then
					IntRetCD = DisplayMsgBox("970023","X","실지급월","적용시작월")
					Frm1.vspdData.Col = C_PAY_MM
					Frm1.vspdData.Text = ""
					Frm1.vspdData.Action = 0
				End If
			End If
		Case C_TO_MM
			To_MM = Frm1.vspdData.Text
			Frm1.vspdData.Col = C_FROM_MM
			From_MM = Frm1.vspdData.Text
			If Trim(From_MM) <> "" Then
				If Trim(From_MM) > Trim(To_MM) Then
					IntRetCD = DisplayMsgBox("970025","X","적용시작월","적용종료월")
					Frm1.vspdData.Col = C_TO_MM
					Frm1.vspdData.Text = ""
					Frm1.vspdData.Action = 0
				End If
			End If
		Case C_FROM_MM
			From_MM = Frm1.vspdData.Text
			Frm1.vspdData.Col = C_TO_MM
			To_MM = Frm1.vspdData.Text
			Frm1.vspdData.Col = C_PAY_MM
			PAY_MM = Frm1.vspdData.Text
			If Trim(To_MM) <> "" Then
				If Trim(From_MM) > Trim(To_MM) Then
					IntRetCD = DisplayMsgBox("970025","X","적용시작월","적용종료월")
					Frm1.vspdData.Col = C_FROM_MM
					Frm1.vspdData.Text = ""
					Frm1.vspdData.Action = 0
				End If	
			End If
			If Trim(PAY_MM) <> "" Then
				If Trim(From_MM) > Trim(PAY_MM) Then
					IntRetCD = DisplayMsgBox("970023","X","실지급월","적용시작월")
					Frm1.vspdData.Col = C_FROM_MM
					Frm1.vspdData.Text = ""
					Frm1.vspdData.Action = 0
				End If
			End If
	End Select			
				
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)


	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0

    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

 If Row = 0 Then
  ggoSpread.Source = frm1.vspdData
  If lgSortKey = 1 Then
   ggoSpread.SSSort
   lgSortKey = 2
  Else
   ggoSpread.SSSort ,lgSortKey
   lgSortKey = 1
  End If    
 End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)
End Sub


'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    Dim iColumnName
 	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
  If Row >= NewRow Then
      Exit Sub
  End If
    End With
End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevKeyIndex <> "" Then
      	   Call DisableToolBar(parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
        End If
    End if
End Sub

Sub fpdtWk_yyyy_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yyyy.Action = 7
		Call SetFocusToDocument("M")
		Frm1.fpdtWk_yyyy.Focus
	End If
End Sub

'=======================================================================================================
'   Event Name : fpdtWk_yyyy_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub fpdtWk_yyyy_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>



<BODY TABINDEX="-1" SCROLL="No">
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>상여기준정보등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>년도</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/a5968ma1_fpDateTime3_fpdtWk_yyyy.js'></script>
									</TD>
									<TD NOWRAP CLASS="TD5">상여종류</TD>
									<TD NOWRAP CLASS="TD6">
										<INPUT TYPE=TEXT   NAME="txtBonusCd" SIZE=10 MAXLENGTH=20 tag="15XXXU" ALT="상여종류" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBonus" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBonus()">
                                        <INPUT TYPE=TEXT   NAME="txtBonus" TAG="14XXU" SIZE=22 MAXLENGTH="50">
									</TD>
                                </TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_30%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/a5968ma1_vspdData_vspdData.js'></script>
								    </TD>
								</TR>
							</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD NOWRAP WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"   TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"  TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

