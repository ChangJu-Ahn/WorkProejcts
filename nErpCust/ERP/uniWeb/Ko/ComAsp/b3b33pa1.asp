<%@ LANGUAGE="VBSCRIPT" %>
<!--
=========================================================================================================
'*  1. Module Name          : Production																*
'*  2. Function Name        : Popup Class Item																*
'*  3. Program ID           : b3b33pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Class Item Popup														        *
'*  7. Modified date(First) : 2003/02/07					 											*
'*  8. Modified date(Last)  : 																            *
'*  9. Modifier (First)     : Lee Woo Guen   															*
'* 10. Modifier (Last)      :       																    *
'* 11. Comment              :																			*
'=======================================================================================================-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../inc/incSvrCcm.inc" -->
<!-- #Include file="../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">  <!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_ID = "b3b33pb1.asp"							<% '☆: 비지니스 로직 ASP명 %>

Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_ClassCd
Dim C_ClassNm
Dim C_CharValuedCd1
Dim C_CharValuedNm1
Dim C_CharValuedCd2
Dim C_CharValuedNm2
Dim C_BasicUnit
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_ItemGroupCd
Dim C_ItemGroupNm
Dim C_BaseItemCd
Dim C_BaseItemNm
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_UnitWeight
Dim C_UnitOfWeight
Dim C_DrawNo
Dim C_HsCd
Dim C_HsUnit
Dim C_ItemImageFlg
Dim C_FormalNm
Dim C_ValidFlg

<!-- #Include file="../inc/lgVariables.inc" -->

Dim strReturn                                              '--- Return Parameter Group %>

Dim lgCurDate
Dim lgIntConFlg
Dim IsOpenPop

Dim gblnWinEvent                                             '~~~ ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
                                                             'PopUp Window가 사용중인지 여부를 나타내는 variable %>
Dim arrReturn
Dim arrParent
Dim arrParam					
Dim arrField
Dim PlantCd
Dim lgCharCd1
Dim lgCharCd2

Dim PopupParent
				
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
arrField = arrParent(2)

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ItemCd		= 1
	C_ItemNm		= 2
	C_Spec			= 3
	C_ClassCd		= 4
	C_ClassNm		= 5
	C_CharValuedCd1	= 6
	C_CharValuedNm1	= 7
	C_CharValuedCd2	= 8
	C_CharValuedNm2	= 9
	C_BasicUnit		= 10
	C_ItemAcct		= 11
	C_ItemAcctNm	= 12
	C_ItemGroupCd	= 13
	C_ItemGroupNm	= 14
	C_BaseItemCd	= 15
	C_BaseItemNm	= 16
	C_ValidFromDt	= 17
	C_ValidToDt		= 18	
	C_UnitWeight	= 19
	C_UnitOfWeight	= 20
	C_DrawNo		= 21
	C_HsCd			= 22
	C_HsUnit		= 23	
	C_ItemImageFlg	= 24
	C_FormalNm		= 25
	C_ValidFlg		= 26

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
	lgStrPrevKeyIndex = ""										<%'initializes Previous Key%>
	lgIntConFlg = 0
	lgIntFlgMode = PopupParent.OPMD_CMODE
	hItemCd.value = ""
	hItemNm.value = ""
	hClassCd.value = ""
	hItemAccount.value = ""
	hItemGroup.value = ""
	
    lgSortKey = 1                                       '⊙: initializes sort direction
	gblnWinEvent = False
	
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "*", "NOCOOKIE", "PA")%>
End Sub

'========================================================================================================
' Name : InitComboBox()	
' Desc : Initialize combo value
'========================================================================================================
Sub InitComboBox()
	Call CommonQueryRs(" A.MINOR_CD, A.MINOR_NM ", _
				" B_MINOR A, B_ITEM_ACCT_INF B ", _
				" A.MINOR_CD = B.ITEM_ACCT AND A.MAJOR_CD = " & Filtervar("P1001", "''", "S") &  "  ORDER BY A.MINOR_CD " , _
				 lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(cboItemAccount, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	txtItemCd.value = arrParam(0)
	txtClassCd.value = arrParam(2)
	txtBaseDt.text = StartDate
	
	If arrParam(4) = "" Then
		lgCurDate = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, "1900","01","01")
	Else
		lgCurDate = arrParam(4)
	End If
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Dim i
    
	Call InitSpreadPosVariables()

    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

    vspdData.ReDraw = False
	    
'    vspdData.OperationMode = 3

    vspdData.MaxCols = C_ValidFlg + 1
    vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")
    
	ggoSpread.SSSetEdit C_ItemCd,			"품목",20
	ggoSpread.SSSetEdit C_ItemNm,			"품목명",25
	ggoSpread.SSSetEdit C_Spec,				"규격",25
	ggoSpread.SSSetEdit C_ClassCd,			"클래스",20
	ggoSpread.SSSetEdit C_ClassNm,			"클래스명",25
	ggoSpread.SSSetEdit C_CharValuedCd1,	"사양값1",20
	ggoSpread.SSSetEdit C_CharValuedNm1,	"사양값명1",25	
	ggoSpread.SSSetEdit C_CharValuedCd2,	"사양값2",20
	ggoSpread.SSSetEdit C_CharValuedNm2,	"사양값명2",25	
	ggoSpread.SSSetEdit C_BasicUnit,		"기준단위",12		
	ggoSpread.SSSetEdit C_ItemAcct,			"품목계정",12
	ggoSpread.SSSetEdit C_ItemAcctNm,		"품목계정명",20
	ggoSpread.SSSetEdit C_ItemGroupCd,		"품목그룹",14
	ggoSpread.SSSetEdit C_ItemGroupNm,		"품목그룹명",25
	ggoSpread.SSSetEdit C_BaseItemCd,		"기준품목",20
	ggoSpread.SSSetEdit C_BaseItemNm,		"기준품목명",25
	ggoSpread.SSSetEdit C_ValidFromDt,		"시작일",10, 2
	ggoSpread.SSSetEdit C_ValidToDt,		"종료일",10, 2	
	ggoSpread.SSSetEdit	C_UnitWeight,		"단위중량",12	
	ggoSpread.SSSetEdit	C_UnitOfWeight,		"중량단위",12
	ggoSpread.SSSetEdit	C_DrawNo,			"도면번호",12
	ggoSpread.SSSetEdit	C_HsCd,				"HS코드",12			
	ggoSpread.SSSetEdit	C_HsUnit,			"HS단위",12	
	ggoSpread.SSSetEdit C_ItemImageFlg,		"사진유무", 12, 2
	ggoSpread.SSSetEdit C_FormalNm,			"품목정식명칭", 30			
	ggoSpread.SSSetEdit	C_ValidFlg,			"유효구분", 12, 2		

	Call ggoSpread.SSSetColHidden(C_ItemAcct, C_ItemAcct, True)
	Call ggoSpread.SSSetColHidden(C_BaseItemNm, C_BaseItemNm, True)
	Call ggoSpread.SSSetColHidden(C_ItemGroupNm, C_ItemGroupNm, True)
	Call ggoSpread.SSSetColHidden(C_UnitWeight, C_UnitWeight, True)
	Call ggoSpread.SSSetColHidden(C_UnitOfWeight, C_UnitOfWeight, True)
	Call ggoSpread.SSSetColHidden(C_DrawNo, C_DrawNo, True)	
	Call ggoSpread.SSSetColHidden(C_HsCd, C_HsCd, True)
	Call ggoSpread.SSSetColHidden(C_HsUnit, C_HsUnit, True)
	Call ggoSpread.SSSetColHidden(C_ValidFlg, C_ValidFlg, True)	
	
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)

	ggoSpread.SSSetSplit2(1)										'frozen 기능추가 

	vspdData.ReDraw = True
	
	Call SetSpreadLock()

End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method lock spreadsheet
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_ClassCd		= iCurColumnPos(4)
			C_ClassNm		= iCurColumnPos(5)
			C_CharValuedCd1	= iCurColumnPos(6)
			C_CharValuedNm1	= iCurColumnPos(7)
			C_CharValuedCd2	= iCurColumnPos(8)
			C_CharValuedNm2	= iCurColumnPos(9)	
			C_BasicUnit		= iCurColumnPos(10)
			C_ItemAcct		= iCurColumnPos(11)
			C_ItemAcctNm	= iCurColumnPos(12)
			C_ItemGroupCd	= iCurColumnPos(13)
			C_ItemGroupNm	= iCurColumnPos(14)
			C_BaseItemCd	= iCurColumnPos(15)
			C_BaseItemNm	= iCurColumnPos(16)
			C_ValidFromDt	= iCurColumnPos(17)
			C_ValidToDt		= iCurColumnPos(18)
			C_UnitWeight	= iCurColumnPos(19)
			C_UnitOfWeight	= iCurColumnPos(20)
			C_DrawNo		= iCurColumnPos(21)
			C_HsCd			= iCurColumnPos(22)
			C_HsUnit		= iCurColumnPos(23)
			C_ItemImageFlg	= iCurColumnPos(24)
			C_FormalNm		= iCurColumnPos(25)
			C_ValidFlg		= iCurColumnPos(26)
    End Select    

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
	Set gActiveSpdSheet = vspdData
    Call SetPopupMenuItemInf("0000111111")

    If vspdData.MaxRows <= 0 Then Exit Sub
   	    
    If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              ' 타이틀 cell을 dblclick했거나....
	   Exit Function
	End If

	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyDown
'   Event Desc :
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
	If KeyAscii = 27 Then
 		Call CancelClick()
	ElseIf KeyAscii = 13 and vspdData.ActiveRow > 0 Then
		Call OkClick()
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With vspdData
		If Row >= NewRow Then
			Exit Sub
		End If
		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				If DbQuery = False Then
					Exit Sub
				End If
			End If
		End If
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End If
		End If
	End If
End Sub

'------------------------------------------  OpenClassCd()  -------------------------------------------------
'	Name : OpenClassCd()
'	Description : Class PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenClassCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "클래스팝업"					<%'팝업 명칭 %>
	arrParam(1) = "B_CLASS"							<%'TABLE 명칭 %>
	arrParam(2) = Trim(txtClasscd.Value)			<%'Code Condition%>
	arrParam(3) = ""								<%'Name Cindition%>
	arrParam(4) = ""								<%'Where Condition%>
	arrParam(5) = "클래스"						<%'TextBox 명칭 %>
	
    arrField(0) = "CLASS_CD"						<%'Field명(0)%>
    arrField(1) = "CLASS_NM"						<%'Field명(1)%>
    arrField(2) = "CLASS_DIGIT"						<%'Field명(2)%>
    
    arrHeader(0) = "클래스"						<%'Header명(0)%>
    arrHeader(1) = "클래스명"					<%'Header명(1)%>
    arrHeader(2) = "클래스자리수"				<%'Header명(2)%>
    
	arrRet = window.showModalDialog("../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetClassCd(arrRet)
	End If
	
	Call SetFocusToDocument("P")
	txtClasscd.focus
	
End Function

'------------------------------------------  OpenCharValueCd1()  -------------------------------------------------
'	Name : OpenCharValueCd1()
'	Description : Characteristic Value PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCharValueCd1()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = lgCharCd1							' Characteristic Code
	arrParam(1) = Trim(txtCharValueCd1.value)	' Characteristic Value Code
	arrParam(2) = ""								' ----------
	arrParam(3) = ""								' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 								' Field명(0) : "Char_CD"
    arrField(1) = 2 								' Field명(1) : "Char_Value_CD"
    
	iCalledAspName = AskPRAspName("B3B32PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B32PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
		"dialogWidth=490px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCharValueCd1(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtCharValueCd1.focus
	
End Function


'------------------------------------------  OpenCharValueCd2()  -------------------------------------------------
'	Name : OpenCharValueCd2()
'	Description : Characteristic Value PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCharValueCd2()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = lgCharCd2							' Characteristic Code
	arrParam(1) = Trim(txtCharValueCd2.value)	' Characteristic Value Code
	arrParam(2) = ""								' ----------
	arrParam(3) = ""								' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 								' Field명(0) : "Char_CD"
    arrField(1) = 2 								' Field명(1) : "Char_Value_CD"
    
	iCalledAspName = AskPRAspName("B3B32PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B32PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
		"dialogWidth=490px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCharValueCd2(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtCharValueCd2.focus
	
End Function

'========================================================================================================
'	Name : OpenItemGroup()
'	Desc : Open Item group popup 
'========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"				<%'팝업 명칭 %>
	arrParam(1) = "B_ITEM_GROUP"					<%'TABLE 명칭 %>
	arrParam(2) = Trim(txtItemGroup.Value)			<%'Code Condition%>
	arrParam(3) = ""								<%'Name Cindition%>
	arrParam(4) = ""								<%'Where Condition%>
	arrParam(5) = "품목그룹"					<%'TextBox 명칭 %>
	
    arrField(0) = "ITEM_GROUP_CD"					<%'Field명(0)%>
    arrField(1) = "ITEM_GROUP_NM"					<%'Field명(1)%>
    
    arrHeader(0) = "품목그룹"					<%'Header명(0)%>
    arrHeader(1) = "품목그룹명"					<%'Header명(1)%>
    
	arrRet = window.showModalDialog("../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If
	
	Call SetFocusToDocument("P")
	txtItemGroup.focus
		
End Function

'------------------------------------------  SetClassCd()  ------------------------------------------------
'	Name : SetClassCd()
'	Description : Class Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetClassCd(byval arrRet)
	txtClassCd.Value    = arrRet(0)		
	txtClassNm.Value    = arrRet(1)
	
	Call LookupCharCd()
	
	txtClassCd.focus
	Set gActiveElement = document.activeElement 		
End Function

'------------------------------------------  SetCharValueCd1()  --------------------------------------------------
'	Name : SetCharValueCd1()
'	Description : Characteristic Value Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCharValueCd1(byval arrRet)
	txtCharValueCd1.Value	= arrRet(0)	
	txtCharValueNm1.Value   = arrRet(1)
	
	txtCharValueCd1.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetCharValueCd2()  --------------------------------------------------
'	Name : SetCharValueCd2()
'	Description : Characteristic Value Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCharValueCd2(byval arrRet)
	txtCharValueCd2.Value	= arrRet(0)	
	txtCharValueNm2.Value   = arrRet(1)
	
	txtCharValueCd2.focus
	Set gActiveElement = document.activeElement
End Function

'========================================================================================================
'	Name : SetItemGroup()
'	Desc : Set Item Data
'========================================================================================================
Function SetItemGroup(byval arrRet)
	txtItemGroup.Value    = arrRet(0)		
	txtItemGroup.focus 		
End Function

'=======================================================================================================
'   Event Name : txtClassCd_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtClassCd_onChange()
	If txtClassCd.value = "" Then
		txtClassNm.value = ""
		lgCharCd1 = ""
		lgCharCd2 = ""
	Else
		Call LookUpCharCd()
	End If
End Sub

'========================================================================================================
'	Name : OKClick()
'	Desc : handle ok icon click event
'========================================================================================================
Function OKClick()
	Dim i, iCurColumnPos
	
	If vspdData.MaxRows > 0 Then
		
		Redim arrReturn(UBound(arrField))

        ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 
			
		For i = 0 To UBound(arrField)
			If arrField(i) <> "" Then
				vspddata.Col = iCurColumnPos(CInt(arrField(i)))
				arrReturn(i) = vspdData.Text
			End If
		Next
		
		Self.Returnvalue = arrReturn
	End If

	Self.Close()
				
End Function

'========================================================================================================
'	Name : CancelClick()
'	Desc : handle  Cancel click event
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
'	Name : MousePointer()
'	Desc : 
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        txtBaseDt.Action = 7
        Call SetFocusToDocument("P")
		txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call FncQuery()
	End If
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../CShared/image/Query.gif","../../CShared/image/OK.gif","../../CShared/image/Cancel.gif")
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	
	Call InitVariables
	
	Call InitComboBox()
	Call SetDefaultVal()
	Call InitSpreadSheet()
	
	If DbQuery = False Then
		Exit Sub
	End If
End Sub

'========================================================================================================
'	Name : FncQuery()
'	Desc : 
'========================================================================================================
Function FncQuery()
	FncQuery = False
	Call InitVariables()
		
	If txtItemCd.value = "" And txtItemNm.value  <> "" Then
		lgIntConFlg = 1							'Code로 조회하는 지 이름으로 조회하는 지 구분 
	End If
	
	vspdData.MaxRows = 0						'Grid 초기화 
	
	If DbQuery = False Then
		Exit Function
	End If
	
	FncQuery = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    'Err.Clear         
                                                          '☜: Protect system from crashing
	'-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
    DbQuery = False                                                         '⊙: Processing is NG
	
	Call LayerShowHide(1)													<%'⊙: 작업진행중 표시 %>	
    
    Dim strVal
    
 	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				'☜: 
 	
	strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)					'☆: 조회 조건 데이타	
	strVal = strVal & "&txtItemNm=" & txtItemNm.value
	strVal = strVal & "&txtClassCd=" & Trim(txtClassCd.value)
	strVal = strVal & "&txtCharValueCd1=" & Trim(txtCharValueCd1.value)
	strVal = strVal & "&txtCharValueCd2=" & Trim(txtCharValueCd2.value)
	strVal = strVal & "&txtItemGroup=" & Trim(txtItemGroup.value)					
	strVal = strVal & "&cboItemAccount=" & Trim(cboItemAccount.value)
		
	If rdoValidFlg1.checked = True Then
		strVal = strVal & "&rdoValidFlg= Y" 
	Else
		strVal = strVal & "&rdoValidFlg= N" 
	End If
		
	strVal = strVal & "&lgIntConFlg=" & lgIntConFlg
	strVal = strVal & "&lgCurDate=" & txtBaseDt.Text
	strVal = strVal & "&txtMaxRows="         & vspdData.MaxRows
	strVal = strVal & "&lgStrPrevKeyIndex="  & lgStrPrevKeyIndex   

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True                                                          '⊙: Processing is NG
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
	End If
    lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
End Function

'========================================================================================
' Function Name : LookupCharCd
' Function Desc : Lookup Characteristic Cd
'========================================================================================
Function LookupCharCd() 

    LookupCharCd = False
    
    LayerShowHide(1) 
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
   
	strVal = BIZ_PGM_ID & "?txtMode=" & ""				'☜: 
	strVal = strVal & "&txtClassCd=" & Trim(txtClassCd.value)				'☆: 조회 조건 데이타 
	strVal = strVal & "&txtCharValueCd1=" & Trim(txtCharValueCd1.value)		'☆: 조회 조건 데이타 
	strVal = strVal & "&txtCharValueCd2=" & Trim(txtCharValueCd2.value)		'☆: 조회 조건 데이타 
 
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    LookupCharCd = True

End Function

</SCRIPT>
<!-- #Include file="../inc/Uni2kCMCom.inc" -->
</HEAD>
<%
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>
			<TR>
				<TD CLASS=TD5 NOWRAP>품목</TD>
				<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=22 MAXLENGTH=18 tag="11XXXU" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 MAXLENGTH=40 tag="11" ALT="품목명"></TD>
				<TD CLASS=TD5 NOWRAP>클래스</TD>
				<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtClassCd" SIZE=20 MAXLENGTH=16 tag="11XXXU"  ALT="클래스"><IMG SRC="../../CShared/image/btnPopup.gif" NAME="btnClassCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenClassCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtClassNm" SIZE=20 tag="14"></TD>
			</TR>
			<TR>
				<TD CLASS=TD5 NOWRAP>사양값1</TD>
				<TD CLASS=TD6 NOWRAP>
					<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharValueCd1" SIZE=18 MAXLENGTH=16 tag="11XXXU"  ALT="사양값1"><IMG SRC="../../CShared/image/btnPopup.gif" NAME="btnCharValueCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharValueCd1()">
					<INPUT TYPE=TEXT NAME="txtCharValueNm1" SIZE=20 tag="14"></TD>
				<TD CLASS=TD5 NOWRAP>사양값2</TD>
				<TD CLASS=TD6 NOWRAP>
					<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharValueCd2" SIZE=18 MAXLENGTH=16 tag="11XXXU"  ALT="사양값2"><IMG SRC="../../CShared/image/btnPopup.gif" NAME="btnCharValueCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharValueCd2()">
					<INPUT TYPE=TEXT NAME="txtCharValueNm2" SIZE=20 tag="14"></TD>
			</TR>
			<TR>
				<TD CLASS=TD5 NOWRAP>품목그룹</TD>
				<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemGroup" SIZE=20 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../CShared/image/btnPopup.gif" NAME="btnItemGroup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">
				<TD CLASS=TD5 NOWRAP>품목계정</TD>
				<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAccount" ALT="품목계정" STYLE="Width: 165px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
			</TR>
			<TR>
				<TD CLASS=TD5 NOWRAP>유효구분</TD>
				<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg1" Value="Y" CLASS="RADIO" tag="1X" CHECKED><LABEL FOR="rdoValidFlg1">예</LABEL>
							<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg2" Value="N" CLASS="RADIO" tag="1X"><LABEL FOR="rdoValidFlg2">아니오</LABEL></TD>
				<TD CLASS=TD5 NOWRAP>기준일</TD>
				<TD CLASS=TD6 NOWRAP>
					<script language =javascript src='./js/b3b33pa1_I422532308_txtBaseDt.js'></script>
				</TD>
			</TR>
		</TABLE></FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/b3b33pa1_vspdData_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK = "FncQuery()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemNm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hClassCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemGroup" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemAccount" tag="24" TABINDEX="-1">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>
