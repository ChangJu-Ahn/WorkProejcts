<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b29ma1.asp
'*  4. Program Name         : Query Item
'*  5. Program Desc         : Query Item Information
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/10
'*  8. Modified date(Last)  :  
'*  9. Modifier (First)     : Lee Woo Guen
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit   

Const BIZ_PGM_QRY_ID	= "b3b29mb1.asp"												'☆: Detail Query 비지니스 로직 ASP명 

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
Dim C_GrossWeight
Dim C_UnitOfGrossWeight
Dim C_CBM
Dim C_CBMDesc
Dim C_DrawNo
Dim C_HsCd
Dim C_HsUnit
Dim C_ItemImageFlg
Dim C_FormalNm
Dim C_ValidFlg


<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgBlnFlgConChg				'☜: Condition 변경 Flag
Dim lgOldRow

Dim IsOpenPop
Dim lgStrPrevKey1

Dim StartDate
Dim lgCharCd1
Dim lgCharCd2
Dim	lgLocalModeFlag

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ItemCd			= 1
	C_ItemNm			= 2
	C_Spec				= 3
	C_ClassCd			= 4
	C_ClassNm			= 5
	C_CharValuedCd1		= 6
	C_CharValuedNm1		= 7
	C_CharValuedCd2		= 8
	C_CharValuedNm2		= 9
	C_BasicUnit			= 10
	C_ItemAcct			= 11
	C_ItemAcctNm		= 12
	C_ItemGroupCd		= 13
	C_ItemGroupNm		= 14
	C_BaseItemCd		= 15
	C_BaseItemNm		= 16
	C_ValidFromDt		= 17
	C_ValidToDt			= 18	
	C_UnitWeight		= 19
	C_UnitOfWeight		= 20
	C_GrossWeight		= 21
	C_UnitOfGrossWeight = 22
	C_CBM				= 23
	C_CBMDesc			= 24
	C_DrawNo			= 25
	C_HsCd				= 26
	C_HsUnit			= 27	
	C_ItemImageFlg		= 28
	C_FormalNm			= 29
	C_ValidFlg			= 30
End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE					'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0									'initializes Group View Size
    
    '---- Coding part--------------------------------------------------------------------
    lgLocalModeFlag	= True
    IsOpenPop = False												
	lgStrPrevKey1 = ""
	lgSortKey = 1
	lgOldRow = 0
	
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "*", "NOCOOKIE", "MA")%>
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtBaseDt.Text	= StartDate
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()    

	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_ValidFlg + 1
		.MaxRows = 0

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
		ggoSpread.SSSetEdit C_ItemAcct,			"품목계정",10
		ggoSpread.SSSetEdit C_ItemAcctNm,		"품목계정",20
		ggoSpread.SSSetEdit C_ItemGroupCd,		"품목그룹",18
		ggoSpread.SSSetEdit C_ItemGroupNm,		"품목그룹명",25
		ggoSpread.SSSetEdit C_BaseItemCd,		"기준품목",20
		ggoSpread.SSSetEdit C_BaseItemNm,		"기준품목명",25
		ggoSpread.SSSetEdit C_ValidFromDt,		"시작일",12, 2
		ggoSpread.SSSetEdit C_ValidToDt,		"종료일",12, 2	
		ggoSpread.SSSetEdit	C_UnitWeight,		"Net중량",12	
		ggoSpread.SSSetEdit	C_UnitOfWeight,		"Net단위",12
		ggoSpread.SSSetEdit	C_GrossWeight,		"Gross중량",15
		ggoSpread.SSSetEdit C_UnitOfGrossWeight,"Gross단위",10
		ggoSpread.SSSetEdit	C_CBM,				"CBM(부피)",15
		ggoSpread.SSSetEdit C_CBMDesc,			"CBM정보", 20
		ggoSpread.SSSetEdit	C_DrawNo,			"도면번호",12
		ggoSpread.SSSetEdit	C_HsCd,				"HS코드",12			
		ggoSpread.SSSetEdit	C_HsUnit,			"HS단위",12	
		ggoSpread.SSSetEdit C_ItemImageFlg,		"사진유무", 12, 2
		ggoSpread.SSSetEdit C_FormalNm,			"품목정식명칭", 30			
		ggoSpread.SSSetEdit	C_ValidFlg,			"유효구분", 12, 2		

		Call ggoSpread.SSSetColHidden(C_ItemAcct, C_ItemAcct, True)
		Call ggoSpread.SSSetColHidden(C_BaseItemNm, C_BaseItemNm, True)
		Call ggoSpread.SSSetColHidden(C_ItemGroupNm, C_ItemGroupNm, True)
		Call ggoSpread.SSSetColHidden(C_UnitWeight, C_CBMDesc, True)
		Call ggoSpread.SSSetColHidden(C_DrawNo, C_DrawNo, True)	
		Call ggoSpread.SSSetColHidden(C_HsCd, C_HsCd, True)
		Call ggoSpread.SSSetColHidden(C_HsUnit, C_HsUnit, True)
		Call ggoSpread.SSSetColHidden(C_ValidFlg, C_ValidFlg, True)	
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SSSetSplit2(1)										'frozen 기능추가 

		.ReDraw = True

		Call SetSpreadLock 

    End With
      
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired 	C_Item,	pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 

Sub InitComboBox()
    On Error Resume Next
    Err.Clear

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))

End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_ItemCd			= iCurColumnPos(1)
			C_ItemNm			= iCurColumnPos(2)
			C_Spec				= iCurColumnPos(3)
			C_ClassCd			= iCurColumnPos(4)
			C_ClassNm			= iCurColumnPos(5)
			C_CharValuedCd1		= iCurColumnPos(6)
			C_CharValuedNm1		= iCurColumnPos(7)
			C_CharValuedCd2		= iCurColumnPos(8)
			C_CharValuedNm2		= iCurColumnPos(9)	
			C_BasicUnit			= iCurColumnPos(10)
			C_ItemAcct			= iCurColumnPos(11)
			C_ItemAcctNm		= iCurColumnPos(12)
			C_ItemGroupCd		= iCurColumnPos(13)
			C_ItemGroupNm		= iCurColumnPos(14)
			C_BaseItemCd		= iCurColumnPos(15)
			C_BaseItemNm		= iCurColumnPos(16)
			C_ValidFromDt		= iCurColumnPos(17)
			C_ValidToDt			= iCurColumnPos(18)
			C_UnitWeight		= iCurColumnPos(19)
			C_UnitOfWeight		= iCurColumnPos(20)
			C_GrossWeight		= iCurColumnPos(21)
			C_UnitOfGrossWeight	= iCurColumnPos(22)
			C_CBM				= iCurColumnPos(23) 
			C_CBMDesc			= iCurColumnPos(24)
			C_DrawNo			= iCurColumnPos(25)
			C_HsCd				= iCurColumnPos(26)
			C_HsUnit			= iCurColumnPos(27)
			C_ItemImageFlg		= iCurColumnPos(28)
			C_FormalNm			= iCurColumnPos(29)
			C_ValidFlg			= iCurColumnPos(30)
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

'========================================================================================
' Function Name : LookupCharCd
' Function Desc : Lookup Characteristic Cd
'========================================================================================
Function LookupCharCd() 

	If gLookUpEnable = False Then Exit Function

    LookupCharCd = False
    
    LayerShowHide(1) 
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal

    With frm1
    
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & ""				'☜: 
		strVal = strVal & "&txtClassCd=" & Trim(.txtClassCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd1=" & Trim(.txtCharValueCd1.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd2=" & Trim(.txtCharValueCd2.value)	'☆: 조회 조건 데이타 
 
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With

    LookupCharCd = True

End Function

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenClassCd()  -------------------------------------------------
'	Name : OpenClassCd()
'	Description : Class PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenClassCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtClasscd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtClassCd.value)	' Class Code
	arrParam(1) = ""							' Class Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "Class_CD"
    arrField(1) = 2 							' Field명(1) : "Class_NM"
	
	iCalledAspName = AskPRAspName("B3B31PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B31PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetClassCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtClassCd.focus
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtClasscd.value = "" Then
		Call DisplayMsgBox("971012", "X", "클래스", "X")
		frm1.txtClasscd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(1) = ""							' Item Name
	arrParam(2) = Trim(frm1.txtClassCd.value)	' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B3B33PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B33PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=900px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
		
End Function

'------------------------------------------  OpenCharValueCd1()  -------------------------------------------------
'	Name : OpenCharValueCd1()
'	Description : Characteristic Value PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCharValueCd1()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtCharValueCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtClasscd.value = "" Then
		Call DisplayMsgBox("971012", "X", "클래스", "X")
		frm1.txtClasscd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = lgCharCd1							' Characteristic Code
	arrParam(1) = Trim(frm1.txtCharValueCd1.value)	' Characteristic Value Code
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

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=490px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCharValueCd1(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCharValueCd1.focus
	
End Function


'------------------------------------------  OpenCharValueCd2()  -------------------------------------------------
'	Name : OpenCharValueCd2()
'	Description : Characteristic Value PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCharValueCd2()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtCharValueCd2.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtClasscd.value = "" Then
		Call DisplayMsgBox("971012", "X", "클래스", "X")
		frm1.txtClasscd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = lgCharCd2							' Characteristic Code
	arrParam(1) = Trim(frm1.txtCharValueCd2.value)	' Characteristic Value Code
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

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=490px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCharValueCd2(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCharValueCd2.focus
	
End Function

'========================================================================================================
'	Name : OpenItemGroup()
'	Desc : Open Item group popup 
'========================================================================================================
Function OpenItemGroup()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "품목그룹"			
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
    
    arrHeader(0) = "품목그룹"		
    arrHeader(1) = "품목그룹명"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If		
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
	
End Function

'------------------------------------------  SetClassCd()  ------------------------------------------------
'	Name : SetClassCd()
'	Description : Class Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetClassCd(byval arrRet)
	frm1.txtClassCd.Value    = arrRet(0)		
	frm1.txtClassNm.Value    = arrRet(1)
	
	Call LookUpCharCd()
	
	frm1.txtClassCd.focus
	Set gActiveElement = document.activeElement 		
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
	
	frm1.txtItemCd.focus
End Function

'------------------------------------------  SetCharValueCd1()  --------------------------------------------------
'	Name : SetCharValueCd1()
'	Description : Characteristic Value Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCharValueCd1(byval arrRet)
	frm1.txtCharValueCd1.Value	= arrRet(0)	
	frm1.txtCharValueNm1.Value   = arrRet(1)
	
	frm1.txtCharValueCd1.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetCharValueCd2()  --------------------------------------------------
'	Name : SetCharValueCd2()
'	Description : Characteristic Value Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCharValueCd2(byval arrRet)
	frm1.txtCharValueCd2.Value	= arrRet(0)	
	frm1.txtCharValueNm2.Value   = arrRet(1)
	
	frm1.txtCharValueCd2.focus
	Set gActiveElement = document.activeElement
End Function

'========================================================================================================
'	Name : SetItemGroup()
'	Desc : Set Item Data
'========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)
	frm1.txtItemGroupNm.Value   = arrRet(1)		

	frm1.txtItemGroupCd.focus 
	Set gActiveElement = document.activeElement		
End Function


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
 Sub Form_Load()

    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
	
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
   
	Call SetDefaultVal
   	Call InitComboBox
    Call InitVariables		
    Call InitSpreadSheet	

	Call SetToolbar("11000000000011")
	
	frm1.txtClassCd.focus
	Set gActiveElement = document.activeElement 

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtClassCd_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtClassCd_onChange()
	If frm1.txtClassCd.value = "" Then
		 frm1.txtClassNm.value = ""
		 lgCharCd1 = ""
		 lgCharCd2 = ""
	Else
		If lgLocalModeFlag = True Then
			Call LookupCharCd()
		Else
			lgLocalModeFlag = True
		End If
	End If
End Sub

Sub txtClassCd_OnKeyDown()
	lgLocalModeFlag = True
End Sub
'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("0000111111")    
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
       
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
    End If
    
	If Row <= 0 Or Col < 0 Then
		ggoSpread.Source = frm1.vspdData
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = "2" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================
Sub vspdData_DblClick(ByVal Col , ByVal Row )

End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row >= NewRow Then Exit Sub
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtClassCd.value = "" Then
		frm1.txtClassNm.value = ""
	End If
		
    If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
	If frm1.txtCharValueCd1.value = "" Then
		frm1.txtCharValueNm1.value = ""
	End If
	
	If frm1.txtCharValueCd2.value = "" Then
		frm1.txtCharValueNm2.value = ""
	End If
			
    If frm1.txtItemGroupCd.value = "" Then
		frm1.txtItemGroupNm.value = ""
	End If
			
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     																'☜: Query db data

    FncQuery = True																'⊙: Processing is OK
   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)											
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
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

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim strAvailableItem
	
	Err.Clear															

	DbQuery = False														

	LayerShowHide(1)
		
	Dim strVal

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				
		strVal = strVal & "&txtClassCd=" & Trim(frm1.hClassCd.value)		
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtCharValueCd1=" & Trim(frm1.hCharValueCd1.value)
		strVal = strVal & "&txtCharValueCd2=" & Trim(frm1.hCharValueCd2.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&cboItemAcct=" & Trim(frm1.hItemAcct.value)
		strVal = strVal & "&txtBaseDt=" & Trim(frm1.hBaseDt.value)		
		strVal = strVal & "&rdoDefaultFlg=" & frm1.hAvailableItem.value	

		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
		strVal = strVal & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtClassCd=" & Trim(frm1.txtClassCd.value)	
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtCharValueCd1=" & Trim(frm1.txtCharValueCd1.value)
		strVal = strVal & "&txtCharValueCd2=" & Trim(frm1.txtCharValueCd2.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&cboItemAcct=" & Trim(frm1.cboItemAcct.value)
		strVal = strVal & "&txtBaseDt=" & Trim(frm1.txtBaseDt.Text)
		
		If frm1.rdoDefaultFlg1.checked = True then
			strAvailableItem = ""
		ElseIf frm1.rdoDefaultFlg2.checked = True then
			strAvailableItem = "Y"
		Else
			strAvailableItem = "N"
		End IF
		strVal = strVal & "&rdoDefaultFlg=" & strAvailableItem
		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
		strVal = strVal & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		
	End If	
	
	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True																					
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()													

 '------ Reset variables area ------
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement		
	End If
	lgIntFlgMode = parent.OPMD_UMODE											
	Call ggoOper.LockField(Document, "Q")								
	Call SetToolbar("11000000000111")
	
End Function



'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>클래스품목조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>클래스</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtClassCd" SIZE=18 MAXLENGTH=16 tag="12XXXU"  ALT="클래스"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnClassCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenClassCd()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
										<INPUT TYPE=TEXT NAME="txtClassNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>사양값1</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharValueCd1" SIZE=18 MAXLENGTH=16 tag="11XXXU"  ALT="사양값1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharValueCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharValueCd1()">
										<INPUT TYPE=TEXT NAME="txtCharValueNm1" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>사양값2</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharValueCd2" SIZE=18 MAXLENGTH=16 tag="11XXXU"  ALT="사양값2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharValueCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharValueCd2()">
										<INPUT TYPE=TEXT NAME="txtCharValueNm2" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=18 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()" >
										<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>품목계정</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" ALT="품목계정" STYLE="Width: 155px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>유효구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg1" CLASS="RADIO" tag="1X" Value="ALL" CHECKED><LABEL FOR="rdoDefaultFlg1">전체</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg2" CLASS="RADIO" tag="1X" Value="Y"><LABEL FOR="rdoDefaultFlg2">예</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg3" CLASS="RADIO" tag="1X" Value="N"><LABEL FOR="rdoDefaultFlg3">아니오</LABEL>
									<TD CLASS=TD5 NOWRAP>기준일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/b3b29ma1_I460757166_txtBaseDt.js'></script> 															
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=* WIDTH=100%>
									<script language =javascript src='./js/b3b29ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TabIndex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="hClassCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hCharValueCd1" tag="24">
<INPUT TYPE=HIDDEN NAME="hCharValueCd2" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24"><INPUT TYPE=HIDDEN NAME="hAvailableItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hBaseDt" tag="24"><INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
</BODY>
</HTML>
