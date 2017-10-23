<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1112MA2
'*  4. Program Name         : 공급처별단가등록(Multi)
'*  5. Program Desc         : 공급처별단가등록(Multi)
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/14
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ahn Jung Je
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" --> 

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit											

CONST BIZ_PGM_ID = "M1112MB2_KO441.ASP"
<!-- #Include file="../../inc/lgvariables.inc" -->	

'@Grid_Column
Dim lgIsOpenPop
Dim C_PlantCd
Dim C_PlantPopup
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemPopup
Dim C_ItemNm
Dim C_OrderUnit
Dim C_OrderUnitPopup
Dim C_Curr
Dim C_CurrPopup
Dim C_CostDt
Dim C_SupplierCd
Dim C_SupplierPopup
Dim C_SupplierNm
Dim C_Cost
Dim C_PrcFlg
Dim C_PrcFlgNm
Dim C_Remark

'@Global_Var
Dim lgSortKey1
Dim IsOpenPop
Dim EndDate, StartDate
   
EndDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gDateFormat)

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
	lgPageNo = ""
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
	Call SetToolBar("111011010011111")				'버튼 툴바 제어 
	
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If

	frm1.txtAppToDt.text = EndDate
	frm1.txtAppfrDt.text = StartDate	

	if lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
  	frm1.txtPlantCd.value = lgPLCd
	End If
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		.ReDraw = false

		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread
		'단가구분 수정 
		.MaxCols = C_Remark + 1
	   '.MaxCols = C_Cost + 1
	   .MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit   C_PlantCd, "공장",  10,,,,2
		ggoSpread.SSSetButton C_PlantPopup
		ggoSpread.SSSetEdit   C_PlantNm, "공장명", 20
		ggoSpread.SSSetEdit   C_ItemCd, "품목", 15,,,,2
		ggoSpread.SSSetButton C_ItemPopup
		ggoSpread.SSSetEdit   C_ItemNm, "품목명", 20
		ggoSpread.SSSetEdit   C_OrderUnit, "발주단위", 10,,,,2
		ggoSpread.SSSetButton C_OrderUnitPopup		
		ggoSpread.SSSetEdit   C_Curr, "화폐", 10,,,,2
		ggoSpread.SSSetButton C_CurrPopup
		ggoSpread.SSSetDate   C_CostDt, "단가적용일", 15,2,gDateFormat
		ggoSpread.SSSetEdit   C_SupplierCd, "공급처", 10,,,,2
		ggoSpread.SSSetButton C_SupplierPopup
		ggoSpread.SSSetEdit   C_SupplierNm, "공급처명", 20
        ggoSpread.SSSetFloat  C_Cost,"단가",20,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, 1,,"Z"
        '단가구분 추가 
        ggoSpread.SSSetCombo  C_PrcFlg , "단가구분" , 10 
        ggoSpread.SSSetCombo  C_PrcFlgNm, "단가구분" , 10
        '20050503 비고 관련추가 
        ggoSpread.SSSetEdit   C_Remark, "비고", 50,0,,240,0

		Call ggoSpread.MakePairsColumn(C_ItemCd,C_ItemPopup)
		Call ggoSpread.MakePairsColumn(C_OrderUnit,C_OrderUnitPopup)
		Call ggoSpread.MakePairsColumn(C_Curr,C_CurrPopup)

		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)	
		
		.ReDraw = True
    End With

    Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
		.ReDraw = False
    
		ggoSpread.Source = frm1.vspdData
		        
		ggoSpread.SpreadLock		-1,			-1
   		ggoSpread.spreadUnlock		C_Cost,		-1,		C_Cost,		-1
		ggoSpread.SSSetRequired		C_Cost,		-1,		-1
		'단가구분 추가 
		ggoSpread.spreadUnlock C_PrcFlg, -1, C_PrcFlg, -1
		ggoSpread.spreadUnlock C_Remark, -1, C_Remark, -1
		ggoSpread.SSSetRequired C_PrcFlg, -1, -1
    
		.ReDraw = True
    End With    
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData
		.ReDraw = False

  		ggoSpread.Source = frm1.vspdData
  
   		ggoSpread.SpreadUnLock		1, pvStartRow, ,pvEndRow
   		
		ggoSpread.SSSetRequired  C_PlantCd,			pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected C_PlantNm,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired  C_ItemCd,			pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected C_ItemNm,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired  C_OrderUnit,		pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired  C_Curr,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired  C_CostDt,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired  C_SupplierCd,		pvStartRow,	pvEndRow	
		ggoSpread.SSSetProtected C_SupplierNm,		pvStartRow,	pvEndRow	
		ggoSpread.SSSetRequired  C_Cost,			pvStartRow,	pvEndRow	
		'단가구분 추가 
		ggoSpread.SSSetRequired  C_PrcFlg,			pvStartRow, pvEndRow			
		ggoSpread.SSSetProtected  C_PrcFlgNm,		pvStartRow, pvEndRow

		If lgPLCd <> "" then 
			ggoSpread.SSSetProtected C_PlantCd,			pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected C_PlantPopup,			pvStartRow,	pvEndRow
		End If			

    
	    .ReDraw = True
    End With
End Sub

'단가구분 추가 
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim strCboCd
	Dim strCboNm
	
	strCboCd = "" & "T" & vbTab & "F"
	ggoSpread.SetCombo strCboCd, C_PrcFlg
	
	strCboNm = "" & "진단가" & vbTab & "가단가"
	ggoSpread.SetCombo strCboNm, C_PrcFlgNm  	  	
End Sub


'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	For intRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = intRow

		frm1.vspdData.Col = C_PrcFlg
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_PrcFlgNm
		frm1.vspdData.value = intindex
	Next
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_PlantCd			=	1
	C_PlantPopup		=	2
	C_PlantNm			=	3
	C_ItemCd			=	4	
	C_ItemPopup			=	5
	C_ItemNm			=	6
	C_OrderUnit			=	7
	C_OrderUnitPopup	=	8
	C_Curr				=	9
	C_CurrPopup			=	10	
	C_CostDt			=	11
	C_SupplierCd		=	12
	C_SupplierPopup		=	13
	C_SupplierNm		=	14
	C_Cost				=	15
	'단가구분 추가 
	C_PrcFlg			=	16
	C_PrcFlgNm			=	17
	C_Remark			=	18
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
			C_PlantCd			=	iCurColumnPos(1)
			C_PlantPopup		=	iCurColumnPos(2)
			C_PlantNm			=	iCurColumnPos(3)
			C_ItemCd			=	iCurColumnPos(4)	
			C_ItemPopup			=	iCurColumnPos(5)
			C_ItemNm			=	iCurColumnPos(6)
			C_OrderUnit			=	iCurColumnPos(7)
			C_OrderUnitPopup	=	iCurColumnPos(8)
			C_Curr				=	iCurColumnPos(9)
			C_CurrPopup			=	iCurColumnPos(10)	
			C_CostDt			=	iCurColumnPos(11)
			C_SupplierCd		=	iCurColumnPos(12)
			C_SupplierPopup		=	iCurColumnPos(13)
			C_SupplierNm		=	iCurColumnPos(14)
			C_Cost				=	iCurColumnPos(15)
			C_PrcFlg			=	iCurColumnPos(16)
			C_PrcFlgNm			=	iCurColumnPos(17)
			C_Remark			=	iCurColumnPos(18)
	End Select    
End Sub

'------------------------------------------  OpenPlant()  ------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
			 
	IsOpenPop = True

	arrParam(0) = "공장" 
	arrParam(1) = "B_Plant"    
	 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	 
	arrParam(4) = ""   
	arrParam(5) = "공장"   
	 
	arrField(0) = "Plant_Cd" 
	arrField(1) = "Plant_NM" 
	    
	arrHeader(0) = "공장"  
	arrHeader(1) = "공장명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)  
		frm1.txtPlantNm.Value= arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If  
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
' Name : OpenItem()
' Description : OpenItem PopUp
'===================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)

	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	= arrRet(0)
		frm1.txtItemNm.Value	= arrRet(1)
		frm1.txtItemCd.focus
	End If
End Function
'------------------------------------------  OpenSupplier()  -------------------------------------------------
' Name : OpenSupplier()
' Description : OpenSupplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSupplier()

	If IsOpenPop = True Or UCase(frm1.txtSuppliercd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "공급처팝업"   
	arrParam(1) = "B_Biz_Partner"  
	 
	arrParam(2) = Trim(frm1.txtSuppliercd.Value) 
	 
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "      
	arrParam(5) = "공급처"       
	 
	arrField(0) = "BP_CD"    
	arrField(1) = "BP_NM"    
	    
	arrHeader(0) = "공급처"   
	arrHeader(1) = "공급처명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSuppliercd.focus
		Exit Function
	Else
		frm1.txtSuppliercd.Value    = arrRet(0)  
		frm1.txtSupplierNm.Value    = arrRet(1)  
		frm1.txtSuppliercd.focus
		Set gActiveElement = document.activeElement
	End If 
End Function

'------------------------------------------  OpenSPlant()  -------------------------------------------------
' Name : OpenSPlant()
' Description : SpreadPlant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSPlant(byval strCon)
	If IsOpenPop = True Then Exit Function

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "공장팝업" 
	arrParam(1) = "B_Plant"    
	arrParam(2) = Trim(strCon)
	arrParam(4) = ""   
	arrParam(5) = "공장"   
	 
	arrField(0) = "Plant_Cd" 
	arrField(1) = "Plant_NM" 
	    
	arrHeader(0) = "공장"  
	arrHeader(1) = "공장명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False

	If arrRet(0) = "" Then 
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_PlantCd
			.text = arrRet(0) 
			.Col = C_PlantNm
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
			.focus
		End With 
	End If 
End Function

'------------------------------------------  OpenSItem()  -------------------------------------------------
' Name : OpenSItem()
' Description : SpreadItem PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSItem(byval strCon)
	If IsOpenPop = True Then Exit Function

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	IsOpenPop = True
	
	With frm1.vspdData
	
		.Row = .ActiveRow 
	
		arrParam(0) = "품목팝업"      
		arrParam(1) = "B_Item_By_Plant,	B_Item"
		arrParam(2) = Trim(strCon)
		arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
		 
		.Col = C_PlantCd
		If Trim(.text) <> "" then
			arrParam(4) = arrParam(4) & "And B_Item_By_Plant.Plant_Cd= " & FilterVar(.text, "''", "S")
		End If 

		arrParam(5) = "품목"       

		arrField(0) = "B_Item_By_Plant.Item_Cd"     
		arrField(1) = "B_Item.Item_NM" 
	
	End With
	
	iCalledAspName = AskPRAspName("m1111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_ItemNm
			.text = arrRet(1)
			.Col = C_ItemCd
			.text = arrRet(0) 
			Call SetFocusToDocument("M") 
			.focus
		End With 
	End If 
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
' Name : OpenUnit()
' Description : SpreadUnit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUnit(byval strCon)  
	If IsOpenPop = True Then Exit Function

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "발주단위"      
	arrParam(1) = "B_Unit_OF_MEASURE"    
	arrParam(2) = Trim(strCon) 
	arrParam(4) = ""      
	arrParam(5) = "발주단위"    
	 
	arrField(0) = "Unit"     
	arrField(1) = "Unit_Nm"     
	    
	arrHeader(0) = "발주단위"   
	arrHeader(1) = "발주단위명"   
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_OrderUnit
			.text = arrRet(0) 
			Call SetFocusToDocument("M") 
			.focus
		End With 
	End If 
End Function

'------------------------------------------  OpenCurr()  -------------------------------------------------
' Name : OpenCurr()
' Description : SpreaCurr PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCurr(byval strCon) 
	If IsOpenPop = True Then Exit Function
 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "화폐"   
	arrParam(1) = "B_Currency"  
	arrParam(2) = Trim(strCon)
	arrParam(4) = ""     
	arrParam(5) = "화폐"    
		 
	arrField(0) = "Currency"   
	arrField(1) = "Currency_Desc"  
		    
	arrHeader(0) = "화폐"   
	arrHeader(1) = "화폐명"   
		    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_Curr
			.text = arrRet(0)
			Call SetFocusToDocument("M") 
			.focus
		End With 
		Call vspdData_Change(C_Curr,frm1.vspdData.ActiveRow)
	End If 
End Function

'------------------------------------------  OpenSSupplier()  -------------------------------------------------
' Name : OpenSSupplier()
' Description : SpreadSupplier PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenSSupplier(byval strCon)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	arrParam(0) = "공급처팝업"   
	arrParam(1) = "B_Biz_Partner"  
	 
	arrParam(2) = Trim(strCon)
	 
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "      
	arrParam(5) = "공급처"       
	 
	arrField(0) = "BP_CD"    
	arrField(1) = "BP_NM"    
	    
	arrHeader(0) = "공급처"   
	arrHeader(1) = "공급처명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_SupplierCd
			.text = arrRet(0) 
			.Col = C_SupplierNm
			.text = arrRet(1) 
			Call SetFocusToDocument("M") 
			.focus
		End With 
	End If 
End Function

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'==========================================================================================================
Sub Form_Load()	'###그리드 컨버전 주의부분###
	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
	Call InitVariables
  Call GetValue_ko441()
	Call SetDefaultVal
	Call InitSpreadSheet                                                        'Setup the Spread sheet1
	'단가구분 추가 
	Call InitComboBox	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)        '  <------변경된 표준 라인 
	
	Select Case Col
		Case C_CURR
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_Cost,"C" ,"X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_Curr,C_Cost,"C" ,"I","X","X")
		Case C_Cost
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_Cost,"C" ,"X","X")		
		Case C_PrcFlg	
			Call InitData()
	End Select
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
 	
 	gMouseClickStatus = "SPC"   
	 	 	
 	Set gActiveSpdSheet = frm1.vspdData
 	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 			
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		Exit Sub
 	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     

    If OldLeft <> NewLeft Then Exit Sub
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '☜: 재쿼리 체크 
		If Trim(lgPageNo) = "" Then Exit Sub
		If lgPageNo > 0 Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End If
End Sub

'========================================================================================
' Function Name : vspdData_ButtonClicked
' Function Desc : 팝업버튼 선택시 
'========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_OrderUnitPopUp Then
			.Col = C_OrderUnit
			.Row = Row
			Call OpenUnit(.text)
		Elseif Row > 0 And Col = C_CurrPopup Then
			.Col = C_Curr
			.Row = Row
			Call OpenCurr(.text)
		Elseif Row > 0 And Col = C_PlantPopup Then
			.Col = C_PlantCd
			.Row = Row
			Call OpenSPlant(.text)
		Elseif Row > 0 And Col = C_ItemPopup Then
			.Col = C_ItemCd
			.Row = Row
			Call OpenSItem(.text)
		Elseif Row > 0 And Col = C_SupplierPopup Then
			.Col = C_SupplierCd
			.Row = Row
			Call OpenSSupplier(.text)		
		End if 
	End With
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
'   Event Name : vspdData_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_Cost
            Call EditModeCheck(frm1.vspdData, Row, C_Curr, C_Cost,    "C" ,"I", Mode, "X", "X")
    End Select
End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    '단가구분 추가 
    Call InitComboBox
    Call InitData()    
    Call ggoSpread.ReOrderingSpreadData
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_Curr,C_Cost,"C" ,"I","X","X")
End Sub 

'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================
Sub txtAppFrDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtAppFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtAppFrDt.Focus
	End If
End Sub

Sub txtAppToDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtAppToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtAppToDt.Focus
	End If
End Sub

Sub txtAppFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtAppToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() '###그리드 컨버전 주의부분###
    Dim IntRetCD     
    FncQuery = False                                                        
	
	If ggoSpread.SSCheckChange = True Then 'lgBlnFlgChgValue = True Or lgBtnClkFlg = True Or 
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")		'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '품목만으로 조회시 메시지 호출(2003.09.04)
    If  Trim(frm1.txtPlantCd.Value) = "" and Trim(frm1.txtItemCd.Value) <> "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData
    Call InitVariables															'Initializes local global variables

   	Call SetToolbar("10000000000111")
   	
    If Check_Input = False Then 
    	Call SetToolBar("111011010011111")				'버튼 툴바 제어 
	    Exit Function
	End If

	If DbQuery = False then	Exit Function
		      
    FncQuery = True	
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    Err.Clear                                                               '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData	   
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    
    FncNew = True  
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    FncSave = False                                                         
    
    If frm1.vspdData.maxrows < 1 then exit function    

	'-----------------------
    'Precheck area
    '-----------------------
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           
        Exit Function
    End If
    
    '----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then Exit Function

   	Call SetToolbar("10000000000111")

    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then Exit Function
	  
    FncSave = True                                                       
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    With frm1.vspdData
		If .maxrows < 1 then exit function
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.ActiveRow,.ActiveRow,C_Curr,C_Cost, "C","I","X","X")
		SetSpreadColor .ActiveRow, .ActiveRow
		.Row = .ActiveRow
		.Col = C_Cost
		.Text = ""
		.Col = C_CostDt
		.Text = ""
		.ReDraw = True
	End With
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	frm1.vspdData.Redraw = False
    If frm1.vspdData.maxrows < 1 then exit function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo 
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_Curr,C_Cost, "C","I","X","X") 
	frm1.vspdData.Redraw = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim iRow
	
	On Error Resume Next
	
	FncInsertRow = False
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End If 
	
    With frm1.vspdData	
		.ReDraw = False
		.focus
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow .ActiveRow, imRow
		SetSpreadColor .ActiveRow, .ActiveRow + imRow - 1


			For iRow = .ActiveRow to .ActiveRow + imRow -1
				If lgPLCd <> "" then 
					call .SetText(C_PlantCd,iRow,lgPLCd)
				End If			
			Next


		.ReDraw = True
    End With

    Set gActiveElement = document.ActiveElement
    
    If Err.number = 0 Then FncInsertRow = True
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows 
	Dim lTempRows 

	If frm1.vspdData.maxrows < 1 then exit function
	
 '----------  Coding part  ------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData	
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 	Call parent.FncExport(Parent.C_MULTI)		
 End Function
 
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
	FncExit = False
	
	Dim IntRetCD
	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If
    
    FncExit = True    
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
	Dim strVal

	DbQuery = False                                                             
	
	Call LayerShowHide(1)
	
	With frm1
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
	    strVal = strVal & "&txtitemcd=" & .hdnitemcd.value
	    strVal = strVal & "&txtSuppliercd=" & .hdnSuppliercd.value
	    strVal = strVal & "&txtAppFrDt=" & .hdnAppFrDt.value
	    strVal = strVal & "&txtAppToDt=" & .hdnAppToDt.value
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value
	    strVal = strVal & "&txtitemcd=" & .txtitemcd.value
	    strVal = strVal & "&txtSuppliercd=" & .txtSuppliercd.value
	    strVal = strVal & "&txtAppFrDt=" & .txtAppFrDt.text
	    strVal = strVal & "&txtAppToDt=" & .txtAppToDt.text
        strVal = strVal & "&lgPageNo="	 & lgPageNo						'☜: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If 
    
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 

	DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbQueryOk()
	Dim ii
	
    lgIntFlgMode = Parent.OPMD_UMODE							'⊙: Indicates that current mode is Update mode
	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("111011110011111")				'버튼 툴바 제어 
	Call SetSpreadLock
	Call InitData()

    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd.focus 
	End if

	if lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
  	frm1.txtPlantCd.value = lgPLCd
	End If
	
	Set gActiveElement = document.activeElement
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow
	Dim lGrpCnt     
	Dim strVal,strDel
	Dim ColSep, RowSep

	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
	Dim ii
	
	ColSep = parent.gColSep               
	RowSep = parent.gRowSep               
	
    DbSave = False                                                          '⊙: Processing is NG

	Call LayerShowHide(1)
	
	frm1.txtMode.value = Parent.UID_M0002
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 0
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferMaxCount = -1 
	iTmpDBufferMaxCount = -1 
	    
	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	'-----------------------
	'Data manipulate area
	'-----------------------
	ggoSpread.source = frm1.vspdData

	With frm1
		For lRow = 1 To .vspdData.MaxRows	'###그리드 컨버전 주의부분###
			Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
				Case ggoSpread.InsertFlag				'☜: 신규 

					strVal = "C" & ColSep	
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_ItemCd,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_SupplierCd,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Curr,lRow,"X","X")) & ColSep
					strVal = strVal & UniConvDate(GetSpreadText(.vspdData,C_CostDt,lRow,"X","X")) & ColSep
					strVal = strVal & UNIConvNum(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"), 0) & ColSep '& lRow & ColSep & RowSep
					'단가구분 추가 
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_PrcFlg,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Remark,lRow,"X","X")) & ColSep & lRow & RowSep

					lGrpCnt = lGrpCnt + 1
										
				Case ggoSpread.UpdateFlag				'☜: 수정 
				
					strVal = "U" & ColSep			'☜: U=Update  0
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_ItemCd,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_SupplierCd,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Curr,lRow,"X","X")) & ColSep
					strVal = strVal & UniConvDate(GetSpreadText(.vspdData,C_CostDt,lRow,"X","X")) & ColSep
					strVal = strVal & UNIConvNum(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"), 0) & ColSep '& lRow & ColSep & RowSep
					'단가구분 추가 
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_PrcFlg,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Remark,lRow,"X","X")) & ColSep & lRow & RowSep

					lGrpCnt = lGrpCnt + 1

				Case ggoSpread.DeleteFlag				'☜: 삭제 
					strDel = "D" & ColSep			
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_ItemCd,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_SupplierCd,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_Curr,lRow,"X","X")) & ColSep
					strDel = strDel & UniConvDate(GetSpreadText(.vspdData,C_CostDt,lRow,"X","X")) & ColSep
					strDel = strDel & UNIConvNum(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"), 0) & ColSep '& lRow & ColSep & RowSep
					'단가구분 추가 
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_PrcFlg,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_Remark,lRow,"X","X")) & ColSep & lRow & RowSep					

					lGrpCnt = lGrpCnt + 1
			End Select

			Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
			       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			End Select   
		Next
	End With

	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			'☜: 비지니스 ASP 를 가동 

	DbSave = True                                                      
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
End Function

'========================================================================================
' Function Name : Check_Input
' Function Desc : 
'========================================================================================
Function Check_Input()
	Check_Input = False
	frm1.txtPlantNm.Value = ""
	frm1.txtItemNm.Value = ""
	frm1.txtSupplierNm.Value = ""

	If Trim(frm1.txtPlantCd.Value) <> "" And Trim(frm1.txtItemcd.Value) <> "" Then

		If 	CommonQueryRs(" B.PLANT_NM, C.ITEM_NM, C.PHANTOM_FLG "," B_ITEM_BY_PLANT A, B_PLANT B, B_ITEM C ", _
		                " A.PLANT_CD = B.PLANT_CD AND A.ITEM_CD = C.ITEM_CD AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S") & " AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("125000","X","X","X")
				frm1.txtPlantNm.Value = ""
				frm1.txtPlantCd.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtPlantNm.Value = lgF0(0)

			If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("122600","X","X","X")
				frm1.txtItemNm.Value = ""
				frm1.txtItemCd.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm.Value = lgF0(0)

			Call DisplayMsgBox("122700","X","X","X")
			frm1.txtItemCd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
		frm1.txtItemNm.Value = lgF1(0)
		
		If Trim(lgF2(0)) <> "N" Then
			Call DisplayMsgBox("181315","X","X","X")
			frm1.txtItemCd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
	
	ElseIf Trim(frm1.txtPlantCd.Value) <> "" Then
		'-----------------------
		'Check Plant CODE		'공장코드가 있는 지 체크 
		'-----------------------
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	
	ElseIf Trim(frm1.txtItemcd.Value) <> "" Then
		'-----------------------
		'Check Item CODE	 '품목코드가 있는 지 체크  
		'-----------------------
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("122600","X","X","X")
			frm1.txtItemNm.Value = ""
			frm1.txtItemCd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtItemNm.Value = lgF0(0)
	End If

	If Trim(frm1.txtSuppliercd.Value) <> "" Then
		'-----------------------
		'Check BPt CODE		'공급처코드가 있는 지 체크 
		'-----------------------
		If 	CommonQueryRs(" BP_NM, BP_TYPE, usage_flag "," B_Biz_Partner ", " BP_CD = " & FilterVar(frm1.txtSuppliercd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("229927","X","X","X")
			frm1.txtSupplierNm.Value = ""
			frm1.txtSuppliercd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		frm1.txtSupplierNm.Value = lgF0(0)

		If Trim(lgF2(0)) <> "Y" Then
			Call DisplayMsgBox("179021","X","X","X")
			frm1.txtSuppliercd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		If Trim(lgF1(0)) <> "S" And Trim(lgF1(0)) <> "CS" Then
			Call DisplayMsgBox("179020","X","X","X")
			frm1.txtSuppliercd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
	End If
	
    If frm1.txtAppFrDt.text <> "" And frm1.txtAppToDt.text <> "" Then
		If ValidDateCheck(frm1.txtAppFrDt, frm1.txtAppToDt) = False Then 
   			frm1.txtAppToDt.focus 
			Set gActiveElement = document.activeElement
			Exit Function
		End If
	End If
	
	Check_Input = True
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow)
	With frm1.vspdData
		.Row = lRow
		.Col = C_ItemCd
		.Action = 0
		Call SetFocusToDocument("M") 
		.focus
	End With
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
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
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공급처별단가(Multi)</font></td>
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
					<FIELDSET CLASS="CLSFLD" >
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>공장</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장"  tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 ALT="공장" tag="14X">
							<TD CLASS="TD5" NOWRAP>품목</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 STYLE="text-transform:uppercase" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
												   <INPUT TYPE=TEXT ALT="품목" NAME="txtItemNm" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>공급처</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSuppliercd" SIZE=10 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSupplier()">&nbsp;<INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="14X"></TD>
							<TD CLASS="TD5" NOWRAP>단가적용일</TD>
							<TD CLASS="TD6" NOWRAP>
								<table cellpadding=0 cellspacing=0>
									<tr>
										<td>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="시작일" NAME="txtAppFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11N" Title="FPDATETIME"></OBJECT>');</SCRIPT>
										</td>
										<td>~</td>
										<td>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="종료일" NAME="txtAppToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N" Title="FPDATETIME"></OBJECT>');</SCRIPT>
										</td>
									</tr>
								</table>
							</TD>
						</TR>
					</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1 TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> ></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1">
			</IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnitemcd"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnSuppliercd"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnAppFrDt"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnAppToDt"  tag="24" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>
 
