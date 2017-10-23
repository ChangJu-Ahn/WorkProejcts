<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Long-term Inventory Analysis
'*  2. Function Name        : 
'*  3. Program ID           : I3111QA1
'*  4. Program Name         : 장기재고현황 
'*  5. Program Desc         :
'*  7. Modified date(First) : 2006/05/25
'*  8. Modified date(Last)  : 2006/05/25
'*  9. Modifier (First)     : KiHong Han
'* 10. Modifier (Last)      : KiHong Han
'* 11. Comment
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

<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->


Const BIZ_PGM_ID = "i3111qb1_ko441.asp"                 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP1_ID = "i3112ma1"
Const BIZ_PGM_JUMP2_ID = "i3112qa1"

Dim C1_ItemCd
Dim C1_ItemNm
Dim C1_Spec
Dim C1_ABCFlag
Dim C1_Unit
Dim C1_InvPrice
Dim C1_StoragePeriod
Dim C1_LastIssueDt
Dim C1_PerniciousStockPeriod
Dim C1_PerniciousStockQty
Dim C1_PerniciousStockAmt
Dim C1_LongtermStockPeriod
Dim C1_LongtermStockQty
Dim C1_LongtermStockAmt

Dim C2_ItemGroupCd
Dim C2_ItemGroupNm
Dim C2_PerniciousStockQty
Dim C2_PerniciousStockAmt
Dim C2_LongtermStockQty
Dim C2_LongtermStockAmt

Dim C3_ABCFlag
Dim C3_PerniciousStockQty
Dim C3_PerniciousStockAmt
Dim C3_LongtermStockQty
Dim C3_LongtermStockAmt

Dim IsOpenPop 

Dim lgStrPrevKey2
Dim lgStrPrevKey3

Dim lgSortKey2
Dim lgSortKey3

'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYM
CompanyYM = UNIMonthClientFormat(UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gAPDateFormat))
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE
	IsOpenPop = False
	lgStrPrevKey = ""
	lgStrPrevKey2 = ""
	lgStrPrevKey3 = ""
End Sub                          

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If ReadCookie("txtPlantCd") = "" Then
		If Parent.gPlant <> "" Then
			frm1.txtPlantCd.value = Ucase(Parent.gPlant)
			frm1.txtPlantNm.value = Parent.gPlantNm
		End If
    Else
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If	
	
	If ReadCookie("txtYYYYMM") = "" Then
		frm1.txtYYYYMM.Text	= CompanyYM
	Else
		frm1.txtYYYYMM.Text = ReadCookie("txtYYYYMM")
	End If	

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
  	frm1.txtPlantCd.value = lgPLCd
	End If
	 
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtYYYYMM", ""
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I", "NOCOOKIE","QA") %>
End Sub


'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	
	Call InitSpreadPosVariables(pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData 
			
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20050202", ,Parent.gAllowDragDropSpread
			
			.ReDraw = False
					
			.MaxCols = C1_LongtermStockAmt + 1											'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit C1_ItemCd, "품목", 18
			ggoSpread.SSSetEdit C1_ItemNm, "품목명", 25
			ggoSpread.SSSetEdit C1_Spec, "규격", 25
			ggoSpread.SSSetEdit C1_ABCFlag, "ABC구분", 10, 2
			ggoSpread.SSSetEdit	C1_Unit, "단위", 10, 2
			ggoSpread.SSSetFloat C1_InvPrice, "재고단가", 15, Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C1_StoragePeriod,"재고기간", 16, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetDate C1_LastIssueDt, "최종출고일자", 10, 2, parent.gDateFormat
			ggoSpread.SSSetFloat C1_PerniciousStockPeriod, "악성재고기준기간", 16, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C1_PerniciousStockQty, "악성재고수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C1_PerniciousStockAmt, "악성재고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C1_LongtermStockPeriod, "장기재고기준기간", 16, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C1_LongtermStockQty, "장기재고수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C1_LongtermStockAmt, "장기재고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			ggoSpread.SpreadLockWithOddEvenRowColor()
			
			.ReDraw = true
    
		End With
    End If

	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
				    
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20050202", ,Parent.gAllowDragDropSpread

			.ReDraw = false
			
			.MaxCols = C2_LongtermStockAmt + 1										'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0

			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit C2_ItemGroupCd, "품목그룹", 18
			ggoSpread.SSSetEdit C2_ItemGroupNm, "품목그룹명", 25
			ggoSpread.SSSetFloat C2_PerniciousStockQty, "악성재고수량", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C2_PerniciousStockAmt, "악성재고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C2_LongtermStockQty, "장기보관재고수량", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C2_LongtermStockAmt, "장기보관재고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

			ggoSpread.SSSetSplit2(2)

			ggoSpread.SpreadLockWithOddEvenRowColor()
			
			.ReDraw = true
    
		End With
	End If	
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		With frm1.vspdData3
		
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit "V20050202", ,Parent.gAllowDragDropSpread
		
			.ReDraw = false
				
			.MaxCols = C3_LongtermStockAmt + 1										'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
		
			Call GetSpreadColumnPos("C")
			
			ggoSpread.SSSetEdit C3_ABCFlag, "ABC구분", 10, 2
			ggoSpread.SSSetFloat C3_PerniciousStockQty, "악성재고수량", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C3_PerniciousStockAmt, "악성재고금액", 15,  Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C3_LongtermStockQty, "장기보관재고수량", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C3_LongtermStockAmt, "장기보관재고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(1)
			
			ggoSpread.SpreadLockWithOddEvenRowColor()
			
			.ReDraw = true

		End With
	End If	
	
End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData) - Order Header
		C1_ItemCd					= 1
		C1_ItemNm					= 2
		C1_Spec						= 3
		C1_ABCFlag					= 4
		C1_Unit						= 5
		C1_InvPrice					= 6
		C1_StoragePeriod			= 7
		C1_LastIssueDt				= 8
		C1_PerniciousStockPeriod	= 9
		C1_PerniciousStockQty		= 10
		C1_PerniciousStockAmt		= 11
		C1_LongtermStockPeriod		= 12
		C1_LongtermStockQty			= 13
		C1_LongtermStockAmt			= 14
	 End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Results
		C2_ItemGroupCd			= 1
		C2_ItemGroupNm			= 2
		C2_PerniciousStockQty	= 3
		C2_PerniciousStockAmt	= 4
		C2_LongtermStockQty		= 5
		C2_LongtermStockAmt		= 6
	End If
		
	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		' Grid 3(vspdData3) - Hidden
		C3_ABCFlag				= 1
		C3_PerniciousStockQty	= 2
		C3_PerniciousStockAmt	= 3
		C3_LongtermStockQty		= 4
		C3_LongtermStockAmt		= 5
	End If	
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==========
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'=================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
 			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			  
			' Grid 1(vspdData) - Order Header
			C1_ItemCd					= iCurColumnPos(1)
			C1_ItemNm					= iCurColumnPos(2)
			C1_Spec						= iCurColumnPos(3)
			C1_ABCFlag					= iCurColumnPos(4)
			C1_Unit						= iCurColumnPos(5)
			C1_InvPrice					= iCurColumnPos(6)
			C1_StoragePeriod			= iCurColumnPos(7)
			C1_LastIssueDt				= iCurColumnPos(8)
			C1_PerniciousStockPeriod	= iCurColumnPos(9)
			C1_PerniciousStockQty		= iCurColumnPos(10)
			C1_PerniciousStockAmt		= iCurColumnPos(11)
			C1_LongtermStockPeriod		= iCurColumnPos(12)
			C1_LongtermStockQty			= iCurColumnPos(13)
			C1_LongtermStockAmt			= iCurColumnPos(14)

		Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			' Grid 2(vspdData2) - Results
			C2_ItemGroupCd			= iCurColumnPos(1)
			C2_ItemGroupNm			= iCurColumnPos(2)
			C2_PerniciousStockQty	= iCurColumnPos(3)
			C2_PerniciousStockAmt	= iCurColumnPos(4)
			C2_LongtermStockQty		= iCurColumnPos(5)
			C2_LongtermStockAmt		= iCurColumnPos(6)
			
		Case "C"
			ggoSpread.Source = frm1.vspdData3
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			' Grid 3(vspdData3) - Results
			C3_ABCFlag				= iCurColumnPos(1)
			C3_PerniciousStockQty	= iCurColumnPos(2)
			C3_PerniciousStockAmt	= iCurColumnPos(3)
			C3_LongtermStockQty		= iCurColumnPos(4)
			C3_LongtermStockAmt		= iCurColumnPos(5)
			
    End Select    
End Sub    

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtPlantCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_Plant"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Condition
	arrParam(4) = ""
	arrParam(5) = "공장"							' TextBox 명칭 

    arrField(0) = "Plant_Cd"					' Field명(0)
    arrField(1) = "Plant_NM"					' Field명(1)
        
    arrHeader(0) = "공장"						' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtPlantCd.Focus
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = Trim(arrRet(0))
		frm1.txtPlantNm.Value = Trim(arrRet(1))
	End If	
	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenQueryTarget()  -------------------------------------------------
'	Name : OpenQueryTarget()
'	Description : OpenQueryTarget PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenQueryTarget()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strQueryTargetClass

	If frm1.rdoQueryTargetClass.rdoQueryTargetClass1.Checked = True Then
		strQueryTargetClass = "1"
	ElseIf frm1.rdoQueryTargetClass.rdoQueryTargetClass2.Checked = True Then
		strQueryTargetClass = "2"
	Else
		strQueryTargetClass = "3"
	End If
	
	Select Case strQueryTargetClass
		Case "1"
			'공장코드가 있는 지 체크 
			If Trim(frm1.txtPlantCd.Value) = "" then 
				Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
				Exit Function
			End If
	
			arrParam(0) = "품목"													' 팝업 명칭 
			arrParam(1) = "B_Item_By_Plant,B_Item"									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtQueryTargetCd.Value)							' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd"
			arrParam(4) = arrParam(4) & "  And B_Item_By_Plant.Plant_Cd = '" & FilterVar(Trim(frm1.txtPlantCd.Value),"''","SNM") & "'" 			' Where Condition
			arrParam(5) = "품목"													' TextBox 명칭 
	
			arrField(0) = "B_Item_By_Plant.Item_Cd"		' Field명(0)
			arrField(1) = "B_Item.Item_NM"				' Field명(1)
			arrField(2) = "B_Item.SPEC"					' Field명(2)
				
			arrHeader(0) = "품목"					' Header명(0)
			arrHeader(1) = "품목명"						' Header명(1)
			arrHeader(2) = "규격"						' Header명(2)
		Case "2"
			arrParam(0) = "품목그룹팝업"	
			arrParam(1) = "B_ITEM_GROUP"				
			arrParam(2) = Trim(frm1.txtQueryTargetCd.Value)
			arrParam(3) = ""
			arrParam(4) = "DEL_FLG = 'N' " 			
			arrParam(5) = "품목그룹"			
	
			arrField(0) = "ITEM_GROUP_CD"	
			arrField(1) = "ITEM_GROUP_NM"	
    
			arrHeader(0) = "품목그룹"		
			arrHeader(1) = "품목그룹명"
		Case "3"
			Exit Function
		
	End Select
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtQueryTargetCd.Focus
	If arrRet(0) <> "" Then
		frm1.txtQueryTargetCd.Value = Trim(arrRet(0))
		frm1.txtQueryTargetNm.Value = Trim(arrRet(1))
	End If	
	
	Set gActiveElement = document.activeElement
End Function

'=============================================  2.5.2 JumpToLongtermInvAnal()  ======================================
'=	Event Name : JumpToLongtermInvAnal
'=	Event Desc : 장기재고현황으로 Jump
'========================================================================================================
Function JumpToLongtermInvAnal()
	With frm1
		'공장코드/명/분석연월 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtYYYYMM", .txtYYYYMM.Text
	End With
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 JumpToLongtermInvChange()  ======================================
'=	Event Name : JumpToLongtermInvChange
'=	Event Desc : 장기재고추이로 Jump
'========================================================================================================
Function JumpToLongtermInvChange()
	With frm1
		'공장코드/명/분석연월 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtYYYYMM", .txtYYYYMM.Text
	End With
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
	
	Call InitVariables	
  Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet("*")
	 
	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어	
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.focus 
    Else
		frm1.txtYYYYMM.focus
	End If
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode ) 
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		
 			lgSortKey = 1
 		End If
 	End If
End Sub

'======================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData2
	
	If frm1.vspdData2.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2 
 		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey2 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey2	
 			lgSortKey2 = 1
 		End If
 	End If
End Sub

'======================================================================================================
'   Event Name : vspdData3_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData3_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData3
	
	If frm1.vspdData3.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData3 
 		If lgSortKey3 = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey3 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey3		
 			lgSortKey3 = 1
 		End If
 	End If
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

'========================================================================================
' Function Name : vspdData2_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
    End If
End Sub 

'========================================================================================
' Function Name : vspdData3_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData3_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
    End If
End Sub 

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData2_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'==========================================================================================
'   Event Name : vspdData3_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData3_GotFocus()
    ggoSpread.Source = frm1.vspdData3
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
' Function Name : vspdData2_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData3_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData3
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

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================
' Function Name : vspdData3_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData3_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("C")
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
Sub PopRestoreSpreadColumnInf()	'###그리드 컨버전 주의부분###
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)	
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	 
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If    
End Sub

'======================================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	 
	'----------  Coding part  -----------------------------
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then
		If lgStrPrevKey2 <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If    
End Sub

'======================================================================================================
'   Event Name : vspdData3_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	 
	'----------  Coding part  -----------------------------
	If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3, NewTop) Then
		If lgStrPrevKey3 <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If    
End Sub


'==========================================================================================
'   Event Name : txtYYYYMM
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtYYYYMM_DblClick(Button)
	If Button = 1 Then
		frm1.txtYYYYMM.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtYYYYMM.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtYYYYMM
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtYYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

'==========================================================================================
'   Event Name : txtPlantCd_onChange
'   Event Desc : 
'==========================================================================================
Function  txtPlantCd_onChange()
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtQueryTargetCd_onChange
'   Event Desc : 
'==========================================================================================
Function  txtQueryTargetCd_onChange()
	If Trim(frm1.txtQueryTargetCd.Value) = "" Then
		frm1.txtQueryTargetNm.Value = ""
	End If
End Function


'==========================================================================================
'   Event Name : rdoQueryTargetClass1_onClick
'   Event Desc : 
'==========================================================================================
Function  rdoQueryTargetClass1_onClick()
	Call ggoOper.SetReqAttr(frm1.txtQueryTargetCd, "D")
End Function

'==========================================================================================
'   Event Name : rdoQueryTargetClass2_onClick
'   Event Desc : 
'==========================================================================================
Function  rdoQueryTargetClass2_onClick()
	Call ggoOper.SetReqAttr(frm1.txtQueryTargetCd, "D")
End Function

'==========================================================================================
'   Event Name : rdoQueryTargetClass3_onClick
'   Event Desc : 
'==========================================================================================
Function  rdoQueryTargetClass3_onClick()
	Call ggoOper.SetReqAttr(frm1.txtQueryTargetCd, "Q")
End Function

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function								'⊙: This function check indispensable field
    
    '-----------------------
    '조회대상분류에 따른 조치 
    '-----------------------
    'Erase contents area
    ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData
	ggoSpread.source = frm1.vspddata2
	ggoSpread.ClearSpreadData 
	ggoSpread.source = frm1.vspddata3
	ggoSpread.ClearSpreadData
	
' Clear & Change Contents Area 
	'frm1.txtLongtermStockCalPeriod.value = ""
	'frm1.txtPerniciousStockCalPeriod.value = ""
	
	If frm1.rdoQueryTargetClass.rdoQueryTargetClass1.Checked = True Then
		QUERYTARGETCLASS2.style.display = "none"
		QUERYTARGETCLASS3.style.display = "none"
		QUERYTARGETCLASS1.style.display = ""
		
	ElseIf frm1.rdoQueryTargetClass.rdoQueryTargetClass2.Checked = True Then
		QUERYTARGETCLASS1.style.display = "none"
		QUERYTARGETCLASS3.style.display = "none"
		QUERYTARGETCLASS2.style.display = ""

	Else
		QUERYTARGETCLASS1.style.display = "none"
		QUERYTARGETCLASS2.style.display = "none"
		QUERYTARGETCLASS3.style.display = ""
		
	End If
	
    Call InitVariables
    
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then Exit Function

    FncQuery = True															'⊙: Processing is OK
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
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
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
	DbQuery = False
	
	Dim strVal
	Dim strQueryTargetClass
    
	Call LayerShowHide(1)
	
	Dim strYear, strMonth, strDay, strYyMm

	Call ExtractDateFrom(frm1.txtYYYYMM.Text,frm1.txtYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	strYyMm = strYear + strMonth
	
    With frm1
    
		If lgIntFlgMode = parent.OPMD_CMODE Then
			If frm1.rdoQueryTargetClass.rdoQueryTargetClass1.Checked = True Then
				strQueryTargetClass = "1"
			ElseIf frm1.rdoQueryTargetClass.rdoQueryTargetClass2.Checked = True Then
				strQueryTargetClass = "2"
			Else
				strQueryTargetClass = "3"
			End If
										
			strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value) _
								& "&txtYr=" & Left(.txtYYYYMM.DateValue,4) _
								& "&txtMnth=" & Mid(.txtYYYYMM.DateValue,5, 2) _
								& "&txtYYYYMM=" & strYyMm _
								& "&txtQueryTargetClass=" & strQueryTargetClass _
								& "&txtQueryTargetCd=" & Trim(.txtQueryTargetCd.value)
		
		Else
			strQueryTargetClass = .hQueryTargetClass.value
			
			strVal = BIZ_PGM_ID & "?txtPlantCd=" & .hPlantCd.value _
								& "&txtYr=" & .hYr.value _
								& "&txtMnth=" & .hMnth.value _
								& "&txtYYYYMM=" & strYyMm _
								& "&txtQueryTargetClass=" & strQueryTargetClass _
								& "&txtQueryTargetCd=" & .hQueryTargetCd.value
		End If
							
		Select Case strQueryTargetClass
			Case "1"
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows _
								& "&lgStrPrevKey=" & lgStrPrevKey 
			Case "2"
				strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows _
								& "&lgStrPrevKey=" & lgStrPrevKey2 
			Case "3"
				strVal = strVal & "&txtMaxRows=" & .vspdData3.MaxRows _
								& "&lgStrPrevKey=" & lgStrPrevKey3
		
		End Select
		
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
    
    Call SetToolbar("11000000000111")
	Set gActiveElement = document.activeElement
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>장기재고현황</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<!--
					<TD WIDTH="*" align=right><button name="btnAutoSel" class="clsmbtn" ONCLICK="PopZAdoConfigGrid()">정렬순서</button></TD>
					-->
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
        							<TD CLASS="TD6" NOWRAP>
        								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14"></TD>								
        							<TD CLASS="TD5" NOWRAP>기간</TD>
									<TD CLASS="TD6">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtYYYYMM name=txtYYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="기간(FROM)" tag="12"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>조회대상분류</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryTargetClass" TAG="1X" ID="rdoQueryTargetClass1" CHECKED><LABEL FOR="rdoQueryTargetClass1">품목</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryTargetClass" TAG="1X" ID="rdoQueryTargetClass2"><LABEL FOR="rdoQueryTargetClass2">품목그룹</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryTargetClass" TAG="1X" ID="rdoQueryTargetClass3"><LABEL FOR="rdoQueryTargetClass3">ABC구분</LABEL>
									</TD>
									<TD CLASS="TD5" NOWRAP>조회대상</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtQueryTargetCd" SIZE=15 MAXLENGTH=18 ALT="조회대상" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnQueryTarget align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenQueryTarget()">
										<INPUT TYPE=TEXT NAME="txtQueryTargetNm" SIZE=20 MAXLENGTH=20 tag="14">
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
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_60%>>
							
							<TR ID=QUERYTARGETCLASS1>
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> id="A" NAME=vspdData WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD"> <PARAM NAME="MAXCOLs" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>	
							</TR>
							<TR ID=QUERYTARGETCLASS2 Style="display:none;">
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> id="B" NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD"> <PARAM NAME="MAXCOLs" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>	
							</TR>
							<TR ID=QUERYTARGETCLASS3 Style="display:none;">
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> id="C" NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD"> <PARAM NAME="MAXCOLs" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpToLongtermInvAnal">장기재고분석</A>&nbsp;|&nbsp;<A href="vbscript:JumpToLongtermInvChange">장기재고추이</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hPlantCd" TAG="14" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hYr" TAG="14" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hMnth" TAG="14" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hYyyyMm" TAG="14" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hQueryTargetClass" TAG="14" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hQueryTargetCd" TAG="14" tabindex=-1>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
