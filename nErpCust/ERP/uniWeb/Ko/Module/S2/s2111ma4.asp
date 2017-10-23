<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2111MA4
'*  4. Program Name         : 조직별 고객판매계획등록 
'*  5. Program Desc         : 조직별 고객판매계획등록 
'*  6. Comproxy List        : PS2G121.dll, PS2G122.dll, PS2G124.dll
'*  7. Modified date(First) : 2000/03/24
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Mr Cho 
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : -2000/03/24 : 3rd 기능구현 및 화면디자인 
'*                            -2000/05/09 : 3rd 표준수정사항 
'*                            -2000/08/10 : 4th 화면 Layout 수정 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

Dim C_ItemCode ' 1           '☆: Spread Sheet의 Column별 변수 
Dim C_ItemPopup ' 2
Dim C_ItemName ' 3
Dim C_PlanUnit ' 4
Dim C_PlanUnitPopup ' 5
Dim C_YearQty ' 6
Dim C_YearAmt ' 7

Dim C_01PlanQty ' 8
Dim C_02PlanQty ' 10
Dim C_03PlanQty ' 12
Dim C_04PlanQty ' 14
Dim C_05PlanQty ' 16
Dim C_06PlanQty ' 18
Dim C_07PlanQty ' 20
Dim C_08PlanQty ' 22
Dim C_09PlanQty ' 24
Dim C_10PlanQty ' 26
Dim C_11PlanQty ' 28
Dim C_12PlanQty ' 30

Dim C_01PlanAmt ' 9
Dim C_02PlanAmt ' 11
Dim C_03PlanAmt ' 13
Dim C_04PlanAmt ' 15
Dim C_05PlanAmt ' 17
Dim C_06PlanAmt ' 19
Dim C_07PlanAmt ' 21
Dim C_08PlanAmt ' 23
Dim C_09PlanAmt ' 25
Dim C_10PlanAmt ' 27
Dim C_11PlanAmt ' 29
Dim C_12PlanAmt ' 31

Dim C_01PlanColor ' 32
Dim C_02PlanColor ' 33
Dim C_03PlanColor ' 34
Dim C_04PlanColor ' 35
Dim C_05PlanColor ' 36
Dim C_06PlanColor ' 37
Dim C_07PlanColor ' 38
Dim C_08PlanColor ' 39
Dim C_09PlanColor ' 40
Dim C_10PlanColor ' 41
Dim C_11PlanColor ' 42
Dim C_12PlanColor ' 43


<!-- #Include file="../../inc/lgvariables.inc" -->
	
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s2111mb4.asp"        '☆: Head Query 비지니스 로직 ASP명 


Dim IsOpenPop       'Popup

Dim lsItemCode       '품목 
Dim lsItemName       '품목명 
Dim lsPlanMonth      '계획월 
Dim lsPlanUnit       '계힉단위 
Dim lsPlanQtyAmt     '계획수량/금액 

Const lsConfirm  = "CONFIRM"	<% '확정처리 %>
Const lsQtyAmt  = "QtyAmt"		<% '수량/금액 자동계산 %>

Const lsInsert = "INSERT"		<% 'Spread Color 지정시 신규입력시 %>
Const lsQuery = "QUERY"			<% 'Spread Color 지정시 조회후 %>
Const lsSelect = "SELECT"		<% 'Spread Color 지정시 라디오버튼 클릭시 %>

Const lsSalesPlanBy  = "C"		<% 'Reference 참조시 - 거래처별 판매계획 %>
Const lsSelectChr = "C"			<% '계획차수 LookUp시 Pad SelectChr - 거래처별 판매계획 %>

'========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCode = 1
    C_ItemPopup = 2
    C_ItemName = 3
    C_PlanUnit = 4
    C_PlanUnitPopup = 5
    C_YearQty = 6
    C_YearAmt = 7

    C_01PlanQty = 8
    C_02PlanQty = 10
    C_03PlanQty = 12
    C_04PlanQty = 14
    C_05PlanQty = 16
    C_06PlanQty = 18
    C_07PlanQty = 20
    C_08PlanQty = 22
    C_09PlanQty = 24
    C_10PlanQty = 26
    C_11PlanQty = 28
    C_12PlanQty = 30

    C_01PlanAmt = 9
    C_02PlanAmt = 11
    C_03PlanAmt = 13
    C_04PlanAmt = 15
    C_05PlanAmt = 17
    C_06PlanAmt = 19
    C_07PlanAmt = 21
    C_08PlanAmt = 23
    C_09PlanAmt = 25
    C_10PlanAmt = 27
    C_11PlanAmt = 29
    C_12PlanAmt = 31

    C_01PlanColor = 32
    C_02PlanColor = 33
    C_03PlanColor = 34
    C_04PlanColor = 35
    C_05PlanColor = 36
    C_06PlanColor = 37
    C_07PlanColor = 38
    C_08PlanColor = 39
    C_09PlanColor = 40
    C_10PlanColor = 41
    C_11PlanColor = 42
    C_12PlanColor = 43
End Sub

'========================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0      
    lgSortKey = 1

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
 Call ggoOper.LockField(Document, "Q") '계획차수관련 
 Call SetToolBar("11000000000011")     
 With frm1
  .txtConSalesOrg.focus
  .txtMode.value = ""
  .btnSplit.disabled = True 
  .txtSelectChr.value = lsSelectChr
  lgBlnFlgChgValue = False

  .txtConCurr.value = Parent.gCurrency  
  .txtCurr.value = Parent.gCurrency

  .txtConSpYear.value = Year(UniConvDateToYYYYMMDD(EndDate,Parent.gDateFormat,Parent.gServerDateType))
  .txtSpYear.value = Year(UniConvDateToYYYYMMDD(EndDate,Parent.gDateFormat,Parent.gServerDateType))

 End With
End Sub

'========================================================================================================= 
Sub SetDefaultVal2()
	Call ggoOper.ClearField(Document, "2")
    Call InitVariables
	Call SetToolBar("11101111001011")    
	frm1.vspdData.MaxRows = 0

	With frm1
		.txtConSalesOrg.focus
		.txtMode.value = ""
		.btnSplit.disabled = True 
		.txtSelectChr.value = lsSelectChr
		lgBlnFlgChgValue = False

		.txtSalesOrg.value = .txtConSalesOrg.value
		.txtSalesOrgNm.value = .txtConSalesOrgNm.value

		.txtSpYear.value = .txtConSpYear.value
		
		.txtPlanTypeCd.value = .txtConPlanTypeCd.value
		.txtPlanTypeNm.value = .txtConPlanTypeNm.value

		.txtDealTypeCd.value = .txtConDealTypeCd.value
		.txtDealTypeNm.value = .txtConDealTypeNm.value
				
		.txtPlanNum.value = .txtConPlanNum.value
		.txtPlanNumNm.value = .txtConPlanNumNm.value
				
		.txtCurr.value = .txtConCurr.value
 End With

End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData

	ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021120",,parent.gAllowDragDropSpread    		

	.ReDraw = False 
 
    .MaxCols = C_12PlanColor+1             '☜: 최대 Columns의 항상 1개 증가시킴 
	.MaxRows = 0
	  
    Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit C_ItemCode, "고객", 20,,,10,2
	ggoSpread.SSSetButton C_ItemPopup
	ggoSpread.SSSetEdit C_ItemName, "고객명", 30
	ggoSpread.SSSetEdit C_PlanUnit, "계획단위", 10,,,3,2
	ggoSpread.SSSetButton C_PlanUnitPopup

	ggoSpread.SSSetFloat C_YearQty,"년 계획량 합계" ,20,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_YearAmt,"년 계획금액 합계",20,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_01PlanQty,"1월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_02PlanQty,"2월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_03PlanQty,"3월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_04PlanQty,"4월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_05PlanQty,"5월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_06PlanQty,"6월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_07PlanQty,"7월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_08PlanQty,"8월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_09PlanQty,"9월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_10PlanQty,"10월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_11PlanQty,"11월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_12PlanQty,"12월계획량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"        
	ggoSpread.SSSetFloat C_01PlanAmt,"1월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_02PlanAmt,"2월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_03PlanAmt,"3월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
	ggoSpread.SSSetFloat C_04PlanAmt,"4월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_05PlanAmt,"5월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_06PlanAmt,"6월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_07PlanAmt,"7월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_08PlanAmt,"8월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
	ggoSpread.SSSetFloat C_09PlanAmt,"9월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_10PlanAmt,"10월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_11PlanAmt,"11월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_12PlanAmt,"12월계획금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
     
    ggoSpread.SSSetEdit C_01PlanColor, "1월배분",3
    ggoSpread.SSSetEdit C_02PlanColor, "2월배분",3
    ggoSpread.SSSetEdit C_03PlanColor, "3월배분",3
    ggoSpread.SSSetEdit C_04PlanColor, "4월배분",3
    ggoSpread.SSSetEdit C_05PlanColor, "5월배분",3
    ggoSpread.SSSetEdit C_06PlanColor, "6월배분",3
    ggoSpread.SSSetEdit C_07PlanColor, "7월배분",3
    ggoSpread.SSSetEdit C_08PlanColor, "8월배분",3
    ggoSpread.SSSetEdit C_09PlanColor, "9월배분",3
    ggoSpread.SSSetEdit C_10PlanColor, "10월배분",3
    ggoSpread.SSSetEdit C_11PlanColor, "11월배분",3
    ggoSpread.SSSetEdit C_12PlanColor, "12월배분",3
 
	Call ggoSpread.MakePairsColumn(C_ItemCode,C_ItemPopup)
	Call ggoSpread.MakePairsColumn(C_PlanUnit,C_PlanUnitPopup)

	Call ggoSpread.SSSetColHidden(C_PlanUnit,C_PlanUnit,True)	
    Call ggoSpread.SSSetColHidden(C_PlanUnitPopup,C_PlanUnitPopup,True)		
		
	Call ggoSpread.SSSetColHidden(C_01PlanColor,C_01PlanColor,True)	
    Call ggoSpread.SSSetColHidden(C_02PlanColor,C_02PlanColor,True)		
	Call ggoSpread.SSSetColHidden(C_03PlanColor,C_03PlanColor,True)	
	Call ggoSpread.SSSetColHidden(C_04PlanColor,C_04PlanColor,True)	
    Call ggoSpread.SSSetColHidden(C_05PlanColor,C_05PlanColor,True)		
	Call ggoSpread.SSSetColHidden(C_06PlanColor,C_06PlanColor,True)	
    Call ggoSpread.SSSetColHidden(C_07PlanColor,C_07PlanColor,True)		
	Call ggoSpread.SSSetColHidden(C_08PlanColor,C_08PlanColor,True)	
    Call ggoSpread.SSSetColHidden(C_09PlanColor,C_09PlanColor,True)		
	Call ggoSpread.SSSetColHidden(C_10PlanColor,C_10PlanColor,True)	
    Call ggoSpread.SSSetColHidden(C_11PlanColor,C_11PlanColor,True)		
	Call ggoSpread.SSSetColHidden(C_12PlanColor,C_12PlanColor,True)	

	Call ggoSpread.SSSetColHidden(C_YearQty,C_YearQty,True)
	
	Call ggoSpread.SSSetColHidden(C_01PlanQty,C_01PlanQty,True)			
    Call ggoSpread.SSSetColHidden(C_02PlanQty,C_02PlanQty,True)		
	Call ggoSpread.SSSetColHidden(C_03PlanQty,C_03PlanQty,True)	
	Call ggoSpread.SSSetColHidden(C_04PlanQty,C_04PlanQty,True)	
    Call ggoSpread.SSSetColHidden(C_05PlanQty,C_05PlanQty,True)		
	Call ggoSpread.SSSetColHidden(C_06PlanQty,C_06PlanQty,True)	
    Call ggoSpread.SSSetColHidden(C_07PlanQty,C_07PlanQty,True)		
	Call ggoSpread.SSSetColHidden(C_08PlanQty,C_08PlanQty,True)	
    Call ggoSpread.SSSetColHidden(C_09PlanQty,C_09PlanQty,True)		
	Call ggoSpread.SSSetColHidden(C_10PlanQty,C_10PlanQty,True)	
    Call ggoSpread.SSSetColHidden(C_11PlanQty,C_11PlanQty,True)		
	Call ggoSpread.SSSetColHidden(C_12PlanQty,C_12PlanQty,True)	

    Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
 
	.ReDraw = True
    
    End With
    
End Sub

'========================================================================================================= 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow,ByVal IGubun)
       
	With frm1

	ggoSpread.Source = .vspdData
	If IGubun = False Then	
		ggoSpread.SSSetRequired C_ItemCode, pvStartRow, pvEndRow
		Call SpreadProtectUnLock(C_01PlanQty,1)
	Else
		ggoSpread.SSSetProtected C_ItemCode, pvStartRow, pvEndRow
	End If
				    
	ggoSpread.SSSetProtected C_ItemName, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_YearQty, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_YearAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_01PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_02PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_03PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_04PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_05PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_06PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_07PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_08PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_09PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_10PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_11PlanAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_12PlanAmt, pvStartRow, pvEndRow

	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_01PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_01PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_02PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_02PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_03PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_03PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_04PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_04PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_05PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_05PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_06PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_06PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_07PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_07PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_08PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_08PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_09PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_09PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_10PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_10PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_11PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_11PlanAmt, pvStartRow, pvEndRow
	frm1.vspdData.Row = 1 : frm1.vspdData.Col = C_12PlanColor
	If frm1.vspdData.Text = "Y" Then ggoSpread.SSSetProtected C_12PlanAmt, pvStartRow, pvEndRow
	 
    End With

End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCode			= iCurColumnPos(1)
			C_ItemPopup			= iCurColumnPos(2)
			C_ItemName			= iCurColumnPos(3)
			C_PlanUnit			= iCurColumnPos(4)
			C_PlanUnitPopup		= iCurColumnPos(5)
			C_YearQty			= iCurColumnPos(6)
			C_YearAmt			= iCurColumnPos(7)
			C_01PlanQty			= iCurColumnPos(8)
			C_02PlanQty			= iCurColumnPos(10)
			C_03PlanQty			= iCurColumnPos(12)
			C_04PlanQty			= iCurColumnPos(14)
			C_05PlanQty			= iCurColumnPos(16)
			C_06PlanQty			= iCurColumnPos(18)
			C_07PlanQty			= iCurColumnPos(20)
			C_08PlanQty			= iCurColumnPos(22)
			C_09PlanQty			= iCurColumnPos(24)
			C_10PlanQty			= iCurColumnPos(26)
			C_11PlanQty			= iCurColumnPos(28)
			C_12PlanQty			= iCurColumnPos(30)
			
			C_01PlanAmt			= iCurColumnPos(9)
			C_02PlanAmt			= iCurColumnPos(11)
			C_03PlanAmt			= iCurColumnPos(13)
			C_04PlanAmt			= iCurColumnPos(15)
			C_05PlanAmt			= iCurColumnPos(17)
			C_06PlanAmt			= iCurColumnPos(19)
			C_07PlanAmt			= iCurColumnPos(21)
			C_08PlanAmt			= iCurColumnPos(23)
			C_09PlanAmt			= iCurColumnPos(25)
			C_10PlanAmt			= iCurColumnPos(27)
			C_11PlanAmt			= iCurColumnPos(29)
			C_12PlanAmt			= iCurColumnPos(31)			

			C_01PlanColor		= iCurColumnPos(32)
			C_02PlanColor		= iCurColumnPos(33)
			C_03PlanColor		= iCurColumnPos(34)
			C_04PlanColor		= iCurColumnPos(35)
			C_05PlanColor		= iCurColumnPos(36)
			C_06PlanColor		= iCurColumnPos(37)
			C_07PlanColor		= iCurColumnPos(38)
			C_08PlanColor		= iCurColumnPos(39)
			C_09PlanColor		= iCurColumnPos(40)
			C_10PlanColor		= iCurColumnPos(41)
			C_11PlanColor		= iCurColumnPos(42)
			C_12PlanColor		= iCurColumnPos(43)
		
	End Select

End Sub	

'=========================================================================== 
Function OpenPlanNumber(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	Case 2    <% '내용부 %>
	 If frm1.txtPlanNum.readOnly = True Then Exit Function
	End Select

	IsOpenPop = True

	arrParam(0) = "계획차수"		<%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"				<%' TABLE 명칭 %>

	Select Case iWhere
	Case 1    <%' 조건부 %>
	 arrParam(2) = Trim(frm1.txtConPlanNum.Value)	<%' Code Condition%>
	Case 2    <%' 내용부 %>
	 arrParam(2) = Trim(frm1.txtPlanNum.Value)		<%' Code Condition%>
	End Select

	arrParam(3) = ""								<%' Name Cindition%>

	arrParam(4) = "MAJOR_CD=" & FilterVar("S2001", "''", "S") & ""    <%' Where Condition%>
	arrParam(5) = "계획차수"		<%' TextBox 명칭 %>
		 
	arrField(0) = "MINOR_CD"			<%' Field명(0)%>
	arrField(1) = "MINOR_NM"			<%' Field명(1)%>
		    
	arrHeader(0) = "계획차수"       <%' Header명(0)%>
	arrHeader(1) = "계획차수명"     <%' Header명(1)%>
	
	Select Case iWhere
	Case 1
		frm1.txtConPlanNum.focus
	Case 2
		frm1.txtPlanNum.focus
	End Select
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then	
		Exit Function	
	Else	
		Call SetPlanNumber(arrRet, iWhere)	
	End If 
 
End Function

'=========================================================================== 
Function SetPlanNumber(Byval arrRet,ByVal iWhere)

 With frm1

  Select Case iWhere
  Case 1
	.txtConPlanNum.value = arrRet(0) 
	.txtConPlanNumNm.value = arrRet(1)
	.txtConPlanNum.focus 
  Case 2
	.txtPlanNum.value = arrRet(0) 
	.txtPlanNumNm.value = arrRet(1)
	.txtPlanNum.focus 
	lgBlnFlgChgValue = True
  End Select
  
 End With

End Function

'=========================================================================== 
Function OpenPlanInfoRef()

 Dim arrRet
 Dim strParam

 On Error Resume Next

'계획차수관련 
If frm1.txtPlanNum.value = "" Then
	Call DisplayMsgBox("900002", "X", "X", "X")
	frm1.txtConSalesOrg.focus
	Exit Function
End IF

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 strParam = ""
 strParam = strParam & Trim(frm1.txtSalesOrg.value) & Parent.gColSep
 strParam = strParam & Trim(frm1.txtSalesOrgNm.value) & Parent.gColSep
 strParam = strParam & Trim(frm1.txtSpYear.value) & Parent.gColSep
 strParam = strParam & Trim(frm1.txtPlanTypeCd.value) & Parent.gColSep
 strParam = strParam & Trim(frm1.txtPlanTypeNm.value) & Parent.gColSep
 strParam = strParam & Trim(frm1.txtDealTypeCd.value) & Parent.gColSep
 strParam = strParam & Trim(frm1.txtDealTypeNm.value) & Parent.gColSep
 strParam = strParam & Trim(frm1.txtCurr.value) & Parent.gColSep
 strParam = strParam & lsSalesPlanBy & Parent.gColSep
 strParam = strParam & "ORG" & Parent.gRowSep

Dim iCalledAspName
Dim IntRetCD

iCalledAspName = AskPRAspName("S2114RA1")
	
If Trim(iCalledAspName) = "" Then
	IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S2114RA1", "X")
	IsOpenPop = False
	Exit Function
End If

arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,strParam), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0,0) = "" Then
  If Err.Number <> 0 Then
   Err.Clear 
  End If
  Exit Function
 Else
  Call SetPlanInfoRef(arrRet)
 End If

End Function

'===========================================================================
Function OpenSalesPlanPopup(Byval strCode, Byval iWhere, Byval Row)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 Select Case iWhere
 Case C_ItemPopUp
  frm1.vspdData.Row = Row
  frm1.vspdData.Col = 0
  If frm1.vspdData.Text <> ggoSpread.InsertFlag And lgIntFlgMode = Parent.OPMD_UMODE Then Exit Function
 End Select

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 Select Case iWhere
 Case C_ItemPopUp     <%' 고객 %>
  arrParam(1) = "B_BIZ_PARTNER"			<%' TABLE 명칭 %>
  arrParam(2) = strCode					<%' Code Condition%>
  arrParam(3) = ""						<%' Name Cindition%>
  arrParam(4) = "BP_TYPE <= " & FilterVar("CS", "''", "S") & ""		<%' Where Condition%>
  arrParam(5) = "고객"				<%' TextBox 명칭 %>
 
  arrField(0) = "BP_CD"					<%' Field명(0)%>
  arrField(1) = "BP_NM"					<%' Field명(1)%>
    
  arrHeader(0) = "고객"				<%' Header명(0)%>
  arrHeader(1) = "고객명"			<%' Header명(1)%>

 Case C_PlanUnitPopup    <%' 단위 %>
  arrParam(1) = "b_unit_of_measure"     <%' TABLE 명칭 %>
  arrParam(2) = strCode					<%' Code Condition%>
  arrParam(3) = ""						<%' Name Cindition%>
  arrParam(4) = ""						<%' Where Condition%>
  arrParam(5) = "단위"				<%' TextBox 명칭 %>
 
  arrField(0) = "unit"					<%' Field명(0)%>
  arrField(1) = "unit_nm"				<%' Field명(1)%>
    
  arrHeader(0) = "단위"				<%' Header명(0)%>
  arrHeader(1) = "단위명"			<%' Header명(1)%>

 End Select

 arrParam(0) = arrParam(5)        <%' 팝업 명칭 %>
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetSalesPlanPopUp(arrRet, iWhere)
 End If 
 
 
End Function

'=========================================================================== 
Function OpenSaleOrg(Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 Select Case iWhere
 Case 2    <% '내용부 %>
  If frm1.txtSalesOrg.readOnly = True Then Exit Function
 End Select

 IsOpenPop = True

 arrParam(0) = "영업조직"      <%' 팝업 명칭 %>
 arrParam(1) = "B_SALES_ORG"       <%' TABLE 명칭 %>

 Select Case iWhere
 Case 1    <% '조건부 %>
  arrParam(2) = Trim(frm1.txtConSalesOrg.Value) <%' Code Condition%>
  frm1.txtConSalesOrg.focus 
 Case 2    <% '내용부 %>
  arrParam(2) = Trim(frm1.txtSalesOrg.Value)	<%' Code Condition%>
  frm1.txtSalesOrg.focus 
 End Select

 arrParam(3) = ""					<%' Name Cindition%>
 arrParam(4) = "END_ORG_FLAG=" & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG=" & FilterVar("Y", "''", "S") & " " <%' Where Condition%>
 arrParam(5) = "영업조직"		<%' TextBox 명칭 %>
 
 arrField(0) = "SALES_ORG"			<%' Field명(0)%>
 arrField(1) = "SALES_ORG_NM"		<%' Field명(1)%>
    
 arrHeader(0) = "영업조직"      <%' Header명(0)%>
 arrHeader(1) = "영업조직명"    <%' Header명(1)%>

 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetSaleOrg(arrRet, iWhere)
 End If 
 
End Function

'=========================================================================== 
Function OpenPlanType(Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 Select Case iWhere
 Case 2    <% '내용부 %>
  If frm1.txtPlanTypeCd.readOnly = True Then Exit Function
 End Select

 IsOpenPop = True

 arrParam(0) = "계획구분"       <%' 팝업 명칭 %>
 arrParam(1) = "B_MINOR"			<%' TABLE 명칭 %>

 Select Case iWhere
 Case 1    <%' 조건부 %>
  arrParam(2) = Trim(frm1.txtConPlanTypeCd.Value)	<%' Code Condition%>
  frm1.txtConPlanTypeCd.focus 
 Case 2    <%' 내용부 %>
  arrParam(2) = Trim(frm1.txtPlanTypeCd.Value)		<%' Code Condition%>
  frm1.txtPlanTypeCd.focus 
 End Select

 arrParam(3) = ""						<%' Name Cindition%>
 arrParam(4) = "MAJOR_CD=" & FilterVar("S4089", "''", "S") & ""		<%' Where Condition%>
 arrParam(5) = "계획구분"			<%' TextBox 명칭 %>
 
 arrField(0) = "MINOR_CD"				<%' Field명(0)%>
 arrField(1) = "MINOR_NM"				<%' Field명(1)%>
    
 arrHeader(0) = "계획구분"			<%' Header명(0)%>
 arrHeader(1) = "계획구분명"		<%' Header명(1)%>

 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetPlanType(arrRet, iWhere)
 End If 
 
End Function

'===========================================================================
Function OpenDealType(Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 Select Case iWhere
 Case 2    <% '내용부 %>
  If frm1.txtDealTypeCd.readOnly = True Then Exit Function
 End Select

 IsOpenPop = True

 arrParam(0) = "거래구분"       <%' 팝업 명칭 %>
 arrParam(1) = "B_MINOR"			<%' TABLE 명칭 %>

 Select Case iWhere
 Case 1    <%' 조건부 %>
  arrParam(2) = Trim(frm1.txtConDealTypeCd.Value)	<%' Code Condition%>
  frm1.txtConDealTypeCd.focus 
 Case 2    <%' 내용부 %>
  arrParam(2) = Trim(frm1.txtDealTypeCd.Value)		<%' Code Condition%>
  frm1.txtDealTypeCd.focus 
 End Select

 arrParam(3) = ""						<%' Name Cindition%>
 arrParam(4) = "MAJOR_CD=" & FilterVar("S4225", "''", "S") & ""		<%' Where Condition%>
 arrParam(5) = "거래구분"			<%' TextBox 명칭 %>
 
 arrField(0) = "MINOR_CD"				<%' Field명(0)%>
 arrField(1) = "MINOR_NM"				<%' Field명(1)%>
    
 arrHeader(0) = "거래구분"			<%' Header명(0)%>
 arrHeader(1) = "거래구분명"		<%' Header명(1)%>

 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetDealType(arrRet, iWhere)
 End If 
 
End Function

'===========================================================================
Function OpenCurr(Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 Select Case iWhere
 Case 2    <% '내용부 %>
  If frm1.txtCurr.readOnly = True Then Exit Function
 End Select

 IsOpenPop = True

 arrParam(0) = "화폐"			<%' 팝업 명칭 %>
 arrParam(1) = "B_CURRENCY"			<%' TABLE 명칭 %>

 Select Case iWhere
 Case 1    <%' 조건부 %>
  arrParam(2) = Trim(frm1.txtConCurr.Value)		<%' Code Condition%>
  frm1.txtConCurr.focus 
 Case 2    <%' 내용부 %>
  arrParam(2) = Trim(frm1.txtCurr.Value)		<%' Code Condition%>
  frm1.txtCurr.focus 
 End Select
 
 arrParam(3) = ""					<%' Name Cindition%>
 arrParam(4) = ""					<%' Where Condition%>
 arrParam(5) = "화폐"			<%' TextBox 명칭 %>
 
 arrField(0) = "CURRENCY"			<%' Field명(0)%>
 arrField(1) = "CURRENCY_DESC"		<%' Field명(1)%>
    
 arrHeader(0) = "화폐"			<%' Header명(0)%>
 arrHeader(1) = "화폐명"        <%' Header명(1)%>

 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Select Case iWhere
  Case 1
   frm1.txtConCurr.value = arrRet(0)
  Case 2
   frm1.txtCurr.value = arrRet(0) 
   lgBlnFlgChgValue = True
  End Select
 End If 
 
End Function


'=========================================================================== 
Function SetPlanInfoRef(Byval arrRet)

 Dim TempRow, I, j
 Dim intLoopCnt
 Dim intCnt
 Dim blnEqualFlg
 Dim strItemCd
 Dim intCntRow
 Dim strJungBokMsg

 With frm1
  .vspdData.focus
  ggoSpread.Source = .vspdData
  .vspdData.ReDraw = False 

  TempRow = .vspdData.MaxRows         <% '☜: 현재까지의 MaxRows %>
  intLoopCnt = Ubound(arrRet, 1)        <% '☜: Reference Popup에서 선택되어진 Row만큼 추가 %>
  intCntRow = 0

  strJungBokMsg = ""

  For intCnt = 1 to intLoopCnt 
   blnEqualFlg = False

   If TempRow <> 0 Then

    strItemCd=""

    <% '---> 판매계획기초자료생성참조시 같은 품목이 있는지 체크한다 %>
    For j = 1 To TempRow
     .vspdData.Row = j
     <% '품목 %>
     .vspdData.Col = C_ItemCode
     strItemCd = .vspdData.text

     If strItemCd = arrRet(intCnt - 1, 0) Then
      blnEqualFlg = True
      strJungBokMsg = strJungBokMsg & Chr(13) & strItemCd
      Exit For
     End If

    Next

   End If
      
   If blnEqualFlg = False then
    intCntRow = intCntRow + 1
    .vspdData.MaxRows = CLng(TempRow) + CLng(intCntRow)
    .vspdData.Row = CLng(TempRow) + CLng(intCntRow)

    .vspdData.Col = 0
    .vspdData.Text = ggoSpread.InsertFlag

    <% '품목 %>
    .vspdData.Col = C_ItemCode
    .vspdData.text = arrRet(intCnt - 1, 0)

    <% '품목명 %>
    .vspdData.Col = C_ItemName
    .vspdData.text = arrRet(intCnt - 1, 1)

    <% '계획단위 %>
    .vspdData.Col = C_PlanUnit
    .vspdData.text = arrRet(intCnt - 1, 2)


    <% '월별 수량,금액 %>
    .vspdData.Col = C_01PlanQty
    .vspdData.text = arrRet(intCnt - 1, 5)
    
    .vspdData.Col = C_02PlanQty
    .vspdData.text = arrRet(intCnt - 1, 7)

    .vspdData.Col = C_03PlanQty
    .vspdData.text = arrRet(intCnt - 1, 9)

    .vspdData.Col = C_04PlanQty
    .vspdData.text = arrRet(intCnt - 1, 11)

    .vspdData.Col = C_05PlanQty
    .vspdData.text = arrRet(intCnt - 1, 13)

    .vspdData.Col = C_06PlanQty
    .vspdData.text = arrRet(intCnt - 1, 15)

    .vspdData.Col = C_07PlanQty
    .vspdData.text = arrRet(intCnt - 1, 17)

    .vspdData.Col = C_08PlanQty
    .vspdData.text = arrRet(intCnt - 1, 19)

    .vspdData.Col = C_09PlanQty
    .vspdData.text = arrRet(intCnt - 1, 21)

    .vspdData.Col = C_10PlanQty
    .vspdData.text = arrRet(intCnt - 1, 23)

    .vspdData.Col = C_11PlanQty
    .vspdData.text = arrRet(intCnt - 1, 25)

    .vspdData.Col = C_12PlanQty
    .vspdData.text = arrRet(intCnt - 1, 27)

    .vspdData.Col = C_01PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 6)
    
    .vspdData.Col = C_02PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 8)

    .vspdData.Col = C_03PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 10)

    .vspdData.Col = C_04PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 12)

    .vspdData.Col = C_05PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 14)

    .vspdData.Col = C_06PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 16)

    .vspdData.Col = C_07PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 18)

    .vspdData.Col = C_08PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 20)

    .vspdData.Col = C_09PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 22)

    .vspdData.Col = C_10PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 24)

    .vspdData.Col = C_11PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 26)

    .vspdData.Col = C_12PlanAmt
    .vspdData.text = arrRet(intCnt - 1, 28)

    Call SetSpreadColor(CLng(TempRow)+CLng(intCntRow),CLng(TempRow)+CLng(intCntRow),False)
  
   End if
  Next
  .vspdData.ReDraw = True

 End With

 Call JungBokMsg(strJungBokMsg,"품목")

 If blnEqualFlg = True And intCntRow = 0 Then Exit Function    <% '참조에 의해 행이 추가되었는지 여부 %>

 Call MonthTotalSum(C_01PlanQty,C_YearQty)
 Call MonthTotalSum(C_01PlanAmt,C_YearAmt)

 lgBlnFlgChgValue = True

End Function

'=========================================================================== 
Function SetSalesPlanPopUp(Byval arrRet,ByVal iWhere)

 With frm1

  Select Case iWhere
  Case C_ItemPopUp  <% '품목 %>
   .vspdData.Col = C_ItemCode
   .vspdData.Text = arrRet(0)
   .vspdData.Col = C_ItemName
   .vspdData.Text = arrRet(1)
  Case C_PlanUnitPopup <% '단위 %>
   .vspdData.Col = C_PlanUnit
   .vspdData.Text = arrRet(0)
  End Select
  
  Call vspdData_Change(.vspdData.Col, .vspdData.Row)  <% ' 변경이 읽어났다고 알려줌 %>

 End With

 lgBlnFlgChgValue = True
 
End Function

'=========================================================================== 
Function SetSaleOrg(Byval arrRet,ByVal iWhere)

 With frm1

  Select Case iWhere
  Case 1
   .txtConSalesOrg.value = arrRet(0) 
   .txtConSalesOrgNm.value = arrRet(1)
   .txtConSalesOrg.focus 
  Case 2
   .txtSalesOrg.value = arrRet(0) 
   .txtSalesOrgNm.value = arrRet(1)
   .txtSalesOrg.focus 
   lgBlnFlgChgValue = True
  End Select
  
 End With

End Function


'=========================================================================== 
Function SetPlanType(Byval arrRet,ByVal iWhere)

 With frm1

  Select Case iWhere
  Case 1
   .txtConPlanTypeCd.value = arrRet(0) 
   .txtConPlanTypeNm.value = arrRet(1)   
   .txtConPlanTypeCd.focus 
  Case 2
   .txtPlanTypeCd.value = arrRet(0) 
   .txtPlanTypeNm.value = arrRet(1)
   .txtPlanTypeCd.focus 
   lgBlnFlgChgValue = True
  End Select
  
 End With

End Function


'=========================================================================== 
Function SetDealType(Byval arrRet,ByVal iWhere)

 With frm1

  Select Case iWhere
  Case 1
   .txtConDealTypeCd.value = arrRet(0) 
   .txtConDealTypeNm.value = arrRet(1)
   .txtConDealTypeCd.focus 
  Case 2
   .txtDealTypeCd.value = arrRet(0) 
   .txtDealTypeNm.value = arrRet(1)
   .txtDealTypeCd.focus 
   lgBlnFlgChgValue = True
  End Select
  
 End With

End Function

<%
'=============================================================================================================
' Function Desc : 수량 라디오버튼을 클릭시 SpreadColor
'=============================================================================================================
%>
Function SetQtySpreadColor(ByVal lRow, KuBun)

    Dim MRow
    
    With frm1

    .vspdData.ReDraw = False

	Select Case KuBun
	Case lsInsert
	 MRow = lRow
	Case lsQuery, lsSelect
	 MRow = .vspdData.MaxRows   
	 Call SpreadProtectUnLock(C_01PlanQty,1)
	End Select

	ggoSpread.Source = .vspdData

    ggoSpread.SSSetRequired C_ItemCode, lRow, MRow
    ggoSpread.SSSetProtected C_ItemName, lRow, MRow
    ggoSpread.SSSetProtected C_YearQty, lRow, MRow
    ggoSpread.SSSetProtected C_YearAmt, lRow, MRow
    ggoSpread.SSSetRequired C_01PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_02PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_03PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_04PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_05PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_06PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_07PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_08PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_09PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_10PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_11PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_12PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_01PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_02PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_03PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_04PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_05PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_06PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_07PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_08PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_09PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_10PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_11PlanAmt, lRow, MRow
    ggoSpread.SSSetProtected C_12PlanAmt, lRow, MRow

    .vspdData.ReDraw = True
    
    End With

End Function


<%
'=============================================================================================================
' Function Desc : 금액 라디오버튼을 클릭시 SpreadColor
'=============================================================================================================
%>
Function SetAmtSpreadColor(ByVal lRow,KuBun)
    
    Dim MRow
    
    With frm1
   
    .vspdData.ReDraw = False

	Select Case KuBun
	Case lsInsert
	 MRow = lRow
	Case lsQuery, lsSelect
	 MRow = .vspdData.MaxRows   
	 Call SpreadProtectUnLock(C_01PlanQty,1)
	End Select

	ggoSpread.Source = .vspdData

    ggoSpread.SSSetRequired C_ItemCode, lRow, MRow
    ggoSpread.SSSetProtected C_ItemName, lRow, MRow
    ggoSpread.SSSetProtected C_YearQty, lRow, MRow
    ggoSpread.SSSetProtected C_YearAmt, lRow, MRow
    ggoSpread.SSSetProtected C_01PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_02PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_03PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_04PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_05PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_06PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_07PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_08PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_09PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_10PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_11PlanQty, lRow, MRow
    ggoSpread.SSSetProtected C_12PlanQty, lRow, MRow
    ggoSpread.SSSetRequired C_01PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_02PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_03PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_04PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_05PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_06PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_07PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_08PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_09PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_10PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_11PlanAmt, lRow, MRow
    ggoSpread.SSSetRequired C_12PlanAmt, lRow, MRow

    .vspdData.ReDraw = True
    
    End With

End Function


<%
'=============================================================================================================
' Function Name : SpreadProtectUnLock
' Function Desc : 현 Spread에서는 Protect에서 UnProtect시킬때 Lock을 초기화하는 부분이 없어서 이 함수를 만들었다..
'=============================================================================================================
%>
Function SpreadProtectUnLock(Col,Row)
 With frm1
     .vspdData.Col = Col
     .vspdData.SetColItemData .vspdData.Col, 2
     .vspdData.Col2 = C_12PlanAmt
     .vspdData.Row = Row
     .vspdData.Row2 = .vspdData.MaxRows
     
     .vspdData.BlockMode = True
     .vspdData.Protect = False
     .vspdData.Lock = False
     .vspdData.BlockMode = False
 End With
End Function


<%
'=======================================================================================================
' Function Name : MonthTotalSum
' Function Desc : 년 판매수량/금액의 합 
'=======================================================================================================
%>
Function MonthTotalSum(GCol,GTotal)

 Dim SumTotal, iMonth, lRow

 ggoSpread.Source = frm1.vspdData 

 For lRow = 1 To frm1.vspdData.MaxRows 

  SumTotal = 0

	if GCol = C_01PlanQty Then
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_01PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_02PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow	
		frm1.vspdData.Col = C_03PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_04PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_05PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_06PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_07PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_08PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_09PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_10PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_11PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_12PlanQty
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

	Else
		
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_01PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_02PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow	
		frm1.vspdData.Col = C_03PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_04PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_05PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_06PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_07PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_08PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_09PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_10PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_11PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_12PlanAmt
	    If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If
	
	End if

  frm1.vspdData.Row = lRow
  frm1.vspdData.Col = GTotal
  frm1.vspdData.Text= UNIFormatNumber(SumTotal,ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
 Next

End Function

<%
'====================================================================================================
' Function Name : CookiePage
' Function Desc : Jump시 해당 화면에 조회값 인자/인수 전달 
'====================================================================================================
%>
Function CookiePage()

 On Error Resume Next

 Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>

 Dim strTemp, arrVal
 
 WriteCookie CookieSplit , frm1.txtSalesOrg.value & Parent.gRowSep & frm1.txtSpYear.value & Parent.gRowSep _
  & frm1.txtPlanTypeCd.value & Parent.gRowSep & frm1.txtDealTypeCd.value & Parent.gRowSep _
  & frm1.txtCurr.value & Parent.gRowSep & frm1.txtPlanNum.value

End Function


<%
'===========================================================================================================
' Function Desc : After Spread Cell Click, Variables initializes
'=======================================================================================================
%>
Function SpreadCellClick(ByVal Col,ByVal Row)

 lsPlanMonth=""

   Select Case Col
 Case C_01PlanQty, C_01PlanAmt : lsPlanMonth = "01"
 Case C_02PlanQty, C_02PlanAmt : lsPlanMonth = "02"
 Case C_03PlanQty, C_03PlanAmt : lsPlanMonth = "03"
 Case C_04PlanQty, C_04PlanAmt : lsPlanMonth = "04"
 Case C_05PlanQty, C_05PlanAmt : lsPlanMonth = "05"
 Case C_06PlanQty, C_06PlanAmt : lsPlanMonth = "06"
 Case C_07PlanQty, C_07PlanAmt : lsPlanMonth = "07"
 Case C_08PlanQty, C_08PlanAmt : lsPlanMonth = "08"
 Case C_09PlanQty, C_09PlanAmt : lsPlanMonth = "09"
 Case C_10PlanQty, C_10PlanAmt : lsPlanMonth = "10"
 Case C_11PlanQty, C_11PlanAmt : lsPlanMonth = "11"
 Case C_12PlanQty, C_12PlanAmt : lsPlanMonth = "12"
 Case Else : Exit Function
 End Select

 frm1.vspdData.Row = Row
 frm1.vspdData.Col = C_ItemCode
 lsItemCode=frm1.vspdData.Text

 frm1.vspdData.Row = Row
 frm1.vspdData.Col = C_ItemName
 lsItemName=frm1.vspdData.Text

 frm1.vspdData.Row = Row
 frm1.vspdData.Col = C_PlanUnit
 lsPlanUnit=frm1.vspdData.Text

 Call SplitCheckbtnName(Col,Row)      <% '해당 월에 맞는 버튼명 %>

End Function


<%
'=======================================================================================================
' Function Desc : Before Batch Button , Requried Value Checking Msg
'=======================================================================================================
%>
Function BatchReqCheckMsg()

 BatchReqCheckMsg = False

 Dim IntRetCD
 ggoSpread.Source = frm1.vspdData 

 <% '변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 %>
 If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
 IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 계속 하시겠습니까?%>
 If IntRetCD = vbNo Then Exit Function
 End If

 <% '변경이 없을때 작업진행여부 체크 %>
 If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
 IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")                <% '작업을 수행하시겠습니까? %>
 If IntRetCD = vbNo Then Exit Function
 End If


 If lsPlanMonth = "" Then
  MsgBox "확정처리 할 월을 선택하세요.", vbExclamation, Parent.gLogoName
  Exit Function
 End If

 BatchReqCheckMsg = True

End Function


<%
'=======================================================================================================
' Function Desc : Before Month OnChange , Requried Value Checking Msg
'=======================================================================================================
%>
Function OnChangeReqCheckMsg()

 OnChangeReqCheckMsg = False

 'Const Parent.SS_ACTION_ACTIVE_CELL = 0

<%  '-----------------------
    'Check content area
    '-----------------------%>

    If Len(Trim(frm1.txtCurr.Value)) = 0 Then
  Call DisplayMsgBox("970021","X",frm1.txtCurr.alt,"X")
  frm1.txtCurr.focus
  Exit Function
 End If

 frm1.vspdData.Row = frm1.vspdData.ActiveRow
 frm1.vspdData.Col = C_ItemCode
 If Len(Trim(frm1.vspdData.Text)) = 0 Then
  frm1.vspdData.Row = 0
  Call DisplayMsgBox("970021","X",frm1.vspdData.Text,"X")

  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL
  
  Exit Function
 End If

 frm1.vspdData.Row = frm1.vspdData.ActiveRow
 frm1.vspdData.Col = C_PlanUnit
 If Len(Trim(frm1.vspdData.Text)) = 0 Then
  frm1.vspdData.Row = 0
  Call DisplayMsgBox("970021","X",frm1.vspdData.Text,"X")

  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL

  Exit Function
 End If

 OnChangeReqCheckMsg = True

End Function


<%
'=======================================================================================================
' Function Desc : Before Month OnChange , Requried Value Checking Msg
'=======================================================================================================
%>
Function SplitFlagMonthColor(strMonth,ColorRow)

 ggoSpread.SSSetRequired CInt(strMonth) + 1, ColorRow, ColorRow
 call SpreadProtectUnLock(CInt(strMonth) + 1, ColorRow) 
 Exit Function

 With frm1
  Select Case .txtRdoSelect.value
  Case .rdoSelectAmt.value
   ggoSpread.SSSetProtected strMonth, ColorRow, ColorRow
   ggoSpread.SSSetRequired CInt(strMonth) + 1, ColorRow, ColorRow
  Case Else
   ggoSpread.SSSetRequired strMonth, ColorRow, ColorRow
   ggoSpread.SSSetProtected CInt(strMonth) + 1, ColorRow, ColorRow
  End Select
 End With

End Function


<%
'=======================================================================================================
' Function Desc : Month Qty/Amt OnChange
'=======================================================================================================
%>
Function UpdateQtyAmtSvr()

 With frm1

  Dim strval

  
  If   LayerShowHide(1) = False Then
             Exit Function 
        End If

  strVal = ""    
  strVal = BIZ_PGM_ID & "?txtMode=" & lsQtyAmt         <%'☜: 비지니스 처리 ASP의 상태 %>
  strVal = strVal & "&lsItemCode=" & lsItemCode         <%'☜: Batch 조건 데이타 %>
  strVal = strVal & "&lsPlanUnit=" & lsPlanUnit
  strVal = strVal & "&lsPlanMonth=" & lsPlanMonth
  strVal = strVal & "&lsPlanQtyAmt=" & lsPlanQtyAmt
  strVal = strVal & "&txtCurr=" & Trim(frm1.txtCurr.value)
  strVal = strVal & "&txtRdoSelect=" & Trim(frm1.txtRdoSelect.value)

  Call RunMyBizASP(MyBizASP, strVal)            <%'☜: 비지니스 ASP 를 가동 %>

 End With

End Function


<%
'=======================================================================================================
' Function Desc : Month Qty/Amt OnChange OK
'=======================================================================================================
%>
Function UpdateQtyAmtSvrOK()

 Call MonthTotalSum(C_01PlanQty,C_YearQty)
 Call MonthTotalSum(C_01PlanAmt,C_YearAmt) 

End Function


<%
'=======================================================================================================
' Function Desc : Before PlanSeq PopUp , Requried Value Checking Msg
'=======================================================================================================
%>
Function PlanSeqCheckMsg()

 PlanSeqCheckMsg = False


<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Len(Trim(frm1.txtConSalesOrg.Value)) = 0 Then
  Call DisplayMsgBox("970021","X",frm1.txtConSalesOrg.alt,"X")
  frm1.txtConSalesOrg.focus
  Exit Function
 End If 

    If Len(Trim(frm1.txtConSpYear.Value)) = 0 Then
  Call DisplayMsgBox("970021","X",frm1.txtConSpYear.alt,"X")
  frm1.txtConSpYear.focus
  Exit Function
 End If

    If Len(Trim(frm1.txtConPlanTypeCd.Value)) = 0 Then
  Call DisplayMsgBox("970021","X",frm1.txtConPlanTypeCd.alt,"X")
  frm1.txtConPlanTypeCd.focus
  Exit Function
 End If

    If Len(Trim(frm1.txtConDealTypeCd.Value)) = 0 Then
  Call DisplayMsgBox("970021","X",frm1.txtConDealTypeCd.alt,"X")
  frm1.txtConDealTypeCd.focus
  Exit Function
 End If
 
    If Len(Trim(frm1.txtConCurr.Value)) = 0 Then
  Call DisplayMsgBox("970021","X",frm1.txtConCurr.alt,"X")
  frm1.txtConCurr.focus
  Exit Function
 End If 
 
 PlanSeqCheckMsg = True

End Function

<%
'=======================================================================================================
' Function Desc : 숫자만 입력받는 형식 체크 
'=======================================================================================================
%>
Function NumericCheck()

 Dim objEl, KeyCode
 
 Set objEl = window.event.srcElement
 KeyCode = window.event.keycode

 Select Case KeyCode
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
 Case Else
  window.event.keycode = 0
 End Select

End Function

<%
'=============================================================================================================
' Function Desc : Reference -> Dual Check
'=============================================================================================================
%>
Function JungBokMsg(strJungBok,strID)

 Dim strJugBokMsg

 If Len(Trim(strJungBok)) Then strJungBok = strID & Chr(13) & String(20,"=") & strJungBok
 If Len(Trim(strJungBok)) Then strJugBokMsg = strJungBok & Chr(13) & Chr(13)
 If Len(Trim(strJugBokMsg)) Then
  strJugBokMsg = strJugBokMsg & "이미 동일한 품목이 존재합니다"
  MsgBox strJugBokMsg, vbInformation, Parent.gLogoName
 End If

End Function

'========================================================================================================= 
Sub Form_Load()

 Call LoadInfTB19029()
 Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
 Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
 Call InitVariables              '⊙: Initializes local global variables
 Call SetDefaultVal 
 Call InitSpreadSheet

 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄/찾기 

'계획차수관련 
	'Call SetToolBar("11101111001011")          '⊙: 버튼 툴바 제어 
	Call SetToolBar("11000000000011")          '⊙: 버튼 툴바 제어 

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


<%
'==========================================================================================
'   Event Name : Radio Button_OnClick()
'   Event Desc : 라디오 버튼 클릭시 이벤트 
'==========================================================================================
%>
Sub rdoSelectQty_OnClick()
 frm1.txtRdoSelect.value = frm1.rdoSelectQty.value
 Call SetQtySpreadColor(1,lsSelect)
End Sub

Sub rdoSelectAmt_OnClick()
 frm1.txtRdoSelect.value = frm1.rdoSelectAmt.value
 Call SetAmtSpreadColor(1,lsSelect)
End Sub

<%
'==========================================================================================
'   Event Name : btnSplit_OnClick()
'   Event Desc : 확정처리을 클릭할 경우 발생 
'==========================================================================================
%>
Sub btnSplit_OnClick()

    Err.Clear                                                               <%'☜: Protect system from crashing%>

	Call SpreadCellClick(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow)
	If BatchReqCheckMsg = False Then Exit Sub        <%' Requried Value Check Msg %>
 
	Call BatchButton(lsConfirm)   <% '확정처리 %>

End Sub

<%
'==========================================================================================
'   Event Name : BatchButton()
'   Event Desc : 버튼을 클릭할 경우 서버단에 넘기는 공통의 값 
'==========================================================================================
%>
Function BatchButton(SKubun)

 Dim strval, strSpread

 
  If   LayerShowHide(1) = False Then
             Exit Function 
        End If
    
 
 strSpread = ""
 strSpread = lsPlanMonth
 frm1.txtSpread.value  = strSpread 

 strVal = ""    
 strVal = BIZ_PGM_ID & "?txtMode=" & SKubun          <%'☜: 비지니스 처리 ASP의 상태 %>
 strVal = strVal & "&txtSalesOrg=" & Trim(frm1.txtSalesOrg.value)    <%'☜: Batch 조건 데이타 %>
 strVal = strVal & "&txtSpYear=" & Trim(frm1.txtSpYear.value)
 strVal = strVal & "&txtPlanTypeCd=" & Trim(frm1.txtPlanTypeCd.value)
 strVal = strVal & "&txtDealTypeCd=" & Trim(frm1.txtDealTypeCd.value)
 strVal = strVal & "&txtCurr=" & Trim(frm1.txtCurr.value)
 strVal = strVal & "&txtPlanNum=" & Trim(frm1.txtPlanNum.value)
 strVal = strVal & "&txtSpread=" & Trim(frm1.txtSpread.value)

 Call RunMyBizASP(MyBizASP, strVal)            <%'☜: 비지니스 ASP 를 가동 %>

End Function


 
<%
'==========================================================================================
'   Function Name : SplitCheckbtnName()
'   Function Desc : 해당월을 클릭할 경우 버튼명 결정 
'==========================================================================================
%>
Function SplitCheckbtnName(ByVal Col, Byval Row)

 With frm1

  If Row < 0 Then
   .btnSplit.disabled = True
   Exit Function
  End If

  Select Case Col  
  Case C_01PlanQty,C_02PlanQty,C_03PlanQty,C_04PlanQty,C_05PlanQty, _
  C_06PlanQty,C_07PlanQty,C_08PlanQty,C_09PlanQty,C_10PlanQty, _
  C_11PlanQty,C_12PlanQty, _
  C_01PlanAmt,C_02PlanAmt,C_03PlanAmt,C_04PlanAmt,C_05PlanAmt, _
  C_06PlanAmt,C_07PlanAmt,C_08PlanAmt,C_09PlanAmt,C_10PlanAmt, _
  C_11PlanAmt,C_12PlanAmt  

  Case Else
    
   .btnSplit.disabled = True
   Exit Function

  End Select  

  .vspdData.Row = Row
  
  Select Case Col
  Case C_01PlanQty, C_01PlanAmt : .vspdData.Col = C_01PlanColor
  Case C_02PlanQty, C_02PlanAmt : .vspdData.Col = C_02PlanColor
  Case C_03PlanQty, C_03PlanAmt : .vspdData.Col = C_03PlanColor
  Case C_04PlanQty, C_04PlanAmt : .vspdData.Col = C_04PlanColor
  Case C_05PlanQty, C_05PlanAmt : .vspdData.Col = C_05PlanColor
  Case C_06PlanQty, C_06PlanAmt : .vspdData.Col = C_06PlanColor
  Case C_07PlanQty, C_07PlanAmt : .vspdData.Col = C_07PlanColor
  Case C_08PlanQty, C_08PlanAmt : .vspdData.Col = C_08PlanColor
  Case C_09PlanQty, C_09PlanAmt : .vspdData.Col = C_09PlanColor
  Case C_10PlanQty, C_10PlanAmt : .vspdData.Col = C_10PlanColor
  Case C_11PlanQty, C_11PlanAmt : .vspdData.Col = C_11PlanColor
  Case C_12PlanQty, C_12PlanAmt : .vspdData.Col = C_12PlanColor
  End Select

  Select Case UCase(Trim(.vspdData.Text))
  Case "Y"
   .btnSplit.value = "확정취소"
   .btnSplit.disabled = False  
  Case "N"
   .btnSplit.value = "확정처리"
   .btnSplit.disabled = False
  Case Else
   .btnSplit.value = "확정처리"
   .btnSplit.disabled = True
  End Select

 End With

End Function


<%
'==========================================================================================
'   Event Desc : btnSplit처리가 성공적일 경우 
'==========================================================================================
%>
Function btnSplit_Ok()
 Call MainQuery()
End Function


'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

 With frm1.vspdData 
 
  ggoSpread.Source = frm1.vspdData
   
  If Row > 0 Then
   Select Case Col
   Case C_ItemPopUp
    .Col = Col - 1
    .Row = Row
    Call OpenSalesPlanPopup(.Text, C_ItemPopUp,Row)
   Case C_PlanUnitPopup
    .Col = Col - 1
    .Row = Row
    Call OpenSalesPlanPopup(.Text, C_PlanUnitPopup,Row)
   End Select
  End If

  Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")
    
 End With
End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
 
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then Exit Sub End If	

	lsPlanMonth=""

	frm1.btnSplit.disabled = True

	Select Case Col  
	Case C_01PlanQty,C_02PlanQty,C_03PlanQty,C_04PlanQty,C_05PlanQty, _
	 C_06PlanQty,C_07PlanQty,C_08PlanQty,C_09PlanQty,C_10PlanQty, _
	 C_11PlanQty,C_12PlanQty, _
	 C_01PlanAmt,C_02PlanAmt,C_03PlanAmt,C_04PlanAmt,C_05PlanAmt, _
	 C_06PlanAmt,C_07PlanAmt,C_08PlanAmt,C_09PlanAmt,C_10PlanAmt, _
	 C_11PlanAmt,C_12PlanAmt  

	  If Row > 0 Then
			Call SpreadCellClick(Col,Row)	
	  End If

	End Select

	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If

		 Exit Sub     
	End If

End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub        
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub        
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub


'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

	Select Case Col  
	Case C_01PlanQty,C_02PlanQty,C_03PlanQty,C_04PlanQty,C_05PlanQty, _
	 C_06PlanQty,C_07PlanQty,C_08PlanQty,C_09PlanQty,C_10PlanQty, _
	 C_11PlanQty,C_12PlanQty, _
	 C_01PlanAmt,C_02PlanAmt,C_03PlanAmt,C_04PlanAmt,C_05PlanAmt, _
	 C_06PlanAmt,C_07PlanAmt,C_08PlanAmt,C_09PlanAmt,C_10PlanAmt, _
	 C_11PlanAmt,C_12PlanAmt  

    Call MonthTotalSum(C_01PlanQty,C_YearQty)
    Call MonthTotalSum(C_01PlanAmt,C_YearAmt)

 End Select

End Sub

'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
 
 Dim IntRetCD 
 
 If Col < 0 Or Row < 0 Or NewCol < 0 Or NewRow < 0 Then Exit Sub
  
 If Col < NewCol Then
  Call SplitCheckbtnName(Col+2,Row)  
 ElseIf Col > NewCol Then
  Call SplitCheckbtnName(Col-2,Row)  
 End If

 If Row < NewRow Then
  Call SplitCheckbtnName(Col,Row+1)
 ElseIf Row > NewRow Then
  Call SplitCheckbtnName(Col,Row-1)
 End If
 
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then 
		If lgStrPrevKey <> "" Then  
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If     
  
			If DBQuery = False Then   
				Exit Sub
			End If
		End if
	End if     
 
End Sub

<%
'=======================================================================================================
' Function Name : Text Numeric_onKeyPress()
' Function Desc : 숫자만 입력받는 TextBox KeyIn 작업시 
'=======================================================================================================
%>
Sub txtConSpYear_onKeyPress()
 Call NumericCheck()
End Sub

Sub txtSpYear_onKeyPress()
 Call NumericCheck()
End Sub

Sub txtConPlanNum_onKeyPress()
 Call NumericCheck()
End Sub

Sub txtPlanNum_onKeyPress()
 Call NumericCheck()
End Sub

'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

<%    '-----------------------
    'Check previous data area
    '----------------------- %>
 '************ 싱글/멀티인 경우 **************
 ggoSpread.Source = frm1.vspdData 
 If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then
      Exit Function
  End If
 End If


    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
    End If

<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")          <%'⊙: Clear Contents  Field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>

<%  '-----------------------
    'Query function call area
    '----------------------- %>
    Call DbQuery                <%'☜: Query db data%>

    FncQuery = True                <%'⊙: Processing is OK%>
        
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Check previous data area
    '-----------------------%>
 '************ 싱글/멀티인 경우 **************
 ggoSpread.Source = frm1.vspdData 
 If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then
      Exit Function
  End If
 End If

<%  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------%>
    Call ggoOper.ClearField(Document, "A")                                      <%'⊙: Clear Condition,Contents Field%>
    Call ggoOper.LockField(Document, "N")                                       <%'⊙: Lock  Suitable  Field%>
    Call SetDefaultVal
    Call InitVariables               <%'⊙: Initializes local global variables%>

 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄 

'계획차수관련 
    'Call SetToolBar("11101111001011")          '⊙: 버튼 툴바 제어 
    Call SetToolBar("11000000000011")          '⊙: 버튼 툴바 제어 

    FncNew = True                <%'⊙: Processing is OK%>

End Function

'========================================================================================
Function FncDelete() 
    
    Exit Function
    Err.Clear                                                               '☜: Protect system from crashing    
    
    FncDelete = False              <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition,Contents Field
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    dim iCount
    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

 If frm1.vspdData.MaxRows < 1 Then
  MsgBox "저장할 품목이 없습니다", vbExclamation, Parent.gLogoName
  Exit Function
 End If
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
 '************ 싱글/멀티인 경우 **************
 ggoSpread.Source = frm1.vspdData 
 If lgBlnFlgChgValue = False Or ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
 End If
 
 ''''' 2002-09-18 수정 : Insert일 경우 연간 금액합계가 0보다 커야 함.
 for iCount = 1 to frm1.vspdData.MaxRows
  frm1.vspdData.Row = iCount
  frm1.vspdData.Col = 0
  If frm1.vspdData.Text = ggoSpread.InsertFlag then
    
   frm1.vspdData.Col = C_YearAmt  
   If Trim(frm1.vspdData.Text ) <= 0 then
    IntRetCD = DisplayMsgBox("202401", "X", iCount & "행", "X")   
    
    frm1.vspdData.ReDraw = False
    frm1.vspdData.Col = C_01PlanAmt
    frm1.vspdData.Action = 0
    frm1.vspdData.EditMode = True
       
    frm1.vspdData.ReDraw = True

 '   frm1.vspdData.focus     
    
    Exit Function    
   End if  
  End if
 Next 
 
 
<%  '-----------------------
    'Check content area
    '-----------------------%>
	ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") Then     <%'⊙: Check contents area%>
       Exit Function
    End If

    If ggoSpread.SSDefaultCheck = False Then     <%'⊙: Check contents area%>
       Exit Function
    End If
<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll DbSave                                                    <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

'========================================================================================
Function FncCopy() 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

 Dim iRow

 With frm1

  .vspdData.ReDraw = False
 
  ggoSpread.Source = frm1.vspdData   
  iRow = .vspdData.ActiveRow
  ggoSpread.CopyRow

  Call SetSpreadColor(iRow + 1,iRow + 1,False)

  .vspdData.Row = .vspdData.ActiveRow
  .vspdData.Col = C_ItemCode
  .vspdData.Text = ""
  .vspdData.Col = C_ItemName
  .vspdData.Text = ""
  .vspdData.Col = C_PlanUnit
  .vspdData.Text = ""

   .vspdData.Col = C_01PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_02PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_03PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_04PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_05PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_06PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_07PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_08PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_09PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_10PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_11PlanColor : .vspdData.Text = ""
   .vspdData.Col = C_12PlanColor : .vspdData.Text = ""
  
  .vspdData.ReDraw = True
 End With
    
End Function

'========================================================================================
Function FncCancel() 
 If frm1.vspdData.MaxRows < 1 Then Exit Function
 ggoSpread.Source = frm1.vspdData 
 ggoSpread.EditUndo                                                  '☜: Protect system from crashing

 Call MonthTotalSum(C_01PlanQty,C_YearQty)
 Call MonthTotalSum(C_01PlanAmt,C_YearAmt)

End Function

'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
 Dim imRow
 Dim GCol

 If IsNumeric(Trim(pvRowCnt)) Then
	imRow = Cint(pvRowCnt)
 Else	
    imRow = AskSpdSheetAddRowCount()
    If imRow = "" Then
        Exit Function
    End If
 End If    

 With frm1
  .vspdData.focus
  ggoSpread.Source = .vspdData

  .vspdData.ReDraw = False

  ggoSpread.InsertRow ,imRow

  Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1, False)
    
  lgBlnFlgChgValue = True

    <% '----------  Coding part  -------------------------------------------------------------%>   

  ggoSpread.Source = .vspdData

   .vspdData.Row = .vspdData.ActiveRow

   .vspdData.Col = C_01PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_02PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_03PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_04PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_05PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_06PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_07PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_08PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_09PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_10PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_11PlanQty : .vspdData.Text = 0
   .vspdData.Col = C_12PlanQty : .vspdData.Text = 0

   .vspdData.Col = C_01PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_02PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_03PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_04PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_05PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_06PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_07PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_08PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_09PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_10PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_11PlanAmt : .vspdData.Text = 0
   .vspdData.Col = C_12PlanAmt : .vspdData.Text = 0

	.vspdData.ReDraw = True

    End With

End Function

'========================================================================================
Function FncDeleteRow() 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
 lDelRows = ggoSpread.DeleteRow

 Call MonthTotalSum(C_01PlanQty,C_YearQty)
 Call MonthTotalSum(C_01PlanAmt,C_YearAmt)
 
    lgBlnFlgChgValue = True
    
    End With
End Function

'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
Function FncExcel() 
 Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function

'========================================================================================
Function FncFind() 
 Call parent.FncFind(Parent.C_SINGLEMULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()

	dbQueryOk()

End Sub


'========================================================================================
Function FncExit()
 Dim IntRetCD
 FncExit = False
 '************ 싱글/멀티인 경우 **************
 ggoSpread.Source = frm1.vspdData 
 If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vb
  If IntRetCD = vbNo Then
   Exit Function
  End If
 End If
    FncExit = True
End Function

'========================================================================================
Function DbQuery() 

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    
	If   LayerShowHide(1) = False Then
        Exit Function 
    End If
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>
    
    Dim strVal
    

    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtConSalesOrg=" & Trim(frm1.HConSalesOrg.value)   <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtConSpYear=" & Trim(frm1.HConSpYear.value)
		strVal = strVal & "&txtConPlanTypeCd=" & Trim(frm1.HPlanTypeCd.value)
		strVal = strVal & "&txtConDealTypeCd=" & Trim(frm1.HConDealTypeCd.value)
		strVal = strVal & "&txtConCurr=" & Trim(frm1.HConCurr.value)
		strVal = strVal & "&txtConPlanNum=" & Trim(frm1.HConPlanNum.value)  
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtConSalesOrg=" & Trim(frm1.txtConSalesOrg.value)   <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtConSpYear=" & Trim(frm1.txtConSpYear.value)
		strVal = strVal & "&txtConPlanTypeCd=" & Trim(frm1.txtConPlanTypeCd.value)
		strVal = strVal & "&txtConDealTypeCd=" & Trim(frm1.txtConDealTypeCd.value)
		strVal = strVal & "&txtConCurr=" & Trim(frm1.txtConCurr.value)
		strVal = strVal & "&txtConPlanNum=" & Trim(frm1.txtConPlanNum.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If 
 
	Call RunMyBizASP(MyBizASP, strVal)          <%'☜: 비지니스 ASP 를 가동 %>
 
    DbQuery = True               <%'⊙: Processing is NG%>

End Function

'========================================================================================
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>
	lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")         <%'⊙: This function lock the suitable field%>
    Call SetToolbar("11101111001111")          '⊙: 버튼 툴바 제어 

	Call MonthTotalSum(C_01PlanQty,C_YearQty)
	Call MonthTotalSum(C_01PlanAmt,C_YearAmt)

	With frm1

		Call SetSpreadColor(-1,-1,True)

		.btnSplit.disabled = False
		.txtConSalesOrg.focus 
	End With

	Call SplitCheckbtnName(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
 
	With frm1
		.txtConSalesOrgNm.value = .txtSalesOrgNm.value 
		.txtConPlanTypeNm.value = .txtPlanTypeNm.value
		.txtConDealTypeNm.value = .txtDealTypeNm.value
		.txtConPlanNumNm.value = .txtPlanNumNm.value
		.vspdData.Focus 
	End With

End Function

'========================================================================================
Function DbSave() 

    Err.Clear                <%'☜: Protect system from crashing%>
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal,strDel
 
 
	If   LayerShowHide(1) = False Then
        Exit Function 
    End If
 
    DbSave = False                                                          '⊙: Processing is NG
    
 With frm1
  .txtMode.value = Parent.UID_M0002
  .txtUpdtUserId.value = Parent.gUsrID
  .txtInsrtUserId.value = Parent.gUsrID
    
  '-----------------------
  'Data manipulate area
  '-----------------------
  lGrpCnt = 0

  strVal = ""
  '-----------------------
  'Data manipulate area
  '-----------------------
  For lRow = 1 To .vspdData.MaxRows
      .vspdData.Row = lRow
      .vspdData.Col = 0

      Select Case .vspdData.Text
          Case ggoSpread.InsertFlag       '☜: 신규 
     strVal = strVal & "C" & Parent.gColSep '☜: C=Create
          Case ggoSpread.UpdateFlag       '☜: 수정 
     strVal = strVal & "U" & Parent.gColSep '☜: U=Update               
    Case ggoSpread.DeleteFlag
     strVal = strVal & "D" & Parent.gColSep '☜: U=Delete                       
   End Select   
   
   
      Select Case .vspdData.Text
          Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag , ggoSpread.DeleteFlag  '☜: 수정, 신규,삭제 

              .vspdData.Col = C_PlanUnit
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

              .vspdData.Col = C_01PlanQty              
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_01PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
              
              .vspdData.Col = C_02PlanQty
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
              
              .vspdData.Col = C_02PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_03PlanQty
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_03PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_04PlanQty
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_04PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_05PlanQty
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_05PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_06PlanQty  
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_06PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_07PlanQty  
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_07PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_08PlanQty  
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_08PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_09PlanQty  
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_09PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
              
              .vspdData.Col = C_10PlanQty  
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_10PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_11PlanQty  
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_11PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_12PlanQty  
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_12PlanAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

              .vspdData.Col = C_ItemCode
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep  & lRow & Parent.gRowSep

              lGrpCnt = lGrpCnt + 1
              
      End Select
  Next

  .txtMaxRows.value = lGrpCnt
  .txtSpread.value =  strVal
 
  Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 
 
 End With
 
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
Function DbSaveOk()               <%'☆: 저장 성공후 실행 로직 %>

 With frm1
  .txtConSalesOrg.value = .txtSalesOrg.value 
  .txtConSalesOrgNm.value = .txtSalesOrgNm.value 
  .txtConSpYear.value = .txtSpYear.value 
  .txtConPlanTypeCd.value = .txtPlanTypeCd.value
  .txtConPlanTypeNm.value = .txtPlanTypeNm.value
  .txtConDealTypeCd.value = .txtDealTypeCd.value
  .txtConDealTypeNm.value = .txtDealTypeNm.value
  .txtConCurr.value = .txtCurr.value
  '.txtPlanNum.value = .txtConPlanNum.value 
 End With

 Call ggoOper.LockField(Document, "N")
    Call InitVariables
 frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
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
     <TD CLASS="CLSMTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTABP"><font color=white>고객별판매계획</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right><A href="vbscript:OpenPlanInfoRef">판매계획기초자료</A></TD>
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
         <TD CLASS="TD5" NOWRAP>영업조직</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConSalesOrg" ALT="영업조직" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSaleOrg 1">&nbsp;<INPUT NAME="txtConSalesOrgNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>계획년도</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConSpYear" ALT="계획년도" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12X1XU"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>계획구분</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConPlanTypeCd" ALT="계획구분" TYPE="Text" MAXLENGTH=1 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanType 1">&nbsp;<INPUT NAME="txtConPlanTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>거래구분</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConDealTypeCd" ALT="거래구분" TYPE="Text" MAXLENGTH=1 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDealType 1">&nbsp;<INPUT NAME="txtConDealTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>계획차수</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConPlanNum" ALT="계획차수" TYPE="Text" MAXLENGTH=2 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanNumber 1">&nbsp;<INPUT NAME="txtConPlanNumNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>화폐</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConCurr" ALT="화폐" TYPE="Text"  MAXLENGTH=3 SiZE=10 tag="14XXXU"></TD>
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
        <TD CLASS="TD5" NOWRAP>영업조직</TD>
        <TD CLASS="TD6"><INPUT NAME="txtSalesOrg" ALT="영업조직" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSaleOrg 2">&nbsp;<INPUT NAME="txtSalesOrgNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
        <TD CLASS="TD5" NOWRAP>계획년도</TD>
        <TD CLASS="TD6"><INPUT NAME="txtSpYear" ALT="계획년도" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="23X1XU"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>계획구분</TD>
        <TD CLASS="TD6"><INPUT NAME="txtPlanTypeCd" ALT="계획구분" TYPE="Text" MAXLENGTH=1 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanType 2">&nbsp;<INPUT NAME="txtPlanTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
        <TD CLASS="TD5" NOWRAP>거래구분</TD>
        <TD CLASS="TD6"><INPUT NAME="txtDealTypeCd" ALT="거래구분" TYPE="Text" MAXLENGTH=1 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDealType 2">&nbsp;<INPUT NAME="txtDealTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>계획차수</TD>
        <TD CLASS="TD6"><INPUT NAME="txtPlanNum" ALT="계획차수" TYPE="Text" MAXLENGTH=2 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanNumber 2">&nbsp;<INPUT NAME="txtPlanNumNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
        <TD CLASS="TD5" NOWRAP>화폐</TD>
        <TD CLASS="TD6"><INPUT NAME="txtCurr" ALT="화폐" TYPE="Text" MAXLENGTH=3  SiZE=10 tag="24XXXU"></TD>
       </TR>
       <TR>
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <script language =javascript src='./js/s2111ma4_I683853173_vspdData.js'></script>
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
 <TR HEIGHT=20>
  <TD WIDTH=100%>
   <TABLE <%=LR_SPACE_TYPE_30%>>
       <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD><BUTTON NAME="btnSplit" CLASS="CLSMBTN">확정처리</BUTTON>
     </TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0   TABINDEX = -1  ></IFRAME></TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"  TABINDEX = -1 >
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtRdoSelect" tag="14"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="txtSelectChr" tag="14"  TABINDEX = -1 >

<INPUT TYPE=HIDDEN NAME="HConSalesOrg" tag="24"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="HConSpYear" tag="24"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="HPlanTypeCd" tag="24"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="HConDealTypeCd" tag="24"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="HConCurr" tag="24"  TABINDEX = -1 >
<INPUT TYPE=HIDDEN NAME="HConPlanNum" tag="24"  TABINDEX = -1 >

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
