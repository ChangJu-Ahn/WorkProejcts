<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : inventory Management
'*  2. Function Name        : 
'*  3. Program ID           : i1812qa1_KO441
'*  4. Program Name         : 창고별 수불현황 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2007/07/31
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Ho Jun
'* 10. Modifier (Last)      : 
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

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit														'☜: indicates that All variables must be declared in advance

<%
EndDate   = GetSvrDate
%>

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop
'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID		= "i1812qb1_KO441.asp"                 '☆: 비지니스 로직 ASP명 

Dim C_SlCd			'창고

Dim C_ItemAcctCd	'품목계정코드
Dim C_ItemAcctNm	'품목계정명

Dim C_ItemCd		'품목
Dim C_ItemNm		'품목명
Dim C_BasicUnit		'단위
Dim C_Price			'단가
Dim C_TransStockQty	'이월재고
Dim C_TransStockAmt	'이월재고금액
Dim C_InProdQty		'생산입고
Dim C_InProdAmt		'생산입고금액
Dim C_InPurQty		'구매입고
Dim C_InPurAmt		'구매입고금액
Dim C_InExecQty		'예외입고
Dim C_InExecAmt		'예외입고금액
Dim C_InStockQty	'재고이동입고
Dim C_InStockAmt	'재고이동입고금액
Dim C_InSumQty		'입고계
Dim C_InSumAmt		'입고계금액
Dim C_OutProdQty	'생산출고
Dim C_OutProdAmt	'생산출고금액
Dim C_OutPurQty		'판매출고
Dim C_OutPurAmt		'판매출고금액
Dim C_OutExecQty	'예외출고
Dim C_OutExecAmt	'예외출고금액
Dim C_OutStockQty	'재고이동출고
Dim C_OutStockAmt	'재고이동출고금액
Dim C_OutSumQty		'출고계
Dim C_OutSumAmt		'출고계(금액)
Dim C_StockQty		'재고량
Dim C_StockAmt		'재고량

'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
                                         '☆: 초기화면에 뿌려지는 시작 날짜 -----
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 

'==========================================  InitComboBox()  ======================================
'	Name : InitComboBox()
'	Description : Init ComboBox
'==================================================================================================
Sub InitComboBox()

	
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
	lgBlnFlgChgValue = False
	IsOpenPop = False
End Sub                          

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()

	Dim StartDate
	StartDate = UNIDateAdd("m", -1,"<%=EndDate%>", parent.gServerDateFormat)
	
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtItemCd.focus 
			Set gActiveElement = document.activeElement 
		Else
			frm1.txtPlantCd.focus 
			Set gActiveElement = document.activeElement 
		End If
	End If

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
  	frm1.txtPlantCd.value = lgPLCd
	End If

	'frm1.txtMovFrDt.Text = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtMovFrDt.Text = UniConvDateAToB("<%=EndDate%>", parent.gServerDateFormat, parent.gDateFormat) 
	frm1.txtMovToDt.Text = UniConvDateAToB("<%=EndDate%>", parent.gServerDateFormat, parent.gDateFormat) 
 	Call ggoOper.FormatDate(frm1.txtMovFrDt, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtMovToDt, parent.gDateFormat, 2)

End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","QA") %>
End Sub


'--------------------------------------------------------------------------------------------------------- 
'	Name : OpenPlant()
'	Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtPlantCd.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "공장코드"		
	arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	

End Function

 '------------------------------------------  OpenItemAcct()  --------------------------------------------------
'	Name : OpenItemAcct()
'	Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목계정 팝업"											' 팝업 명칭 
	arrParam(1) = "B_MINOR"																' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtItemAcct.Value)						' Code Condition
	arrParam(3) = ""																			' Name Cindition
	arrParam(4) = "MAJOR_CD = 'P1001'"										' Where Condition
	arrParam(5) = "품목계정"			
	
	arrField(0) = "MINOR_CD"															' Field명(0)
	arrField(1) = "MINOR_NM"															' Field명(1)
	
	arrHeader(0) = "품목계정코드"															' Header명(0)
	arrHeader(1) = "품목계정명"																' Header명(1)
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemAcct(arrRet)
	End If	
	
End Function

 '------------------------------------------  OpenItemCd()  --------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item Cd
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")   <% '공장정보가 필요합니다 %>
		Exit Function
	End If
	
	'품목계정코드가 있는 지 체크 
	'If Trim(frm1.txtItemAcct.Value) = "" then
	'	Call DisplayMsgBox("169953","X","X","X")  
	'	Exit Function
	'End If              'update 2002/08/08 lsw

	'공장코드 및 품목계정코드 체크 함수 호출 
	'If Plant_Or_ItemAcct_Check = False Then 
	'	Exit Function
	'End If

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목 팝업"											' 팝업 명칭 
	arrParam(1) = "B_ITEM_BY_PLANT P,B_ITEM I"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtItemCd.Value)				        		' Code Condition
	arrParam(3) = ""																			' Name Cindition
	'arrParam(4) = "P.ITEM_CD=I.ITEM_CD AND P.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value,"","S") & " AND P.ITEM_ACCT=" & FilterVar(frm1.txtItemAcct.Value,"","S") ' Where Condition
	arrParam(4) = "P.ITEM_CD=I.ITEM_CD AND P.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value,"","S") ' Where Condition
	arrParam(5) = "품목"			
	
	arrField(0) = "I.ITEM_CD"															' Field명(0)
	arrField(1) = "I.ITEM_NM"															' Field명(1)
	
	arrHeader(0) = "품목코드"															' Header명(0)
	arrHeader(1) = "품목명"																' Header명(1)
	
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If	
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
'	lgBlnFlgChgValue	  	 = True	
End Function

 '------------------------------------------  SetItemAcct()  --------------------------------------------------
'	Name : SetItemAcct()
'	Description : ItemAcct Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemAcct(byval arrRet)
	frm1.txtItemAcct.Value	    =arrRet(0)
	frm1.txtItemAcctNm.Value	=arrRet(1)
	frm1.txtItemAcct.focus
	Set gActiveElement = document.activeElement
End Function

 '------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : ItemAcct Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value	=arrRet(0)
	frm1.txtItemNm.Value	=arrRet(1)
	frm1.txtItemAcct.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenSL()  --------------------------------------------------
'	Name : OpenSL()
'	Description : SL Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

	'-----------------------
	'Check Plant CODE	
	'-----------------------
	'If	CommonQueryRs("	PLANT_NM "," B_PLANT ",	" PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value,"","S"), _
	'	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	= False	Then
		
	'	Call DisplayMsgBox("125000","X","X","X")
	'	frm1.txtPlantNm.Value = ""
	'	frm1.txtPlantCd.focus
	'	Exit function
	'End	If
	'lgF0 = Split(lgF0, Chr(11))
	'frm1.txtPlantNm.Value = lgF0(0)

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	if frm1.txtPlantCd.value <> "" then
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value,"","S")	
	else
	arrParam(4) = ""
	end if
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSLCd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
End Function


'------------------------------------------  SetSL()  --------------------------------------------------
'	Name : SetSL()
'	Description : SL Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSL(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)
	frm1.txtSLCd.focus
'	lgBlnFlgChgValue	  = True
End Function

'========================================================================================
' Function Name : Plant_Or_ItemAcct_Check
' Function Desc : 
'========================================================================================
Function Plant_Or_ItemAcct_Check()
	'-----------------------
	'Check Plant CODE		'공장코드가 있는 지 체크 
	'-----------------------
    'If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value,"","S"), _
	'	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	'	
	'	Call DisplayMsgBox("125000","X","X","X")
	'	frm1.txtPlantNm.Value = ""
	'	frm1.txtPlantCd.focus
	'	Plant_Or_ItemAcct_Check = False
	'	Exit function
	'	
    'End If
	'lgF0 = Split(lgF0, Chr(11))
	'frm1.txtPlantNm.Value = lgF0(0)

	'-----------------------
	'Check ItemAcct CODE	''품목계정코드가 있는 지 체크 
	'-----------------------
    If 	CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = 'P1001' AND MINOR_CD= " & FilterVar(frm1.txtItemAcct.Value,"","S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("169952","X","X","X")
		frm1.txtItemAcctNm.Value = ""
		frm1.txtItemAcct.focus
		Plant_Or_ItemAcct_Check = False
		Exit function
    End If
    lgF0 = Split(lgF0, Chr(11))
	frm1.txtItemAcctNm.Value = lgF0(0)

	'-----------------------
	'Check Item CODE		'품목명 조회 
	'-----------------------
    If frm1.txtItemCd.Value <> "" Then
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value,"","S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm.Value = lgF0(0)
		Else
			frm1.txtItemNm.Value = ""
		End If
	End If

    Plant_Or_ItemAcct_Check = True
End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()

	on error resume next
	
	Dim Ret
	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030804", , Parent.gAllowDragDropSpread

	With frm1.vspdData
	
		.ReDraw = false
		
		.MaxCols = C_StockAmt + 1
		.ColHeaderRows = 2    '헤더를 2줄로
		.Col = .MaxCols
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")

'2008-05-22 1:06오후 :: hanc :: 12->15 &  	ggAmtOfMoneyNo->ggQtyNo 	
 		ggoSpread.SSSetEdit	  C_SlCd,			"창고코드",		10
 		ggoSpread.SSSetEdit	  C_ItemAcctCd,		"품목계정",		10
 		ggoSpread.SSSetEdit	  C_ItemAcctNm,		"품목계정명",	10
		ggoSpread.SSSetEdit	  C_ItemCd,			"품목코드",		15
		ggoSpread.SSSetEdit   C_ItemNm,			"품목명",		20
		ggoSpread.SSSetEdit   C_BasicUnit,		"단위",			8
		ggoSpread.SSSetFloat  C_Price,			"기준단가",		10, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_TransStockQty,	"이월재고",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_TransStockAmt,	"이월재고금액",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InProdQty,		"생산입고수량",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InProdAmt,		"생산입고금액",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InPurQty,		"구매입고수량",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InPurAmt,		"구매입고금액",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InExecQty,		"예외입고(입고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InExecAmt,		"예외입고금액(입고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InStockQty,		"재고이동(입고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InStockAmt,		"재고이동금액(입고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InSumQty,		"계(입고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InSumAmt,		"계금액(입고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutProdQty,		"생산출고(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutProdAmt,		"생산출고금액(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutPurQty,		"판매출고(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutPurAmt,		"판매출고금액(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutExecQty,		"예외출고(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutExecAmt,		"예외출고금액(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutStockQty,	"재고이동(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutStockAmt,	"재고이동금액(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutSumQty,		"계(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutSumAmt,		"계금액(출고)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_StockQty,		"재고량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_StockAmt,		"재고금액",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		
'		ggoSpread.SSSetFloat  C_Price,			"기준단가",		10, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_TransStockQty,	"이월재고",		12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_TransStockAmt,	"이월재고금액",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InProdQty,		"생산입고수량",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InProdAmt,		"생산입고금액",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InPurQty,		"구매입고수량",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InPurAmt,		"구매입고금액",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InExecQty,		"예외입고(입고)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InExecAmt,		"예외입고금액(입고)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InStockQty,		"재고이동(입고)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InStockAmt,		"재고이동금액(입고)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InSumQty,		"계(입고)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InSumAmt,		"계금액(입고)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutProdQty,		"생산출고(출고)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutProdAmt,		"생산출고금액(출고)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutPurQty,		"판매출고(출고)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutPurAmt,		"판매출고금액(출고)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutExecQty,		"예외출고(출고)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutExecAmt,		"예외출고금액(출고)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutStockQty,	"재고이동(출고)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutStockAmt,	"재고이동금액(출고)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutSumQty,		"계(출고)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutSumAmt,		"계금액(출고)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_StockQty,		"재고량",		12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
		
		
		' 그리드 헤더 합침 
		ret = .AddCellSpan(C_SlCd			, -1000, 1, 2)
		ret = .AddCellSpan(C_ItemAcctCd		, -1000, 1, 2)
		ret = .AddCellSpan(C_ItemAcctNm		, -1000, 1, 2)
		ret = .AddCellSpan(C_ItemCd			, -1000, 1, 2)
		ret = .AddCellSpan(C_ItemNm			, -1000, 1, 2)
		ret = .AddCellSpan(C_BasicUnit		, -1000, 1, 2)
		ret = .AddCellSpan(C_Price			, -1000, 1, 2)
		ret = .AddCellSpan(C_TransStockQty	, -1000, 1, 2)
		ret = .AddCellSpan(C_TransStockAmt	, -1000, 1, 2)
		
		ret = .AddCellSpan(C_InProdQty		, -1000, 10, 1)
		ret = .AddCellSpan(C_OutProdQty		, -1000, 10, 1)
		ret = .AddCellSpan(C_StockQty		, -1000, 2, 1)
			
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_SlCd			: .Text = "창고코드"
		.Col = C_ItemAcctCd		: .Text = "품목계정"
		.Col = C_ItemAcctNm		: .Text = "품목계정명"
		.Col = C_ItemCd			: .Text = "품목코드"
		.Col = C_ItemNm			: .Text = "품목명"
		.Col = C_BasicUnit		: .Text = "단위"
		.Col = C_Price			: .Text = "기준단가"
		.Col = C_TransStockQty	: .Text = "이월재고"
		.Col = C_TransStockAmt	: .Text = "이월재고금액"
		.Col = C_InProdQty		: .Text = "입고"
		.Col = C_OutProdQty		: .Text = "출고"
		.Col = C_StockQty		: .Text = "재고"
			
		.Row = -999
		
		.Col = C_SlCd			: .Text = "창고코드"
		.Col = C_ItemAcctCd		: .Text = "품목계정"
		.Col = C_ItemAcctNm		: .Text = "품목계정명"
		.Col = C_ItemCd			: .Text = "품목코드"
		.Col = C_ItemNm			: .Text = "품목명"
		.Col = C_BasicUnit		: .Text = "단위"
		.Col = C_Price			: .Text = "기준단가"	
		.Col = C_TransStockQty	: .Text = "이월재고"
		.Col = C_TransStockAmt	: .Text = "이월재고금액"
		.Col = C_InProdQty		: .Text = "생산입고수량"
		.Col = C_InProdAmt		: .Text = "생산입고금액"
		.Col = C_InPurQty		: .Text = "구매입고수량"
		.Col = C_InPurAmt		: .Text = "구매입고금액"
		.Col = C_InExecQty		: .Text = "예외입고"
		.Col = C_InExecAmt		: .Text = "예외입고금액"
		.Col = C_InStockQty		: .Text = "재고이동"
		.Col = C_InStockAmt		: .Text = "재고이동금액"
		.Col = C_InSumQty		: .Text = "계"
		.Col = C_InSumAmt		: .Text = "계금액"
		.Col = C_OutProdQty		: .Text = "생산출고"
		.Col = C_OutProdAmt		: .Text = "생산출고금액"
		.Col = C_OutPurQty		: .Text = "판매출고" 				
		.Col = C_OutPurAmt		: .Text = "판매출고금액" 					
		.Col = C_OutExecQty		: .Text = "예외출고" 					
		.Col = C_OutExecAmt		: .Text = "예외출고금액" 					
		.Col = C_OutStockQty	: .Text = "재고이동" 			
		.Col = C_OutStockAmt	: .Text = "재고이동금액"
		.Col = C_OutSumQty		: .Text = "계"
		.Col = C_OutSumAmt		: .Text = "계금액"
		.Col = C_StockQty		: .Text = "재고량"
		.Col = C_StockAmt		: .Text = "재고금액"
		
		.RowHeight(-999) = 15	' 높이 재지정	(2줄일 경우, 1줄은 15)	
 	
 		Call ggoSpread.SSSetColHidden(C_Price, C_Price, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
	    'ggoSpread.SSSetSplit2(2)  
	    'Call SetSpreadLock()
		
		.ReDraw = true
		
    End With
End Sub

Sub SetSpreadLock()
   
End Sub

'==========================================  2.6.1 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()

	C_SlCd			= 1			'창고
	C_ItemAcctCd	= 2			'품목계정
	C_ItemAcctNm	= 3			'품목계정명
	C_ItemCd		= 4			'품목
	C_ItemNm		= 5			'품목명
	C_BasicUnit		= 6			'단위
	C_Price			= 7			'단가
	C_TransStockQty	= 8			'이월재고
	C_TransStockAmt	= 9			'이월재고금액
	C_InProdQty		= 10		'생산입고
	C_InProdAmt		= 11		'생산입고금액
	C_InPurQty		= 12		'구매입고
	C_InPurAmt		= 13		'구매입고금액
	C_InExecQty		= 14		'예외입고
	C_InExecAmt		= 15		'예외입고금액
	C_InStockQty	= 16		'재고이동입고
	C_InStockAmt	= 17		'재고이동입고금액
	C_InSumQty		= 18		'계입고
	C_InSumAmt		= 19		'계입고금액
	C_OutProdQty	= 20		'생산출고
	C_OutProdAmt	= 21		'생산출고금액
	C_OutPurQty		= 22		'판매출고
	C_OutPurAmt		= 23		'판매출고금액
	C_OutExecQty	= 24		'예외출고
	C_OutExecAmt	= 25		'예외출고금액
	C_OutStockQty	= 26		'재고이동출고
	C_OutStockAmt	= 27		'재고이동출고금액
	C_OutSumQty		= 28		'계출고
	C_OutSumAmt		= 29		'계출고금액
	C_StockQty		= 30		'재고량
	C_StockAmt		= 31		'재고금액

End Sub

'==========================================  2.6.2 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_SlCd			= iCurColumnPos(1)		'창고코드
		C_ItemAcctCd	= iCurColumnPos(2)		'품목계정
		C_ItemAcctNm	= iCurColumnPos(3)		'품목계정명
		C_ItemCd		= iCurColumnPos(4)		'품목
		C_ItemNm		= iCurColumnPos(5)		'품목명
		C_BasicUnit		= iCurColumnPos(6)		'단위
		C_Price			= iCurColumnPos(7)		'단가
		C_TransStockQty	= iCurColumnPos(8)		'이월재고
		C_TransStockAmt	= iCurColumnPos(9)		'이월재고금액
		C_InProdQty		= iCurColumnPos(10)		'생산입고
		C_InProdAmt		= iCurColumnPos(11)		'생산입고금액
		C_InPurQty		= iCurColumnPos(12)		'구매입고
		C_InPurAmt		= iCurColumnPos(13)		'구매입고금액
		C_InExecQty		= iCurColumnPos(14)		'예외입고
		C_InExecAmt		= iCurColumnPos(15)		'예외입고금액
		C_InStockQty	= iCurColumnPos(16)		'재고이동입고
		C_InStockAmt	= iCurColumnPos(17)		'재고이동입고금액
		C_InSumQty		= iCurColumnPos(18)		'계입고
		C_InSumAmt		= iCurColumnPos(19)		'계입고금액
		C_OutProdQty	= iCurColumnPos(20)		'생산출고
		C_OutProdAmt	= iCurColumnPos(21)		'생산출고금액
		C_OutPurQty		= iCurColumnPos(22)		'판매출고
		C_OutPurAmt		= iCurColumnPos(23)		'판매출고금액
		C_OutExecQty	= iCurColumnPos(24)		'예외출고
		C_OutExecAmt	= iCurColumnPos(25)		'예외출고금액
		C_OutStockQty	= iCurColumnPos(26)		'재고이동출고
		C_OutStockAmt	= iCurColumnPos(27)		'재고이동출고금액
		C_OutSumQty		= iCurColumnPos(28)		'계출고
		C_OutSumAmt		= iCurColumnPos(29)		'계출고금액
		C_StockQty		= iCurColumnPos(30)		'재고량
		C_StockAmt		= iCurColumnPos(31)		'재고금액
				
 	End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	
	Call InitVariables														'⊙: Initializes local global variables
  Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitComboBox()
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어	
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
   	
	frm1.txtIssueReqNo.focus
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
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
    Call InitSpreadSheet
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



'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function								'⊙: This function check indispensable field
    
   	
    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData 

	'If Name_check("A") = False Then
	'	Set gActiveElement = document.activeElement
	'	Exit Function
	'End If
	
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
	Dim strVal
	Dim rdoIssue
	Dim rdoConfirm
	Dim strYear, strMonth, strDay, strYear1, strMonth1, strDay1
    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
	
	Call ExtractDateFrom(frm1.txtMovFrDt.Text,frm1.txtMovFrDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
  	Call ExtractDateFrom(frm1.txtMovToDt.Text,frm1.txtMovToDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtPlantCd="		& Trim(.txtPlantCd.value) & _							
							  "&txtSlCd=" & Trim(.txtSLCd.value) & _
							  "&txtFromYY=" & Trim(strYear) & _
							  "&txtFromMm=" & Trim(strMonth) & _
							  "&txtToYY=" & Trim(strYear1) & _
							  "&txtToMm=" & Trim(strMonth1) & _		
							  "&txtItemAcct=" & Trim(.txtItemAcct.value) & _	
							  "&txtItemCd=" & Trim(.txtItemCd.value) & _		
							  "&txtMaxRows="	& .vspdData.MaxRows & _
							  "&lgStrPrevKey="	& lgStrPrevKey                      '☜: Next key tag
					  
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
    Call SetToolbar("11000000000111")
	lgBlnFlgChgValue = False
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview()
    Dim strYear, strMonth, strDay, strYear1, strMonth1, strDay1
	Dim var1, var2, var3, var4, var5, var6, var7, var8
    Dim condvar 
 	Dim ObjName
   
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then                             '⊙: Check contents area
       Exit Function
    End If
    
    '공장코드 및 품목계정코드 체크 함수 호출 
    'If Plant_Or_ItemAcct_Check = False Then 
	'	Exit Function
	'End If
	
    If ValidDateCheck(frm1.txtMovFrDt, frm1.txtMovToDt) = False Then
		Exit Function
    End If
	
 	Call ExtractDateFrom(frm1.txtMovFrDt.Text,frm1.txtMovFrDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
  	Call ExtractDateFrom(frm1.txtMovToDt.Text,frm1.txtMovToDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
    
    var1 = UCASE(Trim(frm1.txtPlantCd.value))
    If var1 = "" Then
       var1 = "%"
    End If
    var2 = Trim(frm1.txtItemCd.Value)
    var3 = frm1.txtItemAcct.Value
    var4 = strYear
    var5 = strMonth
    var6 = strYear1
    var7 = strMonth1
    var8 = Trim(frm1.txtSLCd.value)
    
	If var2 = "" Then
       var2 = "%"
    End If
    
	If var3 = "" Then
       var3 = "%"
    End If    
    
	If var8 = "" Then
       var8 = "%"
    End If

	ObjName = AskEBDocumentName("I1812OA2_KO441", "ebr")				'창고별
		
	condvar = condvar & "PLANTCD|"    & var1
	condvar = condvar & "|SLCD|"      & var8
    condvar = condvar & "|ITEMCD|"    & var2
    condvar = condvar & "|ITEMACCT|"  & var3
    condvar = condvar & "|FromYy|"    & var4
    condvar = condvar & "|FromMm|"    & var5
    condvar = condvar & "|ToYy|"      & var6
    condvar = condvar & "|ToMm|"      & var7

	Call FncEBRPreview(ObjName, condvar)    
	
End Function

'========================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function FncBtnPrint() 
    Dim strYear, strMonth, strDay, strYear1, strMonth1, strDay1
	Dim var1, var2, var3, var4, var5, var6, var7, var8
    Dim condvar 
 	Dim ObjName
   
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then                             '⊙: Check contents area
       Exit Function
    End If
    
    '공장코드 및 품목계정코드 체크 함수 호출 
    'If Plant_Or_ItemAcct_Check = False Then 
	'	Exit Function
	'End If
	
    If ValidDateCheck(frm1.txtMovFrDt, frm1.txtMovToDt) = False Then
		Exit Function
    End If
	
 	Call ExtractDateFrom(frm1.txtMovFrDt.Text,frm1.txtMovFrDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
  	Call ExtractDateFrom(frm1.txtMovToDt.Text,frm1.txtMovToDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
    
    var1 = UCASE(Trim(frm1.txtPlantCd.value))
    If var1 = "" Then
       var1 = "%"
    End If
    var2 = Trim(frm1.txtItemCd.Value)
    var3 = frm1.txtItemAcct.Value
    var4 = strYear
    var5 = strMonth
    var6 = strYear1
    var7 = strMonth1
    var8 = Trim(frm1.txtSLCd.value)
    
	If var2 = "" Then
       var2 = "%"
    End If
    
    If var3 = "" Then
       var3 = "%"
    End If  
    
	If var8 = "" Then
       var8 = "%"
    End If
     
    ObjName = AskEBDocumentName("I1812OA1_KO441", "ebr")				'창고별
    
	condvar = condvar & "PLANTCD|"    & var1
	condvar = condvar & "|SLCD|"      & var8
	condvar = condvar & "|PLANTCD|"   & var1
    condvar = condvar & "|ITEMCD|"    & var2
    condvar = condvar & "|ITEMACCT|"  & var3
    condvar = condvar & "|FromYy|"    & var4
    condvar = condvar & "|FromMm|"    & var5
    condvar = condvar & "|ToYy|"      & var6
    condvar = condvar & "|ToMm|"      & var7
    
	Call FncEBRprint(EBAction, ObjName, condvar) 	

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>창고별수불현황조회(S)</font></td>
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
									<TD CLASS="TD6">
									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=20 tag="14">
									</TD>
									<TD CLASS="TD5" NOWRAP>수불기간</TD>
									<TD CLASS="TD6">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMM name=txtMovFrDt classid=<%=gCLSIDFPDT%> tag="12x1" ALT="검색시작날짜" VIEWASTEXT><PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""></OBJECT>');</Script>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMM name=txtMovToDt classid=<%=gCLSIDFPDT%> ALT="검색끝날짜" tag="12x1" VIEWASTEXT><PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""></OBJECT>');</Script>
									</TD>
								</TR>
								<TR>									
        							<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6">
									<input TYPE=TEXT NAME="txtItemAcct" SIZE="8" MAXLENGTH="2" tag="11XXXU" ALT="품목계정"  ><IMG align=top height=20 name="btnItemAcct" onclick="vbscript:OpenItemAcct()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 MAXLENGTH=40 tag="14">
									</TD>
									<TD CLASS="TD5">품목</TD>
									<TD CLASS="TD6">
									<input TYPE=TEXT NAME="txtItemCd" SIZE="18" MAXLENGTH="18" ALT="품목" tag="11XXXU" ><IMG align=top height=20 name="btnItemNm" onclick="vbscript:OpenItemCd()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 MAXLENGTH=40 tag="14">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>창고</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE="TEXT" NAME="txtSLCd" SIZE=8 MAXLENGTH=7 tag="11XXXU" ALT="창고" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE="TEXT" NAME="txtSLNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="창고명">
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><TD>
									</TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TR>
							<TD HEIGHT=100% WIDTH=100% Colspan=2>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>	
						</TR>	
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreView()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
    </DIV>
	<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname">
	<input type="hidden" name="dbname">
	<input type="hidden" name="filename">
	<input type="hidden" name="condvar">
	<input type="hidden" name="date">                 
	</FORM>
</BODY>
</HTML>
