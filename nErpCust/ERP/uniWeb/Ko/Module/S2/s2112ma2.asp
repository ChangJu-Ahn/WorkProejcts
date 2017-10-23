<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2112MA2
'*  4. Program Name         : 확정품목별 판매계획 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS2G131.dll, PS2G132.dll, PS2G135.dll
'*  7. Modified date(First) : 2000/03/24
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Mr Cho 
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/03/24 : 3rd 기능구현 및 화면디자인 
'*                            -2000/05/09 : 3rd 표준수정사항 
'*                            -2000/08/10 : 4th 화면 Layout 수정 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                        '☜: Turn on the Option Explicit option.

'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================
Dim C_ItemCode
Dim C_ItemPopup
Dim C_ItemName
Dim C_Spec
Dim C_PlanUnit
Dim C_PlanUnitPopup
Dim C_YearQty

Dim C_01PlanQty
Dim C_02PlanQty
Dim C_03PlanQty
Dim C_04PlanQty
Dim C_05PlanQty
Dim C_06PlanQty
Dim C_07PlanQty
Dim C_08PlanQty
Dim C_09PlanQty
Dim C_10PlanQty
Dim C_11PlanQty
Dim C_12PlanQty

Dim C_01PlanColor
Dim C_02PlanColor
Dim C_03PlanColor
Dim C_04PlanColor
Dim C_05PlanColor
Dim C_06PlanColor
Dim C_07PlanColor
Dim C_08PlanColor
Dim C_09PlanColor
Dim C_10PlanColor
Dim C_11PlanColor
Dim C_12PlanColor

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop		'Popup
Dim lsItemCode      '품목 
Dim lsItemName      '품목명 
Dim lsPlanMonth     '계획월 
Dim lsPlanUnit      '계힉단위 

Dim prDBSYSDate
Dim EndDate ,StartDate

prDBSYSDate = "<%=GetSvrDate%>"

EndDate = UniConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s2112mb2.asp"        '☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "s2112ma1"
Const C_SHEETMAXROWS = 30   'Sheet Max Rows
Const lsSPLIT  = "SPLIT"    <% '공장별배분작업 %>

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_ItemCode = 1           '☆: Spread Sheet의 Column별 상수 
	C_ItemPopup = 2
	C_ItemName = 3
	C_Spec		= 4
	C_PlanUnit = 5
	C_PlanUnitPopup = 6
	C_YearQty = 7

	C_01PlanQty = 8
	C_02PlanQty = 9
	C_03PlanQty = 10
	C_04PlanQty = 11
	C_05PlanQty = 12
	C_06PlanQty = 13
	C_07PlanQty = 14
	C_08PlanQty = 15
	C_09PlanQty = 16
	C_10PlanQty = 17
	C_11PlanQty = 18
	C_12PlanQty = 19

	C_01PlanColor = 20
	C_02PlanColor = 21
	C_03PlanColor = 22
	C_04PlanColor = 23
	C_05PlanColor = 24
	C_06PlanColor = 25
	C_07PlanColor = 26
	C_08PlanColor = 27
	C_09PlanColor = 28
	C_10PlanColor = 29
	C_11PlanColor = 30
	C_12PlanColor = 31

End Sub

'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    'lgLngCurRows = 0      
	lgSortKey = 1

End Sub

'========================================================================================================= 
Sub SetDefaultVal()

	With frm1
		
		.txtConSpYear.focus
		.txtMode.value = ""
		.btnSplit.disabled = True 
		.txtConSpYear.value = Year(UniConvDateToYYYYMMDD(EndDate,parent.gDateFormat,parent.gServerDateType))
		lgBlnFlgChgValue = False

	End With

End Sub

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %> 
End Sub

'========================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()
	
	With frm1.vspdData
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.Spreadinit "V20021120",,parent.gAllowDragDropSpread    
	
	.ReDraw = false
	  
	.MaxCols = C_12PlanColor+1             '☜: 최대 Columns의 항상 1개 증가시킴 
	.MaxRows = 0

	 Call GetSpreadColumnPos("A")		
	
	ggoSpread.SSSetEdit C_ItemCode, "품목", 20,,,18,2
	ggoSpread.SSSetButton C_ItemPopup
	ggoSpread.SSSetEdit C_ItemName, "품목명", 30
	ggoSpread.SSSetEdit C_Spec,		"품목규격",20,0
	ggoSpread.SSSetEdit C_PlanUnit, "계획단위", 10,,,3,2
	ggoSpread.SSSetButton C_PlanUnitPopup

	ggoSpread.SSSetFloat C_YearQty,"년 계획량 합계" ,20,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_01PlanQty,"1월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_02PlanQty,"2월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_03PlanQty,"3월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_04PlanQty,"4월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_05PlanQty,"5월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_06PlanQty,"6월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_07PlanQty,"7월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_08PlanQty,"8월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_09PlanQty,"9월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_10PlanQty,"10월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_11PlanQty,"11월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_12PlanQty,"12월계획량" ,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"        

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
	 
	Call ggoSpread.SSSetColHidden(C_ItemPopup,C_ItemPopup,True)
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
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
	.ReDraw = true
		    
	End With
    
End Sub


'======================================================================================================
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False

    ggoSpread.SSSetRequired C_ItemCode, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Spec, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_PlanUnit, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_YearQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_01PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_02PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_03PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_04PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_05PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_06PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_07PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_08PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_09PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_10PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_11PlanQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_12PlanQty, pvStartRow, pvEndRow

    .vspdData.ReDraw = True
    
    End With

End Sub

Sub SetQuerySpreadColor()
	
	Dim lRow
    
    With frm1

    .vspdData.ReDraw = False

	ggoSpread.source = frm1.vspdData
	
	
		For lRow = 1 To .vspdData.MaxRows 
			ggoSpread.SSSetProtected C_ItemCode, lRow, lRow
			ggoSpread.SSSetProtected C_ItemName, lRow, lRow
			ggoSpread.SSSetProtected C_Spec, lRow, lRow			
			ggoSpread.SSSetProtected C_PlanUnit, lRow, lRow
			ggoSpread.SSSetProtected C_PlanUnitPopup, lRow, lRow    
			ggoSpread.SSSetProtected C_YearQty, lRow, lRow
						
			
			frm1.vspdData.Col = C_01PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_01PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_01PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_02PlanColor			
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_02PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_02PlanQty, lRow, lRow
			End if
						
			frm1.vspdData.Col = C_03PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_03PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_03PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_04PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_04PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_04PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_05PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_05PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_05PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_06PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_06PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_06PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_07PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_07PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_07PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_08PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_08PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_08PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_09PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_09PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_09PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_10PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_10PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_10PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_11PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_11PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_11PlanQty, lRow, lRow
			End if
			
			frm1.vspdData.Col = C_12PlanColor
			If UCase(frm1.vspdData.Text) = "N" then
				ggoSpread.SSSetRequired C_12PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_12PlanQty, lRow, lRow
			End if
								
		Next

    .vspdData.ReDraw = True
    
    End With

End Sub


<% '******************************************  2.4 POP-UP 처리함수  ****************************************
' 기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'       하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* %>

'======================================================================================================
Function OpenITEMPopup()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목"			<%' 팝업 명칭 %>
	arrParam(1) = "b_item item,b_item_by_plant item_plant" <%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtConItemCd.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "item.item_cd=item_plant.item_cd" <%' Where Condition%>
	arrParam(5) = "품목"						<%' TextBox 명칭 %>
	 
	arrField(0) = "item.item_cd"		<%' Field명(0)%>
	arrField(1) = "item.item_nm"		<%' Field명(1)%>
	arrField(2) = "item_plant.plant_cd"
	    
	arrHeader(0) = "품목"			<%' Header명(0)%>
	arrHeader(1) = "품목명"			<%' Header명(1)%>
	arrHeader(2) = "공장"

	frm1.txtConItemCd.focus 
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=520px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtConItemCd.value = arrRet(0)
		frm1.txtConItemNm.value = arrRet(1)
	End If
 
End Function

'======================================================================================================
Function OpenSalesPlanPopup(Byval strCode, Byval iWhere, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	Select Case iWhere
	
	Case C_ItemPopUp
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = 0
		If frm1.vspdData.Text <> ggoSpread.InsertFlag And lgIntFlgMode = parent.OPMD_UMODE Then Exit Function
	End Select

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	
	Case C_ItemPopUp     <%' 품목 %>
		arrParam(1) = "b_item item,b_item_by_plant item_plant"     <%' TABLE 명칭 %>
		arrParam(2) = strCode        <%' Code Condition%>
		arrParam(3) = ""         <%' Name Cindition%>
		arrParam(4) = "item.item_cd=item_plant.item_cd"  <%' Where Condition%>
		arrParam(5) = "품목"       <%' TextBox 명칭 %>
		 
		arrField(0) = "item.item_cd"      <%' Field명(0)%>
		arrField(1) = "item.item_nm"      <%' Field명(1)%>
		    
		arrHeader(0) = "품목"       <%' Header명(0)%>
		arrHeader(1) = "품목명"       <%' Header명(1)%>

	Case C_PlanUnitPopup    <%' 단위 %>
		arrParam(1) = "b_unit_of_measure"     <%' TABLE 명칭 %>
		arrParam(2) = strCode        <%' Code Condition%>
		arrParam(3) = ""         <%' Name Cindition%>
		arrParam(4) = ""         <%' Where Condition%>
		arrParam(5) = "단위"       <%' TextBox 명칭 %>
		 
		arrField(0) = "unit"        <%' Field명(0)%>
		arrField(1) = "unit_nm"        <%' Field명(1)%>
		    
		arrHeader(0) = "단위"       <%' Header명(0)%>
		arrHeader(1) = "단위명"       <%' Header명(1)%>

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


'======================================================================================================
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


<% '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ %>
<%
'======================================================================================================
' Function Desc : 수량 SpreadColor
'=============================================================================================================
%>
Function SetQtySpreadColor(ByVal lRow)

    'Dim MRow
    
    With frm1

    .vspdData.ReDraw = False

	ggoSpread.Source = .vspdData

    ggoSpread.SSSetRequired C_ItemCode, lRow, .vspdData.MaxRows
    ggoSpread.SSSetProtected C_ItemName, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_PlanUnit, lRow, .vspdData.MaxRows
    ggoSpread.SSSetProtected C_YearQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetProtected C_YearAmt, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_01PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_02PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_03PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_04PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_05PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_06PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_07PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_08PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_09PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_10PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_11PlanQty, lRow, .vspdData.MaxRows
    ggoSpread.SSSetRequired C_12PlanQty, lRow, .vspdData.MaxRows

    .vspdData.ReDraw = True
    
    End With

End Function

<%
'======================================================================================================
' Function Desc : 년 판매수량/금액의 합 
'=======================================================================================================
%>
Function MonthTotalSum()

	Dim SumTotal, iMonth, lRow

		ggoSpread.Source = frm1.vspdData 

		For lRow = 1 To frm1.vspdData.MaxRows 

			SumTotal = 0

			frm1.vspdData.Row = lRow

			frm1.vspdData.Col = C_01PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_02PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_03PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_04PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_05PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_06PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_07PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_08PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_09PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_10PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_11PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If

			frm1.vspdData.Col = C_12PlanQty
			If IsNumeric(UNICDbl(frm1.vspdData.Text)) = True Then
				SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			End If


			frm1.vspdData.Row = lRow
			frm1.vspdData.Col = C_YearQty
			frm1.vspdData.Text= UNIFormatNumber(SumTotal,ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
		Next

End Function

<%
'======================================================================================================
' Function Desc : Jump시 해당 화면에 조회값 인자/인수 전달 
'====================================================================================================
%>
Function CookiePage(KuBun)

	On Error Resume Next

	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>

	Dim strTemp, arrVal

	Select Case Kubun
	
	Case "ReadCookie"

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" then Exit Function

		frm1.txtConSpYear.value =  arrVal(0)
		frm1.txtConItemCd.value =  arrVal(1) 
		frm1.txtConItemNm.value =  arrVal(2)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function
		End If
	  
		Call MainQuery

		WriteCookie CookieSplit , ""


	Case "WriteCookie"

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie CookieSplit , frm1.txtConSpYear.value & parent.gRowSep & lsPlanMonth _ 
		& parent.gRowSep & lsItemCode & parent.gRowSep & lsItemName

	End Select
	 
End Function

<%
'======================================================================================================
' Function Desc : 월별 배분 판별 함수 
'=======================================================================================================
%>
Function MonthlyPlan()

	Dim IntRetCD
	ggoSpread.Source = frm1.vspdData 
	
	If ggoSpread.SSCheckChange = True Then
		
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	If frm1.vspdData.Row < 1 Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If

	With frm1.vspdData 
	.Row = .ActiveRow 

	Select Case .ActiveCol
		Case C_01PlanQty : .Col = C_01PlanColor 
		Case C_02PlanQty : .Col = C_02PlanColor 
		Case C_03PlanQty : .Col = C_03PlanColor 
		Case C_04PlanQty : .Col = C_04PlanColor 
		Case C_05PlanQty : .Col = C_05PlanColor 
		Case C_06PlanQty : .Col = C_06PlanColor 
		Case C_07PlanQty : .Col = C_07PlanColor 
		Case C_08PlanQty : .Col = C_08PlanColor 
		Case C_09PlanQty : .Col = C_09PlanColor
		Case C_10PlanQty : .Col = C_10PlanColor 
		Case C_11PlanQty : .Col = C_11PlanColor 
		Case C_12PlanQty : .Col = C_12PlanColor 
	Case Else
		MsgBox "조정할 품목의 계획월을 선택하세요.", vbExclamation, parent.gLogoName
		
		Exit Function
	
	End Select

	'  If .Text = "Y" Then
	Call CookiePage("WriteCookie")
	Call PgmJump(BIZ_PGM_JUMP_ID)
	'  ElseIf .Text = "N" Then
	'   Call DisplayMsgBox("202197","X","X","X")
	'   Exit Function
	'  End If

	End With

End Function


<%
'======================================================================================================
' Function Desc : After Spread Cell Click, Variables initializes
'=======================================================================================================
%>
Function SpreadCellClick(ByVal Col,ByVal Row)

	lsPlanMonth=""

	Select Case Col
		Case C_01PlanQty : lsPlanMonth = "01"
		Case C_02PlanQty : lsPlanMonth = "02"
		Case C_03PlanQty : lsPlanMonth = "03"
		Case C_04PlanQty : lsPlanMonth = "04"
		Case C_05PlanQty : lsPlanMonth = "05"
		Case C_06PlanQty : lsPlanMonth = "06"
		Case C_07PlanQty : lsPlanMonth = "07"
		Case C_08PlanQty : lsPlanMonth = "08"
		Case C_09PlanQty : lsPlanMonth = "09"
		Case C_10PlanQty : lsPlanMonth = "10"
		Case C_11PlanQty : lsPlanMonth = "11"
		Case C_12PlanQty : lsPlanMonth = "12"
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

	Call SpreadCellColor()

End Function

<%
'======================================================================================================
' Function Desc : After Spread Cell Click, Split Color Check
'=======================================================================================================
%>
Function SpreadCellColor()

	Dim i, strItem, cntItem
	
	With frm1.vspdData
		  
		strItem = "" : cntItem = 0
		
		For i = 1 To .MaxRows
		  
			.Row = i

			Select Case lsPlanMonth
				Case "01" : .Col = C_01PlanColor
				Case "02" : .Col = C_02PlanColor
				Case "03" : .Col = C_03PlanColor
				Case "04" : .Col = C_04PlanColor
				Case "05" : .Col = C_05PlanColor
				Case "06" : .Col = C_06PlanColor
				Case "07" : .Col = C_07PlanColor
				Case "08" : .Col = C_08PlanColor
				Case "09" : .Col = C_09PlanColor
				Case "10" : .Col = C_10PlanColor
				Case "11" : .Col = C_11PlanColor
				Case "12" : .Col = C_12PlanColor
			End Select

			If .Text = "N" Then
				.Col = C_ItemCode
				strItem = strItem & Trim(.Text) & parent.gColSep
				cntItem = cntItem + 1    
			End If

		Next
		  
		frm1.txtItemArrary.value = strItem
		frm1.txtItemCount.value = cntItem
	End With

End Function

<%
'======================================================================================================
' Function Desc : Before Batch Button , Requried Value Checking Msg
'=======================================================================================================
%>
Function BatchReqCheckMsg()

	BatchReqCheckMsg = False

	Call SpreadCellClick(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If lsPlanMonth = "" Then
		Call DisplayMsgBox("202151","X","X","X")
		'Msgbox "배분할 월을 클릭하세요"
		Exit Function
	End If

	If Len(Trim(frm1.txtItemCount.value)) = 0 Or CStr(Trim(frm1.txtItemCount.value)) = CStr(0) Then
		MsgBox "공장배분할 해당월의 품목수량이 없습니다", vbExclamation, parent.gLogoName
		Exit Function
	End If

	Dim IntRetCD
	
	ggoSpread.Source = frm1.vspdData 

	<% '변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 %>
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 계속 하시겠습니까?%>
		If IntRetCD = vbNo Then Exit Function
	End If

	<% '변경이 없을때 작업진행여부 체크 %>
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")                <% '작업을 수행하시겠습니까? %>
		If IntRetCD = vbNo Then Exit Function
	End If

	BatchReqCheckMsg = True

End Function


<%
'======================================================================================================
' Function Desc : Optinal Column Protect
'=======================================================================================================
%>
Sub OptinalProtected(ByVal Col, ByVal Row, ByVal Row2)

	With frm1.vspdData

    
	.Col = Col
	.SetColItemData .Col, 2
	.Col2 = Col
	.Row = Row
	.Row2 = Row2
	    
	.BlockMode = True
	.Protect = True
	.Lock = True
	.BlockMode = False

	End With
    
End Sub

<%
'======================================================================================================
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

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ItemCode		= iCurColumnPos(1)
			C_ItemPopup		= iCurColumnPos(2)
			C_ItemName		= iCurColumnPos(3)
			C_Spec			= iCurColumnPos(4)
			C_PlanUnit		= iCurColumnPos(5)
			C_PlanUnitPopup	= iCurColumnPos(6)
			C_YearQty		= iCurColumnPos(7)
			C_01PlanQty		= iCurColumnPos(8)
			C_02PlanQty		= iCurColumnPos(9)
			C_03PlanQty		= iCurColumnPos(10)
			C_04PlanQty		= iCurColumnPos(11)
			C_05PlanQty		= iCurColumnPos(12)
			C_06PlanQty		= iCurColumnPos(13)
			C_07PlanQty		= iCurColumnPos(14)
			C_08PlanQty		= iCurColumnPos(15)
			C_09PlanQty		= iCurColumnPos(16)
			C_10PlanQty		= iCurColumnPos(17)
			C_11PlanQty		= iCurColumnPos(18)
			C_12PlanQty		= iCurColumnPos(19)
			C_01PlanColor	= iCurColumnPos(20)
			C_02PlanColor	= iCurColumnPos(21)
			C_03PlanColor	= iCurColumnPos(22)
			C_04PlanColor	= iCurColumnPos(23)
			C_05PlanColor	= iCurColumnPos(24)
			C_06PlanColor	= iCurColumnPos(25)
			C_07PlanColor	= iCurColumnPos(26)
			C_08PlanColor	= iCurColumnPos(27)
			C_09PlanColor	= iCurColumnPos(28)
			C_10PlanColor	= iCurColumnPos(29)
			C_11PlanColor	= iCurColumnPos(30)
			C_12PlanColor	= iCurColumnPos(31)
    End Select    
End Sub


'========================================================================================================
'========================================================================================================
'                        5.2 Common Group-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029              '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
	Call InitVariables              '⊙: Initializes local global variables
	Call SetDefaultVal 
	Call InitSpreadSheet
 
    Call SetToolbar("11000000000011")          '⊙: 버튼 툴바 제어 
	Call CookiePage("ReadCookie")

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

<% '**************************  3.2 HTML Form Element & Object Event처리  **********************************
' Document의 TAG에서 발생 하는 Event 처리 
' Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
' Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* %>

<% '******************************  3.2.1 Object Tag 처리  *********************************************
' Window에 발생 하는 모든 Even 처리 
'********************************************************************************************************* %>

<%
'==========================================================================================
'   Event Desc : 공장별배분을 클릭할 경우 발생 
'==========================================================================================
%>
Sub btnSplit_OnClick()

    Err.Clear                                                               <%'☜: Protect system from crashing%>

	If BatchReqCheckMsg = False Then Exit Sub        <%' Requried Value Check Msg %>
 
	Call BatchButton(lsSPLIT)   <% '공장별배분작업 %>

End Sub

'==========================================================================================
Function BatchButton(SKubun)

	Dim strval

	If   LayerShowHide(1) = False Then
		Exit Function 
    End If

	strVal = ""    
	strVal = BIZ_PGM_ID & "?txtMode=" & SKubun          <%'☜: 비지니스 처리 ASP의 상태 %>
	'= strVal = strVal & "&lsItemCode=" & Trim(lsItemCode)        <%'☜: Batch 조건 데이타 %>
	strVal = strVal & "&HItemCd=" & Trim(frm1.HItemCd.value)   <%'☜: Batch 조건 데이타 %>
	strVal = strVal & "&txtItemArrary=" & Trim(frm1.txtItemArrary.value)   <%'☜: Batch 조건 데이타 %>
	strVal = strVal & "&txtItemCount=" & Trim(frm1.txtItemCount.value)
	strVal = strVal & "&lsPlanMonth=" & Trim(lsPlanMonth)
	strVal = strVal & "&lsPlanUnit=" & Trim(lsPlanUnit)
	strVal = strVal & "&HConSpYear=" & Trim(frm1.HConSpYear.value)

	Call RunMyBizASP(MyBizASP, strVal)            <%'☜: 비지니스 ASP 를 가동 %>

End Function

'==========================================================================================
Function btnSplit_Ok()
	
	Call MainQuery()

End Function


'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
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
		Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")
		End If
	End With

End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If  

	lsPlanMonth=""

	Select Case Col  
	Case C_01PlanQty,C_02PlanQty,C_03PlanQty,C_04PlanQty,C_05PlanQty,C_06PlanQty, _
		 C_07PlanQty,C_08PlanQty,C_09PlanQty,C_10PlanQty,C_11PlanQty,C_12PlanQty
	  If Row > 0 Then
			Call SpreadCellClick(Col,Row)	
	  End If
	End Select
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    	
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
Sub vspdData_Change(Col , Row)

    Dim iDx
       
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    lgBlnFlgChgValue = True
    
    Select Case Col  
		
		Case C_01PlanQty,C_02PlanQty,C_03PlanQty,C_04PlanQty,C_05PlanQty, _
		C_06PlanQty,C_07PlanQty,C_08PlanQty,C_09PlanQty,C_10PlanQty, _
		C_11PlanQty,C_12PlanQty

		Call MonthTotalSum()

	End Select
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
       
    If Row <= 0 Then
   	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub


'========================================================================================================
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
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


<% '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
' 설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* %>
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear             
        
    FncQuery = False                                                        <%'⊙: Processing is NG%>
 
	<%    '-----------------------
	'Check previous data area
	'----------------------- %>
	'************ 멀티인 경우 **************
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
    End If


<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")          <%'⊙: Clear Contents  Field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>

	If frm1.rdoCfmFlagY.checked = True Then
		frm1.txtCfmFlag.value = frm1.rdoCfmFlagY.value
	Else
		frm1.txtCfmFlag.value = frm1.rdoCfmFlagN.value 
	End If

<%  '-----------------------
    'Query function call area
    '----------------------- %>
    Call DbQuery                <%'☜: Query db data%>

    If Err.number = 0 Then	
       FncQuery = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
        
End Function

'========================================================================================
Function FncNew() 
    
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False																  '☜: Processing is NG
    
	<%  '-----------------------
	'Check previous data area
	'-----------------------%>
	'************ 멀티인 경우 **************
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
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

    Call SetToolbar("11000000000011")          '⊙: 버튼 툴바 제어 

    If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================
Function FncSave() 
    
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False                                                                <%'☜: Protect system from crashing%>

	If frm1.vspdData.MaxRows < 1 Then
		MsgBox "저장할 품목이 없습니다", vbExclamation, parent.gLogoName
		Exit Function
	End If
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
 '************ 멀티인 경우 **************
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If

<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Not chkField(Document, "2") Then   <%'⊙: Check contents area%>
       Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSDefaultCheck = False Then     <%'⊙: Check contents area%>
       Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll DbSave                                                    <%'☜: Save db data%>
    
    If DbSave = False Then                                                        '☜: Query db data
       Exit Function
    End If
    
    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
Function FncCopy() 

	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	With frm1

		.vspdData.ReDraw = False
		
	 	ggoSpread.Source = frm1.vspdData
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

		.vspdData.Col = C_ItemCode
		.vspdData.Text = ""
		.vspdData.Col = C_ItemName
		.vspdData.Text = ""
		.vspdData.Col = C_PlanUnit
		.vspdData.Text = ""
	  
		.vspdData.ReDraw = True
	
	End With
	
	If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================
Function FncCancel() 

	Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG
        
	If frm1.vspdData.MaxRows < 1 Then Exit Function
   
	ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo                                     '☜: Protect system from crashing
	
	Call MonthTotalSum()
	
	If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
	
End Function

'========================================================================================
Function FncInsertRow() 

	Dim imRow
	Dim GCol
 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
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

		ggoSpread.InsertRow
		SetSpreadColor .vspdData.ActiveRow,.vspdData.ActiveRow

		lgBlnFlgChgValue = True

		<% '----------  Coding part  -------------------------------------------------------------%>   

		For GCol = C_01PlanQty To C_12PlanQty
			.vspdData.Col = GCol
			.vspdData.Text = 0
		Next

		.vspdData.ReDraw = True
	
	End With
	
	If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
	Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
Function FncDeleteRow() 
	
	Dim lDelRows
	Dim iDelRowCnt, i
	
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False    
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    
    With frm1  

		.vspdData.focus
		ggoSpread.Source = .vspdData 
    
		<% '----------  Coding part  -------------------------------------------------------------%>   
		lDelRows = ggoSpread.DeleteRow

		Call MonthTotalSum()
 
		lgBlnFlgChgValue = True
    
    End With
    
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_MULTI, False)
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
	Call SetQuerySpreadColor()
End Sub

'========================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	'************ 멀티인 경우 **************
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
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
    
    
	If LayerShowHide(1) = False Then
		Exit Function 
    End If
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>
    
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtConSpYear=" & Trim(frm1.HConSpYear.value)    <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtConItemCd=" & Trim(frm1.HItemCd.value)
		strVal = strVal & "&lsPlanUnit=" & Trim(frm1.HPlanUnit.value)
		strVal = strVal & "&txtCfmFlag=" & Trim(frm1.txtCfmFlag.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtConSpYear=" & Trim(frm1.txtConSpYear.value)    <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtConItemCd=" & Trim(frm1.txtConItemCd.value)
		strVal = strVal & "&lsPlanUnit=" & Trim(lsPlanUnit)
		strVal = strVal & "&txtCfmFlag=" & Trim(frm1.txtCfmFlag.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If
 
	Call RunMyBizASP(MyBizASP, strVal)          <%'☜: 비지니스 ASP 를 가동 %>
 
    DbQuery = True               <%'⊙: Processing is NG%>

End Function

'========================================================================================
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>
	lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")         <%'⊙: This function lock the suitable field%>
    Call SetToolbar("11101011000111")          '⊙: 버튼 툴바 제어 

	Call MonthTotalSum()

 
	If frm1.vspdData.MaxRows > 0 Then
		frm1.btnSplit.disabled = False
		frm1.vspdData.Focus
	Else
		frm1.btnSplit.disabled = True
		frm1.txtConSpYear.focus
	End If

End Function

'========================================================================================
Function DbSave() 

    Err.Clear                <%'☜: Protect system from crashing%>
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal,strDel
 
 
	If LayerShowHide(1) = False Then
        Exit Function 
    End If
 
    DbSave = False                                                          '⊙: Processing is NG
    
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
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
				strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep'☜: C=Create
		    Case ggoSpread.UpdateFlag       '☜: 수정 
				strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep'☜: U=Update
			End Select
		 
		    Select Case .vspdData.Text

		    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag  '☜: 수정, 신규 

		        .vspdData.Col = C_ItemCode
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		        .vspdData.Col = C_PlanUnit
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		        .vspdData.Col = C_01PlanQty
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_02PlanQty
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_03PlanQty
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_04PlanQty
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_05PlanQty
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_06PlanQty  
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_07PlanQty  
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_08PlanQty  
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_09PlanQty  
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_10PlanQty  
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_11PlanQty  
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

		        .vspdData.Col = C_12PlanQty  
		        strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gRowSep

		        lGrpCnt = lGrpCnt + 1
		            
		    Case ggoSpread.DeleteFlag       '☜: 삭제 

				strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep

		        .vspdData.Col = C_ItemCode
		        strDel = strDel & Trim(.vspdData.Text) & parent.gColSep

		        .vspdData.Col = C_PlanUnit
		        strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

		        lGrpCnt = lGrpCnt + 1
		    End Select
		            
		Next

		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strDel & strVal
 
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 
 
	End With
 
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
Function DbSaveOk()               <%'☆: 저장 성공후 실행 로직 %>

    Call InitVariables
	frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function


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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>확정품목별판매계획</font></td>
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
         <TD CLASS="TD5" NOWRAP>계획년도</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConSpYear" ALT="계획년도" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12X1"></TD>
         <TD CLASS="TD5" NOWRAP>품목</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConItemCd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSalesPlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenITEMPopup()">&nbsp;<INPUT NAME="txtConItemNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
        </TR>
        <TR>
         <TD CLASS=TD5 NOWRAP>공장별배분구분</TD>
         <TD CLASS=TD6 NOWRAP>
          <input type=radio CLASS="RADIO" name="rdoCfmFlag" id="rdoCfmFlagY" value="Y" TAG="11X">
           <label for="rdoCfmFlagY">배분</label>&nbsp;&nbsp;&nbsp;&nbsp;
          <input type=radio CLASS = "RADIO" name="rdoCfmFlag" id="rdoCfmFlagN" value="N" TAG="11X" checked>
           <label for="rdoCfmFlagN">미배분</label></TD>
         <TD CLASS="TD5" NOWRAP></TD>
         <TD CLASS="TD6" NOWRAP></TD>
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
       <TABLE <%=LR_SPACE_TYPE_20%>>
        <TR>
         <TD HEIGHT="100%">
          <script language =javascript src='./js/s2112ma2_OBJECT1_vspdData.js'></script>
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
     <TD><BUTTON NAME="btnSplit" CLASS="CLSMBTN">공장별배분</BUTTON></TD>
     <TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:MonthlyPlan()">공장별판매계획조정</a></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtItemArrary" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtItemCount" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtCfmFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HConSpYear" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HPlanUnit" tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
  </DIV>
</BODY>
</HTML>
