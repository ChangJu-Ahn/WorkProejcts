<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : s1935ma1_ko441
'*  4. Program Name         : 고객별공정단가등록 
'*  5. Program Desc         : 고객별공정단가등록
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2008/07/29
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'=============================================================================f==========================
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                 '☜: indicates that All variables must be declared in advance

Dim prDBSYSDate

Dim EndDate ,StartDate

prDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company

StartDate = UniDateAdd("m", -1, EndDate,parent.gDateFormat)


Const BIZ_PGM_ID      = "s1935mb1_ko441.asp"            '☆: 비지니스 로직 ASP명 

Dim C_Bp_Cd
Dim C_Bp_Cd_Popup
Dim C_Bp_Nm
Dim C_Deal_type
Dim C_Deal_type_Popup
Dim C_Deal_type_nm
Dim C_Part_type
Dim C_Part_type_Popup
Dim C_Part_type_nm
Dim C_Valid_Dt
Dim C_Unit
Dim C_Unit_Popup
Dim C_Cur
Dim C_Cur_Popup
Dim C_Item_Price
Dim C_Fr_Qty
Dim C_To_Qty
Dim C_Remark

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim gblnWinEvent

'========================================================================================================
Sub initSpreadPosVariables()  
    Dim i
    i = 1
    
	C_Bp_Cd               = i : i = i + 1
	C_Bp_Cd_Popup         = i : i = i + 1
	C_Bp_Nm               = i : i = i + 1
	C_Deal_type           = i : i = i + 1
	C_Deal_type_Popup     = i : i = i + 1
	C_Deal_type_nm        = i : i = i + 1
	C_Part_type           = i : i = i + 1
	C_Part_type_Popup     = i : i = i + 1
	C_Part_type_nm        = i : i = i + 1
	C_Valid_Dt            = i : i = i + 1
	C_Unit                = i : i = i + 1
	C_Unit_Popup          = i : i = i + 1
	C_Cur                 = i : i = i + 1
	C_Cur_Popup           = i : i = i + 1
	C_Item_Price          = i : i = i + 1
	C_Fr_Qty              = i : i = i + 1
	C_To_Qty              = i : i = i + 1
	C_Remark			  = i : i = i + 1


End Sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
 frm1.txtconBp_cd.focus 
 lgBlnFlgChgValue = False
End Sub


'======================================================================================== 
<% '== 등록 == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
 	With frm1.vspdData

       ggoSpread.Spreadinit "V20050503",,parent.gAllowDragDropSpread    

       .MaxCols   = C_Remark														' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True
       .MaxRows = 0                                                                  ' ☜: Clear spreadsheet data 
      Call GetSpreadColumnPos("A")
	   .ReDraw = false
 			      
		ggoSpread.SSSetEdit     C_Bp_Cd,                "고객" ,12, 0,,10,2
		ggoSpread.SSSetButton   C_Bp_Cd_Popup    
		ggoSpread.SSSetEdit     C_Bp_Nm,                "고객명", 20, 0 
		ggoSpread.SSSetEdit     C_Deal_type,            "공정유형", 8, 0,,15,2        
		ggoSpread.SSSetButton   C_Deal_type_Popup
		ggoSpread.SSSetEdit     C_Deal_type_Nm,         "공정유형명", 12, 0
		ggoSpread.SSSetEdit     C_Part_type,            "생산구분", 8, 0,,15,2        
		ggoSpread.SSSetButton   C_Part_type_Popup
		ggoSpread.SSSetEdit     C_Part_type_Nm,         "생산구분명", 10, 0
		ggoSpread.SSSetDate     C_Valid_Dt,             "적용일", 10, 2, Parent.gDateFormat   
		ggoSpread.SSSetEdit     C_unit,                 "단위", 6, 0,,3,2
		ggoSpread.SSSetButton   C_unit_Popup                 
		ggoSpread.SSSetEdit     C_Cur,                  "화폐", 6, 0,,3,2
		ggoSpread.SSSetButton   C_Cur_Popup 
		ggoSpread.SSSetFloat    C_Item_Price,           "단가",12, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Fr_Qty,               "수량(FR)",12, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_To_Qty,               "수량(TO)",12, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit     C_Remark,				"비고", 30, 0,,240
		     
		.ReDraw = true   

		call ggoSpread.MakePairsColumn(C_Bp_Cd,C_Bp_Cd_Popup)
		call ggoSpread.MakePairsColumn(C_Deal_type,C_Deal_type_Popup)
		call ggoSpread.MakePairsColumn(C_Part_type,C_Part_type_Popup)
		call ggoSpread.MakePairsColumn(C_unit,C_unit_Popup)
		call ggoSpread.MakePairsColumn(C_Cur,C_Cur_Popup)

       'Call ggoSpread.SSSetColHidden(C_Fr_Qty,C_Fr_Qty,True)		

		Call SetSpreadLock 
	    
	End With
    
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False

	ggoSpread.spreadlock		C_Bp_Cd, -1, C_Cur_Popup
	ggoSpread.SSSetRequired		C_Item_Price,  -1, C_Item_Price
	ggoSpread.spreadlock		C_Fr_Qty, -1, C_To_Qty
    
    
    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired    C_Bp_Cd,              pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Bp_Nm  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Deal_type,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Deal_type_Nm,       pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Part_type,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Part_type_Nm,       pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Valid_Dt,           pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Unit,               pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Cur,                pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Item_Price,         pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Fr_Qty,             pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_To_Qty,             pvStartRow, pvEndRow
	    
    .vspdData.ReDraw = True
    
    End With

End Sub

'========================================================================================
Sub SetSpreadColor1(ByVal lRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetProtected    C_Bp_Cd,              lRow, lRow
    ggoSpread.SSSetProtected    C_Bp_Nm  ,            lRow, lRow
    ggoSpread.SSSetProtected    C_Deal_type,          lRow, lRow
    ggoSpread.SSSetProtected    C_Deal_type_Nm,       lRow, lRow
    ggoSpread.SSSetProtected    C_Part_type,          lRow, lRow
    ggoSpread.SSSetProtected    C_Part_type_Nm,       lRow, lRow
	ggoSpread.SSSetProtected    C_Valid_Dt,           lRow, lRow
	ggoSpread.SSSetProtected    C_Unit,               lRow, lRow
	ggoSpread.SSSetProtected    C_Cur,                lRow, lRow
	ggoSpread.SSSetRequired     C_Item_Price,         lRow, lRow
	ggoSpread.SSSetRequired     C_Fr_Qty,             lRow, lRow
	ggoSpread.SSSetRequired     C_To_Qty,             lRow, lRow
    
    .vspdData.ReDraw = True    
'	.vspdData.Col = C_Bp_Cd 
'	.vspdData.Row = .vspdData.ActiveRow
'	.vspdData.Action = 0
'	.vspdData.EditMode = True

    End With

End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    Dim i
    i = 1
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Bp_Cd				= iCurColumnPos(i) : i = i + 1
			C_Bp_Cd_Popup       = iCurColumnPos(i) : i = i + 1
			C_Bp_Nm				= iCurColumnPos(i) : i = i + 1   
			C_Deal_type			= iCurColumnPos(i) : i = i + 1
			C_Deal_type_Popup	= iCurColumnPos(i) : i = i + 1
			C_Deal_type_nm		= iCurColumnPos(i) : i = i + 1
			C_Part_type			= iCurColumnPos(i) : i = i + 1
			C_Part_type_Popup	= iCurColumnPos(i) : i = i + 1
			C_Part_type_nm		= iCurColumnPos(i) : i = i + 1
			C_Valid_Dt		    = iCurColumnPos(i) : i = i + 1
			C_Unit				= iCurColumnPos(i) : i = i + 1
			C_Unit_Popup		= iCurColumnPos(i) : i = i + 1
			C_Cur				= iCurColumnPos(i) : i = i + 1
			C_Cur_Popup			= iCurColumnPos(i) : i = i + 1
			C_Item_Price		= iCurColumnPos(i) : i = i + 1
			C_Fr_Qty		    = iCurColumnPos(i) : i = i + 1
			C_To_Qty		    = iCurColumnPos(i) : i = i + 1
			C_Remark			= iCurColumnPos(i) : i = i + 1			
    End Select    
End Sub
'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()

    Dim strCD
    Dim strVal		    
	

    
End Sub
'===========================================================================
Function OpenConSItemDC(Byval strCode, Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If gblnWinEvent = True Then Exit Function

 gblnWinEvent = True

 Select Case iWhere
 Case 0
  arrParam(1) = "b_item"                           <%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconItem_cd.Value)  <%' Code Condition%>
  arrParam(3) = ""                           <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "품목"       <%' TextBox 명칭 %>
 
  arrField(0) = "item_cd"        <%' Field명(0)%>
  arrField(1) = "item_nm"        <%' Field명(1)%>
  arrField(2) = "spec"        <%' Field명(1)%>
    
  arrHeader(0) = "품목"       <%' Header명(0)%>
  arrHeader(1) = "품목명"       <%' Header명(1)%> 
  arrHeader(2) = "규격"       <%' Header명(1)%>  

 Case 1
  arrParam(1) = "B_UNIT_OF_MEASURE"     <%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconSales_unit.Value)  <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "판매단위"      <%' TextBox 명칭 %>
 
  arrField(0) = "UNIT"        <%' Field명(0)%>
  arrField(1) = "UNIT_NM"        <%' Field명(1)%>
    
  arrHeader(0) = "단위"      <%' Header명(0)%>
  arrHeader(1) = "단위명"      <%' Header명(1)%>
  frm1.txtconSales_unit.focus 
 Case 2
  arrParam(1) = "B_CURRENCY"       <%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconCurrency.Value)  <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "화폐"       <%' TextBox 명칭 %>
 
  arrField(0) = "CURRENCY"       <%' Field명(0)%>
  arrField(1) = "CURRENCY_DESC"      <%' Field명(1)%>
    
  arrHeader(0) = "화폐"       <%' Header명(0)%>
  arrHeader(1) = "화폐명"       <%' Header명(1)%>
  frm1.txtconCurrency.focus 
 Case 3
  arrParam(1) = "b_user_defined_minor"        <%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconDeal_type.Value)  <%' Code Condition%>
  arrParam(3) = ""                                 <%' Name Cindition%>
  arrParam(1) = "b_user_defined_minor"        <%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconDeal_type.Value)  <%' Code Condition%>
  arrParam(3) = ""                                 <%' Name Cindition%>
  arrParam(4) = "ud_major_cd= 'zz505'"      <%' Where Condition%>
  arrParam(5) = "공정유형"      <%' TextBox 명칭 %>
 
  arrField(0) = "ud_minor_cd"        <%' Field명(0)%>
  arrField(1) = "ud_minor_nm"        <%' Field명(1)%>

  arrHeader(0) = "공정유형"      <%' Header명(0)%>
  arrHeader(1) = "공정유형명"      <%' Header명(1)%>

  frm1.txtconDeal_type.focus 
 Case 4
  arrParam(1) = "B_MINOR"        <%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconPay_meth.Value)  <%' Code Condition%>
  arrParam(3) = ""                                 <%' Name Cindition%>
  arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & ""     <%' Where Condition%>
  arrParam(5) = "결제방법"      <%' TextBox 명칭 %>
 
  arrField(0) = "MINOR_CD"       <%' Field명(0)%>
  arrField(1) = "MINOR_NM"       <%' Field명(1)%>
    
  arrHeader(0) = "결제방법"      <%' Header명(0)%>
  arrHeader(1) = "결제방법명"      <%' Header명(1)%>
  frm1.txtconPay_meth.focus 
    Case 5
  arrParam(1) = "B_BIZ_PARTNER"      <%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconBp_cd.Value) <%' Code Condition%>
  arrParam(3) = ""                                    <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"             <%' Where Condition%>
  arrParam(5) = "고객"          <%' TextBox 명칭 %>
 
  arrField(0) = "BP_CD"           <%' Field명(0)%>
  arrField(1) = "BP_NM"           <%' Field명(1)%>
    
  arrHeader(0) = "고객"          <%' Header명(0)%>
  arrHeader(1) = "고객명"            <%' Header명(1)%>
  frm1.txtconBp_cd.focus 
 End Select
    
    arrParam(3) = "" 
 arrParam(0) = arrParam(5)        <%' 팝업 명칭 %>

	Select Case iWhere
		Case 0
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
		Case Else
   arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
   "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

 gblnWinEvent = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetConSItemDC(arrRet, iWhere)
 End If 
 
End Function

'===========================================================================
 Function  OpenBp_cd(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "고객"       <%' 팝업 명칭 %>
  arrParam(1) = "B_Biz_Partner"           <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"         <%' Where Condition%>
  arrParam(5) = "고객"       <%' TextBox 명칭 %>

  arrField(0) = "Bp_cd"        <%' Field명(0)%>
  arrField(1) = "Bp_nm"        <%' Field명(1)%>

  arrHeader(0) = "고객"       <%' Header명(0)%>
  arrHeader(1) = "고객명"       <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetBp_cd(arrRet)
  End If
 End Function


'===========================================================================
 Function  OpenItem_cd(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "품목"       <%' 팝업 명칭 %>
  arrParam(1) = "B_ITEM"           <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "품목"       <%' TextBox 명칭 %>

  arrField(0) = "Item_cd"        <%' Field명(0)%>
  arrField(1) = "Item_nm"        <%' Field명(1)%>  
	arrField(2) = "Spec"	        <%' Field명(2)%>
	'arrField(3) = "HH" & parent.gColSep & "Basic_Unit"	        <%' Field명(3)%>

  arrHeader(0) = "품목"       <%' Header명(0)%>
  arrHeader(1) = "품목명"       <%' Header명(1)%>
	arrHeader(2) = "규격"       <%' Header명(2)%>
'	arrHeader(3) = "단위"       <%' Header명(3)%>
	   
  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetItem_cd(arrRet)
  End If
 End Function


'------------------------------------------  OpenAttend()  -------------------------------------------
'	Name : OpenIssueDt()
'	Description : 적용일 PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenIssueDt()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

     If gblnWinEvent = True Then Exit Function

     gblnWinEvent = True

	arrParam(0) = "과거 적용일 팝업"			' 팝업 명칭 
	arrParam(1) = "(select distinct VALID_FROM_DT from S_Basic_Price_Cust_KO441) a"						' TABLE 명칭 
	if frm1.txtconBp_cd.value <> "" then
	    arrParam(1) = "(select distinct VALID_FROM_DT from S_Basic_Price_Cust_KO441 where BP_CD = " &FilterVar(frm1.txtconBp_cd.value, "''", "S") & " ) a"						' TABLE 명칭 
	End if
	arrParam(2) = ""				' Code Condition
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""							' Where Condition%>
	arrParam(5) = "적용일"			' 조건필드의 라벨 명칭 
	
    arrField(0) = "DD30"  & Parent.gColSep &  "VALID_FROM_DT"					' Field명(0)

		
    arrHeader(0) = "적용일"			' Header명(0)
 
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=480px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtValid_dt.focus
	Else
		With frm1
			.txtValid_dt.text = arrRet(0)
			.txtValid_dt.focus
		End With
	End If	
	
End Function


'===========================================================================
 Function  OpenDeal_type(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "공정유형"       <%' 팝업 명칭 %>
  arrParam(1) = "b_user_defined_minor"                  <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)           <%' Code Condition%>
  arrParam(3) = ""             <%' Name Cindition%>
  arrParam(4) = "ud_major_cd= 'zz505'"         <%' Where Condition%>
  arrParam(5) = "공정유형"       <%' TextBox 명칭 %>

  arrField(0) = "ud_minor_cd"        <%' Field명(0)%>
  arrField(1) = "ud_minor_nm"        <%' Field명(1)%>

  arrHeader(0) = "공정유형"       <%' Header명(0)%>
  arrHeader(1) = "공정유형명"          <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetDeal_type(arrRet)
  End If
 End Function
 
 
'===========================================================================
 Function  OpenPart_type(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "생산구분"       <%' 팝업 명칭 %>
  arrParam(1) = "b_user_defined_minor"                  <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)           <%' Code Condition%>
  arrParam(3) = ""             <%' Name Cindition%>
  arrParam(4) = "ud_major_cd= 'zz506'"         <%' Where Condition%>
  arrParam(5) = "생산구분"       <%' TextBox 명칭 %>

  arrField(0) = "ud_minor_cd"        <%' Field명(0)%>
  arrField(1) = "ud_minor_nm"        <%' Field명(1)%>

  arrHeader(0) = "생산구분"       <%' Header명(0)%>
  arrHeader(1) = "생산구분명"          <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetPart_type(arrRet)
  End If
 End Function

'===========================================================================
 Function  OpenUnit(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "단위"       <%' 팝업 명칭 %>
  arrParam(1) = "B_UNIT_OF_MEASURE"      <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "단위"       <%' TextBox 명칭 %>

  arrField(0) = "UNIT"        <%' Field명(0)%>
  arrField(1) = "UNIT_NM"        <%' Field명(1)%>

  arrHeader(0) = "단위"       <%' Header명(0)%>
  arrHeader(1) = "단위명"      <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetUnit(arrRet)
  End If
 End Function

'===========================================================================
 Function  OpenCur(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "화폐"       <%' 팝업 명칭 %>
  arrParam(1) = "B_CURRENCY"      <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "화폐"       <%' TextBox 명칭 %>

  arrField(0) = "CURRENCY"        <%' Field명(0)%>
  arrField(1) = "CURRENCY_DESC"        <%' Field명(1)%>

  arrHeader(0) = "화폐"       <%' Header명(0)%>
  arrHeader(1) = "화폐명"      <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetCur(arrRet)
  End If
 End Function

'===========================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)

 With frm1

  Select Case iWhere
  Case 0
   .txtconItem_cd.value = arrRet(0) 
   .txtconItem_nm.value = arrRet(1)   
  Case 1
   .txtconSales_unit.value = arrRet(0)
   .txtconSales_unit_nm.value = arrRet(0) 
  Case 2
   .txtconCurrency.value = arrRet(0) 
  Case 3
   .txtconDeal_type.value = arrRet(0) 
   .txtconDeal_type_nm.value = arrRet(1)   
  Case 4
   .txtconPay_meth.value = arrRet(0) 
   .txtconPay_meth_nm.value = arrRet(1)   
  Case 5
   .txtconBp_cd.value = arrRet(0) 
   .txtconBp_nm.value = arrRet(1)    
  End Select

 End With
 
End Function

'===========================================================================
Function SetBp_cd(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Bp_cd
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_Bp_nm
  .vspdData.Text = arrRet(1)
 End With
End Function

'===========================================================================
Function SetItem_cd(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Item_cd
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_Item_nm
  .vspdData.Text = arrRet(1)
 End With
End Function

'===========================================================================
Function SetDeal_type(Byval arrRet)  
 With frm1
  .vspdData.Col =  C_Deal_type
  .vspdData.Text = arrRet(0)
  .vspdData.Col =  C_Deal_type_nm
  .vspdData.Text = arrRet(1)
 End With
End Function


'===========================================================================
Function SetPart_type(Byval arrRet)  
 With frm1
  .vspdData.Col =  C_Part_type
  .vspdData.Text = arrRet(0)
  .vspdData.Col =  C_Part_type_nm
  .vspdData.Text = arrRet(1)
 End With
End Function
'===========================================================================
Function SetUnit(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Unit
  .vspdData.Text = arrRet(0)
 End With
End Function

'===========================================================================
Function SetCur(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Cur
  .vspdData.Text = arrRet(0) 
 End With
 Call vspdData_Change(C_Cur,frm1.vspdData.ActiveRow)
End Function


Function txtconBp_cd_OnChange()
    txtconBp_cd_OnChange = true
    
    If  frm1.txtconBp_cd.value = "" Then
        frm1.txtconBp_nm.value = ""
        frm1.txtconBp_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" BP_NM "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(frm1.txtconBp_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","거래처","x")

            frm1.txtconBp_nm.value = ""
	        frm1.txtconBp_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtconBp_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function

'========================================================================================================= 
Sub Form_Load()
 
 Call LoadInfTB19029()
 Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)   '⊙: Format Contents  Field
 Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)   '⊙: Format Contents  Field
 
 Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
 
 Call InitSpreadSheet 
 Call SetDefaultVal
 Call InitVariables     
'⊙: Initializes local global variables

    '----------  Coding part  -------------------------------------------------------------
    Call InitSpreadComboBox
 
    Call SetToolBar("1110110100101111")          '⊙: 버튼 툴바 제어 

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
 
		ggoSpread.Source = frm1.vspdData

		If Row > 0 Then
			Select Case Col
			Case C_Bp_Cd_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenBp_Cd (.text)
			Case C_Deal_type_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenDeal_type (.Text)
			Case C_Part_type_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenPart_type (.Text)
			Case C_Unit_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenUnit (.Text)
			Case C_Cur_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenCur (.Text)
			End Select
		 
			Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")  
		End If

	End With
End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"   

	Set gActiveSpdSheet = frm1.vspdData
	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
	    
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If
	   	    
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.vspdData.Row = Row
	'---frm1.vspdData.Col = C_MajorCd
		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim	strFlag
	
	With frm1.vspdData
	
		.Row = Row
		Select Case Col	
		End Select		
    End With

End Sub
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

    Select Case Col
        Case  C_Cur
            ' Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Cur,C_Item_Price,"C" ,"X","X")
            ' Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_Cur,C_Item_Price,"C" ,"I","X","X")         
         Case  C_Item_Price                               
             '   Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Cur,C_Item_Price,"C" ,"X","X")
    End Select    
   
  ' Call CheckMinNumSpread(frm1.vspdData, Col, Row)
   ggoSpread.Source = frm1.vspdData
   ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_Item_Price
            Call EditModeCheck(frm1.vspdData, Row, C_Cur, C_Item_Price, "C" ,"I", Mode, "X", "X")        
    End Select
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then 
		If lgStrPrevKey <> "" Then  
			If CheckRunningBizProcess Then Exit Sub
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery()
		End if
	End if     
    
End Sub

<%
'==========================================================================================
'   Event Name : OCX_DbClick()
'   Event Desc : OCX_DbClick() 시 Calendar Popup
'==========================================================================================
%>
Sub txtconValid_from_dt_DblClick(Button)
 If Button = 1 Then
  frm1.txtconValid_from_dt.Action = 7
  Call SetFocusToDocument("M")
  Frm1.txtconValid_from_dt.Focus
 End If
End Sub

<%
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 조회조건부의 OCX_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
%>
Sub txtconValid_from_dt_KeyDown(KeyCode, Shift)

 If KeyCode = 13 Then Call MainQuery()

End Sub

'========================================================================================================
 Function FncQuery()
  Dim IntRetCD

  FncQuery = False             <% '⊙: Processing is NG %>

  Err.Clear               <% '☜: Protect system from crashing %>

  <% '------ Check previous data area ------ %>
  ggoSpread.Source = frm1.vspdData  
  If ggoSpread.SSCheckChange = True and lgBlnFlgChgValue=true Then   
   IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")   <% '⊙: "Will you destory previous data" %>
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  <% '------ Erase contents area ------ %>
  Call ggoOper.ClearField(Document, "2")        <% '⊙: Clear Contents  Field %>
  Call InitVariables             <% '⊙: Initializes local global variables %>

  <% '------ Check condition area ------ %>
  If Not chkField(Document, "1") Then       <% '⊙: This function check indispensable field %>
   Exit Function
  End If

  <% '------ Query function call area ------ %>
  Call DbQuery()              <% '☜: Query db data %>

  FncQuery = True              <% '⊙: Processing is OK %>
 End Function
 
'========================================================================================================
 Function FncNew()
  Dim IntRetCD 

  FncNew = False              <% '☜: Protect system from crashing %>

  <% '------ Check previous data area ------ %>
  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
   IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  <% '------ Erase condition area ----- %>
  <% '------ Erase contents area ------ %>
  Call ggoOper.ClearField(Document, "A")        <%'⊙: Clear Condition Field%>
  Call ggoOper.LockField(Document, "N")        <%'⊙: Lock  Suitable  Field%>
  Call SetDefaultVal
  Call SetToolBar("1110110100101111")          '⊙: 버튼 툴바 제어 
  Call InitVariables             <%'⊙: Initializes local global variables%>

  FncNew = True              <%'⊙: Processing is OK%>

 End Function
 
 
'========================================================================================================
 Function FncSave()
  Dim IntRetCD
  
  FncSave = False                  <% '⊙: Processing is NG %>
  
  Err.Clear                   <% '☜: Protect system from crashing %>
  
  <% '------ Precheck area ------ %>
  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = False Then        <% 'Check if there is retrived data %>   
      IntRetCD = DisplayMsgBox("900001","x","x","x")     <% '⊙: No data changed!! %>
      Exit Function
  End If
  
  <% '------ Check contents area ------ %>
  ggoSpread.Source = frm1.vspdData

  If Not chkField(Document, "2") Then  <% '⊙: Check contents area %>
   Exit Function
  End If

  If Not ggoSpread.SSDefaultCheck Then  <% '⊙: Check contents area %>
   Exit Function
  End If
  
  <% '------ Save function call area ------ %>
  Call DbSave                   <% '☜: Save db data %>
  
  FncSave = True                  <% '⊙: Processing is OK %>
 End Function

'========================================================================================================
 Function FncCopy()
	frm1.vspdData.ReDraw = False
	  
	if frm1.vspdData.maxrows < 1 then exit function
	  
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.CopyRow
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_Cur,C_Item_Price,"C" ,"I","X","X")         	
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
 End Function

'========================================================================================================
 Function FncCancel() 
   
  if frm1.vspdData.maxrows < 1 then exit function
  
  frm1.vspdData.ReDraw = False    
  ggoSpread.Source = frm1.vspdData
  ggoSpread.EditUndo              <%'☜: Protect system from crashing%>
  Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_Cur,C_Item_Price,"C" ,"I","X","X")         
  frm1.vspdData.ReDraw = True
 End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
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
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		lgBlnFlgChgValue = True
		.vspdData.Col= C_Valid_Dt
		.vspdData.Text= EndDate
		.vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
 End Function

'========================================================================================================
 Function FncDeleteRow()
  Dim lDelRows
  Dim iDelRowCnt, i
  
  if frm1.vspdData.maxrows < 1 then exit function
 
  With frm1.vspdData 
   If .MaxRows = 0 Then
    Exit Function
   End If

   .focus
   ggoSpread.Source = frm1.vspdData

   lDelRows = ggoSpread.DeleteRow

   lgBlnFlgChgValue = True
  End With
 End Function

'========================================================================================================
 Function FncPrint()
     ggoSpread.Source = frm1.vspdData
  Call parent.FncPrint()             <%'☜: Protect system from crashing%>
 End Function

'========================================================================================================
 Function FncExcel() 
  Call parent.FncExport(Parent.C_MULTI)
 End Function

'========================================================================================================
 Function FncFind() 
  Call parent.FncFind(Parent.C_MULTI, False)
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
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor(-1)
   ' Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Cur,C_Item_Price,"C","I","X","X")
End Sub

'========================================================================================================
 Function FncExit()
  Dim IntRetCD

  FncExit = False

  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
   IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   <%'⊙: "Will you destory previous data"%>
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  FncExit = True
 End Function


'========================================================================================================
 Function DbQuery()
  Err.Clear               <%'☜: Protect system from crashing%>

  DbQuery = False              <%'⊙: Processing is NG%>

  Dim strVal

  
  If   LayerShowHide(1) = False Then
             Exit Function 
  End If  
  
  
  If lgIntFlgMode = Parent.OPMD_UMODE Then
   strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                      <%'☜: 비지니스 처리 ASP의 상태 %>

   strVal = strVal & "&txtconBp_cd=" & Trim(frm1.txtconBp_cd.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&txtValid_dt=" & Trim(frm1.txtValid_dt.text)        
   strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
   strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows  
     
  Else
   strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                      <%'☜: 비지니스 처리 ASP의 상태 %>

   strVal = strVal & "&txtconBp_cd=" & Trim(frm1.txtconBp_cd.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&txtValid_dt=" & Trim(frm1.txtValid_dt.text)        
   strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
   strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
  End If


  Call RunMyBizASP(MyBizASP, strVal)         <%'☜: 비지니스 ASP 를 가동 %>
 
  DbQuery = True              <%'⊙: Processing is NG%>
  frm1.vspdData.Focus 
 End Function
 
'========================================================================================================
 Function DbSave() 
  Dim lRow
  Dim lGrpCnt
  Dim strVal, strDel 

  DbSave = False              <% '⊙: Processing is OK %>
    
  On Error Resume Next            <% '☜: Protect system from crashing %>

  
  If   LayerShowHide(1) = False Then
             Exit Function 
        End If

  With frm1
   .txtMode.value = Parent.UID_M0002
   .txtUpdtUserId.value = Parent.gUsrID
   .txtInsrtUserId.value = Parent.gUsrID

   lGrpCnt = 1

   strVal = ""
   strDel = ""

 

   For lRow = 1 To .vspdData.MaxRows
    .vspdData.Row = lRow
    .vspdData.Col = 0            
 
    
    Select Case .vspdData.Text
     Case ggoSpread.InsertFlag       
										  strVal = strVal & "C" & parent.gColSep 
										  strVal = strVal & lRow & parent.gColSep
		.vspdData.Col = C_Bp_Cd		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Deal_type	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Part_type	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Valid_Dt		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Unit			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Cur    		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Fr_Qty        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_To_Qty        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Remark        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Item_Price    : strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gRowSep

		lGrpCnt = lGrpCnt + 1
  
     Case ggoSpread.UpdateFlag      
										  strVal = strVal & "U" & parent.gColSep 
										  strVal = strVal & lRow & parent.gColSep
		.vspdData.Col = C_Bp_Cd		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Deal_type	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Part_type	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Valid_Dt		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Unit			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Cur    		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Fr_Qty        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_To_Qty        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Remark        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Item_Price    : strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gRowSep
     
		lGrpCnt = lGrpCnt + 1
 
     Case ggoSpread.DeleteFlag  
										  strDel = strDel & "D" & parent.gColSep 
										  strDel = strDel & lRow & parent.gColSep
		.vspdData.Col = C_Bp_Cd		    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Deal_type	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Part_type	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Valid_Dt		: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Unit			: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Cur    		: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Fr_Qty        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_To_Qty        : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

		lGrpCnt = lGrpCnt + 1
    End Select    
   Next

   .txtMaxRows.value = lGrpCnt-1
   .txtSpread.value =  strDel & strVal

   
   Call ExecMyBizASP(frm1, BIZ_PGM_ID)      <% '☜: 비지니스 ASP 를 가동 %>

  End With

  DbSave = True              <% '⊙: Processing is NG %>
 End Function
 
'========================================================================================================
 Function DbQueryOk()             <% '☆: 조회 성공후 실행로직 %>
  <% '------ Reset variables area ------ %>
  lgIntFlgMode = Parent.OPMD_UMODE           <% '⊙: Indicates that current mode is Update mode %>
  lgBlnFlgChgValue = False
  
  Call ggoOper.LockField(Document, "Q")        <% '⊙: This function lock the suitable field %>
  Call SetToolBar("1110111100111111") 
  frm1.vspdData.Focus     
        <% '⊙: 버튼 툴바 제어 %>
 End Function
 
'========================================================================================================
 Function DbSaveOk()              <%'☆: 저장 성공후 실행 로직 %>
  Call ggoOper.ClearField(Document, "2")
  Call InitVariables     
  Call MainQuery()
  
 End Function
 
</SCRIPT>
<!-- #Include file="../../inc/UNI2kCM.inc" --> 
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
     <TD CLASS="CLSLTAB">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>고객별공정유형단가등록</font></td>
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
         <TD CLASS="TD5" NOWRAP>고객사</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="고객사" TYPE="Text" MAXLENGTH=10 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconBp_cd.value,5">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
         <TD CLASS=TD5 NOWRAP>적용일</TD>
         <TD CLASS=TD6 NOWRAP><OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtValid_dt" CLASS=FPDTYYYYMM TAG="111X1" Title="FPDATETIME" ALT="적용일"  id=OBJECT4></OBJECT><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenIssueDt()"></TD>

        </TR>
      <!--  <TR>
         <TD CLASS="TD5" NOWRAP>품목</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconItem_cd.value,0">&nbsp;<INPUT NAME="txtconItem_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
        
        </TR>-->
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
          <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
         </TD>
        </TR>
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR>
  <TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
 </TR>  
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%>  FRAMEBORDER=0 SCROLLING=no noresize framespacing=0  TABINDEX = -1></IFRAME>
  </TD>
 </TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconBiz_partner" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconItem_cd" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconDeal_type" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconPay_meth" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconValid_from_dt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconSales_unit" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconCurrency" tag="24" TABINDEX = -1>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>