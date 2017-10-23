<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : s4153ma1_ko441
'*  4. Program Name         : po번호별환율적용정보등록 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2008/08/13
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


Const BIZ_PGM_ID      = "s4153mb1_ko441.asp"            '☆: 비지니스 로직 ASP명 

Dim C_Po_No
Dim C_Bp_Cd
Dim C_Bp_Nm
Dim C_Po_Dt
Dim C_Plant_Cd
Dim C_Plant_Nm
Dim C_Item_Cd
Dim C_Item_Nm
Dim C_Spec
Dim C_Po_Unit
Dim C_Po_Qty
Dim C_Po_Price
Dim C_Cur
Dim C_Material_Prc
Dim C_xch_Cur
Dim C_Cur_Popup
Dim C_xch_rate
Dim C_Remark


<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim gblnWinEvent

'========================================================================================================
Sub initSpreadPosVariables()  
    Dim i
    i = 1
    
	C_Po_No				  = i : i = i + 1
	C_Bp_Cd               = i : i = i + 1
	C_Bp_Nm               = i : i = i + 1
	C_Po_Dt				  = i : i = i + 1
	C_Plant_Cd			  = i : i = i + 1
	C_Plant_Nm			  = i : i = i + 1
	C_Item_Cd             = i : i = i + 1
	C_Item_Nm             = i : i = i + 1
	C_Spec                = i : i = i + 1
	C_Po_Unit             = i : i = i + 1
	C_Po_Qty              = i : i = i + 1
	C_Po_Price            = i : i = i + 1
	C_Cur				  = i : i = i + 1
	C_Material_Prc        = i : i = i + 1
	C_xch_Cur             = i : i = i + 1
	C_Cur_Popup			  = i : i = i + 1
	C_xch_rate            = i : i = i + 1
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
 			      
		ggoSpread.SSSetEdit     C_Po_No,                "PO번호" ,10, 0
		ggoSpread.SSSetEdit     C_Bp_Cd,                "고객" ,12, 0,,10,2
		ggoSpread.SSSetEdit     C_Bp_Nm,                "고객명", 14, 0 
        ggoSpread.SSSetDate     C_Po_Dt,				"발주일", 10,2, parent.gDateFormat   'Lock->Unlock/ Date
		ggoSpread.SSSetEdit     C_Plant_Cd,             "공장" ,12, 0,,18,2
		ggoSpread.SSSetEdit     C_Plant_Nm,             "공장명", 10, 0 
		ggoSpread.SSSetEdit     C_Item_Cd,              "품목" ,12, 0,,18,2
		ggoSpread.SSSetEdit     C_Item_Nm,              "품목명", 16, 0 
		ggoSpread.SSSetEdit     C_Spec,                 "규격", 16, 0 
		ggoSpread.SSSetEdit     C_Po_Unit,              "발주단위", 8, 0 
	    ggoSpread.SSSetFloat    C_Po_Qty,				"발주량",8, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Po_Price,				"발주단가",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit     C_Cur,                  "발주환종", 7, 0,,3,2
		ggoSpread.SSSetFloat    C_Material_Prc,			"자재단가",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit     C_xch_Cur,              "변환환종", 7, 0,,3,2
		ggoSpread.SSSetButton   C_Cur_Popup 
		ggoSpread.SSSetFloat    C_xch_rate,				"환율", 10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit     C_Remark,				"비고", 20, 0,,240

		.ReDraw = true   

		call ggoSpread.MakePairsColumn(C_xch_Cur,C_Cur_Popup)

        Call ggoSpread.SSSetColHidden(C_Bp_Cd,C_Bp_Cd,True)		
        Call ggoSpread.SSSetColHidden(C_Plant_Cd,C_Plant_Cd,True)		

		Call SetSpreadLock 
	    
	End With
    
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False

	ggoSpread.spreadlock		C_Po_No, -1, C_Cur
	ggoSpread.SSSetRequired		C_Material_Prc,  -1, C_Material_Prc
	ggoSpread.SSSetRequired		C_xch_Cur,  -1, C_xch_Cur
	ggoSpread.SSSetRequired		C_xch_rate,  -1, C_xch_rate
    
    
    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetProtected   C_Po_No  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Bp_Cd  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Bp_Nm  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Po_Dt  ,			 pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Plant_Cd  ,         pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Plant_Nm  ,         pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Item_Cd  ,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Item_Nm  ,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Spec  ,             pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Po_Unit  ,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Po_Qty  ,           pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Po_Price  ,         pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Cur  ,				 pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Material_Prc,       pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_xch_Cur,			 pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_xch_rate,           pvStartRow, pvEndRow

    .vspdData.ReDraw = True
    
    End With

End Sub

'========================================================================================
Sub SetSpreadColor1(ByVal lRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetProtected    C_Bp_Cd,              lRow, lRow
    ggoSpread.SSSetProtected    C_Bp_Nm  ,            lRow, lRow
    ggoSpread.SSSetProtected    C_Item_Cd,            lRow, lRow
    ggoSpread.SSSetProtected    C_Item_Nm  ,          lRow, lRow
    ggoSpread.SSSetProtected    C_Spec  ,             lRow, lRow
    ggoSpread.SSSetProtected    C_Plant_Cd,          lRow, lRow
    ggoSpread.SSSetProtected    C_Plant_Nm,       lRow, lRow
    
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
			C_Po_No				= iCurColumnPos(i) : i = i + 1
			C_Bp_Cd				= iCurColumnPos(i) : i = i + 1
			C_Bp_Nm				= iCurColumnPos(i) : i = i + 1   
			C_Po_Dt				= iCurColumnPos(i) : i = i + 1
			C_Plant_Cd			= iCurColumnPos(i) : i = i + 1
			C_Plant_Nm			= iCurColumnPos(i) : i = i + 1
			C_Item_Cd			= iCurColumnPos(i) : i = i + 1
			C_Item_Nm			= iCurColumnPos(i) : i = i + 1
			C_Spec              = iCurColumnPos(i) : i = i + 1
			C_Po_Unit           = iCurColumnPos(i) : i = i + 1
			C_Po_Qty            = iCurColumnPos(i) : i = i + 1
			C_Po_Price          = iCurColumnPos(i) : i = i + 1
			C_Cur			    = iCurColumnPos(i) : i = i + 1
			C_Material_Prc        = iCurColumnPos(i) : i = i + 1
			C_xch_Cur             = iCurColumnPos(i) : i = i + 1
			C_Cur_Popup           = iCurColumnPos(i) : i = i + 1
			C_xch_rate            = iCurColumnPos(i) : i = i + 1
			C_Remark			  = iCurColumnPos(i) : i = i + 1

    End Select    
End Sub
'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()

    Dim strCD
    Dim strVal		    
	
	'****************************
	'List Minor code(Price Flag Code)
	'****************************	

    
End Sub


'======================================================================================================
'	Name : Open???()
'	Description : Ref 화면을 call한다. 
'======================================================================================================
Function OpenRefOpenNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("s4153ra1_ko441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s4153ra1_ko441", "X")
		gblnWinEvent = False
		Exit Function
	End If

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	With frm1
		arrParam(0) = .txtconBp_cd.value											' 검색조건이 있을경우 파라미터 
		arrParam(1) = .txtconBp_nm.Value			
	End With

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=760px; dialogHeight=460px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpen(arrRet)
	End If
End Function

'======================================================================================================
'	Name : SetRefOpen()
'	Description : Popup에서 Return되는 값 setting
'======================================================================================================
Function SetRefOpen(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	Dim ii
	Dim iOpenNo 
	
	With frm1
		.vspdData.focus		
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False	
	
		TempRow = .vspdData.MaxRows												'☜: 현재까지의 MaxRows

		For I = TempRow To TempRow + Ubound(arrRet, 1)
			sFindFg	= "N"
			For x = 1 To TempRow
				.vspdData.Row = x
				.vspdData.Col = C_Po_No
				iOpenNo = .vspdData.text
							
				If UCase(Trim(iOpenNo)) = UCase(Trim(arrRet(I - TempRow, 0)))  Then
					sFindFg	= "Y"
				End If
			Next
			If 	sFindFg	= "N" Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
				.vspdData.Row = I + 1				
				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag

				.vspdData.Col = C_Po_No
				.vspdData.text = arrRet(I - TempRow,0)
				.vspdData.Col = C_Bp_Cd
				.vspdData.text = arrRet(I - TempRow,1)
				.vspdData.Col = C_Bp_Nm
				.vspdData.text = arrRet(I - TempRow,2)
				.vspdData.Col = C_Po_Dt
				.vspdData.text = arrRet(I - TempRow,3)
				.vspdData.Col = C_Plant_Cd
				.vspdData.text = arrRet(I - TempRow,4)
				.vspdData.Col = C_Plant_Nm
				.vspdData.text = arrRet(I - TempRow,5)
				.vspdData.Col = C_Item_Cd
				.vspdData.text = arrRet(I - TempRow,6)
				.vspdData.Col = C_Item_Nm
				.vspdData.text = arrRet(I - TempRow,7)
				.vspdData.Col = C_Po_Unit
				.vspdData.text = arrRet(I - TempRow,8)
				.vspdData.Col = C_Po_Qty
				.vspdData.text = arrRet(I - TempRow,9)
				.vspdData.Col = C_Po_Price
				.vspdData.text = arrRet(I - TempRow,10)
				.vspdData.Col = C_Material_Prc
				.vspdData.text = arrRet(I - TempRow,10)				
				.vspdData.Col = C_Cur
				.vspdData.text = arrRet(I - TempRow,11)				
				.vspdData.Col = C_xch_Cur
				.vspdData.text = arrRet(I - TempRow,11)				
				.vspdData.Col = C_xch_rate
				.vspdData.text = arrRet(I - TempRow,12)				

			End If	
		Next

		If TempRow + 1 <= .vspdData.MaxRows Then
			Call SetSpreadColor(TempRow + 1,.vspdData.MaxRows)
		End If

		.vspdData.ReDraw = True
    End With
    
End Function
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
  arrParam(4) = "" '"BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"             <%' Where Condition%>
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
	 With frm1
	  .vspdData.Col = C_xch_Cur
	  .vspdData.Text = arrRet(0) 
	 End With
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


Function txtconItem_cd_OnChange()
   txtconItem_cd_OnChange = true
    
    If  frm1.txtconItem_cd.value = "" Then
        frm1.txtconItem_nm.value = ""
        frm1.txtconItem_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" ITEM_NM "," B_ITEM "," ITEM_CD = " & FilterVar(frm1.txtconItem_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","품목코드","x")

            frm1.txtconItem_nm.value = ""
	        frm1.txtconItem_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtconItem_nm.value = Replace(lgF0, Chr(11), "")
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
 
    Call SetToolBar("1100101100011111")          '⊙: 버튼 툴바 제어 
 
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
			Case  C_Fr_Qty			
				.Col = Col
				strFlag = .Text
				If StrFlag = "T" Then
					.Col = C_To_Qty
					.Text = "진단가"
				Else
					.Col = C_To_Qty
					.Text = "가단가"
				End If
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

   ggoSpread.Source = frm1.vspdData
   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

    Select Case Col
   End Select    
   
  ' Call CheckMinNumSpread(frm1.vspdData, Col, Row)
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
  Call SetToolBar("1100101100011111")          '⊙: 버튼 툴바 제어 
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
	'Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_Cur,C_Item_Price,"C" ,"I","X","X")         	
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

   strVal = strVal & "&txtconBp_cd=" & Trim(frm1.txtHconBp_cd.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&txtPlantCode=" & Trim(frm1.txtHPlantCode.value)
   strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtHconItem_cd.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&txtPoNo=" & Trim(frm1.txtHPoNo.value)  
   strVal = strVal & "&txtconFr_dt=" & Trim(frm1.txtHconFr_dt.value)   
   strVal = strVal & "&txtconTo_dt=" & Trim(frm1.txtHconTo_dt.value)   
   strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
   strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows  
     
  Else
   strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                      <%'☜: 비지니스 처리 ASP의 상태 %>

   strVal = strVal & "&txtconBp_cd=" & Trim(frm1.txtconBp_cd.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&txtPlantCode=" & Trim(frm1.txtPlantCode.value)
   strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtconItem_cd.value)   
   strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)  
   strVal = strVal & "&txtconFr_dt=" & Trim(frm1.txtconFr_dt.text)   
   strVal = strVal & "&txtconTo_dt=" & Trim(frm1.txtconTo_dt.text)  
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
		.vspdData.Col = C_Po_No		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Bp_Cd		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Po_Dt			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Plant_Cd	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Item_Cd		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Po_Unit		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Po_Qty		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Po_Price		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Cur			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Material_Prc	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_xch_Cur		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_xch_rate		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Remark        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		lGrpCnt = lGrpCnt + 1
  
     Case ggoSpread.UpdateFlag      
										  strVal = strVal & "U" & parent.gColSep 
										  strVal = strVal & lRow & parent.gColSep
		.vspdData.Col = C_Po_No		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Material_Prc	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_xch_Cur		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_xch_rate		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Remark        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
     
		lGrpCnt = lGrpCnt + 1
 
     Case ggoSpread.DeleteFlag  
										  strDel = strDel & "D" & parent.gColSep 
										  strDel = strDel & lRow & parent.gColSep
		.vspdData.Col = C_Po_No			: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

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
  Call SetToolBar("1100101100011111") 
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>PO번호별환율적용정보등록</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
		<TD WIDTH=*>&nbsp;</TD>
		<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenRefOpenNo()">발주내역참조</A></TD>
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
         <TD CLASS="TD5" NOWRAP>고객사</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="고객사" TYPE="Text" MAXLENGTH=10 SiZE=12  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconBp_cd.value,5">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>발주일자</TD>
         <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtconFr_dt" CLASS=FPDTYYYYMMDD tag="11X1X" ALT="발주시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtconTo_dt" CLASS=FPDTYYYYMMDD tag="11X1X" ALT="발주종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>       
        </TR>
        <TR>
		 <TD CLASS="TD5" NOWRAP>공장</TD>
		 <TD CLASS="TD6"><INPUT NAME="txtPlantCode" TYPE="Text" ALT="공장" MAXLENGTH=10 SiZE=12 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtPlantCode.value,4">&nbsp;<INPUT NAME="txtPlantName" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>품목</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconItem_cd.value,0">&nbsp;<INPUT NAME="txtconItem_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>PO번호</TD>
         <TD CLASS="TD6"><INPUT NAME="txtPoNo" ALT="PO번호" TYPE="Text" MAXLENGTH=18 SiZE=20  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtPoNo.value,1">&nbsp;</TD>
         <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
         <TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
<INPUT TYPE=HIDDEN NAME="txtHconValid_from_dt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconSales_unit" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconCurrency" tag="24" TABINDEX = -1>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>