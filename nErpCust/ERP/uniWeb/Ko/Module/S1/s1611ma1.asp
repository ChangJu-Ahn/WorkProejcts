<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        :  
'*  3. Program ID           : S1611MA1
'*  4. Program Name         : 적립금정보등록 
'*  5. Program Desc         : 적립금정보등록 
'*  6. Comproxy List        : PS1G119.dll, PS1G120.dll
'*  7. Modified date(First) : 2002/03/07
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/11/15 : UI성능 적용 
'*                            2002/11/23 : Grid성능 적용, Kang Jun Gu
'*                            2002/12/10 : INCLUDE 다시 성능 적용, Kang Jun Gu
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                 '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "s1611mb1.asp"            '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "s1611mb1.asp" 

Dim C_Bp_Cd
Dim C_Bp_Cd_Popup
Dim C_Bp_Nm
Dim C_Item_Cd
Dim C_Item_Cd_Popup
Dim C_Item_Nm
Dim C_ItemSpec
Dim C_Unit
Dim C_Unit_Popup
Dim C_Valid_from_dt
Dim C_Cur
Dim C_Cur_Popup
Dim C_Item_Price
Dim C_ChgFlg

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim gblnWinEvent

'========================================================================================================
Sub initSpreadPosVariables()  
	C_Bp_Cd               = 1
	C_Bp_Cd_Popup         = 2
	C_Bp_Nm               = 3
	C_Item_Cd             = 4
	C_Item_Cd_Popup       = 5
	C_Item_Nm             = 6
	C_ItemSpec            = 7
	C_Unit                = 8
	C_Unit_Popup          = 9
	C_Valid_from_dt       = 10
	C_Cur                 = 11
	C_Cur_Popup           = 12
	C_Item_Price          = 13
	C_ChgFlg    = 14

End Sub

'========================================================================================================
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
 frm1.txtconBiz_partner.focus
 lgBlnFlgChgValue = False
End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %> 
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData

	    ggoSpread.Spreadinit "V20030301",,parent.gAllowDragDropSpread

	    .MaxCols   = C_ChgFlg														' ☜:☜: Add 1 to Maxcols
		.Col       = .MaxCols        : .ColHidden = True
	    .MaxRows = 0                                                                  ' ☜: Clear spreadsheet data 
	    Call GetSpreadColumnPos("A")
		.ReDraw = false
		      
		ggoSpread.SSSetEdit     C_Bp_Cd,                "거래처" ,15, 0,,10,2
		ggoSpread.SSSetButton   C_Bp_Cd_Popup    
		ggoSpread.SSSetEdit     C_Bp_Nm,                "거래처명", 25, 0 
		ggoSpread.SSSetEdit     C_Item_Cd,              "품목" ,15, 0,,18,2
		ggoSpread.SSSetButton   C_Item_Cd_Popup    
		ggoSpread.SSSetEdit     C_Item_Nm,              "품목명", 25, 0 
		ggoSpread.SSSetEdit		C_ItemSpec,				"규격",			20
		ggoSpread.SSSetEdit     C_unit,                 "단위", 10, 0,,3,2
		ggoSpread.SSSetButton   C_unit_Popup                 
		ggoSpread.SSSetDate     C_Valid_from_dt,        "적용일", 10, 2, Parent.gDateFormat   
		ggoSpread.SSSetEdit     C_Cur,                  "화폐", 10, 0,,3,2
		ggoSpread.SSSetButton   C_Cur_Popup 
		ggoSpread.SSSetFloat    C_Item_Price,           "적립단가",15, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit  C_ChgFlg, "Chgfg", 1, 2  
		     
		call ggoSpread.MakePairsColumn(C_Bp_Cd,C_Bp_Cd_Popup)
		call ggoSpread.MakePairsColumn(C_Item_Cd,C_Item_Cd_Popup)
		call ggoSpread.MakePairsColumn(C_unit,C_unit_Popup)
		call ggoSpread.MakePairsColumn(C_Cur,C_Cur_Popup)
		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)

		.ReDraw = true   

		Call SetSpreadLock 
	    
	End With
    
End Sub

'===========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.spreadlock    C_Bp_Nm, -1
    ggoSpread.spreadUnlock  C_Item_Cd , -1
    ggoSpread.spreadlock    C_Item_Nm, -1
    ggoSpread.spreadlock    C_ItemSpec, -1
    ggoSpread.spreadUnlock  C_Valid_from_dt, -1    
    
    .vspdData.ReDraw = True

    End With
End Sub

'===========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired    C_Bp_Cd,              pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Bp_Nm  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Item_Cd,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Item_Nm  ,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_ItemSpec  ,          pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Unit,               pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Valid_from_dt,      pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Cur,                pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Item_Price,         pvStartRow, pvEndRow
    
    .vspdData.ReDraw = True
    
    End With

End Sub

'===========================================================================================================
Sub SetSpreadColor1(ByVal lRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetProtected    C_Bp_Cd,              lRow, lRow
    ggoSpread.SSSetProtected    C_Bp_Nm  ,            lRow, lRow
    ggoSpread.SSSetProtected    C_Item_Cd,            lRow, lRow
    ggoSpread.SSSetProtected    C_Item_Nm  ,          lRow, lRow
    ggoSpread.SSSetProtected    C_ItemSpec  ,          lRow, lRow
    ggoSpread.SSSetProtected    C_Unit,               lRow, lRow
    ggoSpread.SSSetProtected    C_Valid_from_dt,      lRow, lRow
	ggoSpread.SSSetProtected    C_Cur,                lRow, lRow
	ggoSpread.SSSetRequired     C_Item_Price,         lRow, lRow
    
    .vspdData.ReDraw = True
    
    End With

End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Bp_Cd               = iCurColumnPos(1)
			C_Bp_Cd_Popup         = iCurColumnPos(2)
			C_Bp_Nm               = iCurColumnPos(3)
			C_Item_Cd             = iCurColumnPos(4)
			C_Item_Cd_Popup       = iCurColumnPos(5)
			C_Item_Nm             = iCurColumnPos(6)
			C_ItemSpec            = iCurColumnPos(7)
			C_Unit                = iCurColumnPos(8)
			C_Unit_Popup          = iCurColumnPos(9)
			C_Valid_from_dt       = iCurColumnPos(10)
			C_Cur                 = iCurColumnPos(11)
			C_Cur_Popup           = iCurColumnPos(12)
			C_Item_Price          = iCurColumnPos(13)
			C_ChgFlg    = iCurColumnPos(14)
    End Select    
End Sub

'===========================================================================
Function OpenConSItemDC(Byval strCode, Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If gblnWinEvent = True Then Exit Function

 gblnWinEvent = True

 Select Case iWhere
 Case 0
  arrParam(1) = "b_item"							<%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconItem_cd.Value)		<%' Code Condition%>
  arrParam(3) = ""									<%' Name Cindition%>
  arrParam(4) = ""									<%' Where Condition%>
  arrParam(5) = "품목"							<%' TextBox 명칭 %>
 
  arrField(0) = "item_cd"							<%' Field명(0)%>
  arrField(1) = "item_nm"							<%' Field명(1)%>
	   arrField(2) = "Spec"	         <%' Field명(1)%>
    
  arrHeader(0) = "품목"							<%' Header명(0)%>
  arrHeader(1) = "품목명"						<%' Header명(1)%>
	   arrHeader(2) = "규격"        <%' Header명(1)%>

 Case 1
  arrParam(1) = "B_UNIT_OF_MEASURE"					<%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconSales_unit.Value)	<%' Code Condition%>
  arrParam(3) = ""									<%' Name Cindition%>
  arrParam(4) = ""									<%' Where Condition%>
  arrParam(5) = "단위"							<%' TextBox 명칭 %>
 
  arrField(0) = "UNIT"								<%' Field명(0)%>
  arrField(1) = "UNIT_NM"							<%' Field명(1)%>
    
  arrHeader(0) = "단위"							<%' Header명(0)%>
  arrHeader(1) = "단위명"						<%' Header명(1)%>
  frm1.txtconSales_unit.focus 
 Case 2
  arrParam(1) = "B_CURRENCY"						<%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconCurrency.Value)		<%' Code Condition%>
  arrParam(3) = ""									<%' Name Cindition%>
  arrParam(4) = ""									<%' Where Condition%>
  arrParam(5) = "화폐"							<%' TextBox 명칭 %>
 
  arrField(0) = "CURRENCY"							<%' Field명(0)%>
  arrField(1) = "CURRENCY_DESC"						<%' Field명(1)%>
    
  arrHeader(0) = "화폐"							<%' Header명(0)%>
  arrHeader(1) = "화폐명"						<%' Header명(1)%>
  frm1.txtconCurrency.focus 
 Case 5
  arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
  arrParam(2) = Trim(frm1.txtconBiz_partner.Value)	<%' Code Condition%>
  arrParam(3) = ""                                  <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"             <%' Where Condition%>
  arrParam(5) = "거래처"						<%' TextBox 명칭 %>
 
  arrField(0) = "BP_CD"								<%' Field명(0)%>
  arrField(1) = "BP_NM"								<%' Field명(1)%>
    
  arrHeader(0) = "거래처"						<%' Header명(0)%>
  arrHeader(1) = "거래처명"						<%' Header명(1)%>
  frm1.txtconBiz_partner.focus 
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

'========================================================================================================
 Function  OpenBp_cd(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "거래처"       <%' 팝업 명칭 %>
  arrParam(1) = "B_Biz_Partner"           <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"         <%' Where Condition%>
  arrParam(5) = "거래처"       <%' TextBox 명칭 %>

  arrField(0) = "Bp_cd"        <%' Field명(0)%>
  arrField(1) = "Bp_nm"        <%' Field명(1)%>

  arrHeader(0) = "거래처"       <%' Header명(0)%>
  arrHeader(1) = "거래처명"       <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetBp_cd(arrRet)
  End If
 End Function


'========================================================================================================
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
	arrField(2) = "Spec"	         <%' Field명(1)%>

  arrHeader(0) = "품목"       <%' Header명(0)%>
  arrHeader(1) = "품목명"       <%' Header명(1)%>
	arrHeader(2) = "규격"        <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
   "dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetItem_cd(arrRet)
  End If
 End Function

'========================================================================================================
 Function  OpenDeal_type(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "판매유형"       <%' 팝업 명칭 %>
  arrParam(1) = "B_minor"                  <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)           <%' Code Condition%>
  arrParam(3) = ""             <%' Name Cindition%>
  arrParam(4) = "MAJOR_CD=" & FilterVar("S0001", "''", "S") & ""         <%' Where Condition%>
  arrParam(5) = "판매유형"       <%' TextBox 명칭 %>

  arrField(0) = "minor_cd"        <%' Field명(0)%>
  arrField(1) = "minor_nm"        <%' Field명(1)%>

  arrHeader(0) = "판매유형"       <%' Header명(0)%>
  arrHeader(1) = "판매유형명"          <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetDeal_type(arrRet)
  End If
 End Function

'========================================================================================================
 Function  OpenPay_meth(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "결제방법"       <%' 팝업 명칭 %>
  arrParam(1) = "B_MINOR"                  <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)           <%' Code Condition%>
  arrParam(3) = ""             <%' Name Cindition%>
  arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & ""         <%' Where Condition%>
  arrParam(5) = "결제방법"       <%' TextBox 명칭 %>

  arrField(0) = "MINOR_CD"        <%' Field명(0)%>
  arrField(1) = "MINOR_NM"        <%' Field명(1)%>

  arrHeader(0) = "결제방법"       <%' Header명(0)%>
  arrHeader(1) = "결제방법명"          <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetPay_meth(arrRet)
  End If
 End Function

'========================================================================================================
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

'========================================================================================================
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

'========================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)

 With frm1

  Select Case iWhere
  Case 0
   .txtconItem_cd.value = arrRet(0) 
   .txtconItem_nm.value = arrRet(1)   
  Case 1
   .txtconSales_unit.value = arrRet(0)
   .txtconSales_unit_nm.value = arrRet(0) 
   .txtconSales_unit.focus 
  Case 2
   .txtconCurrency.value = arrRet(0) 
   .txtconCurrency.focus 
  Case 5
   .txtconBiz_partner.value = arrRet(0) 
   .txtconBiz_partner_nm.value = arrRet(1)    
   .txtconBiz_partner.focus 
  End Select

 End With
 
End Function

'========================================================================================================
Function SetBp_cd(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Bp_cd
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_Bp_nm
  .vspdData.Text = arrRet(1)
 End With
End Function

'========================================================================================================
Function SetItem_cd(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Item_cd
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_Item_nm
  .vspdData.Text = arrRet(1)
  .vspdData.Col = C_ItemSpec
  .vspdData.Text = arrRet(2)
 End With
End Function




'========================================================================================================
' Name : SetUnit()
' Description : OpenConSItemDC Popup에서 Return되는 값 setting
'========================================================================================================
Function SetUnit(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Unit
  .vspdData.Text = arrRet(0)
  
 End With
End Function

'========================================================================================================
' Name : SetCur()
' Description : OpenConSItemDC Popup에서 Return되는 값 setting
'========================================================================================================
Function SetCur(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Cur
  .vspdData.Text = arrRet(0)
  
 End With
 Call vspdData_Change(C_Cur,frm1.vspdData.ActiveRow) 
End Function


'========================================================================================================
Sub Form_Load()

 Call LoadInfTB19029              '⊙: Load table , B_numeric_format

 Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)   '⊙: Format Contents  Field
 
 Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
 Call InitSpreadSheet
 Call SetDefaultVal
 Call InitVariables     
                                                     '⊙: Initializes local global variables

    '----------  Coding part  -------------------------------------------------------------
  Call SetToolBar("1110110100101111")          '⊙: 버튼 툴바 제어 

End Sub

'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
 With frm1.vspdData 
 
  ggoSpread.Source = frm1.vspdData

  If Row > 0 Then
	Select Case Col
	Case C_Bp_Cd_Popup
		.Col = Col - 1
		.Row = Row
		Call OpenBp_Cd (.text)
	Case C_Item_Cd_Popup
		.Col = Col - 1
		.Row = Row
		Call OpenItem_Cd (.text) 
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

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

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

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

    Select Case Col
        Case  C_Cur
             Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Cur,C_Item_Price,"C" ,"X","X")
             Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_Cur,C_Item_Price,"C" ,"I","X","X")         
         Case  C_Item_Price                               
                Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Cur,C_Item_Price,"C" ,"X","X")
    End Select    

   If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
   End If
   ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_Item_Price
            Call EditModeCheck(frm1.vspdData, Row, C_Cur, C_Item_Price, "C" ,"I", Mode, "X", "X")        
    End Select
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then 
			If CheckRunningBizProcess Then Exit Sub
 
			Call DisableToolBar(Parent.TBC_QUERY)
			Call DbQuery()
		End if
	End if     
End Sub
      

'========================================================================================================
 Function FncQuery()
  Dim IntRetCD

  FncQuery = False             <% '⊙: Processing is NG %>

  Err.Clear               <% '☜: Protect system from crashing %>

  <% '------ Check previous data area ------ %>
  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
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
  Call ggoOper.ClearField(Document, "A")        <%'⊙: Clear Contents  Field%>
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
    frm1.vspdData.Redraw = False    
  ggoSpread.Source = frm1.vspdData
  ggoSpread.EditUndo              <%'☜: Protect system from crashing%>
 
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_Cur,C_Item_Price,"C" ,"I","X","X")         
    frm1.vspdData.Redraw = True 
 End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
  With frm1
		FncInsertRow = False                                                         '☜: Processing is NG
		If Not chkField(Document, "2") Then
		Exit Function
		End If

		If IsNumeric(Trim(pvRowCnt)) Then
		    imRow = CInt(pvRowCnt)
		Else
		    imRow = AskSpdSheetAddRowCount()
		    If imRow = "" Then
		        Exit Function
		    End If
		End If
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow

		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.Row= .vspdData.ActiveRow
		.vspdData.ReDraw = True
  End With
 	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    Set gActiveElement = document.ActiveElement   
   
   lgBlnFlgChgValue = True
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
    Call parent.FncPrint()    <%'☜: Protect system from crashing%>
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
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor1(-1)
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Cur,C_Item_Price,"C","I","X","X")
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
   strVal = strVal & "&txtconBiz_partner=" & Trim(frm1.txtHconBiz_partner.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtHconItem_cd.value) 
   strVal = strVal & "&txtconSales_unit=" & Trim(frm1.txtHconSales_unit.value)
   strVal = strVal & "&txtconValid_from_dt=" & Trim(frm1.txtHconValid_from_dt.value)   
   strVal = strVal & "&txtconCurrency=" & Trim(frm1.txtHconCurrency.value)
   strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
   strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
  
  Else
            strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                      <%'☜: 비지니스 처리 ASP의 상태 %>
   strVal = strVal & "&txtconBiz_partner=" & Trim(frm1.txtconBiz_partner.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtconItem_cd.value) 
   strVal = strVal & "&txtconSales_unit=" & Trim(frm1.txtconSales_unit.value)
      
   strVal = strVal & "&txtconCurrency=" & Trim(frm1.txtconCurrency.value)
   strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
   strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
  
  End If
  Call RunMyBizASP(MyBizASP, strVal)         <%'☜: 비지니스 ASP 를 가동 %>
 
  DbQuery = True              <%'⊙: Processing is NG%>
 End Function
 
'========================================================================================================
 Function DbSave() 
  Dim lRow
  Dim lGrpCnt
  Dim strVal, strDel
  Dim intInsrtCnt
  Dim TotDocAmt, dblQty, dblPrice, dblOldQty

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
     Case ggoSpread.InsertFlag        <% '☜: 신규 %>
      strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep <% '☜: C=Create, Row위치 정보 %>

      .vspdData.Col = C_Bp_Cd            <% '2 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Item_Cd        <% '3 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
       
      .vspdData.Col = C_Unit                      <% '7 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Cur    <% '7 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Valid_from_dt                <% '6 %>
      strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & Parent.gColSep

      .vspdData.Col = C_Item_Price    <% '7 %>
      strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & Parent.gRowSep
                        

      lGrpCnt = lGrpCnt + 1
  
     Case ggoSpread.UpdateFlag        <% '☜: Update %>
      strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep <% '☜: U=Update, Row위치 정보 %>
      
         .vspdData.Col = C_Bp_Cd        <% '2 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Item_Cd        <% '3 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Unit         <% '7 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Cur    <% '7 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Valid_from_dt                <% '6 %>
      strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & Parent.gColSep

      .vspdData.Col = C_Item_Price    <% '7 %>
      strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & Parent.gRowSep
                        
     
      lGrpCnt = lGrpCnt + 1
 
     Case ggoSpread.DeleteFlag        <% '☜: 삭제 %>
      strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep <% '☜: D=Update, Row위치 정보 %>

      .vspdData.Col = C_Bp_Cd        <% '2 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Item_Cd        <% '3 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Unit        <% '7 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Cur        <% '7 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Valid_from_dt                <% '6 %>
      strDel = strDel & UNIConvDate(Trim(.vspdData.Text)) & Parent.gColSep

      .vspdData.Col = C_Item_Price    <% '7 %>
      strDel = strDel & UNICDbl(Trim(.vspdData.Text)) & Parent.gRowSep
      
      lGrpCnt = lGrpCnt + 1

    End Select
   Next

   .txtMaxRows.value = lGrpCnt-1
   .txtSpread.value = strDel & strVal
   
   Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)      <% '☜: 비지니스 ASP 를 가동 %>

  End With

  DbSave = True              <% '⊙: Processing is NG %>
 End Function
 
 
'========================================================================================================
 Function DbQueryOk()             <% '☆: 조회 성공후 실행로직 %>
  <% '------ Reset variables area ------ %>
  lgIntFlgMode = Parent.OPMD_UMODE           <% '⊙: Indicates that current mode is Update mode %>
  lgBlnFlgChgValue = False
  
  Call ggoOper.LockField(Document, "Q")        <% '⊙: This function lock the suitable field %>
  Call SetToolBar("1110111100111111")         <% '⊙: 버튼 툴바 제어 %>

	If frm1.vspdData.MaxRows > 0 Then 
		frm1.vspdData.Focus		
	Else
		frm1.txtconBiz_partner.focus 
	End If
	
 End Function
 
'========================================================================================================
 Function DbSaveOk()              <%'☆: 저장 성공후 실행 로직 %>
  Call ggoOper.ClearField(Document, "2")
  Call InitVariables
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
     <TD CLASS="CLSLTAB">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>적립금 정보</font></td>
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
         <TD CLASS="TD5" NOWRAP>거래처</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconBiz_partner" ALT="거래처" TYPE="Text" MAXLENGTH=10 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconBiz_partner.value,5">&nbsp;<INPUT NAME="txtconBiz_partner_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="24"></TD>
         <TD CLASS="TD5" NOWRAP>품목</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconItem_cd.value,0">&nbsp;<INPUT NAME="txtconItem_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="24"></TD>
         
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>단위</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconSales_unit" ALT="단위" TYPE="Text" MAXLENGTH=3 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconSales_unit.value,1">&nbsp;<INPUT NAME="txtconSales_unit_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="24"></TD>
         <TD CLASS="TD5" NOWRAP>화폐</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconCurrency" ALT="화폐" TYPE="Text" MAXLENGTH=3 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconCurrency.value,2"></TD>
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
          <script language =javascript src='./js/s1611ma1_I525038880_vspdData.js'></script>
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
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconBiz_partner" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconItem_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconDeal_type" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconPay_meth" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconValid_from_dt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconSales_unit" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconCurrency" tag="24">
</FORM>


<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
