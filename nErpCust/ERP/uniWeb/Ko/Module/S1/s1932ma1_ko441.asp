<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : s1932ma1_ko441
'*  4. Program Name         : 단가처리config등록(품목) 
'*  5. Program Desc         : 단가처리config등록(품목) 
'*  6. Comproxy List        : PS1G103.dll, PS1G104.dll
'*  7. Modified date(First) : 2008/07/29
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'=======================================================================================================
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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                 '☜: indicates that All variables must be declared in advance

Dim prDBSYSDate

Dim EndDate ,StartDate
Dim lgStrComDateType		'Company Date Type을 저장(년월 Mask에 사용함.)

prDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company

StartDate = UniDateAdd("m", -1, EndDate,parent.gDateFormat)


Const BIZ_PGM_ID      = "s1932mb1_ko441.asp"            '☆: 비지니스 로직 ASP명 

Dim C_Apply_yymm
Dim C_Item_Cd
Dim C_Item_Cd_Popup
Dim C_Item_Nm
Dim C_Cur
Dim C_Cur_Popup
Dim C_Apply_Opt
Dim C_Apply_Opt_nm
Dim C_Retest_Rate
Dim C_Efr_Rate
Dim C_Qei_Rate
Dim C_Ad_Sec
Dim C_Basic_Opt
Dim C_Basic_Opt_Nm
Dim C_Operator
Dim C_Output_rate
Dim C_Rate_Opt
Dim C_Rate_Opt_Nm

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim gblnWinEvent

'========================================================================================================
Sub initSpreadPosVariables()  
    Dim i 
	i = 1

	C_Apply_yymm		  = i : i = i + 1
	C_Item_Cd               = i : i = i + 1
	C_Item_Cd_Popup         = i : i = i + 1
	C_Item_Nm               = i : i = i + 1
	C_Cur                 = i : i = i + 1
	C_Cur_Popup           = i : i = i + 1
	C_Apply_Opt           = i : i = i + 1
	C_Apply_Opt_nm        = i : i = i + 1
	C_Retest_Rate         = i : i = i + 1 
	C_Efr_Rate			  = i : i = i + 1
	C_Qei_Rate            = i : i = i + 1
	C_Ad_Sec			  = i : i = i + 1
	C_Basic_Opt           = i : i = i + 1
	C_Basic_Opt_Nm        = i : i = i + 1
	C_Operator			  = i : i = i + 1
	C_Output_rate         = i : i = i + 1
	C_Rate_Opt			  = i : i = i + 1
	C_Rate_Opt_Nm		  = i : i = i + 1

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
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()
	Dim strMaskYM
	Dim ii, Tempval

	If Date_DefMask(strMaskYM) = False Then
		strMaskYM = "9999" & lgStrComDateType & "99"
	End If	
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
 	With frm1.vspdData

       ggoSpread.Spreadinit "V20050503",,parent.gAllowDragDropSpread    

       .MaxCols   = C_Rate_Opt_Nm														' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True
       .MaxRows = 0                                                                  ' ☜: Clear spreadsheet data 
      Call GetSpreadColumnPos("A")
	   .ReDraw = false
 			      
		ggoSpread.SSSetMask	    C_Apply_yymm,			"적용년월",   9,2, strMaskYM
		ggoSpread.SSSetEdit     C_Item_Cd,              "품목코드" ,12, 0,,10,2
		ggoSpread.SSSetButton   C_Item_Cd_Popup    
		ggoSpread.SSSetEdit     C_Item_Nm,              "품목명", 22, 0 
		ggoSpread.SSSetEdit     C_Cur,                  "단가환종", 8, 0,,3,2
		ggoSpread.SSSetButton   C_Cur_Popup 
		ggoSpread.SSSetCombo    C_Apply_Opt,            "적용옵션", 10, 0    
		ggoSpread.SSSetCombo    C_Apply_Opt_Nm,         "적용옵션", 10, 0
		ggoSpread.SSSetFloat    C_Retest_Rate,          "Retest"&Chr(13)&"비중(%)",7, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,,"0","999"	
		ggoSpread.SSSetFloat    C_Efr_Rate,				"EFR비중"&Chr(13)&"(CP)",7, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,,"0","999"
		ggoSpread.SSSetFloat    C_Qei_Rate,				"QEI비중"&Chr(13)&"(%)",7, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,,"0","999"	
		ggoSpread.SSSetFloat    C_Ad_Sec,				"Ad.Sec",8, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	

		ggoSpread.SSSetCombo    C_Basic_Opt,			"기초단가Opt", 12, 0     
		ggoSpread.SSSetCombo    C_Basic_Opt_Nm,			"기초단가Opt", 12, 0       
		ggoSpread.SSSetCombo    C_Operator,				"수율"&Chr(13)&"계산자", 7, 0        
		ggoSpread.SSSetFloat    C_Output_rate,          "수율(%)",8, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,,"0","999"
		ggoSpread.SSSetCombo    C_Rate_Opt,				"환율Option", 14, 0       
		ggoSpread.SSSetCombo    C_Rate_Opt_Nm,			"환율Option", 14, 0       
		     
		.ReDraw = true   

		call ggoSpread.MakePairsColumn(C_Item_Cd,C_Item_Cd_Popup)
		call ggoSpread.MakePairsColumn(C_Cur,C_Cur_Popup)

		Call ggoSpread.SSSetColHidden(C_Apply_Opt,C_Apply_Opt,True)
		Call ggoSpread.SSSetColHidden(C_Basic_Opt,C_Basic_Opt,True)
		Call ggoSpread.SSSetColHidden(C_Rate_Opt,C_Rate_Opt,True)


	    .ColHeaderRows = 2
 		For ii = 8 to .MaxCols
			.Col = ii
			.Row = 0
			TempVal =  .text
			.row = 1
			.Text = Tempval
			.Row = 0
			.Text =  ""
		Next 

		Call .AddCellSpan(0,-1000, 1, 2)					'컬럼제목을 -1000 은 합치고 -1000+1 은 놔둔다
		Call .AddCellSpan(1,-1000, 1, 2) 
		Call .AddCellSpan(2,-1000, 1, 2) 
		Call .AddCellSpan(3,-1000, 1, 2) 
		Call .AddCellSpan(4,-1000, 1, 2) 
		Call .AddCellSpan(5,-1000, 1, 2) 
		Call .AddCellSpan(6,-1000, 1, 2) 
		Call .AddCellSpan(7,-1000, 1, 2) 

		Call .AddCellSpan(8,-1000, 5, 1)
		.Row = -1000 : .Col = 8 : .Text = "Test 단가처리 Rule"
		Call .AddCellSpan(13,-1000, 6, 1)
		.Row = -1000 : .Col = 13 : .Text = "TCP/COF 단가처리 Rule"

		.RowHeight(-1000+1) = 20

		Call SetSpreadLock 
	    
	End With
    
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    
	ggoSpread.spreadlock		C_Apply_yymm, -1, C_Item_Nm
	ggoSpread.SSSetRequired		C_Cur,  -1, C_Cur
	ggoSpread.SSSetRequired		C_Apply_Opt_Nm,  -1, C_Apply_Opt_Nm
	ggoSpread.SSSetRequired		C_Retest_Rate,  -1, C_Retest_Rate
	ggoSpread.SSSetRequired		C_Efr_Rate,  -1, C_Efr_Rate
	ggoSpread.SSSetRequired		C_Qei_Rate,  -1, C_Qei_Rate
	ggoSpread.SSSetRequired		C_Ad_Sec,  -1, C_Ad_Sec
	ggoSpread.SSSetRequired		C_Basic_Opt_Nm,  -1, C_Basic_Opt_Nm
	ggoSpread.SSSetRequired		C_Operator,  -1, C_Operator
	ggoSpread.SSSetRequired		C_Output_rate,  -1, C_Output_rate
	ggoSpread.SSSetRequired		C_Rate_Opt_Nm,  -1, C_Rate_Opt_Nm
    
    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired    C_Apply_yymm,         pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Item_Cd,              pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Item_Nm  ,            pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Cur,                pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Apply_Opt_Nm,       pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Retest_Rate,        pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Efr_Rate  ,         pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Qei_Rate  ,         pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Ad_Sec  ,			 pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Basic_Opt_Nm,       pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Operator,			 pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Output_rate,		 pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Rate_Opt_Nm,        pvStartRow, pvEndRow
		
    .vspdData.ReDraw = True
    
    End With

End Sub

'========================================================================================
Sub SetSpreadColor1(ByVal lRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetProtected    C_Item_Cd,              lRow, lRow
    ggoSpread.SSSetProtected    C_Item_Nm  ,            lRow, lRow
    ggoSpread.SSSetProtected    C_Retest_Rate,            lRow, lRow
    ggoSpread.SSSetProtected    C_Qei_Rate  ,          lRow, lRow
    ggoSpread.SSSetProtected    C_Ad_Sec  ,     lRow, lRow
    ggoSpread.SSSetProtected    C_Apply_Opt,          lRow, lRow
    ggoSpread.SSSetProtected    C_Apply_Opt_Nm,       lRow, lRow
	ggoSpread.SSSetProtected    C_Operator,      lRow, lRow
	ggoSpread.SSSetProtected    C_Output_rate,               lRow, lRow
	ggoSpread.SSSetProtected    C_Cur,                lRow, lRow
    
    .vspdData.ReDraw = True    
'	.vspdData.Col = C_Item_Cd 
'	.vspdData.Row = .vspdData.ActiveRow
'	.vspdData.Action = 0
'	.vspdData.EditMode = True

    End With

End Sub


'========================================================================================================
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
Dim i,j
Dim ArrMask,StrComDateType

	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split(parent.gDateFormat,parent.gComDateType)

	If parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType = parent.gComDateType
	End If
		
	If IsArray(ArrMask) Then
		For i=0 To Ubound(ArrMask)		
			If Instr(UCase(ArrMask(i)),"D") = False Then
				If strMaskYM <> "" Then
					strMaskYM = strMaskYM & lgStrComDateType
				End If
				If Instr(UCase(ArrMask(i)),"M") And Len(ArrMask(i)) >= 3 Then
					strMaskYM = strMaskYM & "U"
					For j=0 To Len(ArrMask(i)) - 2
						strMaskYM = strMaskYM & "L"
					Next
				Else
					strMaskYM = strMaskYM & ArrMask(i)
				End If
			End If
		Next		
	Else
		Date_DefMask = False
		Exit Function
	End If	

	strMaskYM = Replace(UCase(strMaskYM),"Y","9")
	strMaskYM = Replace(UCase(strMaskYM),"M","9")

	Date_DefMask = True 

End Function
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	Dim i 
	i = 1
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Apply_yymm		  = iCurColumnPos(i) : i = i + 1
			C_Item_Cd             = iCurColumnPos(i) : i = i + 1
			C_Item_Cd_Popup       = iCurColumnPos(i) : i = i + 1
			C_Item_Nm             = iCurColumnPos(i) : i = i + 1
			C_Cur                 = iCurColumnPos(i) : i = i + 1
			C_Cur_Popup           = iCurColumnPos(i) : i = i + 1
			C_Apply_Opt           = iCurColumnPos(i) : i = i + 1
			C_Apply_Opt_nm        = iCurColumnPos(i) : i = i + 1
			C_Retest_Rate         = iCurColumnPos(i) : i = i + 1 
			C_Efr_Rate			  = iCurColumnPos(i) : i = i + 1
			C_Qei_Rate            = iCurColumnPos(i) : i = i + 1
			C_Ad_Sec			  = iCurColumnPos(i) : i = i + 1
			C_Basic_Opt           = iCurColumnPos(i) : i = i + 1
			C_Basic_Opt_Nm        = iCurColumnPos(i) : i = i + 1
			C_Operator			  = iCurColumnPos(i) : i = i + 1
			C_Output_rate         = iCurColumnPos(i) : i = i + 1
			C_Rate_Opt			  = iCurColumnPos(i) : i = i + 1
			C_Rate_Opt_Nm		  = iCurColumnPos(i) : i = i + 1
    End Select    
End Sub
'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()
    Dim strCD
    Dim strVal		    

	ggoSpread.Source = frm1.vspdData

    Call CommonQueryRs(" ud_MINOR_CD, ud_MINOR_NM "," b_user_defined_minor "," ud_MAJOR_CD = 'ZZ501' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    ggoSpread.SetCombo Replace(lgF0 ,Chr(11),vbTab), C_Apply_Opt
    ggoSpread.SetCombo Replace(lgF1 ,Chr(11),vbTab), C_Apply_Opt_Nm

    Call CommonQueryRs(" ud_MINOR_CD, ud_MINOR_NM "," b_user_defined_minor "," ud_MAJOR_CD = 'ZZ502' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    ggoSpread.SetCombo Replace(lgF0 ,Chr(11),vbTab), C_Basic_Opt
    ggoSpread.SetCombo Replace(lgF1 ,Chr(11),vbTab), C_Basic_Opt_Nm

    Call CommonQueryRs(" ud_MINOR_CD, ud_MINOR_NM "," b_user_defined_minor "," ud_MAJOR_CD = 'ZZ503' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    ggoSpread.SetCombo Replace(lgF0 ,Chr(11),vbTab), C_Rate_Opt
    ggoSpread.SetCombo Replace(lgF1 ,Chr(11),vbTab), C_Rate_Opt_Nm

	
	strCD = "*" & vbTab & "/" 
    ggoSpread.SetCombo Replace(strCD ,Chr(11),vbTab), C_Operator

    
End Sub


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
'===========================================================================
Function OpenConSItemDC(Byval strCode, Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If gblnWinEvent = True Then Exit Function

 gblnWinEvent = True

 Select Case iWhere
 Case 5
  arrParam(1) = "B_BIZ_PARTNER"      <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode) <%' Code Condition%>
  arrParam(3) = ""                                    <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"             <%' Where Condition%>
  arrParam(5) = "고객"          <%' TextBox 명칭 %>
 
  arrField(0) = "BP_CD"           <%' Field명(0)%>
  arrField(1) = "BP_NM"           <%' Field명(1)%>
    
  arrHeader(0) = "고객"          <%' Header명(0)%>
  arrHeader(1) = "고객명"            <%' Header명(1)%>
  frm1.txtconBp_cd.focus 
 Case 0
  arrParam(1) = "b_item"                           <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)  <%' Code Condition%>
  arrParam(3) = ""                           <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "품목"       <%' TextBox 명칭 %>
 
  arrField(0) = "item_cd"        <%' Field명(0)%>
  arrField(1) = "item_nm"        <%' Field명(1)%>
  arrField(2) = "spec"        <%' Field명(1)%>
    
  arrHeader(0) = "품목"       <%' Header명(0)%>
  arrHeader(1) = "품목명"       <%' Header명(1)%> 
  arrHeader(2) = "규격"       <%' Header명(1)%>  
  frm1.txtconItem_cd.focus 

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
	 With frm1
	  Select Case iWhere
	  Case 0
	   .txtconItem_cd.value = arrRet(0) 
	   .txtconItem_nm.value = arrRet(1)   
	  Case 5
	   .txtconBp_cd.value = arrRet(0) 
	   .txtconBp_Nm.value = arrRet(1)    
	  End Select
	 End With
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
	 With frm1
	  .vspdData.Col = C_Item_cd
	  .vspdData.Text = arrRet(0)
	  .vspdData.Col = C_Item_nm
	  .vspdData.Text = arrRet(1)
	 End With
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
'	arrField(3) = "HH" & parent.gColSep & "BasiC_Output_rate"	        <%' Field명(3)%>

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
	 With frm1
	  .vspdData.Col = C_Item_cd
	  .vspdData.Text = arrRet(0)
	  .vspdData.Col = C_Item_Nm
	  .vspdData.Text = arrRet(1)
	 End With
  End If
 End Function

'===========================================================================
 Function  OpenCur(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
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
	  .vspdData.Col = C_Cur
	  .vspdData.Text = arrRet(0) 
	 End With
'	 Call vspdData_Change(C_Cur,frm1.vspdData.ActiveRow)

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
			Case C_Item_Cd_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenItem_cd (.text)
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
	
	With frm1.vspdData
		
		.Row = Row
        Select Case Col
            Case C_Apply_Opt_Nm
                .Col = Col
                intIndex = .Value
				.Col = C_Apply_Opt
				.Value = intIndex
            Case C_Basic_Opt_Nm
                .Col = Col
                intIndex = .Value
				.Col = C_Basic_Opt
				.Value = intIndex
            Case C_Rate_Opt_Nm
                .Col = Col
                intIndex = .Value
				.Col = C_Rate_Opt
				.Value = intIndex
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow

			.Col = C_Apply_Opt
			intIndex = .value
			.col = C_Apply_Opt_Nm
			.value = intindex

			.Col = C_Rate_Opt
			intIndex = .value
			.col = C_Rate_Opt_Nm
			.value = intindex

			.Col = C_Basic_Opt
			intIndex = .value
			.col = C_Basic_Opt_Nm
			.value = intindex
    	Next	
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
   Dim iDx
   Dim IntRetCd

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

    Select Case Col
        Case  C_Cur

         Case  C_Item_Cd
            iDx = Frm1.vspdData.value
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_Item_Nm
                Frm1.vspdData.value = ""
            Else
                IntRetCd = CommonQueryRs(" item_nm "," b_item "," item_cd =  " & FilterVar(iDx , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
                If IntRetCd = false then
			        Call DisplayMsgBox("800054","X","X","X")	'거래처정보에 없음
  	                Frm1.vspdData.Col = C_Item_Nm
                    Frm1.vspdData.value = ""
                Else
					Frm1.vspdData.Col = C_Item_Nm
					Frm1.vspdData.Text = Trim(Replace(lgF0, Chr(11), ""))
                End if 
            End if 
	
	End Select    
   
   Call CheckMinNumSpread(frm1.vspdData, Col, Row)
   ggoSpread.Source = frm1.vspdData
   ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
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
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
 End Function

'========================================================================================================
 Function FncCancel() 
   
  if frm1.vspdData.maxrows < 1 then exit function
  
  frm1.vspdData.ReDraw = False    
  ggoSpread.Source = frm1.vspdData
  ggoSpread.EditUndo              <%'☜: Protect system from crashing%>
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
	Call SetSpreadColor1(-1)
	Call InitData()

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
   strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtHconItem_cd.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
   strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows  
     
  Else
   strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                      <%'☜: 비지니스 처리 ASP의 상태 %>
   strVal = strVal & "&txtconBp_cd=" & Trim(frm1.txtconBp_cd.value)        <%'☆: 조회 조건 데이타 %>
   strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtconItem_cd.value)        <%'☆: 조회 조건 데이타 %>
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
										  strVal = strVal & .txtconBp_cd.value & parent.gColSep
		.vspdData.Col = C_Apply_yymm    : strVal = strVal & Trim(Replace(.vspdData.Text,"-","")) & parent.gColSep
		.vspdData.Col = C_Item_Cd		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Cur    		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Apply_Opt	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Retest_Rate	: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Efr_Rate		: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Qei_Rate		: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Ad_Sec		: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Basic_Opt		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Operator		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Output_rate	: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Rate_Opt      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		lGrpCnt = lGrpCnt + 1
  
     Case ggoSpread.UpdateFlag      
										  strVal = strVal & "U" & parent.gColSep 
										  strVal = strVal & lRow & parent.gColSep
										  strVal = strVal & .txtconBp_cd.value & parent.gColSep
		.vspdData.Col = C_Apply_yymm    : strVal = strVal & Trim(Replace(.vspdData.Text,"-","")) & parent.gColSep
		.vspdData.Col = C_Item_Cd		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Cur    		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Apply_Opt	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Retest_Rate	: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Efr_Rate		: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Qei_Rate		: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Ad_Sec		: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Basic_Opt		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Operator		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Output_rate	: strVal = strVal & Trim(UNIConvNum(.vspdData.Text, 0)) & parent.gColSep
		.vspdData.Col = C_Rate_Opt      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
     
		lGrpCnt = lGrpCnt + 1
 
     Case ggoSpread.DeleteFlag  
										  strDel = strDel & "D" & parent.gColSep 
										  strDel = strDel & lRow & parent.gColSep
										  strDel = strDel & .txtconBp_cd.value & parent.gColSep
		.vspdData.Col = C_Apply_yymm    : strDel = strDel & Trim(Replace(.vspdData.Text,"-","")) & parent.gColSep
		.vspdData.Col = C_Item_Cd		: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

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

  lgIntFlgMode = Parent.OPMD_UMODE           <% '⊙: Indicates that current mode is Update mode %>
  lgBlnFlgChgValue = False
  Call InitData()
  
  Call ggoOper.LockField(Document, "Q")        <% '⊙: This function lock the suitable field %>
  Call SetToolBar("1110111100111111") 
  frm1.vspdData.Focus     

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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>단가처리Config등록(품목별)</font></td>
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
         <TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="고객사" TYPE="Text" MAXLENGTH=10 SiZE=10  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconBp_cd.value,5">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>품목</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconItem_cd.value,0">&nbsp;<INPUT NAME="txtconItem_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
        </TR>
<!--        <TR>
         <TD CLASS="TD5" NOWRAP>적용일</TD>
         <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtconValid_from_dt" CLASS=FPDTYYYYMMDD tag="11X1X" ALT="적용일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>       
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
<INPUT TYPE=HIDDEN NAME="txtHconBp_cd" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconItem_cd" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconValid_from_dt" tag="24" TABINDEX = -1>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>