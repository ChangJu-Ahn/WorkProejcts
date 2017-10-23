<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2112MA1
'*  4. Program Name         : 공장별 판매계획조정 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS2G133.dll, PS2G134.dll, PS2G136.dll
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

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                        '☜: Turn on the Option Explicit option.

Dim C_SPDT
Dim C_ITEM_CD
Dim C_ITEM_POP
Dim C_ITEM_NM
Dim C_SPEC
Dim C_PlanCfmQty
Dim C_PlanBunitCfmQty
Dim C_PlanQty
Dim C_PlanBasicQty
Dim C_ReqSts

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim prDBSYSDate
Dim EndDate ,StartDate

prDBSYSDate = "<%=GetSvrDate%>"

EndDate = UniConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s2112mb1.asp"            '☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "s2111ma1"
Const CID_BATCH  = 2580		' 배치 

Dim IsOpenPop				' Popup
Dim lsClickFlag

'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_SPDT = 1           '☆: Spread Sheet의 Column별 상수 
	C_ITEM_CD = 2
	C_ITEM_POP = 3
	C_ITEM_NM = 4
	C_SPEC = 5
	C_PlanCfmQty = 6
	C_PlanBunitCfmQty = 7
	C_PlanQty = 8
	C_PlanBasicQty = 9
	C_ReqSts = 10

End Sub

'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
End Sub

'========================================================================================
Sub SetDefaultVal()
	
	frm1.txtPlanFromDt.Text = StartDate
	frm1.txtPlanToDt.Text = EndDate
	frm1.txtItem_code.focus
	frm1.btnConfirm.disabled = True
	lgBlnFlgChgValue = False
 
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

		ggoSpread.Spreadinit "V20021119",,parent.gAllowDragDropSpread    
		
		.ReDraw = false
			
		.MaxCols = C_ReqSts+1            '☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
	
		 Call GetSpreadColumnPos("A")		
	
		 ggoSpread.SSSetDate C_SPDT,		"계획일", 10,2,parent.gDateFormat
		 ggoSpread.SSSetEdit C_ITEM_CD,		"품목", 15,0,,18,2
		 ggoSpread.SSSetButton C_ITEM_POP
		 ggoSpread.SSSetEdit C_ITEM_NM,		"품목명", 25
		 ggoSpread.SSSetEdit C_SPEC,		"품목규격",20,0		 
		 ggoSpread.SSSetFloat C_PlanCfmQty,	"생산반영수량" ,20,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		 ggoSpread.SSSetFloat C_PlanBunitCfmQty,"생산반영재고수량" ,20,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		 ggoSpread.SSSetFloat C_PlanQty,	"계획수량" ,20,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		 ggoSpread.SSSetFloat C_PlanBasicQty,"재고단위계획수량" ,20,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		 ggoSpread.SSSetEdit C_ReqSts,		"생산요청상태", 15,2
		
		 Call ggoSpread.MakePairsColumn(C_ITEM_CD,C_ITEM_POP)

		  Call ggoSpread.SSSetColHidden(C_ReqSts,C_ReqSts,True)
		  Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
		.ReDraw = true
   
	End With
    
End Sub

'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
		.vspdData.ReDraw = False
		
		ggoSpread.SSSetRequired C_SPDT, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_ITEM_CD, pvStartRow, pvEndRow
		'ggoSpread.SSSetRequired C_ITEM_POP, lRow, lRow
		ggoSpread.SSSetProtected C_ITEM_NM, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanCfmQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanBunitCfmQty, pvStartRow, pvEndRow    
		ggoSpread.SSSetRequired C_PlanQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanBasicQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ReqSts, pvStartRow, pvEndRow
		
		.vspdData.ReDraw = True
    
    End With

End Sub

'======================================================================================================
Sub SetProtectColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
		ggoSpread.SSSetProtected C_SPDT, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ITEM_CD, pvStartRow, pvEndRow
		'ggoSpread.SSSetRequired C_ITEM_POP, lRow, lRow
		ggoSpread.SSSetProtected C_ITEM_NM, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanCfmQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanBunitCfmQty, pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected C_PlanQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanBasicQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ReqSts, pvStartRow, pvEndRow
    
    .vspdData.ReDraw = True
    
    End With

End Sub

'======================================================================================================
Sub SetRequiredColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
		ggoSpread.SSSetProtected C_SPDT, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ITEM_CD, pvStartRow, pvEndRow
		'ggoSpread.SSSetRequired C_ITEM_POP, lRow, lRow
		ggoSpread.SSSetProtected C_ITEM_NM, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanCfmQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanBunitCfmQty, pvStartRow, pvEndRow    
		ggoSpread.SSSetRequired  C_PlanQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanBasicQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ReqSts, pvStartRow, pvEndRow
    
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
	
			C_SPDT				= iCurColumnPos(1)
			C_ITEM_CD			= iCurColumnPos(2)
			C_ITEM_POP			= iCurColumnPos(3)    
			C_ITEM_NM			= iCurColumnPos(4)
			C_SPEC				= iCurColumnPos(5)
			C_PlanCfmQty		= iCurColumnPos(6)
			C_PlanBunitCfmQty	= iCurColumnPos(7)
			C_PlanQty			= iCurColumnPos(8)
			C_PlanBasicQty		= iCurColumnPos(9)
			C_ReqSts			= iCurColumnPos(10)
			
    End Select    
End Sub


'===========================================================================
Function OpenConPlantPopup(ByVal strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	 
	arrParam(0) = "공장"			<%' 팝업 명칭 %>
	arrParam(1) = "b_plant plant"		<%' TABLE 명칭 %>
	arrParam(2) = strCode				<%' Code Condition%>
	arrParam(3) = ""					<%' Name Cindition%>
	arrParam(4) = ""					<%' Where Condition%>  
	arrParam(5) = "공장"			<%' TextBox 명칭 %>
	 
	arrField(0) = "plant.plant_cd"		<%' Field명(0)%>
	arrField(1) = "plant.plant_nm"		<%' Field명(1)%>
	    
	arrHeader(0) = "공장"			<%' Header명(0)%>
	arrHeader(1) = "공장명"			<%' Header명(1)%>
	   
	frm1.txtPlant_code.focus 
	  
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlant_code.value = arrRet(0) 
		frm1.txtPlant_code_nm.value = arrRet(1)
	End If 
 
End Function


'===========================================================================
Function OpenConItemPopup(ByVal strCode)

	 Dim arrRet
	 Dim arrParam(5), arrField(6), arrHeader(6)

	 If IsOpenPop = True Then Exit Function

	 IsOpenPop = True

	 arrParam(0) = "품목"			<%' 팝업 명칭 %>
	 arrParam(1) = "b_item item,b_item_by_plant item_plant"     <%' TABLE 명칭 %>
	 arrParam(2) = strCode				<%' Code Condition%>
	 arrParam(3) = ""					<%' Name Cindition%>
	 arrParam(4) = "item.item_cd=item_plant.item_cd"			<%' Where Condition%>
	 arrParam(5) = "품목"			<%' TextBox 명칭 %>
	  
	 arrField(0) = "item.item_cd"		<%' Field명(0)%>
	 arrField(1) = "item.item_nm"		<%' Field명(1)%>
	 arrField(2) = "item_plant.plant_cd"
	     
	 arrHeader(0) = "품목"			<%' Header명(0)%>
	 arrHeader(1) = "품목명"		<%' Header명(1)%>
	 arrHeader(2) = "공장"
	 
	 frm1.txtItem_code.focus 
		
	 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	   "dialogWidth=520px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	 IsOpenPop = False

	 If arrRet(0) = "" Then
		Exit Function
	 Else
		frm1.txtItem_code.value = arrRet(0) 
		frm1.txtItem_code_nm.value = arrRet(1)   
	 End If 
 
End Function

'===========================================================================
Function OpenItem(ByVal strCode)

	 Dim arrRet
	 Dim arrParam(5), arrField(6), arrHeader(6)

	 If IsOpenPop = True Then Exit Function

	 IsOpenPop = True

	 arrParam(0) = "품목"			<%' 팝업 명칭 %>
	 arrParam(1) = "b_item item,b_item_by_plant item_plant"     <%' TABLE 명칭 %>
	 arrParam(2) = strCode				<%' Code Condition%>
	 arrParam(3) = ""					<%' Name Cindition%>
	 arrParam(4) = "item.item_cd=item_plant.item_cd"			<%' Where Condition%>
	 arrParam(5) = "품목"			<%' TextBox 명칭 %>
	  
	 arrField(0) = "item.item_cd"		<%' Field명(0)%>
	 arrField(1) = "item.item_nm"		<%' Field명(1)%>
	     
	 arrHeader(0) = "품목"			<%' Header명(0)%>
	 arrHeader(1) = "품목명"		<%' Header명(1)%>

	 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	   "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	 IsOpenPop = False

	 If arrRet(0) = "" Then
		Exit Function
	 Else
		frm1.vspdData.Col = C_ITEM_CD
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_ITEM_NM
		frm1.vspdData.Text = arrRet(1)
		Call vspdData_Change(C_ITEM_CD, frm1.vspdData.Row)  <% ' 변경이 읽어났다고 알려줌 %>
	 End If 
 
End Function

'===========================================================================
Sub SetQuerySpreadColor(ByVal lRow)

    With frm1

    .vspdData.ReDraw = False

	ggoSpread.source = frm1.vspdData
	
	If lRow = 0 Then
		For lRow = 1 To .vspdData.MaxRows 
			ggoSpread.SSSetProtected C_SPDT, lRow, lRow
			ggoSpread.SSSetProtected C_ITEM_CD, lRow, lRow
			ggoSpread.SSSetProtected C_ITEM_POP, lRow, lRow
			ggoSpread.SSSetProtected C_ITEM_NM, lRow, lRow
			ggoSpread.SSSetProtected C_PlanCfmQty, lRow, lRow
			ggoSpread.SSSetProtected C_PlanBunitCfmQty, lRow, lRow
			ggoSpread.SSSetProtected C_PlanQty, lRow, lRow
			ggoSpread.SSSetProtected C_PlanBasicQty, lRow, lRow
			ggoSpread.SSSetProtected C_ReqSts, lRow, lRow
		Next
	ElseIf lRow = 1 Then  
		For lRow = 1 To .vspdData.MaxRows 
			ggoSpread.SSSetProtected C_SPDT, lRow, lRow
			ggoSpread.SSSetProtected C_ITEM_CD, lRow, lRow
			ggoSpread.SSSetProtected C_ITEM_POP, lRow, lRow
			ggoSpread.SSSetProtected C_ITEM_NM, lRow, lRow    
			ggoSpread.SSSetProtected C_PlanCfmQty, lRow, lRow
			ggoSpread.SSSetProtected C_PlanBunitCfmQty, lRow, lRow
			
			frm1.vspdData.Col = C_PlanQty
			
			If UNICDbl(frm1.vspdData.Text) > 0 then
				ggoSpread.SSSetRequired C_PlanQty, lRow, lRow
			Else
				ggoSpread.SSSetProtected C_PlanQty, lRow, lRow
			End if
				
			ggoSpread.SSSetProtected C_PlanBasicQty, lRow, lRow
			ggoSpread.SSSetProtected C_ReqSts, lRow, lRow
		Next
	End If

    .vspdData.ReDraw = True
    
    End With

End Sub


'===========================================================================
Sub CookiePage(Byval Kubun)

	On Error Resume Next

	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>

	Dim strTemp, arrVal

	If Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Sub 

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" then Exit Sub

		frm1.txtPlanFromDt.year = arrVal(0)
		frm1.txtPlanFromDt.Month= arrVal(1)
		frm1.txtPlanFromDt.Day = "01"
  
		frm1.txtPlanToDt.text = UNIDateAdd("d",-1, UNIDateAdd("m", 1, frm1.txtPlanFromDt.text, parent.gDateFormat), parent.gDateFormat)
  
		frm1.txtItem_code.value =  arrVal(2) 


		If Err.number <> 0 Then
	
			Err.Clear
			WriteCookie CookieSplit , ""
			
			Exit Sub
	
		End If

		Call OpenConPlantPopup(frm1.txtPlant_code.value)  
		Call MainQuery

		WriteCookie CookieSplit , ""

	End IF
 
End Sub

<% '======================================== BtnSpreadCheck()  ========================================
' Description : Before Button Click, Spread Check Function
'==================================================================================================== %>
Function BtnSpreadCheck()

	BtnSpreadCheck = False

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

	If UNICDbl(frm1.txtPlan_buom_qty.value) = 0 Then
		MsgBox "생산에 반영할 수량이 없습니다", vbExclamation, parent.gLogoName
		Exit Function
	End If

	BtnSpreadCheck = True

End Function

<%
'==================================== 2.5.14 Numeric Check() ===========================================
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


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029()              '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
	Call InitVariables              '⊙: Initializes local global variables
	Call InitSpreadSheet
	Call SetDefaultVal

    Call SetToolbar("11000000000011")          '⊙: 버튼 툴바 제어 

	Call CookiePage(0)

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


<%
'==========================================================================================
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
%>
Sub btnConfirm_OnClick()
	
	Dim strVal

	If BtnSpreadCheck = False Then Exit Sub

	If   LayerShowHide(1) = False Then
	     Exit Sub
	End If
	
	strVal = ""    
	strVal = BIZ_PGM_ID & "?txtMode=" & CID_BATCH         <%'☜: 비지니스 처리 ASP의 상태 %>

	strVal = strVal & "&txtPlanFromYear=" & Trim(frm1.strHFromYear.value)      <%'☜: 조회 조건 데이타 %>
	strVal = strVal & "&txtPlanFromMonth=" & Trim(frm1.strHFromMonth.value)      <%'☜: 조회 조건 데이타 %>
	strVal = strVal & "&txtPlanFromDt=" & Trim(frm1.txtHPlanFromDt.value)
		   
	strVal = strVal & "&txtPlanToYear=" & Trim(frm1.strHToYear.value)       <%'☜: 조회 조건 데이타 %>
	strVal = strVal & "&txtPlanToMonth=" & Trim(frm1.strHToMonth.value)       <%'☜: 조회 조건 데이타 %>
	strVal = strVal & "&txtPlanToDt=" & Trim(frm1.txtHPlanToDt.value)        <%'☜: 조회 조건 데이타 %>



	strVal = strVal & "&txtItem_code=" & Trim(frm1.HItemCd.value)
	strVal = strVal & "&txtPlant_code=" & Trim(frm1.HPlantCd.value)

	strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)
	
	Call RunMyBizASP(MyBizASP, strVal)            <%'☜: 비지니스 ASP 를 가동 %>
 
End Sub

Function btnConfirm_Ok()
	
	Call SetQuerySpreadColor(0)

End Function


'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)
	
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
   
		If Row > 0 And Col = C_Item_Pop Then
		    .Col = Col - 1
		    .Row = Row
			Call OpenItem(.Text)
			Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")
		End If

	End With
		
End Sub


'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
       
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    lgBlnFlgChgValue = True
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    If UNICdbl(frm1.strHTemp.value) <= 0 Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("1111111111")
	End if	

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
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
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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


'==========================================================================================
Sub txtPlanFromDt_DblClick(Button)

	If Button = 1 Then
		frm1.txtPlanFromDt.Action = 7
		Call SetFocusToDocument("M")   
		Frm1.txtPlanFromDt.Focus
	End If

End Sub

Sub txtPlanToDt_DblClick(Button)
	
	If Button = 1 Then
		frm1.txtPlanToDt.Action = 7
		Call SetFocusToDocument("M")   
		Frm1.txtPlanToDt.Focus
	End If
	
End Sub

<%
'==========================================================================================
'   Event Desc : 조회조건부의 OCX_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
%>
Sub txtPlanFromDt_KeyDown(KeyCode, Shift)
	
	If KeyCode = 13 Then Call MainQuery()
	
End Sub

Sub txtPlanToDt_KeyDown(KeyCode, Shift)
	
	If KeyCode = 13 Then Call MainQuery()

End Sub


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

<%    '-----------------------
    ' Valid Check 
    '----------------------- %>
 
	If ValidDateCheck(frm1.txtPlanFromDt, frm1.txtPlanToDt) = False Then Exit Function

<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")          <%'⊙: Clear Contents  Field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>

<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
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
Function FncDelete() 
    
    Exit Function
    Err.Clear                                                               '☜: Protect system from crashing    
    
    FncDelete = False              <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
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
    If Not chkField(Document, "2") Then     <%'⊙: Check contents area%>
		Exit Function
    End If

    If ggoSpread.SSDefaultCheck = False Then     <%'⊙: Check contents area%>
        Exit Function
    End If
    
	<%  '-----------------------
	    'Save function call area
	    '-----------------------%>
    
    CAll DbSave				                                                <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
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
	
   
	ggoSpread.Source = Frm1.vspdData
	
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		 End If
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
	
	If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
	
End Function

'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
		
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
	 
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False
		ggoSpread.InsertRow ,imRow
		.vspdData.ReDraw = True
		SetSpreadColor .vspdData.ActiveRow , .vspdData.ActiveRow + imRow - 1
		   
		lgBlnFlgChgValue = True
		   <% '----------  Coding part  -------------------------------------------------------------%>   
		.vspdData.Col = C_PlanQty
		.vspdData.Text = 0

		.vspdData.Col = C_ReqSts
		.vspdData.Text = "N"

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
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
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
	
	Call SetQuerySpreadColor(1)

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
Function DbDelete() 
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>
End Function

'========================================================================================
Function DbDeleteOk()              <%'☆: 삭제 성공후 실행 로직 %>
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>
End Function

'========================================================================================
Function DbQuery() 

	Err.Clear                                                               <%'☜: Protect system from crashing%>
    
	If   LayerShowHide(1) = False Then
		Exit Function 
	End If
	    
	DbQuery = False                                                         <%'⊙: Processing is NG%>
	    
	Dim strVal
	Dim strFromYear,strFromMonth,strFromDay
	Dim strToYear,strToMonth,strToDay
	Dim strHFromYear,strHFromMonth,strHFromDay
	Dim strHToYear,strHToMonth,strHToDay


	'<조회부날짜>   
	Call ExtractDateFrom(frm1.txtPlanFromDt.text,parent.gDateFormat, parent.gComDateType,strFromYear,strFromMonth,strFromDay)
	Call ExtractDateFrom(frm1.txtPlanToDt.text,parent.gDateFormat, parent.gComDateType,strToYear,strToMonth,strToDay)

	' lgBlnFlgChgValue = True  
	    
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>

		strVal = strVal & "&txtPlanFromYear=" & Trim(frm1.strHFromYear.value)      <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlanFromMonth=" & Trim(frm1.strHFromMonth.value)      <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlanFromDt=" & Trim(frm1.txtHPlanFromDt.value)
		   
		strVal = strVal & "&txtPlanToYear=" & Trim(frm1.strHToYear.value)       <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlanToMonth=" & Trim(frm1.strHToMonth.value)       <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlanToDt=" & Trim(frm1.txtHPlanToDt.value)        <%'☜: 조회 조건 데이타 %>

		strVal = strVal & "&txtItem_code=" & Trim(frm1.HItemCd.value)    <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlant_code=" & Trim(frm1.HPlantCd.value)   <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		  
		'2002-09-02 확정버튼 
		strVal = strVal & "&strTemp=" & Trim(frm1.strHTemp.value)
	 
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>

		strVal = strVal & "&txtPlanFromYear=" & Trim(strFromYear)      <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlanFromMonth=" & Trim(strFromMonth)      <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlanFromDt=" & Trim(frm1.txtPlanFromDt.text)
		   
		strVal = strVal & "&txtPlanToYear=" & Trim(strToYear)       <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlanToMonth=" & Trim(strToMonth)       <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlanToDt=" & Trim(frm1.txtPlanToDt.text)        <%'☜: 조회 조건 데이타 %>

		strVal = strVal & "&txtItem_code=" & Trim(frm1.txtItem_code.value)    <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtPlant_code=" & Trim(frm1.txtPlant_code.value)   <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		  
		strVal = strVal & "&strTemp=" & 0
	   
	End If

	Call RunMyBizASP(MyBizASP, strVal)            <%'☜: 비지니스 ASP 를 가동 %>
	 
	DbQuery = True                 <%'⊙: Processing is NG%>

End Function

'========================================================================================
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>


	lgBlnFlgChgValue = False
	'2002-08-29일 수정 
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtItem_code.focus
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
    
    'On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1

		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		  
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
     
				.vspdData.Col = C_SPDT  
				strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
				              
				.vspdData.Col = C_ITEM_CD  
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				           
				.vspdData.Col = C_PlanQty  
				strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

				.vspdData.Col = C_PlanBasicQty
				strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

				.vspdData.Col = C_ReqSts  
				strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

				lGrpCnt = lGrpCnt + 1
              
			Case ggoSpread.DeleteFlag       '☜: 삭제 

				strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep

				.vspdData.Col = C_SPDT  
				strDel = strDel & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
				              
				.vspdData.Col = C_ITEM_CD  
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep              
				              
				.vspdData.Col = C_PlanQty  
				strDel = strDel & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

				.vspdData.Col = C_PlanBasicQty
				strDel = strDel & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep

				.vspdData.Col = C_ReqSts  
				strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

				lGrpCnt = lGrpCnt + 1
			End Select
		Next
 
		.txtMaxRows.value = lGrpCnt-1
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공장별판매계획</font></td>
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
         <TD CLASS="TD5" NOWRAP>품목</TD>
         <TD CLASS="TD6"><INPUT NAME="txtItem_code" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSalesPlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConItemPopup frm1.txtItem_code.value">&nbsp;<INPUT NAME="txtItem_code_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>공장</TD>
         <TD CLASS="TD6"><INPUT NAME="txtPlant_code" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSalesPlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPlantPopup frm1.txtPlant_code.value">&nbsp;<INPUT NAME="txtPlant_code_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
        </TR>
        <TR>       
         <TD CLASS="TD5" NOWRAP>계획일</TD>
         <TD CLASS="TD6" NOWRAP>
          <TABLE CELLSPACING=0 CELLPADDING=0>
           <TR>
            <TD>
             <script language =javascript src='./js/s2112ma1_fpDateTime1_txtPlanFromDt.js'></script>
            </TD>
            <TD>
             &nbsp;~&nbsp;
            </TD>
            <TD>
             <script language =javascript src='./js/s2112ma1_fpDateTime2_txtPlanToDt.js'></script>
            </TD>
           </TR>
          </TABLE>
         </TD>
         <TD CLASS=TD5 NOWRAP></TD>
         <TD CLASS=TD6 NOWRAP></TD> 
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
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <script language =javascript src='./js/s2112ma1_I713659006_vspdData.js'></script>
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
     <TD><BUTTON NAME="btnConfirm" CLASS="CLSMBTN">확정처리</BUTTON></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtRadio" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="strHFromYear" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="strHFromMonth" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPlanFromDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="strHToYear" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="strHToMonth" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPlanToDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HPlantCd" tag="24" TABINDEX="-1">


<INPUT TYPE=HIDDEN NAME="txtProduct_buom_qty"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPlan_buom_qty"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtProduct_BasicQty"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPlanBasicQty"  tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtPlanQty"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="strHTemp"  tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
  </DIV>
</BODY>
</HTML>
