<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : mc700ma1
'*  4. Program Name         : 납입지시마감 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/03/13
'*  8. Modified date(Last)  : 2003/03/13
'*  9. Modifier (First)     : Kang Su Hwan
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*							  
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit		

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "mc700mb1.asp"											
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_Check
Dim C_ProdtOrderNo		'제조오더번호 
Dim C_ItemCd			'품목 
Dim C_ItemNm			'품목명 
Dim C_ItemSpec			'규격 
Dim C_ReqDt				'필요일 
Dim C_DoQty				'납입지시량 
Dim C_RcptQty			'입고수량 
Dim C_BaseUnit			'재고단위 
Dim C_SupplierCd		'공급업체 
Dim C_SupplierNm		'공급업체명 
Dim C_TrackingNo		'TrackingNo
Dim C_OprNo				'공정 
Dim C_PoNo				'발주번호 
Dim C_PoSeqNo			'발주순번 
Dim C_Seq				'부품예약일련번호 
Dim C_SubSeq			'납입지시순번 
Dim C_DoStatusDesc		'납입지시상태 

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim StartDate,EndDate
 
EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("d", -7, EndDate, Parent.gServerDateFormat)
EndDate = UNIDateAdd("d", +7, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop          
'==========================================   Selection()  ======================================
'	Name : Selection()
'	Description : 일괄선택버튼의 Event 합수 
'================================================================================================
Sub Selection()
	Dim index,Count
 
 	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_Check
		
		if frm1.vspdData.Text = "1" then
			frm1.vspdData.Text = "0"
		else
			frm1.vspdData.Text = "1"
		End if
		
		frm1.vspdData.Col = 0 
		
		if ggoSpread.UpdateFlag = frm1.vspdData.Text then    
			frm1.vspdData.Text=""
	    else
	    	ggoSpread.UpdateRow Index
		End if
	Next 
	
	frm1.vspdData.ReDraw = true
End Sub


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE  
    lgBlnFlgChgValue = False   
    lgIntGrpCount = 0          
    lgStrPrevKey = ""          
    lgLngCurRows = 0           
    frm1.vspdData.MaxRows = 0
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Set gActiveElement = document.activeElement
	frm1.txtPlantCd.value = Parent.gPlant
	frm1.txtPlantNm.value = Parent.gPlantNm
	
    frm1.txtFrDt.Text = StartDate
    frm1.txtToDt.Text = EndDate
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True

	Call SetToolbar("1110000000001111")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_Check			= 1
	C_ProdtOrderNo	= 2			'제조오더번호 
	C_ItemCd		= 3			'품목 
	C_ItemNm		= 4			'품목명 
	C_ItemSpec		= 5			'규격 
	C_ReqDt			= 6
	C_DoQty			= 7			'납입지시량 
	C_RcptQty		= 8			'입고수량 
	C_BaseUnit		= 9 		'재고단위 
	C_SupplierCd	= 10		'공급업체 
	C_SupplierNm	= 11		'공급업체명 
	C_TrackingNo	= 12		'TrackingNo
	C_OprNo			= 13		'공정 
	C_PoNo			= 14		'발주번호 
	C_PoSeqNo		= 15		'발주순번 
	C_Seq			= 16		'부품예약일련번호 
	C_SubSeq		= 17		'납입지시순번 
	C_DoStatusDesc  = 18
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData

    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030303",,Parent.gAllowDragDropSpread  

	.ReDraw = false	
	
    .MaxCols = C_DoStatusDesc + 1														
	.Col = .MaxCols:    .ColHidden = True											
    .MaxRows = 0

	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetCheck C_Check, "선택",10,,,true
	ggoSpread.SSSetEdit C_ProdtOrderNo,"제조오더번호",18
	ggoSpread.SSSetEdit C_ItemCd, "품목",18
	ggoSpread.SSSetEdit C_ItemNm, "품목명",20
	ggoSpread.SSSetEdit C_ItemSpec, "품목규격",20
    ggoSpread.SSSetDate C_ReqDt	,"필요일", 10, 2, Parent.gDateFormat
    SetSpreadFloatLocal	C_DoQty, "납입지시량", 15, 1, 3
    SetSpreadFloatLocal	C_RcptQty, "입고수량", 15, 1, 3
    ggoSpread.SSSetEdit C_BaseUnit, "단위", 6
    ggoSpread.SSSetEdit C_SupplierCd, "공급처", 10
    ggoSpread.SSSetEdit C_SupplierNm, "공급처명", 20
    ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 25
	ggoSpread.SSSetEdit C_OprNo,"공정",7
	ggoSpread.SSSetEdit C_PoNo,"발주번호",18
    ggoSpread.SSSetEdit C_PoSeqNo, "발주순번", 10
    ggoSpread.SSSetEdit C_Seq, "부품예약일련번호", 10
    ggoSpread.SSSetEdit C_SubSeq, "납입지시 순번", 10
    ggoSpread.SSSetEdit	C_DoStatusDesc, "납입지시상태", 12
    
	Call ggoSpread.SSSetColHidden(C_Seq,		C_Seq,		True)
	Call ggoSpread.SSSetColHidden(C_SubSeq,		C_SubSeq,		True)

	.ReDraw = true
	
    Call SetSpreadLock 
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock	-1, -1
    ggoSpread.SpreadUnLock	C_Check, -1, C_Check, -1
    
    .vspdData.ReDraw = True

    End With
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

			C_Check			= iCurColumnPos(1)
			C_ProdtOrderNo	= iCurColumnPos(2)			'제조오더번호 
			C_ItemCd		= iCurColumnPos(3)			'품목 
			C_ItemNm		= iCurColumnPos(4)			'품목명 
			C_ItemSpec		= iCurColumnPos(5)			'규격 
			C_ReqDt			= iCurColumnPos(6)			'필요일 
			C_DoQty			= iCurColumnPos(7)			'납입지시량 
			C_RcptQty		= iCurColumnPos(8)			'입고수량 
			C_BaseUnit		= iCurColumnPos(9)		'재고단위 
			C_SupplierCd	= iCurColumnPos(10)		'공급업체 
			C_SupplierNm	= iCurColumnPos(11)		'공급업체명 
			C_TrackingNo	= iCurColumnPos(12)		'TrackingNo
			C_OprNo			= iCurColumnPos(13)		'공정 
			C_PoNo			= iCurColumnPos(14)		'발주번호 
			C_PoSeqNo		= iCurColumnPos(15)		'발주순번 
			C_Seq			= iCurColumnPos(16)		'부품예약일련번호 
			C_SubSeq		= iCurColumnPos(17)		'납입지시순번 
			C_DoStatusDesc  = iCurColumnPos(18)		'납입지시상태 
	End Select

End Sub	

'------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"						
	arrParam(1) = "B_PLANT"      					
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)		
	arrParam(4) = ""								
	arrParam(5) = "공장"						
	
    arrField(0) = "PLANT_CD"						
    arrField(1) = "PLANT_NM"						
    
    arrHeader(0) = "공장"						
    arrHeader(1) = "공장명"						
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim iCalledAspName,IntRetCD
	Dim arrParam(5), arrField(2)

	If IsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then	
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function

'------------------------------------------  OpenSupplier()  -------------------------------------------------
'	Name : OpenSupplier()
'	Description : OpenSupplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"						
	arrParam(1) = "B_BIZ_PARTNER"					

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		frm1.txtSupplierCd.focus
	End If	
End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 그리드의 숫자 부분이 변경된면 이 함수를 변경 해야함.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              'Lot 순번 Maker Lot 순번 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6"				  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
    End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                         
    Call ggoOper.LockField(Document, "N")       
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                        
    Call InitVariables                          
    Call SetDefaultVal
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtSupplierCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If
End Sub

'==========================================================================================
'   Event Name : btnPosting_OnClick()
'   Event Desc : 출고처리 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnSelect_OnClick()
	Dim i
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Check
			frm1.vspdData.Row = i
			frm1.vspdData.value = 1
			Call vspdData_ButtonClicked(C_Check, i, 1)
		Next	
	End If
End Sub

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 출고처리취소 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Check
			frm1.vspdData.Row = i
			frm1.vspdData.value = 0

			Call vspdData_ButtonClicked(C_Check, i, 0)
		Next	
	End If
End Sub

'==========================================================================================
'   Event Name : txtFrDt
'   Event Desc :
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtToDt
'   Event Desc :
'==========================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
<%
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
%>
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_Check And Row > 0 Then
	    Select Case ButtonDown
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True		
	    Case 0

			ggoSpread.Source = frm1.vspdData
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row 
			frm1.vspdData.text = "" 
			lgBlnFlgChgValue = False					
	    End Select
	End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	Call SetPopupMenuItemInf("0000111111")
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
		If lgStrPrevKey <> "" Then							
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    Err.Clear                                               
	
	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    														
   '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then						
       Exit Function
    End If
    
	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) and Trim(.txtFrDt.text)<>"" and Trim(.txtToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")			
			Exit Function
		End if   
	End with
	
	If Not chkConditions Then
		Exit Function
	End If
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
    Set gActiveElement = document.ActiveElement   
    FncQuery = True											
    
End Function

'========================================================================================
' Function Name : chkConditions
' Function Desc : 
'========================================================================================
Function chkConditions()
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	chkConditions=False

	If Len(Trim(frm1.txtPlantCd.Value))  Then
		Call CommonQueryRs(" PLANT_NM ", " B_PLANT ", " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iNameArr = Split(lgF0, Chr(11))
		If Err.number <> 0 Then
			MsgBox Err.Description, vbInformation, parent.gLogoName
			Err.Clear 
			Exit Function
		End If
		If lgF0="" Then 
			Call DisplayMsgBox("970000", "X","공장", "X")	
			Exit Function
		End If
		frm1.txtPlantNm.Value = iNameArr(0)
	End If
	If Len(Trim(frm1.txtSupplierCd.Value)) Then
		Call CommonQueryRs(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD =  " & FilterVar(frm1.txtSupplierCd.Value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iNameArr = Split(lgF0, Chr(11))
		If Err.number <> 0 Then
			MsgBox Err.Description, vbInformation, parent.gLogoName
			Err.Clear 
			Exit Function
		End If
		If lgF0="" Then 
			Call DisplayMsgBox("970000", "X","공급처", "X")	
			Exit Function
		End If
		frm1.txtSupplierNM.Value = iNameArr(0)
	End If
	If Len(Trim(frm1.txtItemCd.Value)) Then
		Call CommonQueryRs(" ITEM_NM ", " B_ITEM ", " ITEM_CD =  " & FilterVar(frm1.txtItemCd.Value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iNameArr = Split(lgF0, Chr(11))
		If Err.number <> 0 Then
			MsgBox Err.Description, vbInformation, parent.gLogoName
			Err.Clear 
			Exit Function
		End If
		If lgF0="" Then 
			Call DisplayMsgBox("970000", "X","품목", "X")	
			Exit Function
		End If
		frm1.txtItemNm.Value = iNameArr(0)
	End If

	chkConditions=True	
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    Err.Clear                                               
    'On Error Resume Next                                   
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call InitVariables                                      
    Call SetDefaultVal
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtSupplierCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If
    FncNew = True                                                           

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       
    
    Err.Clear                                                               
    'On Error Resume Next                                                   
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                         
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then Exit Function
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  
    Call ggoOper.ClearField(Document, "2")                                  
    
    Set gActiveElement = document.ActiveElement   
    FncDelete = True                                                        
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear    
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                                                           
    
  
	ggoSpread.Source = frm1.vspdData                         
    If ggoSpread.SSCheckChange = False Then                  
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")    
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                         
    If Not ggoSpread.SSDefaultCheck  Then              
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
    Set gActiveElement = document.ActiveElement   
    FncSave = True                                     
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
	if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow

    lgBlnFlgChgValue = True
    
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_MULTI)						
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False)                
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")              
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
    
    Err.Clear                                         

	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode = " & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd	=" & .hdnPlantCd.value 
	    strVal = strVal & "&txtSupplier	=" & .hdnSupplier.value 
		strVal = strVal & "&txtFrDt		=" & .hdnFrDt.value
		strVal = strVal & "&txtToDt		=" & .hdnToDt.value
	    strVal = strVal & "&txtItemCd	=" & .hdnItemCd.value 
		strVal = strVal & "&txtMaxRows	=" & frm1.vspdData.MaxRows
	else
	    strVal = BIZ_PGM_ID & "?txtMode	=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd	=" & .txtPlantCd.value 
	    strVal = strVal & "&txtSupplier	=" & .txtSupplierCd.value 
		strVal = strVal & "&txtFrDt		=" & .txtFrDt.Text
		strVal = strVal & "&txtToDt		=" & .txtToDt.Text
	    strVal = strVal & "&txtItemCd	=" & .txtItemCd.value 
		strVal = strVal & "&txtMaxRows	=" & frm1.vspdData.MaxRows
	end if 
	
	Call RunMyBizASP(MyBizASP, strVal)				
        
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
    lgIntFlgMode = Parent.OPMD_UMODE							
    
    Call ggoOper.LockField(Document, "Q")				
	Call SetSpreadLock
	frm1.btnSelect.disabled = False
	frm1.btnDisSelect.disabled = False
	Call SetToolBar("11101000000111")														'⊙: 버튼 툴바 제어 
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim igColSep, igRowSep
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
	
    DbSave = False    
    
    If LayerShowHide(1) = False Then Exit Function
    
	With frm1
		.txtMode.value = Parent.UID_M0002
    
		igColSep = parent.gColSep
		igRowSep = parent.gRowSep

		iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

		ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
		ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

		iTmpCUBufferMaxCount = -1 
		iTmpDBufferMaxCount = -1 
		    
		iTmpCUBufferCount = -1
		iTmpDBufferCount = -1

		strCUTotalvalLen = 0
		strDTotalvalLen  = 0
		lGrpCnt = 1
		strVal = ""

		For lRow = 1 To .vspdData.MaxRows
		    If Trim(GetSpreadText(frm1.vspdData,C_Check,lRow,"X","X")) = "1" Then
				strVal = Trim(GetSpreadText(frm1.vspdData,C_ProdtOrderNo,lRow,"X","X")) & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OprNo,lRow,"X","X")) & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Seq,lRow,"X","X")) & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SubSeq,lRow,"X","X")) & igColSep				
				strVal = strVal & lRow & igRowSep
				lGrpCnt = lGrpCnt + 1
			End If
			
			Select Case Trim(GetSpreadText(frm1.vspdData,C_Check,lRow,"X","X"))
			    Case "1"
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
			End Select   
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)     
		End If   

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					
	
	End With
	
    DbSave = True                                       
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()										
	Call InitVariables()
	Call MainQuery()
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납입지시마감</font></td>
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
					<FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
					    <TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
											   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>공급처</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="공급처"  NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
           									   <INPUT TYPE=TEXT AlT="공급처" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
					</TR>					   
					<TR>
						<TD CLASS=TD5 NOWRAP>필요일</TD>
						<TD CLASS=TD6 NOWRAP>								
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/mc700ma1_fpDateTime1_txtFrDt.js'></script>
									</td>
									<td NOWRAP>&nbsp;~&nbsp;</td>
									<td NOWRAP>
										<script language =javascript src='./js/mc700ma1_fpDateTime1_txtToDt.js'></script>
									</td>
								<tr>
							</table>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">
											   <INPUT TYPE=TEXT Alt="품목" NAME="txtItemNm" SIZE=20 tag="14"></TD>
					</TR>
					</TABLE></FIELDSET></TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%">
						    <script language =javascript src='./js/mc700ma1_I150130379_vspdData.js'></script>
						</TD>
					</TR></TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
	    <td WIDTH="100%">
	    	<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td WIDTH="*" align="left">
					    <button name="btnSelect" class="clsmbtn">일괄선택</button>&nbsp;
					    <BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">일괄선택취소</BUTTON>
					</td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
	    	</table>
	    </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">

</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
