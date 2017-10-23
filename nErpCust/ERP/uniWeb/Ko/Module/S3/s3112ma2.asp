<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3112ma2.asp	
'*  4. Program Name         : 수주마감 
'*  5. Program Desc         : 수주마감 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s3112mb2.asp"												'☆: Head Query 비지니스 로직 ASP명 

'☆: Spread Sheet의 Column별 상수 
Dim C_Select			'선택 
Dim C_CloseFlag			'마감여부 
Dim C_SoNo				'수주번호 
Dim C_SoSeq				'수주순번		
Dim C_ItemCd			'품목 
Dim C_ItemNm			'품목명 
Dim C_Unit				'단위 
Dim C_SoQty				'수량 
Dim C_SoBonusQty		'덤수량 
Dim C_DlvyDt			'납기일 
Dim C_ShipToParty 		'납품처 
Dim C_ShipToPartyNm		'납품처명 
Dim C_LcQty				'L/C수량 
Dim C_ReqQty			'출고요청량 
Dim C_ReqBonusQty		'덤출고요청량 
Dim C_GiQty				'출고량 
Dim C_GiBonusQty		'덤출고량 
Dim C_BillQty			'매출량 
Dim C_SalesGrp			'영업그룹 
Dim C_SalesGrpNm		'영업그룹명 
Dim C_Plant				'공장 
Dim C_PlantNm			'공장명 
Dim C_ItemSpec			'규격 

Dim IsOpenPop			'Popup

'==================================================================================================================
Sub initSpreadPosVariables()  
	C_Select		= 1	
	C_CloseFlag		= 2	
	C_SoNo			= 3	
	C_SoSeq			= 4	
	C_ItemCd		= 5	
	C_ItemNm		= 6	
	C_Unit			= 7	
	C_SoQty			= 8	
	C_SoBonusQty	= 9	
	C_DlvyDt		= 10
	C_ShipToParty	= 11
	C_ShipToPartyNm = 12
	C_LcQty			= 13
	C_ReqQty		= 14
	C_ReqBonusQty	= 15
	C_GiQty			= 16
	C_GiBonusQty	= 17
	C_BillQty		= 18
	C_SalesGrp		= 19
	C_SalesGrpNm	= 20
	C_Plant			= 21
	C_PlantNm		= 22
	C_ItemSpec		= 23	'규격필드추가	
End Sub

'==================================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub

'==================================================================================================================
Sub SetDefaultVal()
	frm1.txtConSoNo.focus
	lgBlnFlgChgValue = False
	frm1.rdoCfmAll.checked = True
	frm1.rdoStatusAll.checked = True
	frm1.txtCfmFlag.value = frm1.rdoCfmAll.value
	frm1.txtStatusFlag.value = frm1.rdoStatusAll.value
	
	
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True
	
	frm1.txtConSoFrDt.text = StartDate
	frm1.txtConSoToDt.text = EndDate

End Sub

'==================================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub


'==================================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021104",, parent.gAllowDragDropSpread
		.ReDraw = false
	    .MaxCols = C_PlantNm+1													'☜: 최대 Columns의 항상 1개 증가시킴 
	    .Col = .MaxCols															'☜: 공통콘트롤 사용 Hidden Column
	    .ColHidden = True

	    .MaxRows = 0
	    
        Call GetSpreadColumnPos("A")	    

		ggoSpread.SSSetCheck C_Select, "선택", 10,,,true		
		ggoSpread.SSSetEdit C_CloseFlag, "마감여부", 10
		ggoSpread.SSSetEdit C_SoNo, "수주번호", 18,,,,2
		ggoSpread.SSSetEdit C_SoSeq, "수주순번", 10,1
	    ggoSpread.SSSetEdit C_ItemCd, "품목", 18,,,18,2
	    ggoSpread.SSSetEdit C_ItemNm, "품목명", 25,,,40
	    ggoSpread.SSSetEdit C_Unit, "단위", 8,,,,2
		ggoSpread.SSSetFloat C_SoQty,"수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_SoBonusQty,"덤수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetDate C_DlvyDt, "납기일",10,2,Parent.gDateFormat
	    ggoSpread.SSSetEdit C_ShipToParty, "주문처", 10,,,,2
	    ggoSpread.SSSetEdit C_ShipToPartyNm, "주문처명", 20
		ggoSpread.SSSetFloat C_LcQty,"L/C수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_ReqQty,"출고요청량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_ReqBonusQty,"덤출고요청량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_GiQty,"출고량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_GiBonusQty,"덤출고량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_BillQty,"매출량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit C_SalesGrp, "영업그룹",10,,,4,2
	    ggoSpread.SSSetEdit C_SalesGrpNm, "영업그룹명",20
	    ggoSpread.SSSetEdit C_Plant, "공장",10,,,4,2
	    ggoSpread.SSSetEdit C_PlantNm, "공장명",20
		ggoSpread.SSSetEdit C_ItemSpec, "규격", 20
		
		.ReDraw = true
   
    End With
    
End Sub

'==================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetProtected C_CloseFlag, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_SoNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SoSeq, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Unit, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SoQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SoBonusQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DlvyDt, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ShipToParty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ShipToPartyNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_LcQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ReqQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ReqBonusQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_GiQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_GiBonusQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BillQty, pvStartRow, pvEndRow        
    ggoSpread.SSSetProtected C_SalesGrp, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SalesGrpNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Plant, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlantNm, pvStartRow, pvEndRow       
	ggoSpread.SSSetProtected C_ItemSpec, pvStartRow, pvEndRow
	
    .vspdData.ReDraw = True    
    
    End With

End Sub

'==================================================================================================================
Function OpenConSoDtl()

	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD				

	If IsOpenPop = True Then Exit Function			
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("S3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3111PA1", "X")			
		IsOpenPop = False
		Exit Function
	End If
		
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.Parent, ""), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtConSoNo.value = strRet
		frm1.txtConSoNo.focus
	End If	
	
End Function

'==================================================================================================================
Function OpenConSoldTP()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "주문처"								
	arrParam(1) = "B_BIZ_PARTNER"							
	arrParam(2) = Trim(frm1.txtSoldToParty.value)			

	arrParam(4) = "BP_TYPE <= " & FilterVar("CS", "''", "S") & ""							
	arrParam(5) = "주문처"								

	arrField(0) = "BP_CD"									
	arrField(1) = "BP_NM"									
	    
	arrHeader(0) = "주문처"								
	arrHeader(1) = "주문처명"							

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtSoldToParty.value = arrRet(0)
		frm1.txtSoldToPartyNm.value = arrRet(1)
		frm1.txtSoldToParty.focus
	End If	

End Function

'==================================================================================================================
Function OpenConSalesGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "영업그룹"					
	arrParam(1) = "B_SALES_GRP"						
	arrParam(2) = Trim(frm1.txtSalesGrp.value)		
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "영업그룹"					
		
	arrField(0) = "SALES_GRP"						
	arrField(1) = "SALES_GRP_NM"					
	    
	arrHeader(0) = "영업그룹"					
	arrHeader(1) = "영업그룹명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtSalesGrp.value = arrRet(0)
		frm1.txtSalesGrpNm.value = arrRet(1)
		frm1.txtSalesGrp.focus
	End If	

End Function


'==================================================================================================================
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"					
	arrParam(1) = "B_PLANT"						
	arrParam(2) = Trim(frm1.txtPlant.value)		
	arrParam(3) = ""
	arrParam(4) = ""					
	arrParam(5) = "공장"					
		
	arrField(0) = "Plant_cd"						
	arrField(1) = "Plant_NM"					
	    
	arrHeader(0) = "공장"					
	arrHeader(1) = "공장명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlant.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)
		frm1.txtPlant.focus
	End If	

End Function

'==================================================================================================================
Sub SetQuerySpreadColor(ByVal lRow)
	
    With frm1

    .vspdData.ReDraw = False

		ggoSpread.SSSetProtected C_CloseFlag, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoSeq, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemCd, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_Unit, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoBonusQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_DlvyDt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ShipToParty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ShipToPartyNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_LcQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ReqQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ReqBonusQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_GiQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_GiBonusQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_BillQty, lRow, .vspdData.MaxRows	
		ggoSpread.SSSetProtected C_SalesGrp, lRow, .vspdData.MaxRows	
		ggoSpread.SSSetProtected C_SalesGrpNm, lRow, .vspdData.MaxRows	
		ggoSpread.SSSetProtected C_Plant, lRow, .vspdData.MaxRows	
		ggoSpread.SSSetProtected C_PlantNm, lRow, .vspdData.MaxRows	
		ggoSpread.SSSetProtected C_ItemSpec, lRow, .vspdData.MaxRows

    .vspdData.ReDraw = True
    
    End With

End Sub

'==================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_Select		= iCurColumnPos(1)
			C_CloseFlag		= iCurColumnPos(2)
			C_SoNo			= iCurColumnPos(3)
			C_SoSeq			= iCurColumnPos(4)
			C_ItemCd		= iCurColumnPos(5)
			C_ItemNm		= iCurColumnPos(6)
			C_Unit			= iCurColumnPos(7)
			C_SoQty			= iCurColumnPos(8)
			C_SoBonusQty	= iCurColumnPos(9)
			C_DlvyDt		= iCurColumnPos(10)
			C_ShipToParty	= iCurColumnPos(11)
			C_ShipToPartyNm = iCurColumnPos(12)
			C_LcQty			= iCurColumnPos(13)
			C_ReqQty		= iCurColumnPos(14)
			C_ReqBonusQty	= iCurColumnPos(15)
			C_GiQty			= iCurColumnPos(16)
			C_GiBonusQty	= iCurColumnPos(17)
			C_BillQty		= iCurColumnPos(18)
			C_SalesGrp		= iCurColumnPos(19)
			C_SalesGrpNm	= iCurColumnPos(20)
			C_Plant			= iCurColumnPos(21)
			C_PlantNm		= iCurColumnPos(22)
			C_ItemSpec		= iCurColumnPos(23)
    End Select    
End Sub


'==================================================================================================================
Sub Form_Load()

	Call InitVariables														'⊙: Initializes local global variables
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet
	Call SetDefaultVal
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
   
End Sub


'==================================================================================================================
Function rdoCfmAll_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmAll.value
End Function

Function rdoCfmYes_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmYes.value
End Function

Function rdoCfmNo_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmNo.value
End Function

Function rdoStatusAll_OnClick()
	frm1.txtStatusFlag.value = frm1.rdoStatusAll.value
End Function

Function rdoStatusSO_OnClick()
	frm1.txtStatusFlag.value = frm1.rdoStatusSO.value
End Function

Function rdoStatusDN_OnClick()
	frm1.txtStatusFlag.value = frm1.rdoStatusDN.value
End Function

Function rdoBOAll_OnClick()
	frm1.txtBOFlag.value = frm1.rdoBOAll.value
End Function

Function rdoBOYes_OnClick()
	frm1.txtBOFlag.value = frm1.rdoBOYes.value
End Function

Function rdoBONo_OnClick()
	frm1.txtBOFlag.value = frm1.rdoBONo.value
End Function

'==================================================================================================================
Sub txtConSoFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConSoFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtConSoFrDt.Focus
	End If
End Sub
Sub txtConSoToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConSoToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtConSoToDt.Focus
	End If
End Sub

'==================================================================================================================
Sub txtConSoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtConSoToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


'==================================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then           
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
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

'==================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
    	Exit Sub
    End If	
End Sub

'==================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub


'==================================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    If Col <= C_Select Or NewCol <= C_Select Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
End Sub


'==================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown) 
	If Col = C_Select And Row > 0 Then
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


'==========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		
    	If lgStrPrevKey <> "" Then		    							
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub


'==================================================================================================================
Sub btnSelect_OnClick()
	Dim i

	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Select
			frm1.vspdData.Row = i
			frm1.vspdData.value = 1
			Call vspdData_ButtonClicked(C_Select, i, 1)
		Next	
		
	End If	

End Sub

'==================================================================================================================
Sub btnDisSelect_OnClick()
	Dim i	
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Select
			frm1.vspdData.Row = i
			frm1.vspdData.value = 0

			Call vspdData_ButtonClicked(C_Select, i, 0)
		Next	
	End If		
End Sub

'==================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                      
    
    Err.Clear      
	
	If ValidDateCheck(frm1.txtConSoFrDt, frm1.txtConSoToDt) = False Then Exit Function

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			  
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables					

    Call DbQuery	

    FncQuery = True	
        
End Function

'==================================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                         
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 	
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "A")                                      
    Call ggoOper.LockField(Document, "N")                                       
    Call SetDefaultVal
    Call InitVariables															

    Call SetToolbar("11000000000011")									
    

    FncNew = True																

End Function

'==================================================================================================================
Function FncDelete() 
    
    Exit Function
    Err.Clear                                                               '☜: Protect system from crashing    
    
    FncDelete = False														
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")       
        Exit Function
    End If

    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If

    Call ggoOper.ClearField(Document, "A")                                  '⊙: Clear Condition Field
        
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

'==================================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          
        Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") Then
		Exit Function		
    End If
    
    If ggoSpread.SSDefaultCheck = False Then     
       Exit Function
    End If

    CAll DbSave				                                             
    
    FncSave = True                                                          
    
End Function

'==================================================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function


'==================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'==================================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function FncNext() 
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'==================================================================================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_MULTI, False)
End Function

'==================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'==================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub


'==================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()          
	Call ggoSpread.ReOrderingSpreadData()	
	Call SetQuerySpreadColor(1)
End Sub


'==================================================================================================================
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

    FncExit = True
End Function


'==================================================================================================================
Function DbDelete() 
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function DbDeleteOk()														
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function DbQuery() 

    Err.Clear                                                               
    
    DbQuery = False                                                         

	If LayerShowHide(1) = False Then
		Exit Function
	End If
	   
    Dim strVal

''2002-09-24 공장, 영업그룹 추가 
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001								
		strVal = strVal & "&txtConSoNo=" & Trim(frm1.txtHSoNo.value)				
		strVal = strVal & "&txtCfmFlag=" & Trim(frm1.txtHCfmFlag.value)
		strVal = strVal & "&txtStatusFlag=" & Trim(frm1.txtHStatusFlag.value)
		strVal = strVal & "&txtSoldToParty=" & Trim(frm1.txtHSoldToParty.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtHPlant.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtHSalesGrp.value)
		strVal = strVal & "&txtBOFlag=" & Trim(frm1.txtHBOFlag.value)

		strVal = strVal & "&txtConSoFrDt=" & Trim(frm1.txtHSoFrDt.value)
		strVal = strVal & "&txtConSoToDt=" & Trim(frm1.txtHSoToDt.value)						
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001								
		strVal = strVal & "&txtConSoNo=" & Trim(frm1.txtConSoNo.value)				
		strVal = strVal & "&txtCfmFlag=" & Trim(frm1.txtCfmFlag.value)
		strVal = strVal & "&txtStatusFlag=" & Trim(frm1.txtStatusFlag.value)
		strVal = strVal & "&txtSoldToParty=" & Trim(frm1.txtSoldToParty.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtSalesGrp.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlant.value)
		strVal = strVal & "&txtBOFlag=" & Trim(frm1.txtBOFlag.value)
		
		strVal = strVal & "&txtConSoFrDt=" & Trim(frm1.txtConSoFrDt.Text)
		strVal = strVal & "&txtConSoToDt=" & Trim(frm1.txtConSoToDt.Text)		
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If    

	Call RunMyBizASP(MyBizASP, strVal)												
	
    DbQuery = True																	

End Function

'==================================================================================================================
Function DbQueryOk()														
	
    lgIntFlgMode = Parent.OPMD_UMODE	 
							
    Call SetToolbar("11101000000111")					   
	Call SetQuerySpreadColor(1)
	lgBlnFlgChgValue = False
	
	If frm1.vspdData.Maxrows > 0 then
		frm1.btnSelect.disabled = False
		frm1.btnDisSelect.disabled = False
	End If
		
End Function

'==================================================================================================================
Function DbSave() 

    Err.Clear																
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	
    DbSave = False                                                    

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
    
		lGrpCnt = 0    
		strVal = ""
    
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		        
			if Trim(.vspdData.Text) <> "" then
				
				'--- 수주번호 
				.vspdData.Col = C_SoNo				
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				
				'--- 수주순번 
				.vspdData.Col = C_SoSeq 		            
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep 

				'--- 마감여부 
				.vspdData.Col = C_CloseFlag 		
						
				If Trim(.vspdData.Text) = "Y" then
					strVal = strVal & "N" & Parent.gColSep
				else
					strVal = strVal & "Y" & Parent.gColSep
				end if
				
				strVal = strVal & lRow & Parent.gRowSep
		            
			    lGrpCnt = lGrpCnt + 1 
			end if
		Next
	
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'==================================================================================================================
Function DbSaveOk()															

    Call InitVariables

	Call ggoOper.ClearField(Document, "2")										
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수주마감</font></td>
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
									<TD CLASS=TD5 NOWRAP>수주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSoDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSoDtl()"></TD>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>주문처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldToParty" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSoldTP()">&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSalesGrp()">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>									
									<TD CLASS=TD5 NOWRAP>납기일</TD>
									<TD CLASS=TD6 NOWRAP>
										
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtConSoFrDt" CLASS=FPDTYYYYMMDD tag="11X1" ALT="납기시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtConSoToDt" CLASS=FPDTYYYYMMDD tag="11X1" ALT="납기종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>											
										</TABLE>										
									</TD>
									<TD CLASS=TD5 NOWRAP>마감여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoCfmflag" id="rdoCfmAll" value="A" tag = "11X" checked>
											<label for="rdoCfmAll">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoCfmflag" id="rdoCfmYes" value="Y" tag = "11X">
											<label for="rdoCfmYes">마감</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoCfmflag" id="rdoCfmNo" value="N" tag = "11X">
											<label for="rdoCfmNo">진행</label>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>진행단계</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusAll" value="A" tag = "11X" checked>
											<label for="rdoStatusAll">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusSO" value="S" tag = "11X">
											<label for="rdoStatusSO">수주</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoStatusflag" id="rdoStatusDN" value="D" tag = "11X">
											<label for="rdoStatusDN">출고요청</label>
									</TD>
									<TD CLASS=TD5 NOWRAP>잔량여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoBOflag" id="rdoBOAll" value="A" tag = "11X" checked>
											<label for="rdoBOAll">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoBOflag" id="rdoBOYes" value="Y" tag = "11X">
											<label for="rdoBOYes">있음</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoBOflag" id="rdoBONo" value="N" tag = "11X">
											<label for="rdoBONo">없음</label>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnSelect" CLASS="CLSMBTN">일괄선택</BUTTON>&nbsp;
						<BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">일괄선택취소</BUTTON>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX = -1></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"  TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = -1>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoNo" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtCfmFlag" tag="14" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHCfmFlag" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24"TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHPlant" tag="24"TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoFrDt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoToDt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtStatusFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHStatusFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtBOFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHBOFlag" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
