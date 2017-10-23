<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2511PA1
'*  4. Program Name         : 
'*  5. Program Desc         : 검사의뢰현황 팝업 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_ID = "q2511pb1.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_InspReqNo
Dim C_ItemCd
Dim C_ItemNm
Dim C_BpCd 
Dim C_BpNm 
Dim C_WcCd 
Dim C_WcNm 
Dim C_InspReqDt
Dim C_LotNo 
Dim C_LotSubNo
Dim C_LotSize
Dim C_PorNo 
Dim C_PorSeq
Dim C_DocumentNo
Dim C_DocumentSeqNo

Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 

Dim hItemCd
Dim hInspReqNo
Dim hBpCd
Dim hCustCd
Dim hWcCd
Dim hFrInspReqDt
Dim hToInspReqDt
Dim ArrParent

Dim arrParam					 '--- First Parameter Group 
ReDim arrParam(4)
Dim arrReturn				 '--- Return Parameter Group 

Dim IsOpenPop          

'------ Set Parameters from Parent ASP ------ 
ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
arrParam(0) = ArrParent(1)
arrParam(1) = ArrParent(2)
arrParam(2) = ArrParent(3)
arrParam(3) = ArrParent(4)

top.document.title = PopupParent.gActivePRAspName
'top.document.title = "검사의뢰현황 팝업"
 '--------------------------------------------- 

Function InitVariables()
	lgSortKey    = 1                            '⊙: initializes sort direction
End Function

Sub initSpreadPosVariables()  
    C_InspReqNo = 1
    C_ItemCd = 2
    C_ItemNm = 3
    C_BpCd = 4
    C_BpNm = 5
    C_WcCd = 6
    C_WcNm = 7
    C_InspReqDt = 8
    C_LotNo = 9
    C_LotSubNo = 10
    C_LotSize = 11
    C_PorNo = 12
    C_PorSeq = 13
    C_DocumentNo = 14
    C_DocumentSeqNo = 15
End Sub

Sub SetDefaultVal()
	txtPlantCd.Value = arrParam(0)
	txtPlantNm.Value = arrParam(1)
	txtInspReqNo.Value = arrParam(2)
	cboInspClassCd.Value = arrParam(3)
	
	Self.Returnvalue = Array("")
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q","NOCOOKIE","PA") %>
End Sub

Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(cboInspClassCd , lgF0, lgF1, Chr(11))

End Sub

Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021216",,PopupParent.gAllowDragDropSpread    

	vspdData.ReDraw = False

	vspdData.MaxCols = C_DocumentSeqNo + 1
	vspdData.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit C_InspReqNo,"검사의뢰번호",20
	ggoSpread.SSSetEdit C_ItemCd,"품목코드",15
	ggoSpread.SSSetEdit C_ItemNm,"품목명",15

	Select Case cboInspClassCd.Value
		Case "R" 
			ggoSpread.SSSetEdit C_BpCd,"공급처코드",15
			ggoSpread.SSSetEdit C_BpNm,"공급처명",15	
			ggoSpread.SSSetEdit C_WcCd,"작업장코드",10
			ggoSpread.SSSetEdit C_WcNm,"작업장명",15
			
			Call ggoSpread.SSSetColHidden(C_WcCd, C_WcCd, True)
			Call ggoSpread.SSSetColHidden(C_WcNm, C_WcNm, True)
		Case "P"
			ggoSpread.SSSetEdit C_BpCd,"공급처코드",15
			ggoSpread.SSSetEdit C_BpNm,"공급처명",15	
			ggoSpread.SSSetEdit C_WcCd,"작업장코드",10
			ggoSpread.SSSetEdit C_WcNm,"작업장명",15
			
			Call ggoSpread.SSSetColHidden(C_BpCd, C_BpCd, True)
			Call ggoSpread.SSSetColHidden(C_BpNm, C_BpNm, True)	
		Case "F"
			ggoSpread.SSSetEdit C_BpCd,"거래처코드",15
			ggoSpread.SSSetEdit C_BpNm,"거래처명",15
			ggoSpread.SSSetEdit C_WcCd,"작업장코드",10
			ggoSpread.SSSetEdit C_WcNm,"작업장명",15
			
			Call ggoSpread.SSSetColHidden(C_BpCd, C_BpCd, True)
			Call ggoSpread.SSSetColHidden(C_BpNm, C_BpNm, True)
			Call ggoSpread.SSSetColHidden(C_WcCd, C_WcCd, True)
			Call ggoSpread.SSSetColHidden(C_WcNm, C_WcNm, True)
		Case "S"
			ggoSpread.SSSetEdit C_BpCd,"거래처코드",15
			ggoSpread.SSSetEdit C_BpNm,"거래처명",15
			ggoSpread.SSSetEdit C_WcCd,"작업장코드",10
			ggoSpread.SSSetEdit C_WcNm,"작업장명",15
			
			Call ggoSpread.SSSetColHidden(C_WcCd, C_WcCd, True)
			Call ggoSpread.SSSetColHidden(C_WcNm, C_WcNm, True)
	End Select
	
	ggoSpread.SSSetEdit C_InspReqDt,"검사의뢰일",10, 2
	ggoSpread.SSSetEdit C_LotNo,"LOT NO",12
	ggoSpread.SSSetEdit C_LotSubNo,"LOT SUB NO", 5, 1
	ggoSpread.SSSetFloat C_LotSize,"LOT 크기",10, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	ggoSpread.SSSetEdit C_PorNo,"발주번호",10
	ggoSpread.SSSetEdit C_PorSeq,"순번", 5, 1
	ggoSpread.SSSetEdit C_DocumentNo,"수불번호",10
	ggoSpread.SSSetEdit C_DocumentSeqNo,"순번", 5, 1
	
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	vspdData.ReDraw = True
	
	Call SetSpreadLock()
End Sub

Sub SetSpreadLock()	
    ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()	
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_InspReqNo		= iCurColumnPos(1)
			C_ItemCd		= iCurColumnPos(2)
			C_ItemNm		= iCurColumnPos(3)
			C_BpCd			= iCurColumnPos(4)
			C_BpNm			= iCurColumnPos(5)
			C_WcCd			= iCurColumnPos(6)
			C_WcNm			= iCurColumnPos(7)
			C_InspReqDt		= iCurColumnPos(8)
			C_LotNo			= iCurColumnPos(9)
			C_LotSubNo		= iCurColumnPos(10)
			C_LotSize		= iCurColumnPos(11)
			C_PorNo			= iCurColumnPos(12)
			C_PorSeq		= iCurColumnPos(13)
			C_DocumentNo	= iCurColumnPos(14)
			C_DocumentSeqNo	= iCurColumnPos(15)
			
    End Select    

End Sub

Sub ProtectField()
	
	Select Case cboInspClassCd.Value
		Case "R"
			txtBpCd.Tag = "11"
			txtCustCd.Tag = "14"
			txtWcCd.Tag = "14"
			
		Case "P"
			txtBpCd.Tag = "14"
			txtCustCd.Tag = "14"
			txtWcCd.Tag = "11"
			
		Case "F"
			txtBpCd.Tag = "14"
			txtCustCd.Tag = "14"
			txtWcCd.Tag = "14"
			
		Case "S"
			txtBpCd.Tag = "14"
			txtCustCd.Tag = "11"
			txtWcCd.Tag = "14"
		
	End Select 

End Sub

Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			

   	arrField(0) = "PLANT_CD"	
   	arrField(1) = "PLANT_NM"	

   	arrHeader(0) = "공장코드"		
   	arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	txtPlantCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtPlantCd.value = arrRet(0)
		txtPlantNm.value = arrRet(1)
		txtPlantCd.Focus
	End If	

	Set gActiveElement = document.activeElement
End Function

Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	'공장코드가 있는 지 체크 
	If Trim(txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(txtPlantNm.value)	' Plant Code
	arrParam(2) = Trim(txtItemCd.Value)	' Item Code
	arrParam(3) = Trim(txtItemNm.Value)	' Item Code
	arrParam(4) = Trim(cboInspClassCd.Value)								' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	
	arrField(0) = 1 '"ITEM_CD"					' Field명(0)
    arrField(1) = 2 '"ITEM_NM"					' Field명(1)
    arrField(2) = 9 '"SPECIFICATION"				' Field명(1)
    arrField(3) = 6 '"BASIC_UNIT"					' Field명(1)
	    
  	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	txtItemCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtItemCd.Value    = arrRet(0)		
		txtItemNm.Value    = arrRet(1)		
		txtItemCd.Focus
	End If	

	Set gActiveElement = document.activeElement
End Function

Function OpenBp()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If UCase(txtBpCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE 명칭 
	arrParam(2) = Trim(txtBpCd.Value)					' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "(BP_TYPE = " & FilterVar("CS", "''", "S") & " Or BP_TYPE = " & FilterVar("S", "''", "S") & " )"			' Where Condition	
	
	arrParam(5) = "공급처"						' 조건필드의 라벨 명칭	
	
    	arrField(0) = "BP_CD"								' Field명(0)
    	arrField(1) = "BP_NM"								' Field명(1)
    
    	arrHeader(0) = "공급처코드"					' Header명(0)
    	arrHeader(1) = "공급처명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtBpCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtBpCd.Value = arrRet(0)
		txtBpNm.Value = arrRet(1)
		txtBpCd.Focus
	End If	

	Set gActiveElement = document.activeElement
End Function

Function OpenCust()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If UCase(txtCustCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE 명칭 
	arrParam(2) = Trim(txtCustCd.Value)					' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "(BP_TYPE = " & FilterVar("CS", "''", "S") & " Or BP_TYPE = " & FilterVar("C", "''", "S") & " )"			' Where Condition
	arrParam(5) = "거래처"						' 조건필드의 라벨 명칭	
	
    	arrField(0) = "BP_CD"								' Field명(0)
    	arrField(1) = "BP_NM"								' Field명(1)
    
    	arrHeader(0) = "거래처코드"					' Header명(0)
    	arrHeader(1) = "거래처명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtCustCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtCustCd.Value = arrRet(0)
		txtCustNm.Value = arrRet(1)
		txtCustCd.Focus
	End If	

	Set gActiveElement = document.activeElement
End Function

Function OpenWc()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'공장코드가 있는 지 체크 
	If Trim(txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function
	End If
	
	If UCase(txtWcCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "작업장 팝업"					' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"					' TABLE 명칭 
	arrParam(2) = Trim(txtWcCd.Value)					' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(txtPlantCd.value, "''", "S") 	' Where Condition
	arrParam(5) = "작업장"						' 조건필드의 라벨 명칭	
	
    arrField(0) = "Wc_CD"								' Field명(0)
    arrField(1) = "Wc_NM"								' Field명(1)
    
    arrHeader(0) = "작업장코드"					' Header명(0)
    arrHeader(1) = "작업장명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtWcCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtWcCd.Value = arrRet(0)
		txtWcNm.Value = arrRet(1)
		txtWcCd.Focus
	End If	

	Set gActiveElement = document.activeElement
End Function

Function OKClick()
	Dim intColCnt, iCurColumnPos
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 2)
	
		ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 
				
		For intColCnt = 0 To vspdData.MaxCols - 2
			vspddata.Col = iCurColumnPos(CInt(intColCnt + 1))
			arrReturn(intColCnt) = vspdData.Text
		Next
			
		Self.Returnvalue = arrReturn
	End If
	
	Self.Close()
End Function

Function CancelClick()
	On Error Resume Next
	Self.Close()
End Function

Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format

	Call InitComboBox						
	Call SetDefaultVal()
	Call ProtectField()
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	Call InitVariables
	
	Call InitSpreadSheet()
	
	Call FncQuery()
End Sub

Sub txtFrInspReqDt_DblClick(Button)
    If Button = 1 Then
        txtFrInspReqDt.Action = 7
    End If
End Sub

Sub txtToInspReqDt_DblClick(Button)
    If Button = 1 Then
        txtToInspReqDt.Action = 7
    End If
End Sub

Function FncQuery()
	
	FncQuery = False
   	
   	vspdData.MaxRows = 0
	lgQueryFlag = "1"
	lgStrPrevKey = ""
	
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	if DbQuery = false then
		Exit Function
	End if

	fncQuery = True
End Function

Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")

	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = vspdData

    If vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
   	
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              ' 타이틀 cell을 dblclick했거나....
	   Exit Function
	End If
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Function

Function vspdData_KeyPress(KeyAscii)
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			If DBQuery = False Then
				'Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If   

End Sub

Sub txtFrInspReqDt_KeyPress(KeyAscii)
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

Sub txtToInspReqDt_KeyPress(KeyAscii)
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	vspdData.Redraw = True
End Sub

Function DbQuery()
	Dim strVal
	Dim txtMaxRows
	
	DbQuery = False 	
	
	On Error Resume Next
	
	If ValidDateCheck(txtFrInspReqDt, txtToInspReqDt) = False Then
		Exit Function
	End If
	
	'Show Processing Bar
    Call LayerShowHide(1)  

	txtMaxRows = vspdData.MaxRows
	
	if lgStrPrevKey <> "" Then
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & txtPlantCd.Value
		strVal = strVal & "&txtItemCd=" & hItemCd
		strVal = strVal & "&txtInspClassCd=" & cboInspClassCd.Value
		strVal = strVal & "&txtInspReqNo=" & hInspReqNo
		strVal = strVal & "&txtBpCd=" & hBpCd
		strVal = strVal & "&txtCustCd=" & hCustCd
		strVal = strVal & "&txtWcCd=" & hWcCd
		strVal = strVal & "&txtFrInspReqDt=" & hFrInspReqDt
		strVal = strVal & "&txtToInspReqDt=" & hToInspReqDt
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & txtMaxRows		
	else
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & txtPlantCd.Value
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.Value)
		strVal = strVal & "&txtInspClassCd=" & cboInspClassCd.Value
		strVal = strVal & "&txtInspReqNo=" & Trim(txtInspReqNo.Value)
		strVal = strVal & "&txtBpCd=" & Trim(txtBpCd.Value)
		strVal = strVal & "&txtCustCd=" & Trim(txtCustCd.Value)
		strVal = strVal & "&txtWcCd=" & Trim(txtWcCd.Value)
		strVal = strVal & "&txtFrInspReqDt=" & txtFrInspReqDt.Text
		strVal = strVal & "&txtToInspReqDt=" & txtToInspReqDt.Text
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & txtMaxRows	
	end if                                                        '⊙: Processing is NG
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True 
	
End Function

Function DbQueryOk()								'☆: 조회 성공후 실행로직 

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR HEIGHT=*>
		<TD  WIDTH=100%>
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
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="12XXXU"><IMG align=top height=20 name=btnPlantCd onclick=vbscript:OpenPlant() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>								
									<TD CLASS=TD5 NOWRAP>검사분류</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" TAG="14"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20 MAXLENGTH=18 ALT="검사의뢰번호" tag="11XXXU"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=20 ALT="품목" tag="11XXXU"><IMG align=top height=20 name=btnItemCd onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 ALT="공급처" tag="11XXXU"><IMG align=top height=20 name=btnBpCd onclick=vbscript:OpenBp() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=40 tag="14" ></TD>								
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCustCd" SIZE=10 MAXLENGTH=10 ALT="거래처" tag="11XXXU"><IMG align=top height=20 name=btnCustCd onclick=vbscript:OpenCust() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtCustNm" SIZE=20 MAXLENGTH=40 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=10 MAXLENGTH=7 ALT="작업장" tag="11XXXU"><IMG align=top height=20 name=btnWcCd onclick=vbscript:OpenWc() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 MAXLENGTH=40 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>검사의뢰일</TD>
									<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2511pa1_fpDateTime1_txtFrInspReqDt.js'></script>&nbsp;~&nbsp;
									<script language =javascript src='./js/q2511pa1_fpDateTime2_txtToInspReqDt.js'></script>
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
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD>
									<script language =javascript src='./js/q2511pa1_I785444039_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
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
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/query_d.gif" Style="CURSOR: hand" ALT="Search" NAME="search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>  
