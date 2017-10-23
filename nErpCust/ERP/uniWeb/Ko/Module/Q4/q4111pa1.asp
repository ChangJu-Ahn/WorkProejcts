<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q4111PA1
'*  4. Program Name         : 
'*  5. Program Desc         : 검사의뢰현황 팝업 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2004/07/01
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Hong Chang Ho
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

Const BIZ_PGM_ID = "q4111pb1.asp"							 '☆: 비지니스 로직 ASP명 

Dim C_InspReqNo
Dim C_ProdtOrderNo
Dim C_TrackingNo
Dim C_InspClassNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_ItemSpec
Dim C_SupplierCd
Dim C_SupplierNm
Dim C_RoutNo
Dim C_RoutNoDesc
Dim C_OprNo
Dim C_OprNoDesc
Dim C_WcCd
Dim C_WcNm
Dim C_BpCd
Dim C_BpNm
Dim C_SLCd
Dim C_SLNm
Dim C_StatusFlagNm
Dim C_DecisionNm
Dim C_InspDt
Dim C_LotNo
Dim C_LotSubNo
Dim C_LotSize
Dim C_Unit
Dim C_InspQty
Dim C_DefectQty
Dim C_InspectorCd
Dim C_InspectorNm
Dim C_InspClassCd
Dim C_StatusFlagCd
Dim C_DecisionCd
Dim C_ReleaseDt
'********* rcpt,issue sl_cd add LSW 2005-11-23 ****************
Dim C_SlCdGood
Dim C_SlNmGood
Dim C_SlCdDefect
Dim C_SlNmDefect


Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 
Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim hPlantCd
Dim hInspReqNo
Dim hInspClassCd
Dim hItemCd
Dim hLotNo
Dim hFrInspDt
Dim hToInspDt
Dim hStatusFlagCd
Dim hDecisionCd
Dim hSupplierCd
Dim hRoutNo
Dim hOprNo
Dim hBPCd
Dim hSLCd

Dim ArrParent

Dim arrParam				'--- First Parameter Group 
Dim arrReturn				'--- Return Parameter Group 

Dim IsOpenPop          

<!-- #Include file="../../inc/lgvariables.inc" -->	

'------ Set Parameters from Parent ASP ------ 
ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

Dim i
ReDim arrParam(UBOUND(ArrParent))
For i = 0 To UBOUND(ArrParent) - 1
	arrParam(i) = ArrParent(i + 1)
Next

top.document.title = PopupParent.gActivePRAspName
'--------------------------------------------- 

Function InitVariables()
	lgSortKey = 1                            '⊙: initializes sort direction
	lgQueryFlag = "1"
End Function

Sub initSpreadPosVariables()  
	C_InspReqNo = 1
	C_ProdtOrderNo = 2
	C_TrackingNo = 3
	C_InspClassNm = 4
	C_ItemCd = 5
	C_ItemNm = 6
	C_ItemSpec = 7
	C_SupplierCd = 8
	C_SupplierNm = 9
	C_RoutNo = 10
	C_RoutNoDesc = 11
	C_OprNo = 12
	C_OprNoDesc = 13
	C_WcCd = 14
	C_WcNm = 15
	C_BpCd = 16
	C_BpNm = 17
	C_SLCd = 18
	C_SLNm = 19
	C_SlCdGood = 20
	C_SlNmGood = 21
	C_SlCdDefect = 22
	C_SlNmDefect = 23
	C_StatusFlagNm = 24
	C_DecisionNm = 25
	C_InspDt = 26
	C_LotNo = 27
	C_LotSubNo = 28
	C_LotSize = 29
	C_Unit = 30
	C_InspQty = 31
	C_DefectQty = 32
	C_InspectorCd = 33
	C_InspectorNm = 34
	C_InspClassCd = 35
	C_StatusFlagCd = 36
	C_DecisionCd = 37
	C_ReleaseDt = 38
	
End Sub

Sub SetDefaultVal()
	
	txtPlantCd.Value = arrParam(0)
	txtPlantNm.Value = arrParam(1)
	txtInspReqNo.Value = arrParam(2)
	cboInspClassCd.Value = arrParam(3)
	cboDecision.value = arrParam(4)
	cboStatusFlag.value = arrParam(5)
	
	Self.Returnvalue = Array("")
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q","NOCOOKIE","PA") %>
End Sub

Sub InitComboBox()    
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(cboInspClassCd , lgF0, lgF1, Chr(11))
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0014", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(cboStatusFlag , lgF0, lgF1, Chr(11))
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0010", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(cboDecision , lgF0, lgF1, Chr(11))
End Sub

Sub InitSpreadSheet()
	Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20060118",,PopupParent.gAllowDragDropSpread
	
	With vspdData
		.ReDraw = False
		
		.MaxCols = C_ReleaseDt + 1
		.MaxRows = 0
	End With
	
	Call GetSpreadColumnPos("A")
	
	With ggoSpread
		
		.SSSetEdit C_InspReqNo,"검사의뢰번호", 20
		.SSSetEdit C_ProdtOrderNo,"제조오더번호",20
		.SSSetEdit C_TrackingNo,"Tracking No.",20
		.SSSetEdit C_InspClassNm,"검사분류", 20
		.SSSetEdit C_ItemCd,"품목코드", 15
		.SSSetEdit C_ItemNm,"품목명", 20
		.SSSetEdit C_ItemSpec,"규격", 30
		
		.SSSetEdit C_SupplierCd,"공급처코드",10
		.SSSetEdit C_SupplierNm,"공급처명",15
		.SSSetEdit C_RoutNo,"라우팅",15
		.SSSetEdit C_RoutNoDesc,"라우팅설명",15
		.SSSetEdit C_OprNo,"공정",5
		.SSSetEdit C_OPrNoDesc,"공정작업명",15
		.SSSetEdit C_WCCd,"작업장코드",10
		.SSSetEdit C_WCNm,"작업장명",15
		.SSSetEdit C_BPCd,"거래처코드",10
		.SSSetEdit C_BPNm,"거래처명",15
		.SSSetEdit C_SLCd,"의뢰창고",10
		.SSSetEdit C_SLNm,"의뢰창고명",15
		.SSSetEdit C_SlCdGood,"양품창고",10
		.SSSetEdit C_SlNmGood,"양품창고명",15
		.SSSetEdit C_SlCdDefect,"불량창고",10
		.SSSetEdit C_SlNmDefect,"불량창고명",15
		
		.SSSetEdit C_StatusFlagNm,"검사진행상태", 20
		.SSSetEdit C_DecisionNm,"판정", 20
		.SSSetEdit C_InspDt,"검사일",10, 2
		.SSSetEdit C_LotNo,"로트번호",20
		.SSSetFloat C_LotSubNo,"로트순번", 5, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetFloat C_LotSize,"로트크기", 10, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetEdit C_Unit,"단위", 5, 1
		.SSSetFloat C_InspQty,"검사수", 10, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetFloat C_DefectQty,"불량수", 10, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetEdit C_InspectorCd,"검사자코드", 10
		.SSSetEdit C_InspectorNm,"검사자명", 15
		.SSSetEdit C_InspClassCd,"", 1
		.SSSetEdit C_StatusFlagCd,"", 1
		.SSSetEdit C_DecisionCd,"", 1
		.SSSetEdit C_ReleaseDt,"Release일",10, 2
		
		
	End With
		
	Call ChangingFieldsByInspClass(cboInspClassCd.value)
	
	Call ggoSpread.SSSetColHidden(C_InspClassCd, C_DecisionCd, True)
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
		
	Call SetSpreadLock
	
	vspdData.ReDraw = True
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
			
			C_InspReqNo = iCurColumnPos(1)
			C_ProdtOrderNo = iCurColumnPos(2)
			C_TrackingNo = iCurColumnPos(3)
			C_InspClassNm = iCurColumnPos(4)
			C_ItemCd = iCurColumnPos(5)
			C_ItemNm = iCurColumnPos(6)
			C_ItemSpec = iCurColumnPos(7)
			C_SupplierCd = iCurColumnPos(8)
			C_SupplierNm = iCurColumnPos(9)
			C_RoutNo = iCurColumnPos(10)
			C_RoutNoDesc = iCurColumnPos(11)
			C_OprNo = iCurColumnPos(12)
			C_OprNoDesc = iCurColumnPos(13)
			C_WcCd = iCurColumnPos(14)
			C_WcNm = iCurColumnPos(15)
			C_BpCd = iCurColumnPos(16)
			C_BpNm = iCurColumnPos(17)
			C_SLCd = iCurColumnPos(18)
			C_SLNm = iCurColumnPos(19)
			C_SlCdGood = iCurColumnPos(20)
			C_SlNmGood = iCurColumnPos(21)
			C_SlCdDefect = iCurColumnPos(22)
			C_SlNmDefect = iCurColumnPos(23)
			C_StatusFlagNm = iCurColumnPos(24)
			C_DecisionNm = iCurColumnPos(25)
			C_InspDt = iCurColumnPos(26)
			C_LotNo = iCurColumnPos(27)
			C_LotSubNo = iCurColumnPos(28)
			C_LotSize = iCurColumnPos(29)
			C_Unit = iCurColumnPos(30)
			C_InspQty = iCurColumnPos(31)
			C_DefectQty = iCurColumnPos(32)
			C_InspectorCd = iCurColumnPos(33)
			C_InspectorNm = iCurColumnPos(34)
			C_InspClassCd = iCurColumnPos(35)
			C_StatusFlagCd = iCurColumnPos(36)
			C_DecisionCd = iCurColumnPos(37)
			C_ReleaseDt = iCurColumnPos(38)
	End Select    

End Sub

Sub ChangingFieldsByInspClass(Byval strInspClass)
	With ggoSpread
		Select Case strInspClass
			Case "R"
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, False)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, False)
				Call .SSSetColHidden(C_ProdtOrderNo, C_ProdtOrderNo, True)
				Call .SSSetColHidden(C_RoutNo, C_RoutNo, True)
				Call .SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, True)
				Call .SSSetColHidden(C_OprNo, C_OprNo, True)
				Call .SSSetColHidden(C_OprNoDesc, C_OprNoDesc, True)
				Call .SSSetColHidden(C_WCCd, C_WCCd, True)
				Call .SSSetColHidden(C_WCNm, C_WCNm, True)
				Call .SSSetColHidden(C_BPCd, C_BPCd, True)
				Call .SSSetColHidden(C_BPNm, C_BPNm, True)
				Call .SSSetColHidden(C_SLCd, C_SLCd, False)
				Call .SSSetColHidden(C_SLNm, C_SLNm, False)
				Call .SSSetColHidden(C_SlCdGood, C_SlCdGood, False)
				Call .SSSetColHidden(C_SlNmGood, C_SlNmGood, False)
				Call .SSSetColHidden(C_SlCdDefect, C_SlCdDefect, False)
				Call .SSSetColHidden(C_SlNmDefect, C_SlNmDefect, False)
				
			Case "P"
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, True)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, True)
				Call .SSSetColHidden(C_ProdtOrderNo, C_ProdtOrderNo, False)
				Call .SSSetColHidden(C_RoutNo, C_RoutNo, False)
				Call .SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, False)
				Call .SSSetColHidden(C_OprNo, C_OprNo, False)
				Call .SSSetColHidden(C_OprNoDesc, C_OprNoDesc, False)
				Call .SSSetColHidden(C_WCCd, C_WCCd, False)
				Call .SSSetColHidden(C_WCNm, C_WCNm, False)
				Call .SSSetColHidden(C_BPCd, C_BPCd, True)
				Call .SSSetColHidden(C_BPNm, C_BPNm, True)
				Call .SSSetColHidden(C_SLCd, C_SLCd, True)
				Call .SSSetColHidden(C_SLNm, C_SLNm, True)
				Call .SSSetColHidden(C_SlCdGood, C_SlCdGood, True)
				Call .SSSetColHidden(C_SlNmGood, C_SlNmGood, True)
				Call .SSSetColHidden(C_SlCdDefect, C_SlCdDefect, True)
				Call .SSSetColHidden(C_SlNmDefect, C_SlNmDefect, True)
				
			Case "F"
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, True)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, True)
				Call .SSSetColHidden(C_ProdtOrderNo, C_ProdtOrderNo, False)
				Call .SSSetColHidden(C_RoutNo, C_RoutNo, True)
				Call .SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, True)
				Call .SSSetColHidden(C_OprNo, C_OprNo, True)
				Call .SSSetColHidden(C_OprNoDesc, C_OprNoDesc, True)
				Call .SSSetColHidden(C_WCCd, C_WCCd, True)
				Call .SSSetColHidden(C_WCNm, C_WCNm, True)
				Call .SSSetColHidden(C_BPCd, C_BPCd, True)
				Call .SSSetColHidden(C_BPNm, C_BPNm, True)
				Call .SSSetColHidden(C_SLCd, C_SLCd, False)
				Call .SSSetColHidden(C_SLNm, C_SLNm, False)
				Call .SSSetColHidden(C_SlCdGood, C_SlCdGood, False)
				Call .SSSetColHidden(C_SlNmGood, C_SlNmGood, False)
				Call .SSSetColHidden(C_SlCdDefect, C_SlCdDefect, False)
				Call .SSSetColHidden(C_SlNmDefect, C_SlNmDefect, False)
				
			Case "S"
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, True)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, True)
				Call .SSSetColHidden(C_ProdtOrderNo, C_ProdtOrderNo, True)
				Call .SSSetColHidden(C_RoutNo, C_RoutNo, True)
				Call .SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, True)
				Call .SSSetColHidden(C_OprNo, C_OprNo, True)
				Call .SSSetColHidden(C_OprNoDesc, C_OprNoDesc, True)
				Call .SSSetColHidden(C_WCCd, C_WCCd, True)
				Call .SSSetColHidden(C_WCNm, C_WCNm, True)
				Call .SSSetColHidden(C_BPCd, C_BPCd, False)
				Call .SSSetColHidden(C_BPNm, C_BPNm, False)
				Call .SSSetColHidden(C_SLCd, C_SLCd, True)
				Call .SSSetColHidden(C_SLNm, C_SLNm, True)
				Call .SSSetColHidden(C_SlCdGood, C_SlCdGood, True)
				Call .SSSetColHidden(C_SlNmGood, C_SlNmGood, True)
				Call .SSSetColHidden(C_SlCdDefect, C_SlCdDefect, True)
				Call .SSSetColHidden(C_SlNmDefect, C_SlNmDefect, True)
				
			Case Else
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, False)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, False)
				Call .SSSetColHidden(C_ProdtOrderNo, C_ProdtOrderNo, False)
				Call .SSSetColHidden(C_RoutNo, C_RoutNo, False)
				Call .SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, False)
				Call .SSSetColHidden(C_OprNo, C_OprNo, False)
				Call .SSSetColHidden(C_OprNoDesc, C_OprNoDesc, False)
				Call .SSSetColHidden(C_WCCd, C_WCCd, False)
				Call .SSSetColHidden(C_WCNm, C_WCNm, False)
				Call .SSSetColHidden(C_BPCd, C_BPCd, False)
				Call .SSSetColHidden(C_BPNm, C_BPNm, False)
				Call .SSSetColHidden(C_SLCd, C_SLCd, False)
				Call .SSSetColHidden(C_SLNm, C_SLNm, False)
				Call .SSSetColHidden(C_SlCdGood, C_SlCdGood, False)
				Call .SSSetColHidden(C_SlNmGood, C_SlNmGood, False)
				Call .SSSetColHidden(C_SlCdDefect, C_SlCdDefect, False)
				Call .SSSetColHidden(C_SlNmDefect, C_SlNmDefect, False)
				
		End Select
	End With
End Sub

Sub EnableField(Byval strInspClass)
	Select Case strInspClass
		Case "R"
			Receiving.style.display = ""
			Process.style.display = "none"
			Final.style.display = "none"
			Shipping.style.display = "none"
			
		Case "P"
			Receiving.style.display = "none"
			Process.style.display = ""
			Final.style.display = "none"
			Shipping.style.display = "none"
			
		Case "F"
			Receiving.style.display = "none"
			Process.style.display = "none"
			Final.style.display = ""
			Shipping.style.display = "none"
			
		Case "S"
			Receiving.style.display = "none"
			Process.style.display = "none"
			Final.style.display = "none"
			Shipping.style.display = ""
			
		Case Else
			Receiving.style.display = "none"
			Process.style.display = "none"
			Final.style.display = "none"
			Shipping.style.display = "none"
			
	End Select 

End Sub

Sub cboInspClassCd_onchange()
	Call EnableField(cboInspClassCd.value)
	Call ChangingFieldsByInspClass(cboInspClassCd.value)
End Sub

Function OpenPlant()
	OpenPlant = False
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
		txtPlantCd.Value    = arrRet(0)
		txtPlantNm.Value    = arrRet(1)
		txtPlantCd.Focus
	End If	

	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD

	'공장코드가 있는 지 체크 
	If Trim(txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = Trim(txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(txtItemNm.Value)	' Item Name
	arrParam5 = Trim(cboInspClassCd.Value)
	
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
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
	OpenItem = true
End Function

Function OpenSupplier()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If UCase(txtSupplierCd.ClassName) = UCase(PopupParent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE 명칭 
	arrParam(2) = Trim(txtSupplierCd.Value)					' Code Condition
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
	
	txtSupplierCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtSupplierCd.Value = arrRet(0)
		txtSupplierNm.Value = arrRet(1)
		txtSupplierCd.Focus
	End If	

	Set gActiveElement = document.activeElement	
	OpenSupplier = true
End Function

Function OpenBP()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If UCase(txtBPCd.ClassName) = UCase(PopupParent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE 명칭 
	arrParam(2) = Trim(txtBpCd.Value)					' Code Condition
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
	
	txtBpCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtBpCd.Value = arrRet(0)
		txtBpNm.Value = arrRet(1)
		txtBpCd.Focus
	End If	

	Set gActiveElement = document.activeElement	
	OpenBp = true	
End Function

'====================  OpenRoutNo  ======================================
' Function Name : OpenRoutNo
' Function Desc : OpenRoutNo Reference Popup
'==========================================================================
Function OpenRoutNo()

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	
	If UCase(txtRoutNo.ClassName) = UCase(PopupParent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	If txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "품목","X")
		txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
		
	arrParam(0) = "라우팅 팝업"					' 팝업 명칭 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE 명칭 
	arrParam(2) = Trim(txtRoutNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "P_ROUTING_HEADER.PLANT_CD =" & FilterVar(UCase(txtPlantCd.value), "''", "S") & _
				  "And P_ROUTING_HEADER.ITEM_CD =" & FilterVar(UCase(txtItemCd.value), "''", "S") 	
	arrParam(5) = "라우팅"			
	
    arrField(0) = "ROUT_NO"							' Field명(0)
    arrField(1) = "DESCRIPTION"						' Field명(1)
    arrField(2) = "BOM_NO"							' Field명(1)
    arrField(3) = "MAJOR_FLG"						' Field명(1)
   
    arrHeader(0) = "라우팅"						' Header명(0)
    arrHeader(1) = "라우팅명"					' Header명(1)
    arrHeader(2) = "BOM Type"					' Header명(1)
    arrHeader(3) = "주라우팅"					' Header명(1)        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
	If arrRet(0) <> "" Then
		txtRoutNo.Value		= arrRet(0)		
		txtRoutNoDesc.Value		= arrRet(1)		
	End If
	
	Call SetFocusToDocument("M")
	txtRoutNo.focus
	
End Function

'**************************** Function OpenOprNo() ***********************************8
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If UCase(txtOprNo.ClassName) = UCase(PopupParent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function    

	IsOpenPop = True
	
	If txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "품목","X")
		txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	
	
	If txtRoutNo.value= "" Then
		Call DisplayMsgBox("971012","X", "라우팅","X")
		txtRoutNo.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	

	arrParam(0) = "공정팝업"	
	arrParam(1) = "P_ROUTING_DETAIL A inner join P_WORK_CENTER B on A.wc_cd = B.wc_cd and A.plant_cd = B.plant_cd " & _
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = " & FilterVar("P1006", "''", "S") & ""				
	arrParam(2) = UCase(Trim(txtOprNo.Value))
	arrParam(3) = ""
	arrParam(4) = "A.plant_cd =" & FilterVar(UCase(txtPlantCd.value), "''", "S") & _
				  " and	A.item_cd =" & FilterVar(UCase(txtItemCd.value), "''", "S") & _
				  " and	A.rout_no =" & FilterVar(UCase(txtRoutNo.value), "''", "S") & _
				  "	and	A.rout_order in (" & FilterVar("F", "''", "S") & " ," & FilterVar("I", "''", "S") & " ) "	
	arrParam(5) = "공정"			
	
	arrField(0) = "A.OPR_NO"	
	arrField(1) = "A.WC_CD"
	arrField(2) = "B.WC_NM"
	arrField(3) = "C.MINOR_NM"
	arrField(4) = "A.INSIDE_FLG"
	arrField(5) = "A.MILESTONE_FLG"
	arrField(6) = "A.INSP_FLG"
	
	arrHeader(0) = "공정"		
	arrHeader(1) = "작업장"	
	arrHeader(2) = "작업장명"
	arrHeader(3) = "공정작업명"
	arrHeader(4) = "사내구분"
	arrHeader(5) = "Milestone"
	arrHeader(6) = "검사여부"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		txtOprNo.focus
		Exit Function
	Else
		txtOprNo.Value = arrRet(0)
		txtOprNoDesc.Value = arrRet(3)
	End If	
	
End Function

Function OpenSL()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If UCase(txtSLCd.ClassName) = UCase(PopupParent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If Trim(txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function	
	End If
	
	IsOpenPop = True
	
	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_Storage_Location"				
	arrParam(2) = Trim(txtSLCd.Value)
	arrParam(3) = ""
	arrParam(4) = "Plant_Cd =  " & FilterVar(txtPlantCd.Value, "''", "S") & " And SL_TYPE <> " & FilterVar("E", "''", "S") & " "    ' Where Condition
	arrParam(5) = "창고"			
	
    arrField(0) = "SL_CD"	
    arrField(1) = "SL_NM"	
    
    arrHeader(0) = "창고코드"		
    arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtSLCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtSLCd.value = arrRet(0)   
		txtSLNm.value = arrRet(1) 
		txtSLCd.Focus
	End If	

	Set gActiveElement = document.activeElement	
End Function

Function OKClick()
	Dim intColCnt, iCurColumnPos
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
	
		ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 
				
		For intColCnt = 0 To vspdData.MaxCols - 1
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

Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	
	Call AppendNumberPlace("6", "3","0")
	Call InitComboBox				'순서를 바꾸면 안됨 
	Call SetDefaultVal()
	
	Call EnableField(cboInspClassCd.value)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	Call InitVariables

	
	If arrParam(3) <> "" Then
		Call ggoOper.SetReqAttr(cboInspClassCd, "Q")	'FormatField 다음에 써야함 - D:흰색,Q:회색,N 노란색 
	Else
		Call ggoOper.SetReqAttr(cboInspClassCd, "D")	'FormatField 다음에 써야함 - D:흰색,Q:회색,N 노란색 
	End If	
		
	If arrParam(4) <> "" Then
		Call ggoOper.SetReqAttr(cboDecision, "Q")		'FormatField 다음에 써야함 - D:흰색,Q:회색,N 노란색 
	Else
		Call ggoOper.SetReqAttr(cboDecision, "D")		'FormatField 다음에 써야함 - D:흰색,Q:회색,N 노란색 
	End If	

	If arrParam(5) <> "" Then
		Call ggoOper.SetReqAttr(cboStatusFlag, "Q")		'FormatField 다음에 써야함 - D:흰색,Q:회색,N 노란색 
	Else
		Call ggoOper.SetReqAttr(cboStatusFlag, "D")		'FormatField 다음에 써야함 - D:흰색,Q:회색,N 노란색 
	End If			
	
	Call InitSpreadSheet()
	
	Call FncQuery()
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
	
End Sub

Sub txtFrInspDt_DblClick(Button)
    If Button = 1 Then
        txtFrInspDt.Action = 7
    End If
End Sub

Sub txtToInspDt_DblClick(Button)
    If Button = 1 Then
        txtToInspDt.Action = 7
    End If
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
		If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			If DBQuery = False Then
				Exit Sub
			End If
		End If
	End If
End Sub

Sub txtFrInspDt_KeyPress(KeyAscii)
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

Sub txtToInspDt_KeyPress(KeyAscii)
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

Function FncQuery()
	
	FncQuery = False
   	
   	vspdData.MaxRows = 0
	lgQueryFlag = "1"
	lgStrPrevKey1 = ""
	lgStrPrevKey2 = ""

	Call EnableField(cboInspClassCd.value)
	Call ChangingFieldsByInspClass(cboInspClassCd.value)
	
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

Function DbQuery()
	Dim strVal
	Dim txtMaxRows
	
	DbQuery = False 	
	
	On Error Resume Next
	
	If ValidDateCheck(txtFrInspDt, txtToInspDt) = False Then
		Exit Function
	End If
	
	'Show Processing Bar
    Call LayerShowHide(1)  

	
	txtMaxRows = vspdData.MaxRows
	 
	If lgQueryFlag = "0" Then
		strVal = BIZ_PGM_ID & "?QueryFlag=" & lgQueryFlag _
				& "&txtPlantCd=" & hPlantCd _
				& "&txtInspReqNo=" & lgStrPrevKey1 _
				& "&txtInspResultNo=" & lgStrPrevKey2 _
				& "&txtItemCd=" & hItemCd _
				& "&txtLotNo=" & hLotNo _
				& "&txtInspClassCd=" & hInspClassCd _
				& "&txtFrInspDt=" & hFrInspDt _
				& "&txtToInspDt=" & hToInspDt _
				& "&txtStatusFlagCd=" & hStatusFlagCd _
				& "&txtDecisionCd=" & hDecisionCd
				
		Select Case cboInspClassCd.value
			Case "R"
				strVal = strVal & "&txtSupplierCd=" & hSupplierCd
				
			Case "P"
				strVal = strVal & "&txtRoutNo=" & hRoutNo _
								& "&txtOprNo=" & hOprNo 
				
			Case "F"
				strVal = strVal & "&txtSLCd=" & hSLCd
			Case "S"
				strVal = strVal & "&txtBpCd=" & hBpCd
				
			Case Else
			
		End Select
		
		strVal = strVal &  "&txtMaxRows=" & txtMaxRows		
						
	Else
		strVal = BIZ_PGM_ID & "?QueryFlag=" & lgQueryFlag _
				& "&txtPlantCd=" & Trim(txtPlantCd.Value) _
				& "&txtInspReqNo=" & Trim(txtInspReqNo.Value) _
				& "&txtInspResultNo=" & CStr(1) _
				& "&txtItemCd=" & Trim(txtItemCd.Value) _
				& "&txtInspClassCd=" & cboInspClassCd.Value _
				& "&txtLotNo=" & Trim(txtLotNo.Value) _
				& "&txtFrInspDt=" & txtFrInspDt.Text _
				& "&txtToInspDt=" & txtToInspDt.Text _
				& "&txtStatusFlagCd=" & cboStatusFlag.Value _
				& "&txtDecisionCd=" & cbodecision.Value
				
		Select Case cboInspClassCd.value
			Case "R"
				strVal = strVal & "&txtSupplierCd=" & Trim(txtSupplierCd.Value) 
				
			Case "P"
				strVal = strVal & "&txtRoutNo=" & Trim(txtRoutNo.Value) _
								& "&txtOprNo=" & Trim(txtOprNo.Value)
								
				
			Case "F"
				strVal = strVal & "&txtSLCd=" & Trim(txtSLCd.Value) 
			Case "S"
				strVal = strVal & "&txtBpCd=" & Trim(txtBpCd.Value) 
				
			Case Else
			
		End Select
	End if                                                        '⊙: Processing is NG
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True 
	
End Function

Function DbQueryOk()								'☆: 조회 성공후 실행로직 
	lgQueryFlag = "0"
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
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" TAG="12XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnPlantPopup ONCLICK=vbscript:OpenPlant() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm" TAG="14">
									</TD>
									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE="20" MAXLENGTH="18" ALT="검사의뢰번호" TAG="11XXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE="20" MAXLENGTH="18" ALT="품목" TAG="11XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnItemPopup ONCLICK=vbscript:OpenItem() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtItemNm" TAG="14"></TD>
									<td CLASS="TD5" NOWPAP>검사분류</TD>
									<td CLASS="TD6" NOWPAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" TAG="14"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>규격</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE="40" TAG="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>로트번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE="20" MAXLENGTH="25" ALT="로트번호" TAG="11XXXU"></TD>
									<TD CLASS="TD5" NOWRAP>검사일</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q4111pa1_txtFrInspDt_txtFrInspDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/q4111pa1_txtToInspDt_txtToInspDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWPAP>검사진행상태</TD>
									<TD CLASS="TD6" NOWPAP><SELECT NAME="cboStatusFlag" ALT="검사진행상태" STYLE="WIDTH: 150px" TAG="11"><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWPAP>판정</TD>
									<TD CLASS="TD6" NOWPAP><SELECT NAME="cboDecision" ALT="판정" STYLE="WIDTH: 150px" TAG="11"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								<TR ID="Receiving">
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE="10" MAXLENGTH="10" ALT="공급처" TAG="11XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnSupplierPopup ONCLICK=vbscript:OpenSupplier() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtSupplierNm" TAG="14">
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR ID="Process">
					      			<TD CLASS="TD5" NOWPAP>라우팅</TD>
									<TD CLASS="TD6" NOWPAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="11XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutNo()">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
									<TD CLASS="TD5" NOWPAP>공정</TD>
									<TD CLASS="TD6" NOWPAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="11XXXU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()">&nbsp;<input TYPE=TEXT NAME="txtOprNoDesc" SIZE="30" tag="14"></TD>
								</TR>
								<TR ID="Final">
									<TD CLASS="TD5" NOWRAP>창고</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtSLCd" SIZE="10" MAXLENGTH="7" ALT="창고" TAG="11XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btntxtSLPopup ONCLICK=vbscript:OpenSL() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtSLNm" TAG="14">
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR ID="Shipping">
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtBPCd" SIZE="10" MAXLENGTH="10" ALT="거래처" TAG="11XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnBPPopup ONCLICK=vbscript:OpenBP() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtBPNm" TAG="14">
									</TD>
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
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD>
									<script language =javascript src='./js/q4111pa1_vspdData_vspdData.js'></script>
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
						<IMG SRC="../../../CShared/image/query_d.gif" Style="CURSOR: hand" ALT="Search" NAME="search" OnClick="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
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
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>  
