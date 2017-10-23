<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2512PA1
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

Const BIZ_PGM_ID = "q2512pb1.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_InspReqNo			'검사의뢰번호					0
Dim C_InspClassNm		'검사분류명						1
Dim C_ItemCd			'품목코드						2
Dim C_ItemNm			'품목명							3
Dim C_ItemSpec			'규격(1000T)					4
Dim C_SupplierCd		'공급처코드(00001)				5
Dim C_SupplierNm		'공급처명(333)					6
Dim C_RoutNo			'라우팅							7
Dim C_RoutNoDesc		'라우팅 설명					8
Dim C_OprNo				'공정							9
Dim C_OprNoDesc			'공정작업명						10
Dim C_WcCd				'작업장코드(03)					11
Dim C_WcNm				'작업장명(정제)					12
Dim C_BpCd				'거래처코드						13
Dim C_BpNm				'거래처명						14
Dim C_TrackingNo		'트래킹번호(*)					15
Dim C_LotNo				'로트번호(12)					16
Dim C_LotSubNo			'로트순번(2)					17
Dim C_LotSize			'로트크기(1,000)				18
Dim C_Unit				'단위							19
Dim C_InspReqDt			'검사의뢰일						20
Dim C_InspReqmtDt		'검사요구일						21
Dim C_InspSchdlDt		'검사계획일						22
Dim C_InspStatusNm		'검사현황코드(미검사)			23
Dim C_PRNo				'입고번호(GR20020724055)		24
Dim C_PONo				'발주번호(PS20020724027)		25
Dim C_POSeq				'발주순번(1)					26
Dim C_ProdtNo			'제조오더번호(PD20021224000002)	27
Dim C_ReportSeq			'실적순번(1)					28
Dim C_DocumentNo		'수불번호						29
Dim C_DocumentSeqNo		'수불순번						30
Dim C_DocumentSubNo		'수불지번						31
Dim C_SLCd				'창고코드						32
Dim C_SLNm				'창고							33
Dim C_DNNo				'출하번호						34
Dim C_DNSeq				'출하순번						35
Dim C_InspClassCd		'검사분류코드					36
Dim C_InspStatusCd		'검사현황명						37

Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 

Dim hPlantCd
Dim hInspReqNo
Dim hInspClassCd
Dim hInspstatusCd
Dim hItemCd
Dim hLotNo
Dim hFrInspReqDt
Dim hToInspReqDt

Dim hSupplierCd
Dim hPRNo
Dim hPONo

Dim hRoutNo
Dim hOprNo
Dim hProdtNo1

Dim hSLCd
Dim hProdtNo2

Dim hBPCd
Dim hDNNo

Dim ArrParent

Dim arrParam				'--- First Parameter Group 
ReDim arrParam(5)
Dim arrReturn				'--- Return Parameter Group 


Dim IsOpenPop          
Dim strReqInspClass

'------ Set Parameters from Parent ASP ------ 
ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
arrParam(0) = ArrParent(1)
arrParam(1) = ArrParent(2)
arrParam(2) = ArrParent(3)
arrParam(3) = ArrParent(4)
arrParam(4) = ArrParent(5)

top.document.title = PopupParent.gActivePRAspName
'--------------------------------------------- 

Function InitVariables()
	lgSortKey = 1                            '⊙: initializes sort direction
	lgQueryFlag = "1"
End Function

Sub initSpreadPosVariables()  
	C_InspReqNo = 1
	C_InspClassNm = 2
	C_ItemCd = 3
	C_ItemNm = 4
	C_ItemSpec = 5
	C_SupplierCd = 6
	C_SupplierNm = 7
	C_RoutNo = 8
	C_RoutNoDesc = 9
	C_OprNo = 10
	C_OprNoDesc = 11
	C_WcCd = 12
	C_WcNm = 13
	C_BpCd = 14
	C_BpNm = 15
	C_TrackingNo = 16
	C_LotNo = 17
	C_LotSubNo = 18
	C_LotSize = 19
	C_Unit = 20
	C_InspReqDt = 21
	C_InspReqmtDt = 22
	C_InspSchdlDt = 23
	C_InspStatusNm = 24
	C_PRNo = 25
	C_PONo = 26
	C_POSeq = 27
	C_ProdtNo = 28
	C_ReportSeq = 29
	C_DocumentNo = 30
	C_DocumentSeqNo = 31
	C_DocumentSubNo = 32
	C_SLCd = 33
	C_SLNm = 34
	C_DNNo = 35
	C_DNSeq = 36
	C_InspClassCd = 37
	C_InspStatusCd = 38
End Sub

Sub SetDefaultVal()
	
	txtPlantCd.Value = arrParam(0)
	txtPlantNm.Value = arrParam(1)
	txtInspReqNo.Value = arrParam(2)
	cboInspClassCd.Value = arrParam(3)
	cboInspStatus.value = arrParam(4)
	
	Self.Returnvalue = Array("")
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q","NOCOOKIE","PA") %>
End Sub

Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(cboInspClassCd , lgF0, lgF1, Chr(11))
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0013", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(cboInspStatus , lgF0, lgF1, Chr(11))
End Sub

Sub InitSpreadSheet()
	Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20040518",,PopupParent.gAllowDragDropSpread
	
	With vspdData
		.ReDraw = False
		
		.MaxCols = C_InspStatusCd + 1
		.MaxRows = 0
	End With
	
	Call GetSpreadColumnPos("A")
	
	With ggoSpread
		
		.SSSetEdit C_InspReqNo,"검사의뢰번호", 20
		.SSSetEdit C_InspClassNm,"검사분류", 20
		.SSSetEdit C_ItemCd,"품목코드", 18
		.SSSetEdit C_ItemNm,"품목명", 20
		.SSSetEdit C_ItemSpec,"규격", 30
		
		.SSSetEdit C_SupplierCd,"공급처코드",10
		.SSSetEdit C_SupplierNm,"공급처명",20
		.SSSetEdit C_RoutNo,"라우팅", 15
		.SSSetEdit C_RoutNoDesc,"라우팅설명", 15
		.SSSetEdit C_OprNo,"공정",8
		.SSSetEdit C_OprNoDesc,"공정작업명", 15
		.SSSetEdit C_WCCd,"작업장코드",10
		.SSSetEdit C_WCNm,"작업장명",20
		.SSSetEdit C_BPCd,"거래처코드",10
		.SSSetEdit C_BPNm,"거래처명",20
		
		.SSSetEdit C_TrackingNo,"Tracking No.", 20
		.SSSetEdit C_LotNo,"로트번호",15
		.SSSetFloat C_LotSubNo,"로트순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetFloat C_LotSize,"로트크기", 10, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetEdit C_Unit,"단위", 5, 1
		.SSSetEdit C_InspReqDt,"검사의뢰일",10, 2
		.SSSetEdit C_InspReqmtDt,"검사요구일",10, 2
		.SSSetEdit C_InspSchdlDt,"검사계획일",10, 2
		.SSSetEdit C_InspStatusNm,"검사진행상태", 20
		
		.SSSetEdit C_PRNo,"입고번호",15
		.SSSetEdit C_PONo,"발주번호",15
		.SSSetFloat C_POSeq,"발주순번",8, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		
		.SSSetEdit C_ProdtNo,"제조오더번호", 15
		.SSSetFloat C_ReportSeq,"실적순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetEdit C_DocumentNo,"수불번호", 15
		.SSSetFloat C_DocumentSeqNo,"수불순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetFloat C_DocumentSubNo,"수불지번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		.SSSetEdit C_SLCd,"창고코드",10
		.SSSetEdit C_SLNm,"창고명",20
		.SSSetEdit C_DNNo,"출하번호", 15
		.SSSetFloat C_DNSeq,"출하순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		
		.SSSetEdit C_InspClassCd,"", 1
		.SSSetEdit C_InspStatusCd,"", 1
		
	
		
	End With
		
	Call ChangingFieldsByInspClass(cboInspClassCd.value)
	
	Call ggoSpread.SSSetColHidden(C_InspClassCd, C_InspStatusCd, True)
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
			C_InspClassNm = iCurColumnPos(2)
			C_ItemCd = iCurColumnPos(3)
			C_ItemNm = iCurColumnPos(4)
			C_ItemSpec = iCurColumnPos(5)
			C_SupplierCd = iCurColumnPos(6)
			C_SupplierNm = iCurColumnPos(7)
			C_RoutNo = iCurColumnPos(8)
			C_RoutNoDesc = iCurColumnPos(9)
			C_OprNo = iCurColumnPos(10)
			C_OprNoDesc = iCurColumnPos(11)
			C_WcCd = iCurColumnPos(12)
			C_WcNm = iCurColumnPos(13)
			C_BpCd = iCurColumnPos(14)
			C_BpNm = iCurColumnPos(15)
			C_TrackingNo = iCurColumnPos(16)
			C_LotNo = iCurColumnPos(17)
			C_LotSubNo = iCurColumnPos(18)
			C_LotSize = iCurColumnPos(19)
			C_Unit = iCurColumnPos(20)
			C_InspReqDt = iCurColumnPos(21)
			C_InspReqmtDt = iCurColumnPos(22)
			C_InspSchdlDt = iCurColumnPos(23)
			C_InspStatusNm = iCurColumnPos(24)
			C_PRNo = iCurColumnPos(25)
			C_PONo = iCurColumnPos(26)
			C_POSeq = iCurColumnPos(27)
			C_ProdtNo = iCurColumnPos(28)
			C_ReportSeq = iCurColumnPos(29)
			C_DocumentNo = iCurColumnPos(30)
			C_DocumentSeqNo = iCurColumnPos(31)
			C_DocumentSubNo = iCurColumnPos(32)
			C_SLCd = iCurColumnPos(33)
			C_SLNm = iCurColumnPos(34)
			C_DNNo = iCurColumnPos(35)
			C_DNSeq = iCurColumnPos(36)
			C_InspClassCd = iCurColumnPos(37)
			C_InspStatusCd = iCurColumnPos(38)
			
	End Select    
End Sub

Sub ChangingFieldsByInspClass(Byval strInspClass)
	With ggoSpread
		vspdData.MaxRows = 0
		vspdData.Row = 0
		vspdData.Col = C_InspReqNo
		vspdData.Action = 0
		Select Case strInspClass
			Case "R"
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, False)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, False)
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
				Call .SSSetColHidden(C_PRNo, C_PRNo, False)				'입고번호 
				Call .SSSetColHidden(C_PONo, C_PONo, False)				'발주번호 
				Call .SSSetColHidden(C_POSeq, C_POSeq, False)			'발주순번 
				Call .SSSetColHidden(C_DocumentNo, C_DocumentNo, True)
				Call .SSSetColHidden(C_DocumentSeqNo, C_DocumentSeqNo, True)
				Call .SSSetColHidden(C_DocumentSubNo, C_DocumentSubNo, True)
				Call .SSSetColHidden(C_ProdtNo, C_ProdtNo, True)
				Call .SSSetColHidden(C_ReportSeq, C_ReportSeq, True)
				Call .SSSetColHidden(C_DNNo, C_DNNo, True)
				Call .SSSetColHidden(C_DNSeq, C_DNSeq, True)
				
				
			Case "P"				
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, True)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, True)
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
				Call .SSSetColHidden(C_PRNo, C_PRNo, True)
				Call .SSSetColHidden(C_PONo, C_PONo, True)
				Call .SSSetColHidden(C_POSeq, C_POSeq, True)
				Call .SSSetColHidden(C_DocumentNo, C_DocumentNo, True)
				Call .SSSetColHidden(C_DocumentSeqNo, C_DocumentSeqNo, True)
				Call .SSSetColHidden(C_DocumentSubNo, C_DocumentSubNo, True)
				Call .SSSetColHidden(C_ProdtNo, C_ProdtNo, False)
				Call .SSSetColHidden(C_ReportSeq, C_ReportSeq, False)
				Call .SSSetColHidden(C_DNNo, C_DNNo, True)
				Call .SSSetColHidden(C_DNSeq, C_DNSeq, True)
				
			Case "F"
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, True)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, True)
				Call .SSSetColHidden(C_RoutNo, C_RoutNo, False)
				Call .SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, False)
				Call .SSSetColHidden(C_OprNo, C_OprNo, False)
				Call .SSSetColHidden(C_OprNoDesc, C_OprNoDesc, False)
				Call .SSSetColHidden(C_WCCd, C_WCCd, True)
				Call .SSSetColHidden(C_WCNm, C_WCNm, True)
				Call .SSSetColHidden(C_BPCd, C_BPCd, True)
				Call .SSSetColHidden(C_BPNm, C_BPNm, True)
				Call .SSSetColHidden(C_SLCd, C_SLCd, False)
				Call .SSSetColHidden(C_SLNm, C_SLNm, False)
				Call .SSSetColHidden(C_PRNo, C_PRNo, True)
				Call .SSSetColHidden(C_PONo, C_PONo, True)
				Call .SSSetColHidden(C_POSeq, C_POSeq, True)
				Call .SSSetColHidden(C_DocumentNo, C_DocumentNo, False)
				Call .SSSetColHidden(C_DocumentSeqNo, C_DocumentSeqNo, False)
				Call .SSSetColHidden(C_DocumentSubNo, C_DocumentSubNo, False)
				Call .SSSetColHidden(C_ProdtNo, C_ProdtNo, False)
				Call .SSSetColHidden(C_ReportSeq, C_ReportSeq, False)
				Call .SSSetColHidden(C_DNNo, C_DNNo, True)
				Call .SSSetColHidden(C_DNSeq, C_DNSeq, True)		
				
			Case "S"
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, True)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, True)
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
				Call .SSSetColHidden(C_PRNo, C_PRNo, True)
				Call .SSSetColHidden(C_PONo, C_PONo, True)
				Call .SSSetColHidden(C_POSeq, C_POSeq, True)
				Call .SSSetColHidden(C_DocumentNo, C_DocumentNo, True)
				Call .SSSetColHidden(C_DocumentSeqNo, C_DocumentSeqNo, True)
				Call .SSSetColHidden(C_DocumentSubNo, C_DocumentSubNo, True)
				Call .SSSetColHidden(C_ProdtNo, C_ProdtNo, True)
				Call .SSSetColHidden(C_ReportSeq, C_ReportSeq, True)
				Call .SSSetColHidden(C_DNNo, C_DNNo, False)
				Call .SSSetColHidden(C_DNSeq, C_DNSeq, False)
				
			Case Else
				Call .SSSetColHidden(C_SupplierCd, C_SupplierCd, False)
				Call .SSSetColHidden(C_SupplierNm, C_SupplierNm, False)
				Call .SSSetColHidden(C_RoutNo, C_RoutNo, False)
				Call .SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, False)
				Call .SSSetColHidden(C_OprNo, C_OprNo, True)
				Call .SSSetColHidden(C_OprNoDesc, C_OprNoDesc, False)
				Call .SSSetColHidden(C_WCCd, C_WCCd, False)
				Call .SSSetColHidden(C_WCNm, C_WCNm, False)
				Call .SSSetColHidden(C_BPCd, C_BPCd, False)
				Call .SSSetColHidden(C_BPNm, C_BPNm, False)
				Call .SSSetColHidden(C_SLCd, C_SLCd, False)
				Call .SSSetColHidden(C_SLNm, C_SLNm, False)
				Call .SSSetColHidden(C_PRNo, C_PRNo, False)
				Call .SSSetColHidden(C_PONo, C_PONo, False)
				Call .SSSetColHidden(C_POSeq, C_POSeq, False)
				Call .SSSetColHidden(C_DocumentNo, C_DocumentNo, False)
				Call .SSSetColHidden(C_DocumentSeqNo, C_DocumentSeqNo, False)
				Call .SSSetColHidden(C_DocumentSubNo, C_DocumentSubNo, False)
				Call .SSSetColHidden(C_ProdtNo, C_ProdtNo, False)
				Call .SSSetColHidden(C_ReportSeq, C_ReportSeq, False)
				Call .SSSetColHidden(C_DNNo, C_DNNo, False)
				Call .SSSetColHidden(C_DNSeq, C_DNSeq, False)
				
		End Select
	End With
End Sub

Sub EnableField(Byval strInspClass)
	Select Case strInspClass
		Case "R"
		
			Receiving1.style.display = ""
			Receiving2.style.display = ""
			
			Process1.style.display = "none"
			Process2.style.display = "none"
			
			Final.style.display = "none"
			
			Shipping.style.display = "none"
			
		Case "P"
			Receiving1.style.display = "none"
			Receiving2.style.display = "none"
			
			Process1.style.display = ""
			Process2.style.display = ""
			
			Final.style.display = "none"
			
			Shipping.style.display = "none"
			
		Case "F"
			Receiving1.style.display = "none"
			Receiving2.style.display = "none"
			
			Process1.style.display = "none"
			Process2.style.display = "none"
			
			Final.style.display = ""
			
			Shipping.style.display = "none"
			
		Case "S"
			Receiving1.style.display = "none"
			Receiving2.style.display = "none"
			
			Process1.style.display = "none"
			Process2.style.display = "none"
			
			Final.style.display = "none"
			
			Shipping.style.display = ""
			
		Case Else
			Receiving1.style.display = "none"
			Receiving2.style.display = "none"
			
			Process1.style.display = "none"
			Process2.style.display = "none"
			
			Final.style.display = "none"
			
			Shipping.style.display = "none"
			
	End Select 

End Sub

Sub cboInspClassCd_onchange()
	Call EnableField(cboInspClassCd.value)
	Call ChangingFieldsByInspClass(cboInspClassCd.value)
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

Function OpenRoutNo()


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


'------------------------------------------  OpenRoutNo()  -------------------------------------------------
'	Name : OpenRoutNo()
'	Description : RoutNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRoutNo()

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

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
    
    txtRoutNo.focus
	If arrRet(0) <> "" Then
		txtRoutNo.Value		= arrRet(0)		
		txtRoutNoDesc.Value		= arrRet(1)		
	End If
	
	Set gActiveElement = document.activeElement	
End Function


'------------------------------------------  OpenOprNo()  -------------------------------------------------
'	Name : OpenOprNo()
'	Description : OprNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

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
	
	txtOprNo.focus
	If arrRet(0) <> "" Then
		txtOprNo.Value	= arrRet(0)
		txtOprNoDesc.Value	= arrRet(3)
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
			'msgbox "arrReturn(" & intColCnt & ") : " & arrReturn(intColCnt)
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
	
	Call InitComboBox				'순서를 바꾸면 안됨 
	Call AppendNumberPlace("6", "3","0")
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
	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
	
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

Function FncQuery()
	FncQuery = False
   	
   	vspdData.MaxRows = 0
	lgQueryFlag = "1"
	lgStrPrevKey = ""

	Call EnableField(cboInspClassCd.value)
	Call ChangingFieldsByInspClass(cboInspClassCd.value)
	
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	if DbQuery = false then
		Exit Function
	End if

	FncQuery = True
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
	
	If ValidDateCheck(txtFrInspReqDt, txtToInspReqDt) = False Then
		Exit Function
	End If
	
	'Show Processing Bar
    Call LayerShowHide(1)  

	txtMaxRows = vspdData.MaxRows
	 
	If lgQueryFlag = "0" Then
		strVal = BIZ_PGM_ID & "?QueryFlag=" & lgQueryFlag _
				& "&txtPlantCd=" & hPlantCd _
				& "&txtInspReqNo=" & lgStrPrevKey _
				& "&txtItemCd=" & hItemCd _
				& "&txtInspClassCd=" & hInspClassCd _
				& "&txtLotNo=" & hLotNo _
				& "&txtInspStatusCd=" & hInspstatusCd _
				& "&txtFrInspReqDt=" & hFrInspReqDt _
				& "&txtToInspReqDt=" & hToInspReqDt
				
		Select Case hInspClassCd
			Case "R"
				strVal = strVal & "&txtSupplierCd=" & hSupplierCd _
								& "&txtPRNo=" & hPRNo _
								& "&txtPONo=" & hPONo
				
			Case "P"
				strVal = strVal & "&txtRoutNo=" & hRoutNo _
								& "&txtOprNo=" & hOprNo _
								& "&txtProdtNo=" & hProdtNo1
				
			Case "F"
				strVal = strVal & "&txtSLCd=" & hSLCd _
								& "&txtProdtNo=" & hProdtNo2
			Case "S"
				strVal = strVal & "&txtBpCd=" & hBPCd _
								& "&txtDNNo=" & hDNNo
			Case Else
			
		End Select
	
		strVal = strVal & "&txtMaxRows=" & txtMaxRows		
					
	Else
		strVal = BIZ_PGM_ID & "?QueryFlag=" & lgQueryFlag _
				& "&txtPlantCd=" & Trim(txtPlantCd.Value) _
				& "&txtInspReqNo=" & Trim(txtInspReqNo.Value) _
				& "&txtItemCd=" & Trim(txtItemCd.Value) _
				& "&txtInspClassCd=" & cboInspClassCd.Value _
				& "&txtLotNo=" & Trim(txtLotNo.value) _
				& "&txtInspStatusCd=" & cboInspStatus.Value _
				& "&txtFrInspReqDt=" & txtFrInspReqDt.Text _
				& "&txtToInspReqDt=" & txtToInspReqDt.Text
				
		Select Case cboInspClassCd.value
			Case "R"
				strVal = strVal & "&txtSupplierCd=" & Trim(txtSupplierCd.Value) _
								& "&txtPRNo=" & Trim(txtPRNo.Value) _
								& "&txtPONo=" & Trim(txtPONo.Value)
				
			Case "P"
				strVal = strVal & "&txtRoutNo=" & Trim(txtRoutNo.Value) _
								& "&txtOprNo=" & Trim(txtOprNo.Value) _
								& "&txtProdtNo=" & Trim(txtProdtNo1.Value)
				
			Case "F"
				strVal = strVal & "&txtSLCd=" & Trim(txtSLCd.Value) _
								& "&txtProdtNo=" & Trim(txtProdtNo2.Value)
			Case "S"
				strVal = strVal & "&txtBpCd=" & Trim(txtBpCd.Value) _
								& "&txtDNNo=" & Trim(txtDNNo.Value)
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
									<td CLASS="TD5" NOWPAP>검사진행상태</TD>
									<td CLASS="TD6" NOWPAP><SELECT NAME="cboInspStatus" ALT="검사진행상태" STYLE="WIDTH: 150px" TAG="11"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>로트번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE="20" MAXLENGTH="25" ALT="로트번호" TAG="11XXXU"></TD>
									<TD CLASS="TD5" NOWRAP>검사의뢰일</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q2512pa1_txtFrInspReqDt_txtFrInspReqDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/q2512pa1_txtToInspReqDt_txtToInspReqDt.js'></script>
									</TD>
								</TR>
								<TR ID="Receiving1">
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE="10" MAXLENGTH="10" ALT="공급처" TAG="11XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnSupplierPopup ONCLICK=vbscript:OpenSupplier() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtSupplierNm" TAG="14">
									</TD>
									<TD CLASS="TD5" NOWRAP>입고번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPRNo" SIZE="20" MAXLENGTH="18" ALT="입고번호" TAG="11XXXU" ></TD>
								</TR>
								<TR ID="Receiving2">
									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPONo" SIZE="20" MAXLENGTH="18" ALT="발주번호" TAG="21XXXU"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
				                </TR>
				                <TR ID="Process1">
									<TD CLASS="TD5" NOWRAP>라우팅</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE="20" MAXLENGTH="20" ALT="라우팅" TAG="11XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnRoutNoPopup ONCLICK=vbscript:OpenRoutNo() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtRoutNoDesc" TAG="14"></TD>
									<TD CLASS="TD5" NOWRAP>공정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE="5" MAXLENGTH="3" ALT="공정" TAG="11XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnOprNoPopup ONCLICK=vbscript:OpenOprNo() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtOprNoDesc" TAG="14"></TD>
								</TR>
								<TR ID="Process2">
							    	<TD CLASS="TD5" NOWRAP>제조오더번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProdtNo1" SIZE="20" MAXLENGTH="16" ALT="제조오더번호" TAG="11XXXU"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
							    </TR>
								<TR ID="Final">
									<TD CLASS="TD5" NOWRAP>제조오더번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProdtNo2" SIZE="20" MAXLENGTH="16" ALT="제조오더번호" TAG="11XXXU"></TD>
									<TD CLASS="TD5" NOWRAP>창고</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtSLCd" SIZE="10" MAXLENGTH="7" ALT="창고" TAG="11XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btntxtSLPopup ONCLICK=vbscript:OpenSL() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtSLNm" TAG="14">
									</TD>
								</TR>
							    <TR ID="Shipping">
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtBPCd" SIZE="10" MAXLENGTH="10" ALT="거래처" TAG="11XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnBPPopup ONCLICK=vbscript:OpenBP() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtBPNm" TAG="14">
									</TD>
							    	<TD CLASS="TD5" NOWRAP>출하번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDNNo" SIZE="20" MAXLENGTH="18" ALT="출하번호" TAG="11XXXU"></TD>
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
									<script language =javascript src='./js/q2512pa1_I537726801_vspdData.js'></script>
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
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>  
