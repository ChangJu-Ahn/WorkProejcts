<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M3112MA7
'*  4. Program Name         : 반품발주내역등록 
'*  5. Program Desc         : 반품발주내역등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2005/07/20
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kim Duk Hyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
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
<!-- #Include file="../../inc/incSvrHTML.inc" -->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID 					= "m3112mb7.asp"		
Const BIZ_PGM_JUMP_ID 				= "M3111MA7"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_SeqNo
Dim C_PlantCd
Dim C_Popup1
Dim C_PlantNm
Dim C_itemCd
Dim C_Popup2
Dim C_itemNm
Dim C_SpplSpec
Dim C_OrderQty
Dim C_OrderUnit
Dim C_Popup3
Dim C_Cost
Dim C_CostCon
Dim C_CostConCd
Dim C_NetOrderAmt
Dim C_OrderAmt
Dim C_OrgOrderAmt
Dim C_IOFlg
Dim C_IOFlgCd
Dim C_VatType
Dim C_Popup7
Dim C_VatNm
Dim C_VatRate
Dim C_VatAmt
Dim C_DlvyDT
Dim C_HSCd
Dim C_Popup5
Dim C_HSNm
Dim C_SLCd
Dim C_Popup6
Dim C_SLNm
Dim C_RetCd
Dim C_Popup8
Dim C_RetNm
Dim C_TrackingNo
Dim C_TrackingNoPop
Dim C_Lot_No
Dim C_Popup9
Dim C_Lot_Seq
Dim C_Over
Dim C_Under
Dim C_Bal_Qty
Dim C_Bal_Doc_Amt
Dim C_Bal_Loc_Amt
Dim C_ExRate
Dim C_PrNo
Dim C_MvmtNo
Dim C_PoNo
Dim C_PoSeqNo
Dim C_MaintSeq
Dim C_SoNo
Dim C_SoSeqNo
Dim C_IVNO
Dim C_IVSEQ
Dim C_OrgOrderAmt1
Dim C_OrgNetOrderAmt
Dim C_OrgNetOrderAmt1
Dim C_Stateflg
Dim C_Remrk

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lblnWinEvent
Dim releaseFlg
Dim arrCollectVatType
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop      

Dim EndDate
Dim iDBSYSDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate   = UniConvDateAToB(iDBSYSDate  ,Parent.gServerDateFormat,Parent.gDateFormat)
    
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'==========================================   Release()  ======================================
'	Name : Release()
'	Description : 
'===================================================================================================
Sub Release()
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Sub
	End If                
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Trim(frm1.hdnMode.Value)	
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.Value)
    
    If LayerShowHide(1) = False Then Exit Sub
	Call RunMyBizASP(MyBizASP, strVal)								
End Sub
'==========================================   btnCfm()  ======================================
'	Name : btnCfm()
'	Description : 확정버튼,확정취소버튼의 Event 합수 
'=========================================================================================================
Sub Cfm()
    Dim IntRetCD 
    
    Err.Clear                                                       
    
    if ggoSpread.SSCheckChange = True then
		Call DisplayMsgBox("189217", "X", "X", "X")
		Exit sub
	End if
	
	if Trim(frm1.hdnReleaseflg.Value) = "N" then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		frm1.hdnMode.Value = "Release"
					                                                
	elseif Trim(frm1.hdnReleaseflg.Value) = "Y" then
			
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		
		frm1.hdnMode.Value = "UnRelease"
		
	End if
	
	Call Release()
End Sub

'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)

	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 0 Then

		strTemp = ReadCookie("PoNo")
		
		If strTemp = "" then Exit Function

		frm1.txtPoNo.value =  strTemp
	    
	    if Trim(frm1.txtPoNo.value) <> "" then
			frm1.txtQuerytype.value = "Auto"
			frm1.txthdnPoNo.Value = frm1.txtPoNo.value
			Call dbquery()
	    end if
	    
		WriteCookie "PoNo" , ""
			  
	elseIf Kubun = 1 Then

	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                          
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End if
	    
	    If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
		WriteCookie "PoNo" , frm1.txtPoNo.value
		
		Call PgmJump(BIZ_PGM_JUMP_ID)

	elseIf Kubun = 2 Then
	
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                          
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End if
	    	
	    If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
	    WriteCookie "Process_Step" , "PO"
		WriteCookie "Po_No" , Trim(frm1.txtPoNo.value)
		WriteCookie "Pur_Grp", Trim(frm1.txtGroupCd.Value)
		WriteCookie "Po_Cur", Trim(frm1.txtCurr.Value)
		WriteCookie "Po_Xch", Trim(frm1.hdnXch.Value)
		
		Call PgmJump(BIZ_PGM_JUMP_ID_PUR_CHARGE)
	End IF
End Function

'--------------------------------------------------------------------
'		Name        : SetState()
'		Description : Spread의 Row상태를 "R","C"로 Setting
'					  R-reference 참조      C-InsertRow
'--------------------------------------------------------------------
Sub SetState(byval strState,byval IRow)	
	frm1.vspdData.Row=IRow
	frm1.vspdData.Col=C_Stateflg
	frm1.vspdData.Text=strState
End Sub

'==========================================   ChangeItemPlant()  ======================================
'	Name : ChangeItemPlant()
'	Description : 
'=========================================================================================================
Sub ChangeItemPlant(byVal intStartRow ,byVal IntEndRow)
    Err.Clear                                                       

    Dim strVal
    Dim intIndex
    Dim lGrpCnt
	Dim igColSep,igRowSep
	
	igColSep = Parent.gColSep
	igRowSep = Parent.gRowSep

	with frm1.vspdData
		
		if Trim(frm1.txtHMaintNo.Value) <> "" then Exit Sub
		
    	frm1.txtMode.Value = "LookUpItemPlant"			
	    lGrpCnt = 1
	    strVal = ""
	    
		For intIndex = intStartRow To intEndRow
		
			strVal = strVal & CStr(intIndex) & igColSep
			strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,intIndex,"X","X")) & igColSep
			strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,intIndex,"X","X")) & igColSep
			strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OrderUnit,intIndex,"X","X")) & igRowSep
				
			lGrpCnt = lGrpCnt + 1

			Call .SetText(C_Cost	,	intIndex, "")
			Call .SetText(C_Over	,	intIndex, "")
			Call .SetText(C_Under	,	intIndex, "")
		Next
		
	frm1.txtMaxRows.value = lGrpCnt-1
	frm1.txtSpread.value = strVal

	End with
	
    If LayerShowHide(1) = False Then Exit Sub
    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)				
End Sub

Sub changeItemPlantOK()
	if Trim(frm1.hdnTrackingflg.Value) = "*" then
		ggoSpread.spreadlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
	else
		ggoSpread.spreadUnlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
		ggoSpread.sssetrequired C_TrackingNo, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	end if
End Sub

'==========================================   ChangeItemPlant()  ======================================
'	Name : ChangeItemPlant()
'	Description : 
'=========================================================================================================
Sub ChangeItemPlantForUnit(byVal intStartRow ,byVal IntEndRow)
    Err.Clear                                       

    Dim strVal
    Dim intIndex
    Dim lGrpCnt
	Dim igColSep,igRowSep
	
	igColSep = Parent.gColSep
	igRowSep = Parent.gRowSep

	with frm1.vspdData
		
		if Trim(frm1.txtHMaintNo.Value) <> "" then Exit Sub
		
    	frm1.txtMode.Value = "LookUpItemPlantForUnit"	
	    lGrpCnt = 1
	    strVal = ""
	    
		for intIndex = intStartRow to intEndRow
		
			strVal = strVal & CStr(intIndex) & igColSep
			strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,intIndex,"X","X")) & igColSep
			strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,intIndex,"X","X")) & igColSep
			strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OrderUnit,intIndex,"X","X")) & igRowSep
			
			lGrpCnt = lGrpCnt + 1
		next
		
	frm1.txtMaxRows.value = lGrpCnt-1
	frm1.txtSpread.value = strVal
			
	End with
	
    If LayerShowHide(1) = False Then Exit Sub
    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)					
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

'========================================================================================
' Function Name : FncVatCalc
' Function Desc : VAT 계산 
'========================================================================================
function FncVatCalc(Amt,vatRate)
	dim tmpVatAmt
	tmpVatAmt = Amt * (vatRate/(100 + vatRate))
	tmpVatAmt = UNIConvNumPCToCompanyByCurrency(tmpVatAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")     
	tmpVatAmt = UNICDbl(tmpVatAmt)
	FncVatCalc = tmpVatAmt
end function

'========================================================================================
' Function Name : FncVatToZero
' Function Desc : 해당 행의 VAT 정보를 0으로 세팅(행추가, 구매입고참조,반품출고참조)
'========================================================================================
function FncVatToZero(iRow)
	dim tmpVatAmt
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False

			.vspdData.Col = C_VatRate		'긴급 발주일경우에는 VATRATE = 0
			.vspdData.Text =  UNIConvNumPCToCompanyByCurrency(0, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")

			.vspdData.Col = C_VatAmt		'긴급 발주일경우에는 VATAMT = 0
			.vspdData.Text =  UNIConvNumPCToCompanyByCurrency(0, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")

		.vspdData.ReDraw = True
	end With 
end function

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
    Call SetToolbar("1110000000001111")
    frm1.btnCfmSel.disabled = true
    frm1.btnCfm.value = "확정"
    frm1.txtPoNo.focus 
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_SeqNo 		= 1
	C_PlantCd 		= 2
	C_Popup1		= 3
	C_PlantNm 		= 4
	C_itemCd 		= 5
	C_Popup2 		= 6
	C_itemNm 		= 7
	C_SpplSpec      = 8   
	C_OrderQty		= 9
	C_OrderUnit		= 10
	C_Popup3		= 11
	C_Cost			= 12
	C_CostCon		= 13
	C_CostConCd		= 14
	C_NetOrderAmt	= 15
	C_OrderAmt		= 16
	C_OrgOrderAmt   = 17
	C_IOFlg		    = 18                     
	C_IOFlgCd	    = 19
	C_VatType       = 20
	C_Popup7        = 21
	C_VatNm         = 22
	C_VatRate       = 23
	C_VatAmt        = 24
	C_DlvyDT		= 25
	C_HSCd			= 26
	C_Popup5		= 27
	C_HSNm			= 28
	C_SLCd			= 29
	C_Popup6		= 30
	C_SLNm			= 31
	C_RetCd         = 32
	C_Popup8        = 33
	C_RetNm         = 34
	C_TrackingNo	= 35
	C_TrackingNoPop	= 36
	C_Lot_No        = 37
	C_Popup9        = 38
	C_Lot_Seq       = 39
	C_Over			= 40
	C_Under			= 41
	C_Bal_Qty		= 42
	C_Bal_Doc_Amt	= 43
	C_Bal_Loc_Amt	= 44
	C_ExRate		= 45
	C_PrNo			= 46
	C_MvmtNo		= 47
	C_PoNo			= 48
	C_PoSeqNo		= 49
	C_MaintSeq		= 50
	C_SoNo			= 51
	C_SoSeqNo		= 52
	C_IVNO			= 53
	C_IVSEQ			= 54
	C_OrgOrderAmt1  = 55
	C_OrgNetOrderAmt= 56
	C_OrgNetOrderAmt1= 57
	C_Stateflg		= 58
	C_Remrk			= 59

End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables
	
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021118",,Parent.gAllowDragDropSpread  

    Call AppendNumberPlace("6", "3", "0")

	.ReDraw = false
	
    .MaxCols = C_Remrk+1
    .Col = .MaxCols :	.ColHidden = True
    .MaxRows = 0
	
    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit 		C_PlantCd, "공장", 7,,,4,2
    ggoSpread.SSSetButton 		C_Popup1
    ggoSpread.SSSetEdit 		C_PlantNm, "공장명", 20
    ggoSpread.SSSetEdit 		C_ItemCd, "품목", 18,,,18,2
    ggoSpread.SSSetButton 		C_Popup2
    ggoSpread.SSSetEdit 		C_ItemNm, "품목명", 20
    ggoSpread.SSSetEdit			C_SpplSpec, "품목규격", 20        '품목규격 추가 
    SetSpreadFloatLocal			C_OrderQty, "반품수량",15,1,3
    ggoSpread.SSSetEdit 		C_OrderUnit, "단위", 6,,,3,2
    ggoSpread.sssetButton 		C_Popup3
    SetSpreadFloatLocal			C_Cost, "단가",15,1,4
    ggoSpread.SSSetCombo 		C_CostCon, "단가구분", 10,0,False
    ggoSpread.SSSetCombo 		C_CostConCd, "단가구분코드", 10,0,False
    SetSpreadFloatLocal			C_NetOrderAmt, "발주순금액",15,1,2
    SetSpreadFloatLocal			C_OrgNetOrderAmt, "C_OrgNetOrderAmt",15,1,2
    SetSpreadFloatLocal			C_OrgNetOrderAmt1, "C_OrgNetOrderAmt1",15,1,2        
    SetSpreadFloatLocal			C_OrderAmt, "금액",15,1,2
    SetSpreadFloatLocal			C_OrgOrderAmt, "C_OrgOrderAmt",15,1,2
    SetSpreadFloatLocal			C_OrgOrderAmt1, "C_OrgOrderAmt1",15,1,2
    ggoSpread.SSSetDate 		C_DlvyDt, "납기일", 10, 2, Parent.gDateFormat
    ggoSpread.SSSetEdit 		C_HSCd, "HS부호", 15,,,20,2
    ggoSpread.sssetButton 		C_Popup5
    ggoSpread.SSSetEdit 		C_HSNm, "HS명", 20
    ggoSpread.SSSetEdit 		C_SLCd, "창고", 10,,,7,2
    ggoSpread.SSSetButton 		C_Popup6
    ggoSpread.SSSetEdit 		C_SLNm, "창고명", 20
    ggoSpread.SSSetEdit 		C_TrackingNo, "Tracking No.",  15,,,25,2
    ggoSpread.SSSetButton 		C_TrackingNoPop
    ggoSpread.SSSetEdit 		C_Lot_No, "Lot No.",  20,,,25,2           '13 차 추가 
    ggoSpread.SSSetButton		C_Popup9
	SetSpreadFloatLocal 		C_Lot_Seq, "Lot No.순번", 15,1,6
    
    SetSpreadFloatLocal 		C_Over, "과부족허용율(+)(%)",20,1,6
    SetSpreadFloatLocal 		C_Under,"과부족허용율(-)(%)",20,1,6
    ggoSpread.SSSetCombo		C_IOFlg,"VAT포함여부", 15,2,False               '13 차 추가 
    ggoSpread.SSSetCombo 		C_IOFlgCd, "VAT포함여부코드", 15,2,False
    ggoSpread.SSSetEdit 		C_VatType, "VAT", 7,,,4,2
    ggoSpread.SSSetButton 		C_Popup7
    ggoSpread.SSSetEdit 		C_VatNm, "VAT명", 20 
    SetSpreadFloatLocal			C_VatRate, "VAT율(%)",15,1,5
    SetSpreadFloatLocal			C_VatAmt, "VAT금액",15,1,2
    ggoSpread.SSSetEdit 		C_RetCd , "반품유형", 10,,,5,2
    ggoSpread.SSSetButton 		C_Popup8
    ggoSpread.SSSetEdit 		C_RetNm , "반품유형명", 20 
    SetSpreadFloatLocal			C_Bal_Qty, "Bal. Qty.",15,1,3  
    SetSpreadFloatLocal			C_Bal_Doc_Amt, "Bal. Doc. Amt.",15,1,2  
    SetSpreadFloatLocal			C_Bal_Loc_Amt, "Bal. Loc. Amt.",15,1,2  
    SetSpreadFloatLocal			C_ExRate, "Xch. Rate",15,1,5  
    ggoSpread.SSSetEdit 		C_SeqNo, "순번", 10
    ggoSpread.SSSetEdit 		C_PrNo, "구매요청번호", 20
    ggoSpread.SSSetEdit 		C_MvmtNo, "구매입고번호", 20
 '2003.08 원발주번호와 원발주순번임.
    ggoSpread.SSSetEdit 		C_PoNo, "발주번호", 20
    ggoSpread.SSSetEdit 		C_PoSeqNo, "발주순번", 20
    ggoSpread.SSSetEdit 		C_MaintSeq, "maintseq", 10
	ggoSpread.SSSetEdit 		C_SoNo, "C_SoNo", 10
	ggoSpread.SSSetEdit 		C_SoSeqNo, "C_SoSeqNo", 10
	ggoSpread.SSSetEdit 		C_IVNO, "C_IVNO", 10
	ggoSpread.SSSetEdit 		C_IVSEQ, "C_IVSEQ", 10
    ggoSpread.SSSetEdit 		C_Stateflg, "stateflg", 10
    ggoSpread.SSSetEdit 	C_Remrk, "비고", 20,,,120,2

	Call ggoSpread.MakePairsColumn(C_PlantCd,C_Popup1)
	Call ggoSpread.MakePairsColumn(C_ItemCd,C_Popup2)
	Call ggoSpread.MakePairsColumn(C_OrderUnit,C_Popup3)
	Call ggoSpread.MakePairsColumn(C_HSCd,C_Popup5)
	Call ggoSpread.MakePairsColumn(C_SLCd,C_Popup6)
    Call ggoSpread.MakePairsColumn(C_TrackingNo,C_TrackingNoPop)
	Call ggoSpread.MakePairsColumn(C_Lot_No,C_Popup9)
	Call ggoSpread.MakePairsColumn(C_VatType,C_Popup7)
	Call ggoSpread.MakePairsColumn(C_RetCd,C_Popup8)

	Call ggoSpread.SSSetColHidden(C_HSCd,C_HSNm,True)	
	Call ggoSpread.SSSetColHidden(C_CostCon,C_CostConCd,True)	
	Call ggoSpread.SSSetColHidden(C_IOFlg,C_VatNm,True)	
	Call ggoSpread.SSSetColHidden(C_Over,C_Stateflg,True)	
	Call ggoSpread.SSSetColHidden(C_SeqNo,C_SeqNo,True)	
	Call ggoSpread.SSSetColHidden(C_OrgOrderAmt,C_OrgOrderAmt,True)	
       
    ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
    ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
    ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
    ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd
    
    Call SetSpreadLock
    
	.ReDraw = true
	
    End With
End Sub
'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    ggoSpread.SpreadLock frm1.vspddata.maxcols,-1
    ggoSpread.SpreadLock C_SeqNo , -1
    ggoSpread.SpreadLock C_PlantCd , -1
    ggoSpread.SpreadLock C_Popup1 , -1
    ggoSpread.spreadlock C_PlantNm , -1
    ggoSpread.SpreadLock C_ItemCd, -1
    ggoSpread.spreadlock C_SpplSpec,-1         '품목규격 추가 
    ggoSpread.SpreadLock C_Popup2 , -1
    ggoSpread.spreadlock C_ItemNm , -1
    ggoSpread.SpreadUnLock C_OrderQty, -1
    ggoSpread.sssetrequired C_OrderQty, -1
    ggoSpread.SpreadUnLock C_OrderUnit , -1
    ggoSpread.sssetrequired C_OrderUnit, -1
    ggoSpread.SpreadUnLock C_Popup3 , -1
    ggoSpread.SpreadUnLock C_Cost , -1
    ggoSpread.sssetrequired C_Cost, -1
    ggoSpread.SpreadUnLock C_CostCon, -1
 	ggoSpread.spreadlock C_NetOrderAmt, -1    
    ggoSpread.spreadlock C_OrderAmt, -1
    ggoSpread.spreadlock C_OrgOrderAmt, -1
    ggoSpread.SpreadUnLock C_DlvyDT, -1
    ggoSpread.sssetrequired C_DlvyDT, -1
    ggoSpread.spreadlock C_HSCd, -1
    ggoSpread.spreadlock C_Popup5, -1
    ggoSpread.spreadlock C_HSNm, -1
    ggoSpread.SpreadUnLock C_SLCd , -1
    ggoSpread.sssetrequired C_SLCd, -1
    ggoSpread.SpreadUnLock C_Popup6 , -1
    ggoSpread.spreadlock C_SLNm, -1
	ggoSpread.spreadlock C_VatType , -1
	ggoSpread.spreadlock C_Popup7 , -1
	ggoSpread.spreadlock C_VatNm , -1
	ggoSpread.spreadlock C_VatRate , -1
	ggoSpread.spreadlock C_VatAmt , -1 
	ggoSpread.spreadlock C_IOFlg , -1    '13차추가	
    ggoSpread.SpreadUnLock C_Popup8 , -1
    ggoSpread.spreadlock C_RetNm , -1
    ggoSpread.spreadlock C_Lot_No , -1     '13차추가 
    ggoSpread.spreadlock C_Lot_Seq , -1    '13차추가 
    ggoSpread.spreadlock C_TrackingNo , -1  
    ggoSpread.SpreadUnLock C_Remrk , -1
    
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
			
			C_SeqNo 		= iCurColumnPos(1)
			C_PlantCd 		= iCurColumnPos(2)
			C_Popup1		= iCurColumnPos(3)
			C_PlantNm 		= iCurColumnPos(4)
			C_itemCd 		= iCurColumnPos(5)
			C_Popup2 		= iCurColumnPos(6)
			C_itemNm 		= iCurColumnPos(7)
			C_SpplSpec      = iCurColumnPos(8)
			C_OrderQty		= iCurColumnPos(9)
			C_OrderUnit		= iCurColumnPos(10)
			C_Popup3		= iCurColumnPos(11)
			C_Cost			= iCurColumnPos(12)
			C_CostCon		= iCurColumnPos(13)
			C_CostConCd		= iCurColumnPos(14)
			C_NetOrderAmt	= iCurColumnPos(15)
			C_OrderAmt		= iCurColumnPos(16)
			C_OrgOrderAmt   = iCurColumnPos(17)
			C_IOFlg		    = iCurColumnPos(18)
			C_IOFlgCd	    = iCurColumnPos(19)
			C_VatType       = iCurColumnPos(20)
			C_Popup7        = iCurColumnPos(21)
			C_VatNm         = iCurColumnPos(22)
			C_VatRate       = iCurColumnPos(23)
			C_VatAmt        = iCurColumnPos(24)
			C_DlvyDT		= iCurColumnPos(25)
			C_HSCd			= iCurColumnPos(26)
			C_Popup5		= iCurColumnPos(27)
			C_HSNm			= iCurColumnPos(28)
			C_SLCd			= iCurColumnPos(29)
			C_Popup6		= iCurColumnPos(30)
			C_SLNm			= iCurColumnPos(31)
			C_RetCd         = iCurColumnPos(32)
			C_Popup8        = iCurColumnPos(33)
			C_RetNm         = iCurColumnPos(34)
			C_TrackingNo	= iCurColumnPos(35)
			C_TrackingNoPop	= iCurColumnPos(36)
			C_Lot_No        = iCurColumnPos(37)
			C_Popup9        = iCurColumnPos(38)
			C_Lot_Seq       = iCurColumnPos(39)
			C_Over			= iCurColumnPos(40)
			C_Under			= iCurColumnPos(41)
			C_Bal_Qty		= iCurColumnPos(42)
			C_Bal_Doc_Amt	= iCurColumnPos(43)
			C_Bal_Loc_Amt	= iCurColumnPos(44)
			C_ExRate		= iCurColumnPos(45)
			C_PrNo			= iCurColumnPos(46)
			C_MvmtNo		= iCurColumnPos(47)
			C_PoNo			= iCurColumnPos(48)
			C_PoSeqNo		= iCurColumnPos(49)
			C_MaintSeq		= iCurColumnPos(50)
			C_SoNo			= iCurColumnPos(51)
			C_SoSeqNo		= iCurColumnPos(52)
			C_IVNO			= iCurColumnPos(53)
			C_IVSEQ			= iCurColumnPos(54)
			C_OrgOrderAmt1  = iCurColumnPos(55)
			C_OrgNetOrderAmt= iCurColumnPos(56)
			C_OrgNetOrderAmt1= iCurColumnPos(57)
			C_Stateflg		= iCurColumnPos(58)
			C_Remrk			= iCurColumnPos(59)

	End Select

End Sub	

Sub InitSpreadAfterQuery()
    	frm1.vspdData.ReDraw = False			
    	
    	if UCase(frm1.hdnRetflg.Value)="Y" and UCase(frm1.hdnIVFlg.Value)="Y" then
			frm1.vspdData.Col = C_VatType:		frm1.vspdData.ColHidden = false
			frm1.vspdData.Col = C_Popup7:       frm1.vspdData.ColHidden = false
			frm1.vspdData.Col = C_VatAmt:		frm1.vspdData.ColHidden = false
			frm1.vspdData.Col = C_VatRate:		frm1.vspdData.ColHidden = false
			frm1.vspdData.Col = C_VatNm:		frm1.vspdData.ColHidden = false
			frm1.vspdData.Col = C_IOFlg:  	    frm1.vspdData.ColHidden = false
		else
			frm1.vspdData.Col = C_VatType:		frm1.vspdData.ColHidden = true
			frm1.vspdData.Col = C_Popup7:       frm1.vspdData.ColHidden = true
			frm1.vspdData.Col = C_VatAmt:		frm1.vspdData.ColHidden = true
			frm1.vspdData.Col = C_VatRate:		frm1.vspdData.ColHidden = true
			frm1.vspdData.Col = C_VatNm:		frm1.vspdData.ColHidden = true
			frm1.vspdData.Col = C_IOFlg:  	    frm1.vspdData.ColHidden = true
		end if
    	frm1.vspdData.ReDraw = True	
End Sub

Sub SetSpreadLockAfterQuery()
	Dim index,Count,index1 
	Dim chkMvmt, chkIV 
	Dim strMvmtNo,strIvNo

    With frm1
    
    Call InitSpreadAfterQuery

    .vspdData.ReDraw = False
    
    if .txtRelease.Value = "Y" then
		For index = C_SeqNo to C_Stateflg
			ggoSpread.SpreadLock index , -1
		Next		
	Else
	    For index1 = Cint(.hdnmaxrow.value) + 1 to .vspdData.MaxRows
			ggoSpread.SpreadLock C_SeqNo , index1,C_SeqNo,index1
			ggoSpread.SpreadLock C_PlantCd , index1,C_PlantCd,index1
			ggoSpread.SpreadLock C_Popup1 , index1,C_Popup1,index1
			ggoSpread.spreadlock C_PlantNm , index1,C_PlantNm,index1
			ggoSpread.SpreadLock C_ItemCd, index1,C_ItemCd,index1
			ggoSpread.SpreadLock C_Popup2 , index1,C_Popup2,index1
			ggoSpread.spreadlock C_ItemNm , index1,C_ItemNm,index1
			ggoSpread.spreadlock C_SpplSpec, index1,C_SpplSpec,index1         '품목규격 추가 
			ggoSpread.SpreadUnLock C_OrderQty, index1,C_OrderQty,index1
			ggoSpread.sssetrequired C_OrderQty, index1,index1

			if UCase(frm1.hdnRetflg.Value) = "N" then
				ggoSpread.SpreadUnLock C_OrderUnit , index1,C_OrderUnit,index1
				ggoSpread.sssetrequired C_OrderUnit, index1,index1 
				ggoSpread.SpreadUnLock C_Popup3 ,index1,C_Popup3,index1
				ggoSpread.SpreadUnLock C_DlvyDT, index1,C_DlvyDT,index1
				ggoSpread.SpreadUnLock C_Cost , index1,C_Cost,index1
				ggoSpread.sssetrequired C_Cost, index1,index1
			else
				ggoSpread.SpreadUnLock C_Remrk, index1,C_Remrk,index1
				ggoSpread.SpreadLock C_OrderUnit , index1,C_OrderUnit,index1
				ggoSpread.SpreadLock C_Popup3 , index1,C_Popup3,index1
				ggoSpread.SpreadLock C_DlvyDT, index1,C_DlvyDT,index1
				ggoSpread.SpreadLock C_Cost , index1,C_Cost,index1
			end if		

			ggoSpread.SpreadUnLock C_CostCon, index1,C_CostCon,index1
			ggoSpread.spreadlock C_NetOrderAmt, index1,C_NetOrderAmt,index1    		
			ggoSpread.spreadlock C_OrderAmt, index1,C_OrderAmt,index1
			ggoSpread.spreadlock C_OrgOrderAmt, index1,C_OrgOrderAmt,index1		
			if UCase(frm1.hdnRetflg.Value) <> "N" then
				ggoSpread.SpreadUnLock C_DlvyDT, index1,C_DlvyDT,index1
				ggoSpread.sssetrequired C_DlvyDT, index1,index1
			else
				ggoSpread.SpreadLock C_DlvyDT, index1,C_DlvyDT,index1
			end if	
			if .hdnImportflg.value = "Y" then
				ggoSpread.spreadUnlock C_HSCd , index1,C_HSCd,index1
				ggoSpread.sssetrequired C_HSCd, index1,index1
				ggoSpread.spreadUnlock C_Popup5 , index1,C_Popup5,index1
				ggoSpread.spreadlock C_HSNm , index1,C_HSNm,index1				
			else
				ggoSpread.spreadlock C_HSCd, index1,C_HSCd,index1
				ggoSpread.spreadlock C_Popup5, index1,C_Popup5,index1
				ggoSpread.spreadlock C_HSNm, index1,C_HSNm,index1
			End if	

			ggoSpread.spreadlock C_TrackingNo , index1,C_TrackingNo,index1
			ggoSpread.spreadlock C_VatType , index1,C_VatType,index1
			ggoSpread.spreadlock C_Popup7 , index1,C_Popup7,index1
			ggoSpread.spreadlock C_VatNm , index1,C_VatNm,index1
			ggoSpread.spreadlock C_VatRate , index1,C_VatRate,index1
			ggoSpread.spreadlock C_VatAmt , index1,C_VatAmt,index1 
			ggoSpread.spreadlock C_IOFlg , index1,C_IOFlg,index1    '13차추가	        
			ggoSpread.SpreadUnLock C_SLCd , index1,C_SLCd,index1
			ggoSpread.sssetrequired C_SLCd, index1,index1
			ggoSpread.SpreadUnLock C_Popup6 , index1,C_Popup6,index1
			ggoSpread.spreadlock C_SLNm, index1,C_SLNm,index1
			ggoSpread.spreadUnLock C_RetCd , index1,C_RetCd,index1
			ggoSpread.SpreadUnLock C_Popup8 , index1,C_Popup8,index1
			ggoSpread.spreadlock   C_RetNm , index1,C_RetNm,index1
			ggoSpread.spreadlock C_Lot_No , index1,C_Lot_No,index1       
			ggoSpread.spreadlock C_Lot_Seq , index1,C_Lot_Seq,index1 
		
			.vspdData.Row = index1
			.vspdData.Col = C_TrackingNo
			if Trim(.vspdData.Text) = "*" then
				ggoSpread.spreadlock C_TrackingNo, index1, C_TrackingNoPop, index1
			else
				'입고,출고,매입참조시에는 TrackingNo수정못함. 긴급반품발주인건만 수정가능함. 200309				
				.vspdData.Col = C_MvmtNo				
				strMvmtNo = Trim(.vspdData.Text)
				.vspdData.Col = C_IVNO				
				strIvNo = Trim(.vspdData.Text)
				If  strMvmtNo = "" and strIvNo = "" then 
					ggoSpread.spreadUnlock C_TrackingNo, index1, C_TrackingNoPop, index1
					ggoSpread.sssetrequired C_TrackingNo, index1, index1
				Else
					ggoSpread.spreadlock C_TrackingNo, index1, C_TrackingNoPop, index1
				End if 
			end if	

		    .vspdData.Col = C_Lot_No
			if Trim(.vspdData.Text) = "*" then
				ggoSpread.spreadlock C_Lot_No, index1, C_Lot_Seq, index1
			else
				ggoSpread.spreadUnlock C_Lot_No, index1, C_Lot_Seq, index1
			end if		    
		
			frm1.vspdData.Row = index1
		    frm1.vspdData.Col = C_PrNo
			if Trim(.vspdData.Text) <> "" then
				ggoSpread.spreadlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.spreadlock C_Popup3 , index1, C_Popup3, index1
				ggoSpread.spreadlock C_Cost, index1, C_Cost, index1
				ggoSpread.spreadlock C_DlvyDT, index1, C_DlvyDT, index1
       		else
				ggoSpread.spreadUnlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.sssetrequired C_OrderUnit, index1, index1
				ggoSpread.spreadUnlock C_Popup3 , index1, C_Popup3, index1
				ggoSpread.spreadUnlock C_Cost, index1, C_Cost, index1
				ggoSpread.sssetrequired C_Cost, index1, index1
				ggoSpread.spreadUnlock C_DlvyDT, index1, C_DlvyDT, index1
				ggoSpread.sssetrequired C_DlvyDT, index1, index1
			end if

		    frm1.vspdData.Col = C_MvmtNo
		    chkMvmt = .vspdData.Text
		    frm1.vspdData.Col = C_IVNO
		    chkIV = .vspdData.Text
		    
			If chkMvmt <> "" or chkIV <> "" Then
				ggoSpread.spreadlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.spreadlock C_Popup3 , index1, C_Popup3, index1							
				ggoSpread.spreadlock C_Cost, index1, C_Cost, index1
				ggoSpread.spreadlock C_Lot_No, index1, C_Lot_No, index1
				ggoSpread.spreadlock C_Popup9 , index1, C_Popup9, index1
				ggoSpread.spreadlock C_Lot_Seq, index1, C_Lot_Seq, index1
			End If			
		next
			
	End if
	
    .vspdData.ReDraw = True
    
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SeqNo		 , pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_PlantCd	 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm 	 , pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_ItemCd	 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm	 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SpplSpec	 , pvStartRow, pvEndRow '품목규격 추가 
    ggoSpread.SSSetRequired	C_OrderQty	 , pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_OrderUnit	 , pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_Cost		 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_CostConCd	 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_OrderAmt	 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_OrgOrderAmt, pvStartRow, pvEndRow 
    ggoSpread.SSSetProtected	C_NetOrderAmt, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_DlvyDt	 , pvStartRow, pvEndRow
    
    if Trim(frm1.hdnImportflg.value) <> "Y" then
	    ggoSpread.SSSetProtected	C_HSCd	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_Popup5, pvStartRow, pvEndRow
	else
		ggoSpread.spreadUnlock	C_HSCd	, pvStartRow, C_HSCd, pvEndRow
		ggoSpread.sssetrequired	C_HSCd	, pvStartRow, pvEndRow
		ggoSpread.spreadUnlock	C_Popup5, pvStartRow, C_Popup5, pvEndRow
	end if
	
	ggoSpread.SSSetProtected C_TrackingNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_HSNm		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_SLCd		, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SLNm		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_VatType	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_VatAmt	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_VatRate	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_VatNm		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_IOFlg		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_IOFlgCd	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Lot_No	, pvStartRow, pvEndRow        
	ggoSpread.SSSetProtected C_Lot_Seq	, pvStartRow, pvEndRow       
    
    '******************************************
	if .hdnRetflg.Value <> "Y" then
		ggoSpread.SSSetProtected C_RetCd	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Popup8, pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected C_RetNm	, pvStartRow, pvEndRow
	end if        
	'******************************************
    End With
End Sub

'================================== 2.2.5 SetSpreadColorRef() ==================================================
' Function Name : SetSpreadColorRef
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColorRef(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SeqNo		 , pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_PlantCd	, pvStartRow, C_PlantCd, pvEndRow
	ggoSpread.spreadlock		C_Popup1	, pvStartRow, C_Popup1,  pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm 	 , pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_ItemCd	, pvStartRow, C_ItemCd, pvEndRow
	ggoSpread.spreadlock		C_Popup2	, pvStartRow, C_Popup2,  pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm	 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SpplSpec	 , pvStartRow, pvEndRow '품목규격 추가 
    ggoSpread.SSSetRequired		C_OrderQty	 , pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_OrderUnit	, pvStartRow, C_Popup3, pvEndRow
    ggoSpread.spreadlock		C_Cost	, pvStartRow, C_Cost, pvEndRow
    ggoSpread.SSSetProtected	C_CostConCd	 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_OrderAmt	 , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_OrgOrderAmt, pvStartRow, pvEndRow 
    ggoSpread.SSSetProtected	C_NetOrderAmt, pvStartRow, pvEndRow
    ggoSpread.spreadUnLock	C_DlvyDT ,pvStartRow,C_DlvyDT,pvEndRow
	ggoSpread.SSSetRequired	C_DlvyDT ,pvStartRow, pvEndRow
	
    if Trim(frm1.hdnImportflg.value) <> "Y" then
	    ggoSpread.SSSetProtected	C_HSCd	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_Popup5, pvStartRow, pvEndRow
	else
		ggoSpread.spreadUnlock	C_HSCd	, pvStartRow, C_HSCd, pvEndRow
		ggoSpread.sssetrequired	C_HSCd	, pvStartRow, pvEndRow
		ggoSpread.spreadUnlock	C_Popup5, pvStartRow, C_Popup5, pvEndRow
	end if
	
    ggoSpread.SSSetProtected	C_HSNm		, pvStartRow, pvEndRow
    ggoSpread.spreadUnLock	C_SLCd ,pvStartRow,C_SLCd,pvEndRow
	ggoSpread.SSSetRequired	C_SLCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SLNm		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_VatType	, pvStartRow, pvEndRow
	ggoSpread.SpreadLock	C_Popup7 ,pvStartRow,C_Popup7,pvEndRow
	ggoSpread.SSSetProtected C_VatAmt	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_VatRate	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_VatNm		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_IOFlg		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_IOFlgCd	, pvStartRow, pvEndRow

    ggoSpread.spreadUnLock C_RetCd ,pvStartRow,C_RetCd,pvEndRow
    ggoSpread.SpreadUnLock C_Popup8 ,pvStartRow,C_Popup8,pvEndRow
    ggoSpread.spreadlock   C_RetNm , pvStartRow,C_RetNm,pvEndRow
	ggoSpread.spreadlock C_Lot_No , pvStartRow,C_Lot_No,pvEndRow      
	ggoSpread.spreadlock C_Lot_Seq ,pvStartRow,C_Lot_Seq,pvEndRow

	ggoSpread.spreadLock	C_TrackingNo ,pvStartRow,C_TrackingNoPop,pvEndRow
    ggoSpread.SpreadLock	C_Popup9 ,pvStartRow,C_Popup9,pvEndRow
	
    End With
End Sub


'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If lblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M3111PA7")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "M3111PA7", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,"Y"), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function

'------------------------------------------  OpenRetRef()  -------------------------------------------------
'	Name : OpenRetRef()
'	Description : 반품출고참조 
'---------------------------------------------------------------------------------------------------------
Function OpenRetRef()
	Dim strRet
	Dim arrParam(15)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_CMODE then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtPoNo.focus	
		Exit Function
	End if 
	
	if frm1.txtRelease.Value = "Y" then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		frm1.txtPoNo.focus	
		Exit Function
	End if
	
	if Not(UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "Y") then
		Call DisplayMsgBox("17A012", "X","반품형태" & frm1.txtPotypeCd.Value & "(" & frm1.txtPoTypeNm.value & ")","반품출고참조" )
		frm1.txtPoNo.focus	
		Exit Function
	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	'===============쿨젠====================
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = ""
	arrParam(4) = Trim(frm1.hdnClsflg.value)
	arrParam(5) = Trim(frm1.hdnReleaseflg.value)
	arrParam(6) = Trim(frm1.hdnRcptflg.value)
	arrParam(7) = Trim(frm1.hdnRetflg.value)
	arrParam(8) = "PO" 	'RefType
	arrParam(9) = Trim(frm1.hdnRcptType.value)	'RcptType
	arrParam(10) = ""	'Ivflg
	arrParam(11) = ""	'IvType
	arrParam(12) = Trim(frm1.hdnRefPoNo.value)	'RefPoNo
	arrParam(13) = Trim(frm1.txtCurr.value)		'Currency
	arrParam(14) = Trim(frm1.hdnSubcontraflg.value)	'Subcontraflg
	'===============쿨젠====================

	iCalledAspName = AskPRAspName("M3113RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3113RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		frm1.txtPoNo.focus	
		Exit Function
	Else
		Call SetRetRef(strRet)
	End If	
End Function

Function SetRetRef(strRet)
	Dim Index1,index2,Index3,Count1,Count2
	Dim IntIflg
	Dim strMessage
	Dim intstartRow,intEndRow, TempRow
	Dim comtemp1,comtemp2,temp
	Dim iInsRow
	
	Const C_PlantCd_Ref		= 0
	Const C_PlantNm_Ref		= 1
	Const C_ItemCd_Ref		= 2
	Const C_ItemNm_Ref		= 3
	Const C_PoQty_Ref		= 4
	Const C_MvmtQty_Ref		= 5		'반품출고수량 
	Const C_RetOrdQty_Ref	= 6		'재입고오더수량 
	Const C_TotRetQty_ref	= 7
	Const C_Unit_Ref		= 8
	Const C_MvmtDt_Ref		= 9
	Const C_GrNo_Ref		= 10
	Const C_GmNo_Ref		= 11
	Const C_GmSeqNo_Ref		= 12
	Const C_PoNo_Ref		= 13
	Const C_PoSeq_Ref		= 14
	Const C_TrackingNo_Ref  = 15	
	Const C_Lot_No_Ref		= 16
	Const C_Lot_Seq_Ref		= 17
	Const C_SpplSpec_Ref    = 18
	Const C_Prc_Ref 		= 19
	Const C_Amt_Ref 		= 20
	Const C_IvQty_Ref		= 21
	Const C_SlCd_Ref		= 22
	Const C_SlNm_Ref		= 23
	Const C_Trackingflg_Ref = 24
	Const C_MvmtNo_Ref		= 25

	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	strMessage = ""
	IntIflg=true
	
	with frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		intStartRow = .vspdData.MaxRows + 1
		
		.vspdData.Redraw = False
		
		TempRow = .vspdData.MaxRows					'리스트 max값 
		
	for index1 = 0 to Count1
	
		.vspdData.Row=Index1+1
		If TempRow <> 0 Then
			for Index3=1 to TempRow
				if Trim(GetSpreadText(.vspdData,C_PoNo,Index3,"X","X")) = strRet(index1,C_PoNo_Ref) and GetSpreadText(.vspdData,C_PoSeqNo,Index3,"X","X") = strRet(Index1,C_PoSeq_Ref) then
					strMessage = strMessage & strRet(Index1,C_PoNo_Ref) & "-" & strRet(Index1,C_PoSeq_Ref) & ";"
					intIflg=False
					Exit for
				End if 
			Next
		End If
		
		if IntIflg <> False then
			.vspdData.MaxRows = CLng(TempRow) + CLng(index1) + 1
			iInsRow = CLng(TempRow) + CLng(index1) + 1

			Call .vspdData.SetText(0		,	iInsRow, ggoSpread.InsertFlag)
			'Call .vspdData.SetText(C_DlvyDT,	iInsRow, .hdnDlvydt.value)
			Call .vspdData.SetText(C_DlvyDT,	iInsRow, EndDate)
			'Insert Row 시 헤더의 부가세관련 정보 초기값으로 2002.2.19
			Call .vspdData.SetText(C_VatType,	iInsRow, .hdnVATType.value)
	 
			if Trim(.hdnVATINCFLG.value) = "2" then	'포함 
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,0,"X","X")			
				Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,0,"X","X")			
			else
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,1,"X","X")			
				Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,1,"X","X")			
			end if
	
			if .hdnVATType.value <> "" then
				call SetVatType(iInsRow)
			end if

			Call .vspdData.SetText(C_TrackingNo,	iInsRow, "*")
			Call .vspdData.SetText(C_Lot_No,	iInsRow, "*")
			Call SetSpreadValue(.vspdData,C_CostCon	,iInsRow,1,"X","X")			
			Call SetSpreadValue(.vspdData,C_CostConCd	,iInsRow,1,"X","X")			

			Call SetState("C",iInsRow)
				
			Call .vspdData.SetText(C_PlantCd	,	iInsRow, strRet(index1,C_PlantCd_Ref))
			Call .vspdData.SetText(C_PlantNm	,	iInsRow, strRet(index1,C_PlantNm_Ref))
			Call .vspdData.SetText(C_itemCd		,	iInsRow, strRet(index1,C_ItemCd_Ref))
			Call .vspdData.SetText(C_itemNm		,	iInsRow, strRet(index1,C_ItemNm_Ref))
			Call .vspdData.SetText(C_SpplSpec	,	iInsRow, strRet(index1,C_SpplSpec_Ref))
			Call .vspdData.SetText(C_OrderQty	,	iInsRow, strRet(index1,C_MvmtQty_Ref))
			Call .vspdData.SetText(C_OrderUnit	,	iInsRow, strRet(index1,C_Unit_Ref))
			Call .vspdData.SetText(C_Cost		,	iInsRow, strRet(index1,C_Prc_Ref))
			Call .vspdData.SetText(C_OrderAmt	,	iInsRow, strRet(index1,C_Amt_Ref))
			Call .vspdData.SetText(C_OrgOrderAmt,	iInsRow, strRet(index1,C_Amt_Ref))
			Call .vspdData.SetText(C_SLCd		,	iInsRow, strRet(index1,C_SLCd_Ref))
			Call .vspdData.SetText(C_SLNm		,	iInsRow, strRet(index1,C_SLNm_Ref))
			Call .vspdData.SetText(C_Lot_No		,	iInsRow, strRet(index1,C_Lot_No_Ref))
			Call .vspdData.SetText(C_Lot_Seq	,	iInsRow, strRet(index1,C_Lot_Seq_Ref))
			Call .vspdData.SetText(C_TrackingNo	,	iInsRow, strRet(index1,C_TrackingNo_Ref))
			Call .vspdData.SetText(C_PoNo		,	iInsRow, strRet(index1,C_PoNo_Ref))
			Call .vspdData.SetText(C_PoSeqNo	,	iInsRow, strRet(index1,C_PoSeq_Ref))
			Call .vspdData.SetText(C_MvmtNo		,	iInsRow, strRet(index1,C_MvmtNo_Ref))

'			IF strRet(index1,C_TotRetQty_ref) <> "" Then
'				temp = UNICDbl(GetSpreadText(.vspdData,C_OrderQty,iInsRow,"X","X")) - UNICDbl(strRet(index1,C_TotRetQty_ref))
'				Call .vspdData.SetText(C_OrderQty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
'			End If
			' 반품출고수량 - 재입고오더수량 
			IF strRet(index1,C_RetOrdQty_ref) <> "" Then
				temp= UNICDbl(strRet(index1,C_MvmtQty_ref)) - UNICDbl(strRet(index1,C_RetOrdQty_ref))
				Call .vspdData.SetText(C_OrderQty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			End If
			
			IF strRet(index1,C_IvQty_ref) <> "" Then
				temp = UNICDbl(GetSpreadText(.vspdData,C_OrderQty,iInsRow,"X","X")) - UNICDbl(strRet(index1,C_IvQty_ref))
				Call .vspdData.SetText(C_OrderQty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			End If
			
		Else
			IntIFlg=True
		End if 
	    
	next
	
	intEndRow = .vspdData.MaxRows

	Call SetSpreadColorRef(intStartRow,intEndRow)
	
	if strMessage<>"" then
		Call DisplayMsgBox("17a005", "X",strmessage,"반품번호" & "," & "순번")
		.vspdData.ReDraw = True
		Exit Function
	End if
	
	.vspdData.ReDraw = True
	
	End with
	
	for index1 = intStartRow to intEndRow
		 Call vspdData_Change2(C_Cost, index1)
	next 

End Function

'------------------------------------------  OpenIvRef()  -------------------------------------------------
'	Name : OpenIvRef()
'	Description :매입참조 
'---------------------------------------------------------------------------------------------------------
Function OpenIvRef()
	Dim strRet
	Dim arrParam(4)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_CMODE then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtPoNo.focus	
		Exit Function
	End if 
	
	if frm1.txtRelease.Value = "Y" then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		frm1.txtPoNo.focus	
		Exit Function
	End if
	
	' 대물정산 반품출고 매입참조 가능하게 
	'if Not(UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnIVFlg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "N") then  
	if Not(UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "N") then 
		Call DisplayMsgBox("17A012", "X","반품형태" & frm1.txtPotypeCd.Value & "(" & frm1.txtPoTypeNm.value & ")","매입참조" )
		frm1.txtPoNo.focus	
		Exit Function
	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtCurr.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = Trim(frm1.hdnRefPoNo.value)
	arrParam(4) = Trim(frm1.hdnSubcontraflg.value)

	iCalledAspName = AskPRAspName("M5111RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5111RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		frm1.txtPoNo.focus	
		Exit Function
	Else
		Call SetIvRef(strRet)
	End If
End Function

Function SetIvRef(strRet)
	Dim Index1,index2,Index3,Count1,Count2
	Dim IntIflg,iInsRow
	Dim strMessage
	Dim intstartRow,intEndRow,temp, TempRow

	Const C_ReqNo_Ref 			= 0
	Const C_PlantCd_Ref 		= 0
	Const C_PlantNm_Ref 		= 1
	Const C_ItemCd_Ref 			= 2
	Const C_ItemNm_Ref			= 3
	Const C_Qty_Ref 			= 5						'매입수량 
	Const C_Unit_Ref 			= 6						'단위 
	Const C_IVCOST_Ref 			= 7						'매입단가 
	Const C_IVAMT_Ref	 		= 8						'매입금액 
	Const C_CURR_Ref			= 9						'화폐 
	Const C_IVDT_Ref			= 10					'매입일 
	Const C_IVNo_Ref			= 11					'매입번호	
	Const C_IVSEQ_Ref			= 12					'매입순번 
	Const C_PONo_Ref			= 13					'발주번호 
	const C_POSEQ_Ref 			= 14					'발주순번 
	Const C_Traking_Ref			= 15                    'TRAKING NO
	Const C_Tot_Qty_Ref			= 16                    '총수량 
	Const C_Doc_Ref 			= 17                    'Doc Amt
	Const C_Loc_Ref 			= 18                    'Loc Amt
	Const C_SpplSpec_Ref 		= 19					'품목규격 
	Const C_SlCd_Ref 			= 20                    '창고 
	Const C_SlNm_Ref 			= 21                    '창고명 
	Const C_VAT_TYPE_REF		= 22
	Const C_VAT_RT_REF			= 23
	Const C_VAT_DOC_AMT_REF		= 24
	Const C_VAT_LOC_AMT_REF		= 25
	Const C_VAT_IO_FLG_REF		= 26
	Const C_VAT_INC_FLAG_REF	= 27
	Const C_Mvmt_No_REF			= 28
	
	' === 2005.07.14 Lot No. Lot Sub No. 추가 =============
	Const C_Lot_No_REF			= 29
	Const C_Lot_Sub_No_REF		= 30
	' === 2005.07.14 Lot No. Lot Sub No. 추가 =============

	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	strMessage = ""
	IntIflg=true
	
	with frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		intStartRow = .vspdData.MaxRows + 1
		
		.vspdData.Redraw = False
	
		TempRow = .vspdData.MaxRows					'리스트 max값 
	
	for index1 = 0 to Count1
	
		.vspdData.Row=Index1+1
		
		If TempRow <> 0 Then
			for Index3=1 to TempRow
				if GetSpreadText(.vspdData,C_PrNo,index3,"X","X") = strRet(index1,C_ReqNo_Ref) then
					strMessage = strMessage & strRet(Index1,C_ReqNo_Ref) & ";"
					intIflg=False
					Exit for
				End if
			Next
		End If
		
		if IntIflg <> False then
			.vspdData.MaxRows = CLng(TempRow) + CLng(index1) + 1
			iInsRow = CLng(TempRow) + CLng(index1) + 1

			Call .vspdData.SetText(0		,	iInsRow, ggoSpread.InsertFlag)
			'Call .vspdData.SetText(C_DlvyDT,	iInsRow, .hdnDlvydt.value)
			Call .vspdData.SetText(C_DlvyDT,	iInsRow, EndDate)
			Call .vspdData.SetText(C_VatType,	iInsRow, .hdnVATType.value)
	 
			if Trim(.hdnVATINCFLG.value) = "2" then	'포함 
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,0,"X","X")			
				Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,0,"X","X")			
			else
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,1,"X","X")			
				Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,1,"X","X")			
			end if
	
			if .hdnVATType.value <> "" then
				call SetVatType(iInsRow)
			end if

			Call .vspdData.SetText(C_TrackingNo,	iInsRow, "*")
			' === 2005.07.14 Lot No. Lot Sub No. 추가 =============
'			Call .vspdData.SetText(C_Lot_No,	iInsRow, "*")
			Call .vspdData.SetText(C_Lot_No,	iInsRow, strRet(index1,C_Lot_No_REF))
			Call .vspdData.SetText(C_Lot_Seq,	iInsRow, strRet(index1,C_Lot_Sub_No_REF))
			' === 2005.07.14 Lot No. Lot Sub No. 추가 =============
			
			
			Call SetSpreadValue(.vspdData,C_CostCon		,iInsRow,1,"X","X")			
			Call SetSpreadValue(.vspdData,C_CostConCd	,iInsRow,1,"X","X")			
			
			Call SetState("C",iInsRow)
			
			Call .vspdData.SetText(C_PlantCd,	iInsRow, strRet(index1,C_PlantCd_Ref))

			Call .vspdData.SetText(C_PlantNm,	iInsRow, strRet(index1,C_PlantNm_Ref))
			Call .vspdData.SetText(C_itemCd,	iInsRow, strRet(index1,C_ItemCd_Ref))
			Call .vspdData.SetText(C_itemNm,	iInsRow, strRet(index1,C_ItemNm_Ref))
			Call .vspdData.SetText(C_SpplSpec,	iInsRow, strRet(index1,C_SpplSpec_Ref))
			Call .vspdData.SetText(C_OrderQty,	iInsRow, strRet(index1,C_Qty_Ref))
			Call .vspdData.SetText(C_OrderUnit,	iInsRow, strRet(index1,C_Unit_Ref))
			Call .vspdData.SetText(C_Cost	,	iInsRow, strRet(index1,C_IVCOST_Ref))
			Call .vspdData.SetText(C_OrderAmt,	iInsRow, strRet(index1,C_IVAMT_Ref))
			Call .vspdData.SetText(C_OrgOrderAmt,	iInsRow, strRet(index1,C_IVAMT_Ref))
			Call .vspdData.SetText(C_IVNO,	iInsRow, strRet(index1,C_IVNo_Ref))
			Call .vspdData.SetText(C_IVSEQ,	iInsRow, strRet(index1,C_IVSEQ_Ref))
			Call .vspdData.SetText(C_PoNo,	iInsRow, strRet(index1,C_PONo_Ref))
			Call .vspdData.SetText(C_PoSeqNo,	iInsRow, strRet(index1,C_POSEQ_Ref))
			Call .vspdData.SetText(C_TrackingNo,	iInsRow, strRet(index1,C_Traking_Ref))

			If strRet(index1,C_Tot_Qty_Ref) <> "" Then
						
				 temp = UNICDbl(GetSpreadText(.vspdData,C_OrderQty,iInsRow,"X","X")) - UNICDbl(strRet(index1,C_Tot_Qty_Ref))
				Call .vspdData.SetText(C_OrderQty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			    		
				temp = UNICDbl(GetSpreadText(.vspdData,C_Bal_Qty,iInsRow,"X","X")) - UNICDbl(strRet(index1,C_Tot_Qty_Ref))					
				Call .vspdData.SetText(C_Bal_Qty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			End If

			Call .vspdData.SetText(C_Bal_Doc_Amt,	iInsRow, strRet(index1,C_Doc_Ref))
			Call .vspdData.SetText(C_Bal_Loc_Amt,	iInsRow, strRet(index1,C_Loc_Ref))
			Call .vspdData.SetText(C_SLCd,	iInsRow, strRet(index1,C_SLCd_Ref))
			Call .vspdData.SetText(C_SLNm,	iInsRow, strRet(index1,C_SLNm_Ref))
			Call .vspdData.SetText(C_VatType,	iInsRow, strRet(index1,C_VAT_TYPE_REF))
			Call .vspdData.SetText(C_VatRate,	iInsRow, strRet(index1,C_VAT_RT_REF))
			Call .vspdData.SetText(C_IOFlgCd,	iInsRow, strRet(index1,C_VAT_INC_FLAG_REF))

			If Trim(GetSpreadText(.vspdData,C_IOFlgCd,iInsRow,"X","X")) = "2" Then
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,0,"X","X")			
			Else
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,1,"X","X")			
			End If

			Call .vspdData.SetText(C_MvmtNo,	iInsRow, strRet(index1,C_Mvmt_No_REF))
		Else
			IntIFlg=True
		End if 
	next
	
	intEndRow = .vspdData.MaxRows
	
	Call SetSpreadColorRef(intStartRow,intEndRow)
	
	if strMessage<>"" then
		Call DisplayMsgBox("17a005", "X",strmessage,"구매요청번호")
		.vspdData.ReDraw = True
		Exit Function
	End if
	
	.vspdData.ReDraw = True
	
	End with
	
    for index1 = intStartRow to intEndRow
		Call vspdData_Change2(C_Cost, index1)
	next 
End Function

'------------------------------------------  OpenMvmtRef()  -------------------------------------------------
'	Name : OpenMvmtRef()
'	Description :구매입고참조 
'---------------------------------------------------------------------------------------------------------
Function OpenMvmtRef()
	Dim strRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_CMODE then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtPoNo.focus	
		Exit Function
	End if 
	
	if frm1.txtRelease.Value = "Y" then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		frm1.txtPoNo.focus	
		Exit Function
	End if

	'if Not(UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnIVFlg.Value) = "N" and UCase(frm1.hdnRcptflg.Value) = "N") then 
	if Not(UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "N") then 
		Call DisplayMsgBox("17A012", "X","반품형태" & frm1.txtPotypeCd.Value & "(" & frm1.txtPoTypeNm.value & ")","구매입고참조" )
		frm1.txtPoNo.focus	
		Exit Function
	End if
		
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnMvmtType.value)
	arrParam(1) = Trim(frm1.txtSupplierCd.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = "PO"
	arrParam(4) = ""
	arrParam(5) = Trim(frm1.hdnRefPoNo.value)
	arrParam(6) = Trim(frm1.txtCurr.value)
	arrParam(7) = Trim(frm1.hdnSubcontraflg.value)
	
	iCalledAspName = AskPRAspName("M4111RA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4111RA2", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		frm1.txtPoNo.focus	
		Exit Function
	Else
		Call SetMvmtRef(strRet)
	End If	
		
End Function

Function SetMvmtRef(strRet)
	Dim Index1,index2,Index3,Count1,Count2
	Dim IntIflg
	Dim strMessage
	Dim intstartRow,intEndRow,temp, TempRow
	Dim iInsRow
	
	Const C_PlantCd_ref		= 0
	Const C_PlantNm_ref		= 1
	Const C_ItemCd_ref		= 2
	Const C_ItemNm_ref		= 3
	Const C_Spec_ref        = 4 
	Const C_PoQty_ref		= 5
	Const C_MvmtQty_ref		= 6			'입고수량 
	Const C_RetOrdQty_ref	= 7 		'반품오더수량(위치변경)
	Const C_TotRetQty_ref	= 8			'반품수량(위치변경)
	Const C_Unit_ref		= 9 
	Const C_MvmtDt_ref		= 10		'입고일 
	Const C_GrNo_ref		= 11
	Const C_GmNo_ref		= 12
	Const C_GmSeqNo_ref		= 13
	Const C_PoNo_ref		= 14
	Const C_PoSeq_ref		= 15
	Const C_TrackingNo_ref  = 16	
	Const C_Lot_No_ref      = 17
	Const C_Lot_Seq_ref	    = 18
	Const C_MvmtPrc_ref		= 19
	Const C_MvmtAmt_ref		= 20
	Const C_IvQty_ref		= 21
	Const C_SlCd_ref		= 22
	Const C_SlNm_ref		= 23
	Const C_Trackingflg_ref = 24
	Const C_MvmtNo_ref		= 25


	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	
	strMessage = ""
	IntIflg=true
	
	with frm1

		.vspdData.focus
		ggoSpread.Source = .vspdData
		intStartRow = .vspdData.MaxRows + 1
		.vspdData.Redraw = False
	
		TempRow = .vspdData.MaxRows					'리스트 max값 
		
	for index1 = 0 to Count1
	
		.vspdData.Row=Index1+1
		
		If TempRow <> 0 Then
			for Index3=1 to TempRow
				if GetSpreadText(.vspdData,C_MvmtNo,index3,"X","X") = strRet(index1,C_MvmtNo_ref) then
					strMessage = strMessage & strRet(Index1,C_MvmtNo_ref) & ";"
					intIflg=False
					Exit for
				End if
			Next
		End If
		
		if IntIflg <> False then
			.vspdData.MaxRows = CLng(TempRow) + CLng(index1) + 1
			iInsRow = CLng(TempRow) + CLng(index1) + 1
			
			Call .vspdData.SetText(0		,	iInsRow, ggoSpread.InsertFlag)
			'Call .vspdData.SetText(C_DlvyDT,	iInsRow, .hdnDlvydt.value)
			Call .vspdData.SetText(C_VatType,	iInsRow, .hdnVATType.value)
			Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,1,"X","X")
			Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,1,"X","X")
			'Call .vspdData.SetText(C_DlvyDT,	iInsRow, "")
			Call .vspdData.SetText(C_DlvyDT,	iInsRow, EndDate)
	
			if .hdnVATType.value <> "" then
				call SetVatType(iInsRow)
			end if

			Call .vspdData.SetText(C_TrackingNo,	iInsRow, "*")
			Call .vspdData.SetText(C_Lot_No,	iInsRow, "*")
			Call SetSpreadValue(.vspdData,C_CostCon	,iInsRow,1,"X","X")
			Call SetSpreadValue(.vspdData,C_CostConCd	,iInsRow,1,"X","X")
			
			Call SetState("C",iInsRow)
					
			Call .vspdData.SetText(C_PlantCd,	iInsRow, strRet(index1,C_PlantCd_Ref))
			Call .vspdData.SetText(C_PlantNm,	iInsRow, strRet(index1,C_PlantNm_Ref))
			Call .vspdData.SetText(C_itemCd,	iInsRow, strRet(index1,C_ItemCd_Ref))
			Call .vspdData.SetText(C_itemNm,	iInsRow, strRet(index1,C_ItemNm_Ref))
			Call .vspdData.SetText(C_SpplSpec,	iInsRow, strRet(index1,C_Spec_ref))
			Call .vspdData.SetText(C_OrderQty,	iInsRow, strRet(index1,C_MvmtQty_ref))
			Call .vspdData.SetText(C_TrackingNo,	iInsRow, strRet(index1,C_TrackingNo_ref))
			Call .vspdData.SetText(C_Lot_No,	iInsRow, strRet(index1,C_Lot_No_ref))
			Call .vspdData.SetText(C_Lot_Seq,	iInsRow, strRet(index1,C_Lot_Seq_ref))
			Call .vspdData.SetText(C_Cost,	iInsRow, strRet(index1,C_MvmtPrc_ref))
			Call .vspdData.SetText(C_OrderAmt,	iInsRow, strRet(index1,C_MvmtAmt_ref))
			Call .vspdData.SetText(C_OrgOrderAmt,	iInsRow, strRet(index1,C_MvmtAmt_ref))
			Call .vspdData.SetText(C_OrderUnit,	iInsRow, strRet(index1,C_Unit_Ref))
			Call .vspdData.SetText(C_SlCd,	iInsRow, strRet(index1,C_SlCd_Ref))
			Call .vspdData.SetText(C_SlNm,	iInsRow, strRet(index1,C_SlNm_ref))
			Call .vspdData.SetText(C_MvmtNo,	iInsRow, strRet(index1,C_MvmtNo_ref))
			Call .vspdData.SetText(C_PoNo,	iInsRow, strRet(index1,C_PoNo_Ref))
			Call .vspdData.SetText(C_PoSeqNo,	iInsRow, strRet(index1,C_PoSeq_Ref))

			' 입고수량 - 반품오더수량 
			IF strRet(index1,C_RetOrdQty_ref) <> "" Then
				temp= UNICDbl(strRet(index1,C_MvmtQty_ref)) - UNICDbl(strRet(index1,C_RetOrdQty_ref))
				Call .vspdData.SetText(C_OrderQty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			End If
			
			IF strRet(index1,C_IvQty_ref) <> "" Then
				temp = UNICDbl(GetSpreadText(.vspdData,C_OrderQty,iInsRow,"X","X")) - UNICDbl(strRet(index1,C_IvQty_ref))
				Call .vspdData.SetText(C_OrderQty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			End If
		Else
			IntIFlg=True
		End If 
	
	Next
	
	intEndRow = .vspdData.MaxRows
	
	Call SetSpreadColorRef(intStartRow,intEndRow)
	
	If strMessage<>"" then
		Call DisplayMsgBox("17a005", "X",strmessage,"구매입고번호")
		.vspdData.ReDraw = True
		Exit Function
	End If
	
'	.vspdData.Col 	= C_Stateflg
'	.vspdData.Text = "C"
	
	.vspdData.ReDraw = True
	
	End with

	for index1 = intStartRow to intEndRow
		Call vspdData_Change2(C_Cost, index1)
	next 
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : item(item_by_Plant) PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	frm1.vspdData.Col = C_PlantCd	
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 	 
	
	if  Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Exit Function
	End if

	IsOpenPop = True

	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(0) = Trim(frm1.vspdData.Text)
	
	frm1.vspdData.Col=C_ItemCd
	arrParam(1) = Trim(frm1.vspdData.Text)
	
	if frm1.hdnSubcontraflg.Value <> "Y" then
		arrParam(2) = "36!PP"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
		arrParam(3) = "30!P"
	else
		arrParam(2) = "12!MO"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
		arrParam(3) = "20!M"
	end if
	
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명	
    
    iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_ItemCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_ItemNm
			.vspdData.Text = arrRet(1)
		End With
		Call ChangeReturnCost()
	End If	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"
	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(2) = Trim(frm1.vspdData.text)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else	
		frm1.vspdData.Col=C_ItemCd
		frm1.vspdData.Row=frm1.vspdData.ActiveRow 
		frm1.vspdData.text=""
		
		frm1.vspdData.Col=C_ItemNM
		frm1.vspdData.Row=frm1.vspdData.ActiveRow 
		frm1.vspdData.text=""
		With frm1
			.vspdData.Col = C_PlantCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_PlantNm
			.vspdData.Text = arrRet(1)
		End With
		Call ChangeReturnCost()
	End If	
End Function

'------------------------------------------  OpenHS()  -------------------------------------------------
'	Name : OpenHS()
'	Description : OpenHS PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenHS()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "HS부호"	
	arrParam(1) = "B_HS_code"
	frm1.vspdData.Col=C_HSCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(2) = Trim(frm1.vspdData.text)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "HS부호"			
	
    arrField(0) = "HS_CD"	
    arrField(1) = "HS_NM"	
    
    arrHeader(0) = "HS부호"		
    arrHeader(1) = "HS명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_HSCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_HSNm
			.vspdData.Text = arrRet(1)
		End With
		Call vspdData_Change2(C_HsCd, frm1.vspdData.ActiveRow)
	End If	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위"					
	arrParam(1) = "B_Unit_OF_MEASURE"		
	
	frm1.vspdData.Col=C_OrderUnit
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(2) = Trim(frm1.vspdData.text)	
	
	arrParam(4) = ""						
	arrParam(5) = "단위"					
	
    arrField(0) = "Unit"					
    arrField(1) = "Unit_Nm"					
    
    arrHeader(0) = "단위"				
    arrHeader(1) = "단위명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.text= arrRet(0)		
		Call ChangeReturnCost()
	End If	
End Function

'------------------------------------------  OpenSL()  -------------------------------------------------
'	Name : OpenSL()
'	Description : Storage_Location PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	
	if Trim(frm1.vspdData.Text)="" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Exit function	
	End if 
	
	arrParam(4) = "PLANT_CD= " & FilterVar(Trim(frm1.vspdData.Text), " " , "S") & " "

	IsOpenPop = True

	frm1.vspdData.Col=C_SLCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "창고"					
	arrParam(1) = "B_STORAGE_LOCATION"		
	
	arrParam(2) = Trim(frm1.vspdData.Text)	
	arrParam(5) = "창고"					
	
    arrField(0) = "SL_CD"					
    arrField(1) = "SL_NM"					
    
    arrHeader(0) = "창고"				
    arrHeader(1) = "창고명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_SLCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_SLNm
			.vspdData.Text = arrRet(1)
		End With
		Call vspdData_Change2(C_SLCd, frm1.vspdData.ActiveRow)
	End If	
End Function

'------------------------------------------  OpenVat()  -------------------------------------------------
'	Name : OpenVat()
'	Description : 
'-------------------------------------------------------------------------------------------------------
Function OpenVat()
	Dim price, chk_vat_flg
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function 

	IsOpenPop = True
 
    frm1.vspdData.Col=C_VatType
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 

	arrParam(0) = "VAT형태"				
	arrParam(1) = "B_MINOR,b_configuration"	
	
	arrParam(2) = Trim(frm1.vspdData.Text)		
		
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd "	
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
	arrParam(5) = "VAT형태"					
	
    arrField(0) = "b_minor.MINOR_CD"			
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"	
    
    arrHeader(0) = "VAT형태"					
    arrHeader(1) = "VAT형태명"				
    arrHeader(2) = "VAT율"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
	    With frm1
			.vspdData.Col = C_VatType
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_VatNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_VatRate
			.vspdData.Text = arrRet(2)
			
			.vspdData.Col = C_OrderAmt
			price = UNICDbl(.vspdData.Text)
			'	vat 금액계산 
			' 부가세 포함/불포함 부가세 계산 변경 2002.3.9 L.I.P
			.vspdData.Col = C_IOFlgCd
			chk_vat_flg	= .vspdData.text
			
			.vspdData.Col = C_VatAmt 
			if chk_vat_flg = "2"		Then
				.vspdData.Text = UNIConvNumPCToCompanyByCurrency((price * (arrRet(2)/(100 + arrRet(2)))),frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
			Else
				.vspdData.Text = UNIConvNumPCToCompanyByCurrency((price * (arrRet(2)/100)),frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
			End If		
		End With
	    Call vspdData_Change2(C_VatType, frm1.vspdData.ActiveRow)   
	End If	
End Function

'------------------------------------------  OpenRet()  -------------------------------------------------
'	Name : OpenRet()
'	Description : 
'-------------------------------------------------------------------------------------------------------
Function OpenRet()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function 

	IsOpenPop = True
 
    frm1.vspdData.Col=C_RetCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 

	arrParam(0) = "반품유형"				
	arrParam(1) = "B_MINOR"	
	
	arrParam(2) = Trim(frm1.vspdData.Text)		
		
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9017", "''", "S") & " "	
	arrParam(5) = "반품유형"					
	
    arrField(0) = "MINOR_CD"			
    arrField(1) = "MINOR_NM"
    
    
    arrHeader(0) = "반품유형"					
    arrHeader(1) = "반품유형명"				

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_RetCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_RetNm
			.vspdData.Text = arrRet(1)
	End With
	Call vspdData_Change2(C_RetCd, frm1.vspdData.ActiveRow)   
	End If	
End Function

'------------------------------------------  OpenTrackingNo()  -------------------------------------------
'	Name : OpenTrackingNo()
'	Description : TrackingNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = ""				
	arrParam(1) = ""					

	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	If Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		IsOpenPop = False
		Exit Function
	End if
	
	arrParam(2) = Trim(frm1.vspdData.Text)	
	arrParam(3) = ""
	
	frm1.vspdData.Col=C_SoNo
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	arrParam(4) = ""
	arrParam(5) = " and A.tracking_no not in (" & FilterVar("*", "''", "S") & " ) " 
	arrParam(6) = "M" 

	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_TrackingNo
			.vspdData.Text = arrRet
		End With
		Call vspdData_Change2(C_TrackingNo, frm1.vspdData.ActiveRow)
	End If	
End Function

'------------------------------------------  OpenLotNo()  -------------------------------------------------
'	Name : OpenLotNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenLotNo()
		Dim arrRet
	Dim arrParam(5),arrField(6)
	Dim IntRetCD
	Dim iCalledAspName

	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	
	If IsOpenPop = True Then Exit Function
	
	
	IsOpenPop = True
	
	frm1.vspdData.Col=C_PlantCd
	arrParam(0) = frm1.vspdData.Text
	
	frm1.vspdData.Col= C_PlantNm
	arrParam(1) = frm1.vspdData.Text
	
	If arrParam(0) = "" Then	
		Call DisplayMsgBox("169901","X", "X", "X")   <% '공장정보가 필요합니다 %>
		Exit Function
	End If		
		
    frm1.vspdData.Col= C_SLCd
	arrParam(2) = frm1.vspdData.Text
		
	frm1.vspdData.Col= C_SLNm
	arrParam(3) = frm1.vspdData.Text
	
	frm1.vspdData.Col= C_itemCd
	arrParam(4) = frm1.vspdData.Text
	
	If arrParam(4) = "" Then
		Call DisplayMsgBox("169915","X", "X", "X")   <% '품목코드를 입력하십시오 %>
		Exit Function
	End If		
	
	arrParam(5) = ""
	
	arrField(0) = 1 'ITEM_CD					' Field명(0)
	arrField(1) = 2 'ITEM_NM					' Field명(1)
	arrField(2) = 3	'SPECIFICATION	
	arrField(3) = 4
	arrField(4) = 5
	arrField(5)	= 6
	arrField(6) = 7
	
	iCalledAspName = AskPRAspName("I1211PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "I1211PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspddata
			Call .SetText(C_Lot_No, .ActiveRow, arrRet(5))
			Call .SetText(C_Lot_Seq, .ActiveRow, arrRet(6))
		End With
		Call vspdData_Change2(C_Lot_No, frm1.vspdData.ActiveRow)
	End If	
End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 
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
        Case 6                                                              '과부족허용율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
    End Select
End Sub

'============================================  2.5.1 TotalSum()  ======================================
'=	Name : TotalSum()																					=
'=	Description : Master L/C Header 화면으로부터 넘겨받은 parameter setting(Cookie 사용)				=
'========================================================================================================
Sub TotalSum(ByVal row)
	
    Dim SumTotal, lRow, tmpGrossAmt, tmpVatAmt,tmpamt
	Dim chk_vat_flg,tmprate
	SumTotal = 0
	ggoSpread.source = frm1.vspdData
	SumTotal = UNICDbl(frm1.txtGrossAmt.value)
	frm1.vspdData.Row = row

	frm1.vspdData.Col = C_VatRate		 
	tmprate	= UNICDbl(frm1.vspdData.text)

	frm1.vspdData.Col = C_IOFlgCd
	chk_vat_flg	= UNICDbl(frm1.vspdData.text)

	frm1.vspdData.Col = C_VatAmt							
	tmpVatAmt = UNICDbl(frm1.vspdData.Text)		
	frm1.vspdData.Col = C_OrderAmt				
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)

		if chk_vat_flg = "2" Then	
			frm1.vspdData.Col = C_NetOrderAmt		
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency((tmpGrossAmt - tmpVatAmt), frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
			'frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(tmpGrossAmt - tmpVatAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")				
		Else
			frm1.vspdData.Col = C_NetOrderAmt		
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(tmpGrossAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
		End If
			
		dim tmpNetGrossAmt
		frm1.vspdData.Col = C_NetOrderAmt
		tmpNetGrossAmt = UNICDbl(frm1.vspdData.Text)			
		    
	    frm1.vspdData.Col = C_OrgNetOrderAmt
		SumTotal = SumTotal + (tmpNetGrossAmt-UNICDbl(frm1.vspdData.Text))			
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(tmpNetGrossAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
			
        frm1.txtGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	
End Sub

Sub TotalSum_Copy(ByVal row)
	Dim SumTotal, lRow, tmpGrossAmt, tmpVatAmt,tmpamt
	Dim chk_vat_flg,tmprate
	SumTotal = 0
	ggoSpread.source = frm1.vspdData
		
	SumTotal = UNICDbl(frm1.txtGrossAmt.value)
		
	frm1.vspdData.Row = row

	frm1.vspdData.Col = C_VatRate		 
	tmprate	= UNICDbl(frm1.vspdData.text)

	frm1.vspdData.Col = C_IOFlgCd
	chk_vat_flg	= UNICDbl(frm1.vspdData.text)
		
	frm1.vspdData.Col = C_VatAmt							
	tmpVatAmt = UNICDbl(frm1.vspdData.Text)		
		
	frm1.vspdData.Col = C_OrderAmt				
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)

	
	frm1.vspdData.Col = C_OrgOrderAmt	

	if chk_vat_flg = "2" Then	
		SumTotal = SumTotal + (tmpGrossAmt - tmpVatAmt)
	Else
		SumTotal = SumTotal + tmpGrossAmt '//+ tmpVatAmt
	End If

    frm1.txtGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	
End Sub	

'==========================================================================================
'   Event Name : ChangeReturnCost
'   Event Desc :
'==========================================================================================
Sub ChangeReturnCost()
		
	Dim IntCol, IntRow
	Dim strssTemp1,strssTemp2,strssTemp3
	
	intCol = frm1.vspdData.ActiveCol - 1
	intRow = frm1.vspdData.ActiveRow
	
		if IntCol = C_itemCd or IntCol = C_PlantCd or IntCol = C_OrderUnit then
			
			frm1.vspdData.Col = C_ItemCd
			strssTemp1 = Trim(frm1.vspdData.Text)
			frm1.vspdData.Col = C_PlantCd
			strssTemp2 = Trim(frm1.vspdData.Text)
			frm1.vspdData.Col = C_OrderUnit
			strssTemp3 = Trim(frm1.vspdData.Text)
			
			if strssTemp1 = "" or strssTemp2 = ""  then'or strssTemp3 = "" then
				Exit Sub
			End if
			
			if intCol = C_OrderUnit then
				Call ChangeItemPlantForUnit(IntRow,IntRow)
			else
				Call ChangeItemPlant(IntRow,IntRow)
			end if
			
		End if
		
End Sub
	
'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtGrossAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub



'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_Cost,-1, .txtCurr.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_OrderAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_OrgOrderAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_NetOrderAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
		'VAT금액 
		ggoSpread.SSSetFloatByCellOfCur C_VatAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"

	End With

End Sub	
<%
'================================== =====================================================
' Function Name : InitCollectType
' Function Desc : 소비세유형코드/명/율 저장하기 
' 여기부터 키보드에서 소비세유형코드를 변경시 소비세유형명,소비세율,매입금액,NetAmount를 변경시키는 함수 
'========================================================================================
%>
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, vbInformation, parent.gLogoName
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub

'========================================================================================
' Function Name : GetCollectTypeRef
' Function Desc : 
'========================================================================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)  
		If arrCollectVatType(iCnt, 0) = UCase(VatType) Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = arrCollectVatType(iCnt, 2)
			Exit Sub
		End If
	Next
	VatTypeNm = ""
	VatRate = ""
End Sub

'========================================================================================
' Function Name : SetVatType
' Function Desc : 
'========================================================================================
Sub SetVatType(byval inti)
	Dim VatType, VatTypeNm, VatRate 
	Dim txtVatRate ,txtVatAmt, chk_vat_flg,txtOrderAmt
	     
with frm1.vspdData
      
       .Row = inti
	   .Col = C_VatType
	  
	  VatType = .text
	
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)
  
       .Col = C_VatNm  
       .text = VatTypeNm
      
       .Col = C_VatRate
	   .text = UNIFormatNumber(UNICDbl(VatRate),ggExchRate.DecPoint, -2, 0, Parent.ggExchRate.RndPolicy, Parent.ggExchRate.RndUnit)
	   txtVatRate =  UNICDbl(.text)


	   '	vat 금액계산  
	   ' 부가세 포함/불포함 부가세 계산 변경 2002.3.9 L.I.P
		.Col		= C_IOFlgCd
		chk_vat_flg	= .text
		
       .Col          = C_OrderAmt
		if chk_vat_flg = "2"	Then	
			txtVatAmt    = UNICDbl(.text) * (txtVatRate/(100 + txtVatRate))
		Else
			txtVatAmt    = UNICDbl(.text) * (txtVatRate/100)
		End If

		.Col = C_VatAmt 
		.Text = UNIConvNumPCToCompanyByCurrency(txtVatAmt,frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
		
		.Col = C_OrderAmt
		txtOrderAmt = UNICDbl(.Text)
		
		.Col = C_NetOrderAmt
		if chk_vat_flg = "2"	Then	
			.Text = UNIFormatNumber(txtOrderAmt - txtVatAmt, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		Else
			.Text = UNIFormatNumber(txtOrderAmt, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		End If
	    
	    if not (UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnIVFlg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "N") then 
		    call FncVatToZero(.ActiveRow)
		end if 	
 
 End With
	   
End Sub
'========================================================================================
' Function Name : SetRetCd
' Function Desc : 반납유형 직접 입력시 처리 
'					2002-02-22 추가 L.I.P
'========================================================================================
Sub SetRetCd()
	Dim iRetCd, iRetNm, strQUERY, tmpData
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i

	with frm1.vspdData

		Err.Clear
    
	   .Col = C_RetCd

		strQUERY = " Minor.MAJOR_CD=" & FilterVar("B9017", "''", "S") & " and  Minor.MINOR_CD =  " & FilterVar(Trim( .text), " " , "S") & "  "
    
		Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM ", " B_MINOR Minor ", strQUERY, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Err.number = 0 Then
			
			if lgF0 <> "" then
				iRetNm = Split(lgF1, Chr(11))
			   .Col = C_RetNm  
			   .text = iRetNm(0)
			  else
			   .Col = C_RetNm  
			   .text = ""
			end if
		else
			MsgBox Err.description, VbInformation, parent.gLogoName
			Err.Clear 
			Exit Sub
		End If
     
	End With
	   
End Sub
'========================================================================================
' Function Name : setCVatFlg
' Function Desc : 부가세 포함에 따른 의제매입계산 처리 
' Append		: 2002-03-09  L.I.P
'========================================================================================
Sub setCVatFlg(byval iRow)
	Call setVatType(iRow)
End Sub

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	
	Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")   
    
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                    
    Call SetDefaultVal
    Call InitVariables                      
    Call CookiePage(0)
    
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
       
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		If frm1.txtRelease.Value <> "Y" Then
			Call SetPopupMenuItemInf("1101111111")
		Else
			Call SetPopupMenuItemInf("0000111111")
		End If
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If

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
    Call CurFormatNumSprSheet() 
    Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadLockAfterQuery
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 행복사시 발주번호, 입고번호, 매입번호 값을 넘기지 않음 
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	With frm1.vspdData
		Select Case Col
			Case C_MvmtNo
				.Col = C_MvmtNo
				.Text = ""
			Case C_PoNo
				.Col = C_PoNo
				.Text = ""
			Case C_IVNO
				.Col = C_IVNO
				.Text = ""
			Case Else
				Call vspdData_Change2(Col , Row)
		End Select		
	End With
End Sub
'==========================================================================================
'   Event Name : vspdData_Change2
'   Event Desc : 기존 vspdData_Change -> vspdData_Change2 함수로 수정 
'				 행복사시 발주번호, 입고번호, 매입번호 값을 넘기지 않기 위하여.
'==========================================================================================
Sub vspdData_Change2(ByVal Col , ByVal Row )

	Dim strsstemp1,strsstemp2,strsstemp3
	Dim Qty, Price, DocAmt, VatAmt, VatRate, chk_vat_flg
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
  
    with frm1.vspdData 
		
		.Row = Row
		.Col = 0
		
		if Trim(.Text) = ggoSpread.DeleteFlag  then
		    Exit Sub
		end if 
		
		.Col = C_Stateflg:	.Row = Row
		if Trim(.Text) = "" then
			.Text = "U"
		End if

		if Col = C_itemCd then 
			.Col = C_ItemCd
			strssTemp1 = Trim(.Text)
			.Col = C_PlantCd
			strssTemp2 = Trim(.Text)
			
			if strssTemp1 = "" or strssTemp2 = "" then
				Exit Sub
			End if
			
			Call ChangeItemPlant(Row,Row)
		elseif Col = C_PlantCd then		
			frm1.vspdData.Col=C_ItemCd
			frm1.vspdData.Row=frm1.vspdData.ActiveRow 
			frm1.vspdData.text=""
			
			frm1.vspdData.Col=C_ItemNM
			frm1.vspdData.Row=frm1.vspdData.ActiveRow 
			frm1.vspdData.text=""

			'.Col = C_ItemCd
			'strssTemp1 = Trim(.Text)
			'.Col = C_PlantCd
			'strssTemp2 = Trim(.Text)
			
			'if strssTemp1 = "" or strssTemp2 = "" then
			'	Exit Sub
			'End if
			
			'Call ChangeItemPlant(Row,Row)
		elseif Col = C_OrderUnit then
			
			.Col = C_ItemCd
			strssTemp1 = Trim(.Text)
			.Col = C_PlantCd
			strssTemp2 = Trim(.Text)
			.Col = C_OrderUnit
			strssTemp3 = Trim(.Text)
			
			if strssTemp1 = "" or strssTemp2 = "" or strssTemp3 = "" then
				Exit Sub
			End if
			
			Call ChangeItemPlantForUnit(Row,Row)
		
		elseif Col = C_OrderQty then
			'Call ChangeItemPlantForUnit(Row,Row)
		'jin
		elseif Col = C_VatType then 'or Col = C_VatAmt then
				Call SetVatType(frm1.vspdData.ActiveRow)     ' C_VatNm,C_VatRate 세팅 
				call vspdData_Change2(C_OrderQty ,Row )
				'Call TotalSum					'총품목금액합계 
		elseif Col = C_RetCd	then '반납유형 직접입력시 처리 부분 2002-02-22 L.I.P
				
				Call SetRetCd()
		End if
		
    End With
    
	Select Case col
	Case C_OrderQty,C_Cost',C_VatRate
	
		frm1.vspdData.Col = C_OrderQty
		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
			Qty = 0
		Else
			Qty = UNICDbl(frm1.vspdData.Text)
		End If
		
		frm1.vspdData.Col = C_Cost
		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
			Price = 0
		Else
			Price = UNICDbl(frm1.vspdData.Text)
		End If
		DocAmt = Qty * Price
		frm1.vspdData.Col = C_OrderAmt		
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(DocAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
        DocAmt = UNICDbl(frm1.vspdData.Text)
		'VAT 금액 추가  -->
		frm1.vspdData.Col = C_VatRate ' VAT 율 
		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
			VatRate = 0
		Else
			VatRate = UNICDbl(frm1.vspdData.Text)
		End If

		' 부가세 포함/불포함 부가세 계산 변경 2002.3.9 L.I.P
		frm1.vspdData.Col		= C_IOFlgCD
		chk_vat_flg	= frm1.vspdData.text
		if chk_vat_flg = "2"	Then	
			VatAmt    = DocAmt * (VatRate/(100 + VatRate))
		Else
			VatAmt    = DocAmt * (VatRate/100)
		End If
		'VatAmt = DocAmt * (VatRate /100) 'Vat 금액 계산 
		
		frm1.vspdData.Col = C_VatAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(VatAmt),frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
		
		'<-- VAT 금액 추가 2002.2.18 L.I.P
		Call TotalSum(row)					'총품목금액합계 
		
		frm1.vspdData.Col = 0
		'if  frm1.vspdData.Text = ggoSpread.InsertFlag  then
		    frm1.vspdData.Col = C_OrgOrderAmt		
		    frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(DocAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
		'end if	
	Case C_DlvyDt
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_DlvyDt
		strsstemp1 = frm1.vspdData.Text
		if strsstemp1 = "" then Exit Sub
		strsstemp2 = frm1.txtPoDt.text
		if UniConvDateToYYYYMMDD(strsstemp2,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(strsstemp1,Parent.gDateFormat,"") then
			Call DisplayMsgBox("970023", "X", "납기일", frm1.txtPoDt.Alt)
		end if
	Case C_IOFlg
		frm1.vspdData.Col = C_IOFlg
		
		call setCVatFlg(frm1.vspdData.ActiveRow)	
		call vspdData_Change2(C_OrderQty ,Row )
		'Call TotalSum					'총품목금액합계 
	end select
      
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 
	
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

Dim intIndex 

	With frm1.vspdData
	
		.Row = Row
		.Col = Col

		if Col = C_CostCon then 
				intIndex = .Value
				.Col = C_CostCon+1
				.Value = intIndex
		else  
		        intIndex = .Value
				.Col = C_IOFlg+1
				.Value = intIndex
        end if
 	
  End With
 
End Sub
'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 Then
        .Col = Col
        .Row = Row
        
		Select Case Col 
			
		Case C_Popup1
			Call OpenPlant()
		Case C_Popup2
			Call OpenItem()
		Case C_Popup3
			Call OpenUnit()
		Case C_Popup5
			Call OpenHS()
		Case C_Popup6
			Call OpenSL()
		Case C_TrackingNoPop
			Call OpenTrackingNo()
		case C_Popup7	
			Call OpenVat()
		case C_Popup8
		    Call OpenRet()	
        case C_Popup9
		    Call OpenLotNo()	
		End Select
        
    End If
    
    End With
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

    If NewRow = .MaxRows Then
        'DbQuery
    End if    

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
  
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then	
		If CheckRunningBizProcess = True Then
			Exit Sub
		End If	
			
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if 
    
End Sub	

'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim intIndex
    
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
    
    For intIndex = 1 to frm1.vspdData.MaxCols 
		frm1.vspdData.SetColItemData intindex,0
	Next
	
    frm1.vspdData.MaxRows = 0
    
    Call InitVariables
    										
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then		
       Exit Function
    End If
   
    '-----------------------
    'Query function call area
    '-----------------------
    frm1.txtQuerytype.value = "Query"
    If DbQuery = False Then Exit Function
       
    FncQuery = True	
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
   
    FncNew = False                  
    
    Err.Clear                       
    
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
    Call SetDefaultVal()
    Call InitVariables                     
    FncNew = True                          

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                      

    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then Exit Function
    														
    Err.Clear                              
    
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
    
    'On Error Resume Next                          
    '-----------------------
    'Precheck area
    '-----------------------
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    
	If Not chkField(Document, "2") OR Not ggoSpread.SSDefaultCheck Then		
		Exit Function
	End If
   
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
    FncSave = True                                                          
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    Dim SumTotal,tmpGrossAmt
    if frm1.vspdData.Maxrows < 1	then exit function
    ggoSpread.Source = frm1.vspdData	

    ggoSpread.CopyRow
    
    frm1.vspdData.ReDraw = False
    
    SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
    
    frm1.vspdData.ReDraw = True
    
    Call SetState("C",frm1.vspdData.ActiveRow)
    
    '복사한 것은 긴급반품로 간주.
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_SeqNo
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_PrNo
    frm1.vspdData.Text = ""
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_MvmtNo
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_SoNo
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_SoSeqNo
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_OrgOrderAmt1
    frm1.vspdData.Text = ""
    
    
    call TotalSum_copy(frm1.vspdData.ActiveRow)
    
    if UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnIVFlg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "N" then 
		ggoSpread.spreadUnlock C_IOFlg,frm1.vspdData.Row,C_IOFlg,frm1.vspdData.Row
	    ggoSpread.sssetrequired C_IOFlg,frm1.vspdData.Row,frm1.vspdData.Row
	    ggoSpread.spreadUnlock C_VatType,frm1.vspdData.Row,C_Popup7,frm1.vspdData.Row
	end if
    
    if UCase(frm1.hdnImportflg.value) = "Y" then
		ggoSpread.spreadUnlock C_HsCd,frm1.vspdData.Row,C_HsCd,.vspdData.Row
		ggoSpread.sssetrequired C_HsCd,frm1.vspdData.Row,.vspdData.Row
		ggoSpread.spreadUnlock C_Popup5,frm1.vspdData.Row,C_Popup5,.vspdData.Row
	End if
    
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_TrackingNo
	if Trim(frm1.vspdData.Text) = "*" then
		ggoSpread.spreadlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
	else
		ggoSpread.spreadUnlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
		ggoSpread.sssetrequired C_TrackingNo, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	end if
		
	frm1.vspdData.Col = C_Lot_No
	if Trim(frm1.vspdData.Text) = "*" then
		ggoSpread.spreadlock C_Lot_No, frm1.vspdData.ActiveRow, C_Lot_Seq, frm1.vspdData.ActiveRow
	else
		ggoSpread.spreadUnlock C_Lot_No, frm1.vspdData.ActiveRow, C_Lot_Seq, frm1.vspdData.ActiveRow
	end if	
    
    
    ggoSpread.spreadlock C_RetNm,frm1.vspdData.Row,C_RetNm,frm1.vspdData.Row
    
    frm1.vspdData.ReDraw = True
    
End Function
'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	Dim maxrow,maxrow1,SumTotal,tmpGrossAmt,index,index1,orgtmpGrossAmt
	Dim starindex ,endindex,delflag,chk_vat_flg,tmprate,tmpAmt
	if frm1.vspdData.Maxrows < 1	then exit function
	maxrow = frm1.vspdData.Maxrows
	index1 = 0
	frm1.vspdData.Row		= frm1.vspdData.ActiveRow
	frm1.vspdData.Col		= C_IOFlgCd


	chk_vat_flg	= frm1.vspdData.text	
	
	frm1.vspdData.Col		= C_VatRate		 
	tmprate	= UNICDbl(frm1.vspdData.text)	
	starindex = frm1.vspdData.SelBlockRow
	endindex  = frm1.vspdData.SelBlockRow2
    
    Redim orgtmpNetGrossAmt(endindex - starindex)
    Redim tmpNetGrossAmt(endindex - starindex)
    Redim delflag(endindex - starindex)
    SumTotal = UNICDbl(frm1.txtGrossAmt.value)
	
	for index = starindex to endindex
	    frm1.vspdData.Row = index
	    
	    frm1.vspdData.Col = C_NetOrderAmt
	    tmpNetGrossAmt(index1) = UNICDbl(frm1.vspdData.Text)
	    
	    frm1.vspdData.Col = C_OrgNetOrderAmt1
	    orgtmpNetGrossAmt(index1) = UNICDbl(frm1.vspdData.Text)
	    
	    frm1.vspdData.Col = 0
	    delflag(index1) = frm1.vspdData.Text
	    index1 = index1 + 1
	next

	ggoSpread.Source = frm1.vspdData
	index = frm1.vspdData.ActiveRow - starindex

	if delflag(index) = ggoSpread.UpdateFlag then		
		SumTotal = SumTotal + (orgtmpNetGrossAmt(index)-tmpNetGrossAmt(index))							
	elseif  delflag(index) = ggoSpread.DeleteFlag then							
		SumTotal = SumTotal + orgtmpNetGrossAmt(index)
	elseif delflag(index) = ggoSpread.InsertFlag  then			
		SumTotal = SumTotal - tmpNetGrossAmt(index)
	end if		
			
    frm1.txtGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")

    ggoSpread.EditUndo
    maxrow1 = frm1.vspdData.Maxrows    
    
    frm1.txtGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")     
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow_old() 
	Dim inti
	
	With frm1
	
    .vspdData.ReDraw = False
	.vspdData.focus
	
    ggoSpread.Source = .vspdData
	imRow = AskSpdSheetAddRowCount()
    If imRow = "" Then Exit Function

    ggoSpread.InsertRow ,imRow
    
    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
    
    ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
    ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
    ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
    ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd
    
    Call SetState("C",.vspdData.ActiveRow)
    
    '공장 기본값 추가 
	.vspdData.Col=C_PlantCd
    .vspdData.Text=Parent.gPlant
	
    '.vspdData.Col=C_OrderQty
    '.vspdData.Text="0"
    .vspdData.Col=C_OrderAmt
    .vspdData.Text="0"
    .vspdData.Col=C_Cost
    .vspdData.Text="0"
    .vspdData.Col = C_DlvyDT
    .vspdData.Text = .hdnDlvydt.value
    
    'Insert Row 시 헤더의 부가세관련 정보 초기값으로 2002.2.19
    .vspdData.Col = C_VatType
    .vspdData.Text = .hdnVATType.value
    
    '.vspdData.Col = C_IOFlg
    '.vspdData.value = .hdnVATINCFLG.value
    '.vspdData.value = 0
	'.vspdData.Col = C_IOFlgCd
    '.vspdData.value = .hdnVATINCFLG.value
	'.vspdData.value = 0
	
	'leeip
	if Trim(.hdnVATINCFLG.value) = "2" then	'포함 
    .vspdData.Col = C_IOFlg
    .vspdData.value = 0
	.vspdData.Col = C_IOFlgCd
   	.vspdData.value = 0
	else
	.vspdData.Col = C_IOFlg
    .vspdData.value = 1
	.vspdData.Col = C_IOFlgCd
    .vspdData.value = 1
	end if
	
	if .hdnVATType.value <> "" then
		call SetVatType(frm1.vspdData.ActiveRow)
    end if

    .vspdData.Col = C_VatRate
    .vspdData.Text = .hdnVATRate.value
    
    .vspdData.Col = C_TrackingNo
    .vspdData.Text = "*"
    .vspdData.Col = C_Lot_No
    .vspdData.Text = "*"
    '---------------------------------------------------------
          
	ggoSpread.spreadUnlock C_PlantCd,.vspdData.Row,C_PlantCd,.vspdData.Row
	ggoSpread.sssetrequired C_PlantCd,.vspdData.Row,.vspdData.Row
	ggoSpread.spreadUnlock C_Popup1,.vspdData.Row,C_Popup1,.vspdData.Row
	ggoSpread.spreadUnlock C_ItemCd,.vspdData.Row,C_ItemCd,.vspdData.Row
	ggoSpread.sssetrequired C_ItemCd,.vspdData.Row,.vspdData.Row
	ggoSpread.spreadUnlock C_Popup2,.vspdData.Row,C_Popup2,.vspdData.Row
	ggoSpread.spreadUnlock C_IOFlg,.vspdData.Row,C_IOFlg,.vspdData.Row
	ggoSpread.spreadlock C_RetNm,.vspdData.Row,C_IOFlg,.vspdData.Row
	ggoSpread.spreadlock C_NetOrderAmt.vspdData.Row,C_NetOrderAmt.vspdData.Row
	'ggoSpread.sssetrequired C_IOFlg,.vspdData.Row,.vspdData.Row
	ggoSpread.spreadLock	C_TrackingNo ,.vspdData.Row,C_TrackingNoPop,.vspdData.Row
	if .hdnImportflg.value = "Y" then
		ggoSpread.spreadUnlock C_HsCd,.vspdData.Row,C_HsCd,.vspdData.Row
		ggoSpread.sssetrequired C_HsCd,.vspdData.Row,.vspdData.Row
		ggoSpread.spreadUnlock C_Popup5,.vspdData.Row,C_Popup5,.vspdData.Row
	End if
		
	.vspdData.Col = C_CostCon
	.vspdData.Value = 1
	.vspdData.Col = C_CostConCd
	.vspdData.Value = 1
	'.vspdData.Col = C_IOFlg
	'.vspdData.Value = 1
	'.vspdData.Col = C_IOFlgCd
	'.vspdData.Value = 1
	.vspdData.ReDraw = True
    
    End With
        
End Function

Function FncInsertRow(ByVal pvRowCnt)

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		call FncInsertRow_Ref(pvRowCnt)
	end With
end Function 

Function FncInsertRow_Ref(ByVal pvRowCnt) 
	Dim imRow
	Dim inti
	DIm intStratRow
	Dim intEndRow

	
	With frm1
	
	.vspdData.focus
    ggoSpread.Source = .vspdData
    .vspdData.ReDraw = False
    
    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then Exit Function
	End If	

    ggoSpread.InsertRow ,imRow
    SetSpreadColor .vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
    
    ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
    ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
    ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
    ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd
    
    .vspdData.ReDraw = False
	
	intStratRow = .vspdData.ActiveRow
	intEndRow = .vspdData.ActiveRow + imRow - 1
	for inti=.vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
	 
		.vspdData.row = inti
		Call SetState("C",inti)

		'공장 기본값 추가 
		.vspdData.Col=C_PlantCd
		.vspdData.Text=Parent.gPlant
		.vspdData.Col=C_OrderAmt
		.vspdData.Text="0"
		.vspdData.Col=C_Cost
		.vspdData.Text="0"
		.vspdData.Col = C_DlvyDT
		'.vspdData.Text = .hdnDlvydt.value
		.vspdData.Text = EndDate		' 행추가시 현재일을 기본값으로 표시 
		
		'Insert Row 시 헤더의 부가세관련 정보 초기값으로 2002.2.19
		.vspdData.Col = C_VatType
		.vspdData.Text = .hdnVATType.value
	 
		'leeip
		if Trim(.hdnVATINCFLG.value) = "2" then	'포함 
			.vspdData.Col = C_IOFlg
			.vspdData.value = 0
			.vspdData.Col = C_IOFlgCd
			.vspdData.value = 0
		else
			.vspdData.Col = C_IOFlg
			.vspdData.value = 1
			.vspdData.Col = C_IOFlgCd
			.vspdData.value = 1
		end if
	
		if .hdnVATType.value <> "" then
			call SetVatType(inti)
		end if

		.vspdData.Col = C_TrackingNo
		.vspdData.Text = "*"
		.vspdData.Col = C_Lot_No
		.vspdData.Text = "*"
		.vspdData.Col = C_CostCon
		.vspdData.Value = 1
		.vspdData.Col = C_CostConCd
		.vspdData.Value = 1
	next

	ggoSpread.spreadUnlock	C_PlantCd	,intStratRow,C_PlantCd,intEndRow
	ggoSpread.sssetrequired	C_PlantCd	,intStratRow,intEndRow
	ggoSpread.spreadUnlock	C_Popup1	,intStratRow,C_Popup1,intEndRow
	ggoSpread.spreadUnlock	C_ItemCd	,intStratRow,C_ItemCd,intEndRow
	ggoSpread.sssetrequired	C_ItemCd	,intStratRow,intEndRow
	ggoSpread.spreadUnlock	C_Popup2	,intStratRow,C_Popup2,intEndRow
	ggoSpread.spreadUnlock	C_IOFlg		,intStratRow,C_IOFlg,intEndRow
	ggoSpread.spreadlock		C_RetNm		,intStratRow,C_IOFlg,intEndRow
	ggoSpread.spreadlock		C_NetOrderAmt,intStratRow,C_NetOrderAmt,intEndRow
	ggoSpread.spreadLock	C_TrackingNo ,intStratRow,C_TrackingNoPop,intEndRow
	ggoSpread.spreadLock	C_Lot_No ,intStratRow,C_Lot_Seq,intEndRow
	if .hdnImportflg.value = "Y" then
		ggoSpread.spreadUnlock C_HsCd,intStratRow,C_HsCd,intEndRow
		ggoSpread.sssetrequired C_HsCd,intStratRow,intEndRow
		ggoSpread.spreadUnlock C_Popup5,intStratRow,C_Popup5,intEndRow
	End if
	
	if UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnIVFlg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "N" then 
		ggoSpread.spreadUnlock C_IOFlg,intStratRow,C_IOFlg,intEndRow
	    ggoSpread.sssetrequired C_IOFlg,intStratRow,intEndRow
	    ggoSpread.spreadUnlock C_VatType,intStratRow,C_Popup7,intEndRow

	end if
	.vspdData.ReDraw = True
    
    End With
        
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    Dim index,SumTotal,idel
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow

		SumTotal = UNICDbl(frm1.txtGrossAmt.value)
		for index = .SelBlockRow to .SelBlockRow2
			.Row = index
			.Col = C_Stateflg
			idel = .text
			.Col = 0

			if Trim(.text) <> ggoSpread.InsertFlag and Trim(idel) <> "D" then
			    .Col = C_NetOrderAmt							
		         SumTotal = SumTotal - UNICDbl(.Text)
                 .Col = C_Stateflg
			     frm1.vspdData.text = "D"
                 frm1.txtGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")     
                 
		    end if
		Next
    End With
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                   
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                   
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_MULTI)							
End Function
		
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False)                    
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
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* %>
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
    
    Err.Clear           

	Dim strVal
    
    Call SetToolbar("11100000000111") '조회버튼 누르자마자 행추가 누르는 것을 방지 

    
    With frm1    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    strVal = strVal & "&txtPoNo=" & .txthdnPoNo.value
	    strVal = strVal & "&txtQuerytype=" & .txtQuerytype.value
	 
    else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
	    strVal = strVal & "&txtQuerytype=" & .txtQuerytype.value
    
    end if 
    .hdnmaxrow.value = .vspdData.MaxRows
    If LayerShowHide(1) = False Then Exit Function
    
    Call RunMyBizASP(MyBizASP, strVal)				
   
   
    End With
    
    DbQuery = True
    
End Function

Function ToolBarCtrl()
    if frm1.txtRelease.Value <> "Y" then
		Call SetToolbar("11101111001111")
    else
		Call SetToolbar("11100000000111")
    end if
    
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
	Call SetSpreadLockAfterQuery

	Call ToolBarCtrl()
	if Trim(UCase(frm1.hdnReleaseflg.Value)) = "Y" then
		frm1.btnCfm.value = "확정취소"
	else
		frm1.btnCfm.value = "확정"
	end if
	frm1.btnCfmSel.disabled = False
	
    frm1.vspdData.focus
	Set gActiveElement = document.activeElement
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal,strDel
	Dim ColSep, RowSep
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
    
	ColSep = Parent.gColSep		'";"									
	RowSep = Parent.gRowSep		'"|"									

	With frm1
		.txtMode.value = Parent.UID_M0002
		    
		lGrpCnt = 1    
		strVal = ""
		iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

		ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
		ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

		iTmpCUBufferCount = -1
		iTmpDBufferCount = -1

		strCUTotalvalLen = 0
		strDTotalvalLen  = 0
    
    For lRow = 1 To .vspdData.MaxRows

        Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
            Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag	
 				if Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))=ggoSpread.InsertFlag then
					strVal = "C" & ColSep						'0
				Else
					strVal = "U" & ColSep
				End if      
				
				If Trim(UNICDbl(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))) = "" Or Trim(UNICDbl(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))) = "0" then
					Call DisplayMsgBox("970021", "X","반품수량", "X")
					Call LayerShowHide(0)
					Exit Function
				End if
					
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_itemCd,lRow,"X","X")) & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X")),0) & ColSep
				End If
                   
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X")) & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Cost,lRow,"X","X")),0) & ColSep
				End If
                    
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_CostConCd,lRow,"X","X")) & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_OrderAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderAmt,lRow,"X","X")),0) & ColSep
				End If

                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_IOFlgCd,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_VatType,lRow,"X","X")) & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_VatRate,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_VatRate,lRow,"X","X")),0) & ColSep
				End If
                   
                If Trim(GetSpreadText(.vspdData,C_VatAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_VatAmt,lRow,"X","X")),0) & ColSep
				End If
                   
                strVal = strVal & UNIConvDate(Trim(GetSpreadText(.vspdData,C_DlvyDT,lRow,"X","X"))) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_HSCd,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_SLCd,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_TrackingNo,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_Lot_No,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_Lot_Seq,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_RetCd,lRow,"X","X")) & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_Over,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Over,lRow,"X","X")),0) & ColSep
				End If
					      
                If Trim(GetSpreadText(.vspdData,C_Under,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Under,lRow,"X","X")),0) & ColSep
				End If

                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_PrNo,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_MvmtNo,lRow,"X","X")) & ColSep
                '비고 추가 
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Remrk,lRow,"X","X"))  & ColSep

                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_MaintSeq,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_SoNo,lRow,"X","X")) & ColSep
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_Stateflg,lRow,"X","X")) & ColSep
	                
	            '반품등록 추가 C_IVNO,C_IVSEQ  
				strVal = strVal & UCase(Trim(GetSpreadText(.vspdData,C_IVNO,lRow,"X","X"))) & ColSep

				If Trim(GetSpreadText(.vspdData,C_IVSEQ,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_IVSEQ,lRow,"X","X")),0) & ColSep
				End If

                 If Trim(GetSpreadText(.vspdData,C_NetOrderAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_NetOrderAmt,lRow,"X","X")),0) & ColSep
				End If

				strVal = strVal & "N" & ColSep
				strVal = strVal & lRow & ColSep	
				'2003.08
				strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_PoNo,lRow,"X","X")) & ColSep		'M109_I3_ref_po_no
                strVal = strVal & Trim("" & GetSpreadText(.vspdData,C_PoSeqNo,lRow,"X","X")) & ColSep	'M109_I3_ref_po_seq_no
                
                strVal = strVal & "" & ColSep	'M109_I3_row_num_Amend
				strVal = strVal & "" & ColSep	'M109_I3_ext1_cd
				strVal = strVal & "" & ColSep	'M109_I3_ext1_qty	
				strVal = strVal & "" & ColSep	'M109_I3_ext1_amt				
				strVal = strVal & "" & ColSep	'M109_I3_ext1_rt				
				strVal = strVal & "" & ColSep	'M109_I3_ext2_cd				
				strVal = strVal & "" & ColSep	'M109_I3_ext2_qty				
				strVal = strVal & "" & ColSep	'M109_I3_ext2_amt				
				strVal = strVal & "" & ColSep	'M109_I3_ext2_rt				
				strVal = strVal & "" & ColSep	'M109_I3_ext3_cd
				strVal = strVal & "" & ColSep	'M109_I3_ext3_qty										
				strVal = strVal & "" & ColSep	'M109_I3_ext3_amt
				strVal = strVal & "" & ColSep	'M109_I3_ext3_rt
				strVal = strVal & "" & RowSep	'M109_I3_max_col_num				              								

                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag
                
'                strDel = "D" & ColSep
'				strDel = strDel & Trim(GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")) & ColSep
'				strDel = strDel & Trim(GetSpreadText(.vspdData,C_PrNo,lRow,"X","X")) & ColSep				
'				strDel = strDel & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep
'                strDel = strDel & lRow & RowSep
'                lGrpCnt = lGrpCnt + 1

                strDel = "D" & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")) & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_PrNo,lRow,"X","X")) & ColSep
				strDel = strDel & ColSep
				' 삭제될 수량을 보낸다.
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X")) & ColSep
				strDel = strDel & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep
				'strDel = strDel & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep
'-----------
'add by jt.kim
                strDel = strDel & Trim("" & GetSpreadText(.vspdData,C_MvmtNo,lRow,"X","X")) & ColSep
'-----
				strDel = strDel & ColSep & ColSep & ColSep & ColSep & ColSep '& ColSep
               strDel = strDel & lRow & RowSep
                lGrpCnt = lGrpCnt + 1

        End Select
                
		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
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
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
		          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
		       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
		         iTmpDBuffer(iTmpDBufferCount) =  strDel         
		         strDTotalvalLen = strDTotalvalLen + Len(strDel)
		End Select   
    Next
		
	.txtMaxRows.value = lGrpCnt-1
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	if lGrpCnt > 1 then
		If LayerShowHide(1) = False Then Exit Function
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					
	End if
	
	End With
	
    DbSave = True                                             
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()											
   
	Call InitVariables
	
    'lgIntFlgMode = Parent.OPMD_UMODE		'⊙: Indicates that current mode is Update mode
	'Call ggoOper.LockField(Document, "Q")
    Call MainQuery()

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
End Function
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute 
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>반품내역등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenMvmtRef">구매입고참조</A> | <A href="vbscript:OpenRetRef">반품출고참조</A> | <A href="vbscript:OpenIvRef">매입참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>반품번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo"  SIZE=29 MAXLENGTH=18 ALT="반품번호" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>반품형태</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="반품형태" NAME="txtPoTypeCd" SIZE=10 tag="24X">
													   <INPUT TYPE=TEXT NAME="txtPoTypeNm" SIZE=20 ALT ="반품형태" tag="24X" ></TD>
								<TD CLASS="TD5" NOWRAP>등록일</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3112ma7_fpDateTime2_txtPoDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" SIZE=10 tag="24X">
													   <INPUT TYPE=TEXT NAME="txtSupplierNm" SIZE=20 ALT ="공급처" tag="24X" ></TD>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="구매그룹" NAME="txtGroupCd" SIZE=10 tag="24X">
													   <INPUT TYPE=TEXT NAME="txtGroupNm" SIZE=20 ALT ="구매그룹" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>반품순금액</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3112ma7_fpDoubleSingle1_txtGrossAmt.js'></script></td>
								<TD CLASS="TD5" NOWRAP>화폐</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="화폐" NAME="txtCurr" SIZE=10 tag="24X">
													   <INPUT TYPE=TEXT NAME="txtCurrNm" SIZE=20 ALT ="화폐" tag="24X" ></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m3112ma7_I326293361_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
		
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td align="Left"><a><button name="btnCfmSel" id="btnCfm" class="clsmbtn" ONCLICK="vbscript:Cfm()">확정</button></a></td>					
					<td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">반품발주등록</a> <!--| <a href="VBSCRIPT:CookiePage(2)">경비등록</a></td>-->
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SRC="m3112mb1.asp" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden"  NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRelease" tag="14">
<INPUT TYPE=HIDDEN NAME="txthdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="txtQuerytype" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnDlvyDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnImportFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubContraFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnClsflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnXch" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMode" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnTrackingflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHMaintNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVATType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVATRate" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVATINCFLG" tag="1">
<INPUT TYPE=HIDDEN NAME="hdnXchRateOp"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIVFlg"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow"  tag="14">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
