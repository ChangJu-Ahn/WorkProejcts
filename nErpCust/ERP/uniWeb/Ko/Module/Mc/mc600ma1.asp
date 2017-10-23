<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc600ma1
'*  4. Program Name         : 납입지시입고등록 
'*  5. Program Desc         : 납입지시입고등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/28
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ahn Jung Je
'* 10. Modifier (Last)      : Kang Su Hwan
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
'##########################################################################################################
'******************************************  1.1 Inc 선언   ***************************************** !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   =====================================!-->
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
Const BIZ_PGM_QUERY_ID = "mc600mb1.asp"									
Const BIZ_PGM_SAVE_ID = "mc600mb2.asp"									
Const BIZ_PGM_JUMP_ID	= "M4131MA1"

'==========================================  1.2.1 Global 상수 선언  ======================================

'==========================================  1.2.2 Global 변수 선언  =====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable값 정의  ===============================
Dim IsOpenPop          
Dim lblnWinEvent
Dim lgOpenFlag
Dim lgCurrentDay
Dim interface_Account

Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec   
Dim C_TrackingNo 
Dim C_InspFlg	
Dim C_GrQty	
Dim C_GRUnit
Dim C_SlCd	
Dim C_SlCdPop	
Dim C_SlNm		
Dim C_LotNo		
Dim C_LotSeqNo	
Dim C_MakerLotNo
Dim C_MakerLotSeqNo
Dim C_PoNo
Dim C_PoSeqNo
Dim C_ProdtOrderNo    
Dim C_OprNo
Dim C_Seq
Dim C_SubSeq
Dim C_MvmtNo

'--------------------------------------------------------------------
'		Field의 Tag속성을 Protect로 전환,복구 시키는 함수 
'--------------------------------------------------------------------
Function ChangeTag(Byval Changeflg)
	Dim index

	If Changeflg = true then
	
		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "Q"
	    ggoSpread.SpreadLock -1,	-1
		Call ggoSpread.SSSetColHidden(C_ProdtOrderNo,C_ProdtOrderNo,True)	
		Call ggoSpread.SSSetColHidden(C_OprNo,C_OprNo,True)	

	Else
		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "N"
		Call ggoOper.LockField(Document, "N")	
		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "D"
		Call ggoSpread.SSSetColHidden(C_ProdtOrderNo,C_ProdtOrderNo,False)	
		Call ggoSpread.SSSetColHidden(C_OprNo,C_OprNo,False)	
	End if 
End Function 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_PlantCd		= 1
	C_PlantNm		= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5 
	C_TrackingNo	= 6
	C_InspFlg		= 7
	C_GrQty			= 8
	C_GRUnit		= 9
	C_SlCd			= 10
	C_SlCdPop		= 11
	C_SlNm			= 12
	C_LotNo			= 13
	C_LotSeqNo		= 14
	C_MakerLotNo	= 15
	C_MakerLotSeqNo	= 16
	C_PoNo			= 17
	C_PoSeqNo		= 18
	C_ProdtOrderNo	= 19
	C_OprNo			= 20
	C_Seq			= 21
	C_SubSeq		= 22
	C_MvmtNo		= 23
End Sub
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                
    lgIntGrpCount = 0                        
    
    lgStrPrevKey = ""                        
    lgLngCurRows = 0                         
    
    lgBlnFlgChgValue = False                 
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	lgOpenFlag = False    
	
	lgCurrentDay = UNIConvDateAtoB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtGmDt.Text = lgCurrentDay
	
	frm1.txtGroupCd.Value = Parent.gPurGrp
    
    Call SetToolBar("1110000000001111")
    
    frm1.txtMvmtNo.focus 
    Set gActiveElement = document.activeElement    
    interface_Account = GetSetupMod(Parent.gSetupMod, "a")
	frm1.btnGlSel.disabled = true    
End Sub

'==========================================  2.2.2 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" IO_Type_Cd,IO_Type_NM "," ( select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b where a.rcpt_type = b.io_type_cd    and a.sto_flg = " & FilterVar("N", "''", "S") & "  AND a.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and ((b.RCPT_FLG=" & FilterVar("Y", "''", "S") & "  AND b.RET_FLG=" & FilterVar("N", "''", "S") & " ) or (b.RET_FLG=" & FilterVar("N", "''", "S") & "  And b.SUBCONTRA_FLG=" & FilterVar("N", "''", "S") & " )) ) c ", _
					 " 1 = 1 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboMvmtType ,lgF0  ,lgF1  ,Chr(11))
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021118",,Parent.gAllowDragDropSpread  
		
		.ReDraw = false
		
		.MaxCols = C_MvmtNo + 1
    	.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
			
		ggoSpread.SSSetEdit 		C_PlantCd,	"공장", 10,,,,2
		ggoSpread.SSSetEdit 		C_PlantNm,	"공장명", 20
		ggoSpread.SSSetEdit 		C_ItemCd,	"품목", 10,,,,2
		ggoSpread.SSSetEdit 		C_ItemNm,	"품목명", 20 
		ggoSpread.SSSetEdit 		C_Spec,	    "품목규격", 20 	
		ggoSpread.SSSetEdit 		C_TrackingNo,	"Tracking No.", 25 				
		ggoSpread.SSSetCheck 		C_InspFlg,	"검사품여부",10,,,true
        ggoSpread.SSSetFloat		C_GrQty,	"입고수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 		C_GRUnit,	"단위", 10
		ggoSpread.SSSetEdit			C_SlCd,		"창고", 10,,,,2
		ggoSpread.SSSetButton 		C_SlCdPop
		ggoSpread.SSSetEdit 		C_SlNm,		"창고명", 20	    
		ggoSpread.SSSetEdit 		C_LotNo,	"Lot No.", 20,,,,2    
        ggoSpread.SSSetFloat		C_LotSeqNo, "LOT NO 순번", 20, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
		ggoSpread.SSSetEdit 		C_MakerLotNo,	"MAKER LOT NO.", 20,,,,2    
        ggoSpread.SSSetFloat		C_MakerLotSeqNo, "Maker Lot 순번", 20, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
		ggoSpread.SSSetEdit 		C_PoNo,		"발주번호", 20,,,,2
        ggoSpread.SSSetFloat		C_PoSeqNo, "발주순번", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
 
		ggoSpread.SSSetEdit 		C_ProdtOrderNo,	"제조오더번호", 20
		ggoSpread.SSSetEdit 		C_OprNo,		"공정", 10
        ggoSpread.SSSetFloat		C_Seq,		"부품예약일련번호", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
        ggoSpread.SSSetFloat		C_SubSeq,	"납입지시 순번", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
		ggoSpread.SSSetEdit 		C_MvmtNo,	"Movmt.No.", 20

		Call ggoSpread.MakePairsColumn(C_SlCd,C_SlCdPop)
		Call ggoSpread.SSSetColHidden(C_Seq,.MaxCols,True)	

		.ReDraw = true
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData

    ggoSpread.SpreadLock -1,	-1
    ggoSpread.SpreadUnLock	C_GrQty,	-1,		C_GrQty
    ggoSpread.SpreadUnLock	C_SlCd,		-1,		C_SlCd
    ggoSpread.SpreadUnLock	C_SlCdPop,	-1,		C_SlCdPop
    ggoSpread.SpreadUnLock	C_LotNo,	-1,		C_LotNo
    ggoSpread.SpreadUnLock	C_LotSeqNo,	-1,		C_LotSeqNo
    ggoSpread.SpreadUnLock	C_MakerLotNo,	-1,		C_MakerLotNo
    ggoSpread.SpreadUnLock	C_MakerLotSeqNo,	-1,		C_MakerLotSeqNo
	
	ggoSpread.SSSetRequired 	C_GrQty ,		-1, -1
	ggoSpread.SSSetRequired 	C_SlCd ,		-1, -1
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

			C_PlantCd		= iCurColumnPos(1)
			C_PlantNm		= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5) 
			C_TrackingNo	= iCurColumnPos(6)
			C_InspFlg		= iCurColumnPos(7)
			C_GrQty			= iCurColumnPos(8)
			C_GRUnit		= iCurColumnPos(9)
			C_SlCd			= iCurColumnPos(10)
			C_SlCdPop		= iCurColumnPos(11)
			C_SlNm			= iCurColumnPos(12)
			C_LotNo			= iCurColumnPos(13)
			C_LotSeqNo		= iCurColumnPos(14)
			C_MakerLotNo	= iCurColumnPos(15)
			C_MakerLotSeqNo	= iCurColumnPos(16)
			C_PoNo			= iCurColumnPos(17)
			C_PoSeqNo		= iCurColumnPos(18)
			C_ProdtOrderNo	= iCurColumnPos(19)
			C_OprNo			= iCurColumnPos(20)
			C_Seq			= iCurColumnPos(21)
			C_SubSeq		= iCurColumnPos(22)
			C_MvmtNo		= iCurColumnPos(23)
	End Select
End Sub	

'------------------------------------------  OpenGLRef()  -------------------------------------------------
'	Name : OpenGLRef()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenGLRef()
	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnGlNo.value)
	arrParam(1) = ""
	
   If frm1.hdnGlType.Value = "A" Then               '회계전표팝업 
	   strRet = window.showModalDialog("../../comasp/a5120ra1.asp", Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    Elseif frm1.hdnGlType.Value = "T" Then          '결의전표팝업 
	   strRet = window.showModalDialog("../../comasp/a5130ra1.asp", Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다. 
    End if

	lblnWinEvent = False
End Function

'------------------------------------------  OpenDlvyOrdRef()  -------------------------------------------------
'	Name : OpenDlvyOrdRef()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenDlvyOrdRef()
	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","참조" )
		Exit Function
	End if 

	if Trim(frm1.cboMvmtType.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "입고형태","X")
		frm1.cboMvmtType.focus
		Set gActiveElement = document.activeElement
		Exit Function	
	elseif Trim(frm1.txtSupplierCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공급처","X")
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	
	If lblnWinEvent = True Then Exit Function

	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)

	iCalledAspName = AskPRAspName("MC601RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "MC601RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	lgOpenFlag	= False

	If isEmpty(strRet) Then Exit Function			'페이지를 찾을 수 없는 에러발생시.	

	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetDlvyOrdRef(strRet)
	End If	
End Function

Function SetDlvyOrdRef(strRet)
	Dim Index1,Count1
	Dim Row1
	
	Const C_PlantCd1		= 1
	Const C_PlantNm1		= 2
	Const C_ProdtOrderNo1	= 3											'☆: Spread Sheet의 Column별 상수 
	Const C_ItemCd1			= 4
	Const C_ItemNm1			= 5
	Const C_Spec1			= 6
	Const C_BaseUnit1		= 7
	Const C_DoQty1			= 8
	Const C_RcptQty1		= 9
	Const C_SlCd1			= 10
	Const C_SlNm1			= 11
	Const C_DoDate1			= 12
	Const C_DoTime1			= 13
	Const C_TrackingNo1		= 14
	Const C_InspFlag1		= 15
	Const C_PoNo1			= 16
	Const C_PoSeqNo1		= 17
	Const C_WcCd1			= 18
	Const C_OprNo1			= 19
	Const C_Seq1			= 20
	Const C_SubSeq1			= 21
	Const C_PurGrp1			= 22

	Count1 = Ubound(strRet,1)

	Call Fncinsertrow(Count1+1)
	
	With frm1.vspdData
	
		.Redraw = False
		
		For index1 = 0 to Count1
			Row1 = .ActiveRow + Index1
		
			Call .SetText(C_PlantCd,	Row1, strRet(index1,C_PlantCd1 - 1))
			Call .SetText(C_PlantNm,	Row1, strRet(index1,C_PlantNm1 - 1))
			Call .SetText(C_ItemCd,		Row1, strRet(index1,C_ItemCd1 - 1))
			Call .SetText(C_ItemNm,		Row1, strRet(index1,C_ItemNm1 - 1))
			Call .SetText(C_Spec,		Row1, strRet(index1,C_Spec1 - 1))
			Call .SetText(C_TrackingNo,	Row1, strRet(index1,C_TrackingNo1 - 1))
			Call .SetText(C_GrQty,		Row1, UNIFormatNumber(UNICDbl(strRet(index1,C_DoQty1 - 1)) - UNICDbl(strRet(index1,C_RcptQty1 - 1)),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			Call .SetText(C_GRUnit,	Row1, strRet(index1,C_BaseUnit1 - 1))
			Call .SetText(C_SlCd,	Row1, strRet(index1,C_SlCd1 - 1))
			Call .SetText(C_PoNo,	Row1, strRet(index1,C_PoNo1 - 1))
			Call .SetText(C_PoSeqNo,	Row1, strRet(index1,C_PoSeqNo1 - 1))
			Call .SetText(C_ProdtOrderNo,	Row1, strRet(index1,C_ProdtOrderNo1 - 1))
			Call .SetText(C_OprNo,	Row1, strRet(index1,C_OprNo1 - 1))
			Call .SetText(C_Seq,	Row1, strRet(index1,C_Seq1 - 1))
			Call .SetText(C_SubSeq,	Row1, strRet(index1,C_SubSeq1 - 1))
			
			.Row = Row1
			.Col = C_InspFlg
			If strRet(index1,C_InspFlag1 - 1) = "Y" Then
				.Text = "1"
			Else
				.Text = "0"
			End if
		Next
		
		frm1.txtGroupCd.value = strRet(0,C_PurGrp1 - 1)
		
		.ReDraw = True

		Call setReference()
	End With
End Function

'------------------------------------------  OpenSlCd()  -------------------------------------------------
'	Name : OpenSlCd()
'	Description : Sl PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSlCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1.vspdData

		arrParam(0) = "창고"						
		arrParam(1) = "B_STORAGE_LOCATION"			
		.Row = .ActiveRow	
		.Col = C_SlCd	
		arrParam(2) = Trim(.Text)		
		.Col=C_PlantCd
		arrParam(4) = "PLANT_CD= " & FilterVar(.Text, "''", "S") & " "
		arrParam(5) = "창고"						
	
		arrField(0) = "SL_CD"						
		arrField(1) = "SL_NM"						
    
		arrHeader(0) = "창고"					
		arrHeader(1) = "창고명"					
    
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
		IsOpenPop = False
		If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
		If arrRet(0) = "" Then
			Exit Function
		Else
			.Row = .ActiveRow
			.Col = C_SlCd
			.Text = arrRet(0)
			.Col = C_SlNm
			.Text = arrRet(1)		
		End If
	End With	
End Function

'------------------------------------------  OpenMvmtNo()  -------------------------------------------------
'	Name : OpenMvmtNo()
'	Description : OpenPoNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMvmtNo()
		Dim strRet
		Dim arrParam(3)
		Dim iCalledAspName
		Dim IntRetCD
	
		If lblnWinEvent = True Or UCase(frm1.txtMvmtNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
		lblnWinEvent = True

		arrParam(0) = ""'Trim(frm1.hdnSupplierCd.Value)
		arrParam(1) = ""'Trim(frm1.hdnGroupCd.Value)
		arrParam(2) = ""'Trim(frm1.hdnMvmtType.Value)		
		arrParam(3) = ""'This is for Inspection check, must be nothing.
		
		iCalledAspName = AskPRAspName("MC602PA1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "MC602PA1", "X")
			IsOpenPop = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
		lblnWinEvent = False
		If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
		If strRet(0) = "" Then
			frm1.txtMvmtNo.focus
			Exit Function
		Else
			frm1.txtMvmtNo.value = strRet(0)
			frm1.txtMvmtNo.focus
			Set gActiveElement = document.ActiveElement   
		End If	
End Function

'------------------------------------------  OpenGroup()  ------------------------------------------------
'	Name : OpenGroup()
'	Description : OpenGroup1 PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
	
	arrParam(4) = "B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    arrHeader(2) = "구매조직"		
    arrHeader(3) = "구매조직명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			
	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)		
		frm1.txtGroupCd.focus
		Set gActiveElement = document.ActiveElement   
		lgBlnFlgChgValue = True
	End If	
End Function

'------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenSppl()
'	Description :  OpenSppl PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtSupplierCd.className)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"				
	arrParam(1) = "B_Biz_Partner"
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	arrParam(3) = ""							
	
	arrParam(4) = "Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & " "	
	arrParam(5) = "공급처"				
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					

	arrHeader(0) = "공급처"				
	arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.ActiveElement   
		lgBlnFlgChgValue = True
	End If	
End Function

'==========================================================================================
'   Event Name : setReference()
'   Event Desc : 
'==========================================================================================
Function setReference()
	ggoOper.SetReqAttr	frm1.cboMvmtType, "Q"
	ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"

	Call SetToolBar("11101001000111")
End Function

'==========================================================================================
'   Event Name : CookiePage
'   Event Desc : 
'==========================================================================================
Function CookiePage(Byval Kubun)
	Dim strTemp

	If Kubun = 1 Then
	    
	    WriteCookie "MvmtNo" , Trim(frm1.txtMvmtNo1.value)				
		Call PgmJump(BIZ_PGM_JUMP_ID)
		
	Else
		strTemp = ReadCookie("MvmtNo")
	
		If strTemp = "" then Exit Function
	
		frm1.txtMvmtNo.value = ReadCookie("MvmtNo")
	
		Call WriteCookie("MvmtNo" , "")
	
		MainQuery()
	End if
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029                                                    
    Call ggoOper.LockField(Document, "N")                                  
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call AppendNumberPlace("6", "3", "0")
    Call InitSpreadSheet                                                   
    Call SetDefaultVal
	Call InitVariables
   	Call InitComboBox
    Call CookiePage(0)
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    IF lgIntFlgMode <> Parent.OPMD_UMODE And frm1.vspdData.MaxRows <= 0 Then
		Call SetPopupMenuItemInf("0000111111")
	ElseIf lgIntFlgMode <> Parent.OPMD_UMODE And frm1.vspdData.MaxRows > 0 Then	'참조시 
		Call SetPopupMenuItemInf("0001111111")
	Else
		Call SetPopupMenuItemInf("0101111111")
	End If
   
   gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
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

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	ggoSpread.Source = frm1.vspdData
    If Row > 0 And Col = C_SlCdPop then Call OpenSlCd()
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then gMouseClickStatus = "SPCR"
End Sub    

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
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
    Call ChangeTag(True)
End Sub

'==========================================================================================
'   Event Name : txtGmDt
'   Event Desc :
'==========================================================================================
Sub txtGmDt_DblClick(Button)
	If Button = 1 then 
		frm1.txtGmDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtGmDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtGmDt
'   Event Desc :
'==========================================================================================
Sub txtGmDt_Change()
	lgBlnFlgChgValue = true	
End Sub

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    
    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
    
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
		If lgStrPrevKey <> "" Then	
			If CheckRunningBizProcess = True Then Exit Sub
			
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
    
    On Error Resume Next                                                 
    Err.Clear                                               
    
	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then	Exit Function
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function
  
  	'-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")                  
	Call InitVariables

    '-----------------------
    'Check Delivery Order
    '-----------------------
 	If 	CommonQueryRs(" DLVY_ORD_FLG "," M_PUR_GOODS_MVMT ", " MVMT_RCPT_NO = " & FilterVar(frm1.txtMvmtNo.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
		Call DisplayMsgBox("174100","X","X","X")
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))

	If Trim(lgF0(0)) <> "Y" Then
		Call DisplayMsgBox("17C004","X","X","X")
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Exit function
	End If 

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
    FncQuery = True											
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    On Error Resume Next                                   
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ChangeTag(False)
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")           
    Call SetDefaultVal
    Call InitVariables
        
    Set gActiveElement = document.ActiveElement   
    FncNew = True                     
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim intIndex
    
    FncSave = False                                 
    
    On Error Resume Next                           
    Err.Clear                                       
    
	ggoSpread.Source = frm1.vspdData				
    
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")					
        Exit Function
    End If

    If Not chkField(Document, "2") Then Exit Function
	
	If CompareDateByFormat(frm1.txtGmDt.text,lgCurrentDay,frm1.txtGmDt.Alt,"현재일", _
               "970025",frm1.txtGmDt.UserDefinedFormat,Parent.gComDateType,True) = False  then	
		Exit Function
	End if   
	

    ggoSpread.Source = frm1.vspdData									
    If Not ggoSpread.SSDefaultCheck Then Exit Function
    
    If frm1.vspdData.Maxrows < 1 then Exit Function
    '-----------------------
    'Check content area
    '-----------------------
    For intIndex = 1 to frm1.vspdData.MaxCols 
		frm1.vspdData.SetColItemData intindex,0	
	Next
	    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                      
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	If frm1.vspdData.Maxrows < 1 then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo  
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow

	On Error Resume Next
	Err.Clear
	
	FncInsertRow = False

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End IF
	
	With frm1.vspdData
		.ReDraw = False
		.focus
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow .ActiveRow, imRow
		Call SetSpreadLock
		.ReDraw = True
    End With
	
	If Err.number = 0 Then FncInsertRow = True
	
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    
    ggoSpread.Source = frm1.vspdData
    If frm1.vspdData.Maxrows < 1 then exit function
    
    With frm1.vspdData 
		.focus    
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
	End With
End Function
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_SINGLEMULTI)		
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False) 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")           
		
		If IntRetCD = vbNo Then Exit Function
		
    End If
    
    FncExit = True

End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'*********************************************************************************************************
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
    
    DbQuery = False  
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
  
    If LayerShowHide(1) = False Then Exit Function
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtRcptNo=" & .hdnRcptNo.value
		    strVal = strVal & "&txtMvmtNo=" & .hdnMvmtNo.value
		else
		    strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&txtMvmtNo=" & Trim(.txtMvmtNo.value)
		End if
    
		Call RunMyBizASP(MyBizASP, strVal)									
    End With
    
    DbQuery = True

End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()													
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
	
	'-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE											
    
    Call ggoOper.LockField(Document, "Q")								
	lgBlnFlgChgValue = False	
	
	Call SetToolBar("11101011000111")
	
	Call ChangeTag(True)
	
	if interface_Account = "N" then		
		frm1.btnGlSel.disabled = true
	Else 
		frm1.btnGlSel.disabled = False		
	End if
	frm1.vspdData.focus
End Function
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim igColSep,igRowSep
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

	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode
	igColSep = parent.gColSep
	igRowSep = parent.gRowSep

	With frm1
		lGrpCnt = 1
	
		If lgIntFlgMode = Parent.OPMD_CMODE Then
			
			For lRow = 1 To .vspdData.MaxRows
		    
				strVal = Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_TrackingNo,lRow,"X","X"))    & igColSep

                If Trim(GetSpreadText(frm1.vspdData,C_InspFlg,lRow,"X","X")) = "0" Then
					strVal = strVal & "N" & igColSep
				Else
					strVal = strVal & "Y" & igColSep
				End If
                    
				If UNICDbl(GetSpreadText(frm1.vspdData,C_GrQty,lRow,"X","X")) = "" Or UNICDbl(GetSpreadText(frm1.vspdData,C_GrQty,lRow,"X","X")) = "0" then 
					Call DisplayMsgBox("970021","X","입고수량","X")
					Call LayerShowHide(0)
					Exit Function
				End if

				strVal = strVal & UNIConvNum(Trim(GetSpreadText(frm1.vspdData,C_GrQty,lRow,"X","X")),0)    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GRUnit,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SlCd,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotSeqNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MakerLotNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MakerLotSeqNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PoNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PoSeqNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProdtOrderNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OprNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Seq,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SubSeq,lRow,"X","X"))    & igColSep
				strVal = strVal & lRow & igRowSep


		        lGrpCnt = lGrpCnt + 1

				Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
				    Case ggoSpread.InsertFlag
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
		   	
		Else 

			For lRow = 1 To .vspdData.MaxRows
				
				If Trim(GetSpreadText(frm1.vspdData,0,lRow,"X","X")) = ggoSpread.DeleteFlag Then
			
					strVal = Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow,"X","X"))    & igColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X"))    & igColSep
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(frm1.vspdData,C_GrQty,lRow,"X","X")),0)    & igColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GRUnit,lRow,"X","X"))    & igColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PoNo,lRow,"X","X"))    & igColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PoSeqNo,lRow,"X","X"))    & igColSep
					strVal = strVal & lRow & igRowSep
					    
					lGrpCnt = lGrpCnt + 1
		   		End If               
	
				Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
				    Case ggoSpread.DeleteFlag
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
		End If		
    	frm1.txtMaxRows.value = lGrpCnt-1

		If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)     
		End If   

		If LayerShowHide(1) = False Then Exit Function
	
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								

	End With
	
    DbSave = True                                                       
    
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call ChangeTag(False)
    Call ggoOper.ClearField(Document, "2")                  
    Call SetDefaultVal
	Call InitVariables
	Call MainQuery()
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
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

'========================================================================================
' Function Name : changeSpplCd
' Function Desc : 
'========================================================================================
Function changeSpplCd()

	With frm1
		'-----------------------
		'Check BP CODE		'공급처코드가 있는 지 체크 
		'-----------------------
		If 	CommonQueryRs(" BP_NM, BP_TYPE, usage_flag, in_out_flag "," B_Biz_Partner ", " BP_CD = " & FilterVar(.txtSuppliercd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("229927","X","X","X")
			.txtSupplierNm.Value = ""
			Call SetFocusToDocument("M") 
			.txtSuppliercd.focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		lgF3 = Split(lgF3, Chr(11))
		.txtSupplierNm.Value = lgF0(0)

		If Trim(lgF2(0)) <> "Y" Then
			Call DisplayMsgBox("179021","X","X","X")
			Call SetFocusToDocument("M") 
			.txtSuppliercd.focus
			Exit function
		End If
		If Trim(lgF1(0)) <> "S" and Trim(lgF1(0)) <> "CS" Then
			Call DisplayMsgBox("179020","X","X","X")
			Call SetFocusToDocument("M") 
			.txtSuppliercd.focus
			Exit function
		End If
		If Trim(lgF3(0)) <> "O" Then
			Call DisplayMsgBox("17C003","X","X","X")
			Call SetFocusToDocument("M") 
			.txtSuppliercd.focus
			Exit function
		End If
	End With
End Function

'========================================================================================
' Function Name : changeSpplCd
' Function Desc : 
'========================================================================================
Function changeGroupCd()
	
	changeGroupCd = False
	
	With frm1
		'-----------------------
		'Check  CODE		'구매그룹 있는 지 체크 
		'-----------------------
		If 	CommonQueryRs("PUR_GRP_NM, USAGE_FLG "," B_PUR_GRP ", "PUR_GRP = " & FilterVar(.txtGroupCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
						
			Call DisplayMsgBox("125100","X","X","X") ' 구매그룹이 없다.
			.txtGroupNm.Value = ""
			Call SetFocusToDocument("M") 
			.txtGroupCd.focus 
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		.txtGroupNm.Value = lgF0(0)

		If Trim(lgF1(0)) <> "Y" Then
			Call DisplayMsgBox("125114","X","X","X")
			Call SetFocusToDocument("M") 
			.txtGroupCd.focus
			Exit function
		End If
	End With
	
	changeGroupCd = True

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납입지시입고</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenDlvyOrdRef()">납입지시 입고대상 참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>입고번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo" MAXLENGTH=18 SIZE=32 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMvmtNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()"></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
								<TD CLASS="TD5" NOWRAP>입고형태</TD>
								<TD CLASS="TD6"><SELECT Name="cboMvmtType" ALT="입고형태"  STYLE="WIDTH: 150px" tag="23"><OPTION Value=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>입고일</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/mc600ma1_fpDateTime1_txtGmDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="23XXXU" OnChange="VBScript:changeSpplCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="25XXXU" OnChange="VBScript:changeGroupCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="구매그룹" ID="txtGroupNm" SIZE=20 NAME="arrCond" tag="24X"></TD>								
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>입고번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo1" MAXLENGTH=18 SIZE=34 tag="21XXXU"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/mc600ma1_I721425074_vspdData.js'></script>
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
					<td>						
		         		<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>&nbsp;
					</td>					
					<td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">검사결과등록</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnMvmtNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnMvmtCur" tag="24" TabIndex="-1">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>	
