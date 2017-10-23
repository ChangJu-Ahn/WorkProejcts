<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Sale,Production
'*  2. Function Name        : 
'*  3. Program ID           : m4132ma1
'*  4. Program Name         : 예외입고/반품 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/08/22
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Duk Hyun
'* 10. Modifier (Last)      : 
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
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<!--meta http-equiv="Content-type" content="text/html; charset=euc-kr"-->
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

Const BIZ_PGM_ID = "m4132mb1.asp"			

<!-- #Include file="../../inc/lgvariables.inc" -->
'==============================================================================================================================
Dim IsOpenPop          
Dim lblnWinEvent
Dim interface_Account

Dim C_PlantCd      ' 공장 
Dim C_PlantPop      
Dim C_PlantNm      ' 공장명 
Dim C_ItemCd       ' 품목 
Dim C_ItemPop      
Dim C_ItemNm       ' 품목명 
Dim C_Spec         ' 품목규격 
Dim C_Unit         ' 단위 
Dim C_UnitPop
Dim C_GrQty        ' 수량 
Dim C_GrQtyPop
Dim C_StockQty     ' 재고처리수량 
Dim C_Cur          ' 화폐 
Dim C_CurPop       
Dim C_MvmtPrc      ' 품목단가 
Dim C_DocAmt       ' 품목금액 
Dim C_WorkPrc      ' 가공비단가 
Dim C_WorkLocAmt   ' 가공비금액 
Dim C_SlCd         ' 창고 
Dim C_SlCdPop      
Dim C_SlNm         ' 창고명 
Dim C_LotNo        ' LOT NO
Dim C_LotNoPop
Dim C_LotNoSeq     ' LOT NO 순번 
Dim C_MakerLotNo   ' Maker Lot No
Dim C_MakerLotSeqNo' Maker Lot 순분 
Dim C_RetType	   ' 반품유형 
Dim C_RetTypePop
Dim C_RetTypeNm    ' 반품유형명 
Dim C_RemarkDtl	   ' 비고 
Dim C_TrackingNo   ' Tracking No
Dim C_TrackingNoPop
Dim C_GRNo         ' 재고처리번호 
Dim C_GRSeqNo      ' 재고처리순번 
Dim C_InspFlg      ' 검사품여부 
Dim C_InspSts      ' 검사상태 
Dim C_GRMeth       ' 납입시검사방법 
Dim C_InspReqNo    ' 검사요청번호 
Dim C_InspResultNo ' 검사결과등록번호 
Dim C_MvmtNo
Dim C_RefMvmtNo
Dim	C_ProcurType   ' 조달구분(Hidden)
'==============================================================================================================================
Function ChangeTag(Byval Changeflg)
	
	Dim index

	If Changeflg = true then
		ggoOper.SetReqAttr	frm1.txtGrNo1, "Q"
		ggoOper.SetReqAttr	frm1.txtRemark, "Q"
		frm1.vspdData.ReDraw = false
		For index = 1 to frm1.vspdData.MaxCols
			ggoSpread.SpreadLock index , -1
		Next
		frm1.vspdData.ReDraw = true
	Else
		Call ggoOper.LockField(Document, "N")
		ggoOper.SetReqAttr	frm1.txtGrNo1, "D"
		ggoOper.SetReqAttr	frm1.txtRemark, "D"
	End if 
	
End Function 

'==============================================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0        
    lgStrPrevKey = ""        
    lgLngCurRows = 0         
    frm1.vspdData.MaxRows = 0
    
End Sub
'==============================================================================================================================
Sub SetDefaultVal()
	frm1.txtGmDt.Text = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)
	
	frm1.txtGroupCd.Value = Parent.gPurGrp
    Call SetToolBar("1110010000001111")					 				
    frm1.txtGrNo.focus 
    Set gActiveElement = document.activeElement
    interface_Account = GetSetupMod(Parent.gSetupMod, "a")
	frm1.btnGlSel.disabled = true 
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'==============================================================================================================================
Sub InitSpreadSheet()

	call InitSpreadPosVariables()
	
	with frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20050823",,Parent.gAllowDragDropSpread  
		
		.ReDraw = false
		
		.MaxCols = C_ProcurType + 1
		.MaxRows = 0
		
		Call AppendNumberPlace("6", "5", "0")
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 		C_PlantCd,		"공장", 8, , , 4, 2
		ggoSpread.SSSetButton 		C_PlantPop
		ggoSpread.SSSetEdit 		C_PlantNm,		"공장명", 15
		ggoSpread.SSSetEdit 		C_ItemCd,		"품목", 15, , , 18, 2
		ggoSpread.SSSetButton 		C_ItemPop
		ggoSpread.SSSetEdit 		C_ItemNm,		"품목명", 20 
		ggoSpread.SSSetEdit 		C_Spec,	    	"품목규격", 20 	

		ggoSpread.SSSetEdit 		C_Unit,			"단위", 10, , , 3, 2
		ggoSpread.SSSetButton 		C_UnitPop
		SetSpreadFloatLocal 		C_GrQty,		"수량", 15, 1, 3
		ggoSpread.SSSetButton 		C_GrQtyPop
		SetSpreadFloatLocal 		C_StockQty,		"재고처리수량", 15, 1, 3
		ggoSpread.SSSetEdit 		C_Cur,	    	"화폐", 10, , , 3, 2
		ggoSpread.SSSetButton 		C_CurPop

		SetSpreadFloatLocal			C_MvmtPrc,		"품목단가"	, 15, 1, 4
		SetSpreadFloatLocal 		C_DocAmt,		"품목금액"	, 15, 1, 2
		SetSpreadFloatLocal 		C_WorkPrc,		"가공비단가"	, 15, 1, 4
		SetSpreadFloatLocal 		C_WorkLocAmt, 	"가공비금액"	, 15, 1, 2
		ggoSpread.SSSetEdit			C_SlCd,			"창고", 10 , , , 7, 2
		ggoSpread.SSSetButton 		C_SlCdPop
		ggoSpread.SSSetEdit 		C_SlNm,			"창고명", 20	    

		ggoSpread.SSSetEdit 		C_LotNo,		"Lot No.", 20, , , 25, 2    
		ggoSpread.SSSetButton 		C_LotNoPop
		SetSpreadFloatLocal 		C_LotNoSeq, 	"LOT NO 순번", 20,1,6
		ggoSpread.SSSetEdit 		C_MakerLotNo,	"MAKER LOT NO.", 20,,,12,2    
		SetSpreadFloatLocal 		C_MakerLotSeqNo,"Maker Lot 순번", 20,1,6
		ggoSpread.SSSetEdit  		C_RetType,		"반품유형", 10, , , , 2
		ggoSpread.SSSetButton 		C_RetTypePop
		ggoSpread.SSSetEdit 		C_RetTypeNm,	"반품유형명", 15    
		ggoSpread.SSSetEdit 		C_RemarkDtl,	"비고", 20
		ggoSpread.SSSetEdit 		C_TrackingNo,	"Tracking No.", 15 	
		ggoSpread.SSSetButton 		C_TrackingNoPop
		ggoSpread.SSSetEdit 		C_GRNo,			"재고처리번호", 20
		ggoSpread.SSSetFloat 		C_GRSeqNo,		"재고처리순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		ggoSpread.SSSetCheck 		C_InspFlg,		"검사품여부",10,,,true
		ggoSpread.SSSetEdit 		C_InspSts,		"검사상태", 10
		ggoSpread.SSSetEdit 		C_GRMeth,		"납입시검사방법", 20
		ggoSpread.SSSetEdit 		C_InspReqNo,	"검사요청번호", 20
		ggoSpread.SSSetEdit 		C_InspResultNo,	"검사결과등록번호", 20
		ggoSpread.SSSetEdit 		C_MvmtNo,		"", 10
		ggoSpread.SSSetEdit 		C_RefMvmtNo,		"", 10
		ggoSpread.SSSetEdit 		C_ProcurType,	"", 10

		Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantPop)
		Call ggoSpread.MakePairsColumn(C_ItemCd,C_ItemPop)
		Call ggoSpread.MakePairsColumn(C_Unit,C_UnitPop)
		Call ggoSpread.MakePairsColumn(C_GrQty,C_GrQtyPop)
		Call ggoSpread.MakePairsColumn(C_Cur,C_CurPop)
		Call ggoSpread.MakePairsColumn(C_SlCd,C_SlCdPop)
		Call ggoSpread.MakePairsColumn(C_LotNo,C_LotNoPop)
		Call ggoSpread.MakePairsColumn(C_RetType,C_RetTypePop)
		Call ggoSpread.MakePairsColumn(C_TrackingNo,C_TrackingNoPop)
		Call ggoSpread.SSSetColHidden(C_InspFlg,C_ProcurType ,True)	
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols ,True)	
			
		.ReDraw = true
		
		Call SetSpreadLock()
    End With
    
End Sub
'==============================================================================================================================
Sub InitSpreadPosVariables()
	C_PlantCd      	= 1
	C_PlantPop     	= 2
	C_PlantNm      	= 3
	C_ItemCd       	= 4
	C_ItemPop      	= 5
	C_ItemNm		= 6
	C_Spec         	= 7
	C_Unit         	= 8
	C_UnitPop      	= 9
	C_GrQty        	= 10
	C_GrQtyPop     	= 11
	C_StockQty     	= 12
	C_Cur          	= 13
	C_CurPop       	= 14
	C_MvmtPrc      	= 15
	C_DocAmt        = 16
	C_WorkPrc      	= 17
	C_WorkLocAmt   	= 18
	C_SlCd         	= 19
	C_SlCdPop      	= 20
	C_SlNm         	= 21
	C_LotNo        	= 22
	C_LotNoPop     	= 23
	C_LotNoSeq     	= 24
    C_MakerLotNo   	= 25
    C_MakerLotSeqNo	= 26
    C_RetType      	= 27
    C_RetTypePop	= 28
    C_RetTypeNm    	= 29
    C_RemarkDtl		= 30
    C_TrackingNo   	= 31
    C_TrackingNoPop = 32
    C_GRNo         	= 33
    C_GRSeqNo      	= 34
    C_InspFlg      	= 35
    C_InspSts      	= 36
    C_GRMeth       	= 37
    C_InspReqNo    	= 38
    C_InspResultNo 	= 39
    C_MvmtNo 		= 40
	C_RefMvmtNo		= 41
	C_ProcurType	= 42
End Sub
'==============================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd      	= iCurColumnPos(1)
			C_PlantPop     	= iCurColumnPos(2)
			C_PlantNm      	= iCurColumnPos(3)
			C_ItemCd       	= iCurColumnPos(4)
			C_ItemPop      	= iCurColumnPos(5)
			C_ItemNm		= iCurColumnPos(6)
			C_Spec         	= iCurColumnPos(7)
			C_Unit         	= iCurColumnPos(8)
			C_UnitPop      	= iCurColumnPos(9)
			C_GrQty        	= iCurColumnPos(10)
			C_GrQtyPop     	= iCurColumnPos(11)
			C_StockQty     	= iCurColumnPos(12)
			C_Cur          	= iCurColumnPos(13)
			C_CurPop       	= iCurColumnPos(14)
			C_MvmtPrc      	= iCurColumnPos(15)
			C_DocAmt        = iCurColumnPos(16)
			C_WorkPrc      	= iCurColumnPos(17)
			C_WorkLocAmt   	= iCurColumnPos(18)
			C_SlCd         	= iCurColumnPos(19)
			C_SlCdPop      	= iCurColumnPos(20)
			C_SlNm         	= iCurColumnPos(21)
			C_LotNo        	= iCurColumnPos(22)
			C_LotNoPop     	= iCurColumnPos(23)
			C_LotNoSeq     	= iCurColumnPos(24)
            C_MakerLotNo   	= iCurColumnPos(25)
            C_MakerLotSeqNo	= iCurColumnPos(26)
            C_RetType      	= iCurColumnPos(27)
            C_RetTypePop	= iCurColumnPos(28)
            C_RetTypeNm    	= iCurColumnPos(29)
            C_RemarkDtl		= iCurColumnPos(30)
            C_TrackingNo   	= iCurColumnPos(31)
            C_TrackingNoPop = iCurColumnPos(32)
            C_GRNo         	= iCurColumnPos(33)
            C_GRSeqNo      	= iCurColumnPos(34)
            C_InspFlg      	= iCurColumnPos(35)
            C_InspSts      	= iCurColumnPos(36)
            C_GRMeth       	= iCurColumnPos(37)
            C_InspReqNo    	= iCurColumnPos(38)
            C_InspResultNo 	= iCurColumnPos(39)
            C_MvmtNo		= iCurColumnPos(40)
            C_RefMvmtNo		= iCurColumnPos(41)
            C_ProcurType	= iCurColumnPos(42)
	End Select

End Sub	
'==============================================================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              'Lot 순번 Maker Lot 순번 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
    End Select
         
End Sub

'==============================================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	
    With ggoSpread
    
		.SpreadLock 	C_PlantCd , -1,C_PlantCd , -1
		.SpreadLock 	C_PlantNm , -1,C_PlantNm , -1
		.SpreadLock 	C_ItemCd , -1,C_ItemCd , -1
		.SpreadLock 	C_ItemNm , -1,C_ItemNm , -1
		.SpreadLock 	C_Spec , -1,C_Spec , -1
		.SSSetProtected frm1.vspdData.MaxCols, -1

    End With
    frm1.vspdData.ReDraw = True
End Sub
'==============================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    ggoSpread.Source = frm1.vspdData
    With ggoSpread
        frm1.vspdData.ReDraw = False
		
		.SSSetRequired	C_PlantCd,		pvStartRow, pvEndRow
		.SSSetProtected	C_PlantNm,		pvStartRow, pvEndRow
		.SSSetRequired	C_ItemCd,		pvStartRow, pvEndRow
		.SSSetProtected	C_ItemNm,		pvStartRow, pvEndRow
		.SSSetProtected	C_Spec,        	pvStartRow, pvEndRow
		.SSSetRequired	C_Unit,        	pvStartRow, pvEndRow
		.SSSetRequired	C_GrQty,       	pvStartRow, pvEndRow
		.SSSetProtected	C_StockQty,    	pvStartRow, pvEndRow
		.SSSetRequired	C_Cur,         	pvStartRow, pvEndRow    
		.SSSetRequired	C_MvmtPrc,     	pvStartRow, pvEndRow
		.SSSetProtected	C_DocAmt,      	pvStartRow, pvEndRow
		.SSSetRequired	C_WorkPrc,     	pvStartRow, pvEndRow
		.SSSetProtected	C_WorkLocAmt,  	pvStartRow, pvEndRow
        .SSSetRequired  C_SlCd,			pvStartRow, pvEndRow 
        .SSSetProtected C_SlNm,         pvStartRow, pvEndRow 
        .SSSetProtected C_LotNo,        pvStartRow, pvEndRow 
        .SSSetProtected C_LotNoPop,     pvStartRow, pvEndRow 
        .SSSetProtected C_LotNoSeq,     pvStartRow, pvEndRow 
        .SSSetProtected C_RetTypeNm,    pvStartRow, pvEndRow 
        .SSSetProtected C_TrackingNo,   pvStartRow, pvEndRow 
        .SSSetProtected C_TrackingNoPop,pvStartRow, pvEndRow 
        .SSSetProtected C_GRNo,         pvStartRow, pvEndRow 
        .SSSetProtected C_GRSeqNo,      pvStartRow, pvEndRow 
        .SSSetProtected C_InspFlg,      pvStartRow, pvEndRow 
        .SSSetProtected C_InspSts,      pvStartRow, pvEndRow 
        .SSSetProtected C_GRMeth,       pvStartRow, pvEndRow 
        .SSSetProtected C_InspReqNo,    pvStartRow, pvEndRow
        .SSSetProtected C_InspResultNo, pvStartRow, pvEndRow
                        
		.SSSetProtected frm1.vspdData.MaxCols, pvStartRow, pvEndRow
		                
		frm1.vspdData.ReDraw = True
    End With            
End Sub                 
'==============================================================================================================================
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
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			lblnWinEvent = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "T" Then          '결의전표팝업 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			lblnWinEvent = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다. 
    End if

	lblnWinEvent = False
	
End Function

'------------------------------------------  OpenRetRef()  -------------------------------------------------
'	Name : OpenRetRef()
'	Description : 예외반품출고참조 
'---------------------------------------------------------------------------------------------------------
Function OpenRetRef()
	Dim strRet
	Dim arrParam(15)
	Dim iCalledAspName
	Dim IntRetCD
	
	If Not chkField(Document, "2") Then  Exit Function
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call DisplayMsgBox("K21022", "X","X","X" )
		Exit Function
	End If

	if Not(UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "Y") then
		Call DisplayMsgBox("17A012", "X","입출고유형" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","예외반품출고참조" )
		frm1.txtGrNo.focus	
		Exit Function
	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	'===============쿨젠====================
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.hdnSubcontra2flg.value)
	'===============쿨젠====================

	iCalledAspName = AskPRAspName("M4132RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4132RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		frm1.txtGrNo.focus	
		Exit Function
	Else
		Call SetRetRef(strRet)
	End If	
End Function

'==============================================================================================================================
Function SetRetRef(strRet)
	Dim Index1,index2,Index3,Count1,Count2
	Dim IntIflg
	Dim strMessage
	Dim intstartRow,intEndRow, TempRow
	Dim comtemp1,comtemp2,temp
	Dim iInsRow
	
	Const C_PlantCd_Ref		= 0		' 공장 
	Const C_PlantNm_Ref		= 1
	Const C_ItemCd_Ref		= 2		' 품목 
	Const C_ItemNm_Ref		= 3
	Const C_MvmtQty_Ref		= 4		' 반품출고수량 
	Const C_TotRetQty_ref	= 5		' 재입고수량 
	Const C_IvQty_Ref		= 6 	' 매입수량 
	Const C_MvmtUnit_Ref	= 7 	' 단위 
	Const C_MvmtDt_Ref		= 8 	' 반품출고일 
	Const C_MvmtRcptNo_Ref	= 9 	' 반품출고번호 
	Const C_GmNo_Ref		= 10	' 재고처리번호 
	Const C_GmSeqNo_Ref		= 11	' 재고처리순번 
	Const C_TrackingNo_Ref  = 12	' Tracking No
	Const C_Lot_No_Ref		= 13	' Lot No.
	Const C_Lot_Seq_Ref		= 14	' Lot No. 순번 
	Const C_SpplSpec_Ref    = 15	' 품목규격 
	Const C_MvmtPrc_Ref 	= 16	' 발주단가 
	Const C_MvmtDocAmt_Ref	= 17
	Const C_SlCd_Ref		= 18	' 창고 
	Const C_SlNm_Ref		= 19
	Const C_Trackingflg_Ref = 20	' TRACKINGFLG
	Const C_RefMvmtNo_Ref	= 21	' 출고번호 
	Const C_ItemPrc_Ref		= 22	' 품목단가 

	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	strMessage = ""
	IntIflg=true
	
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		intStartRow = .vspdData.MaxRows + 1
		
		.vspdData.Redraw = False
		
		TempRow = .vspdData.MaxRows					'리스트 max값 
		
		For index1 = 0 to Count1
		
			If TempRow <> 0 Then

				For Index3 = 1 To TempRow				'같은 No가 있으면 Row를 추가하지 않는다.
					.vspdData.Row = Index3
					.vspdData.Col = C_RefMvmtNo
					if .vspdData.Text = strRet(index1,C_RefMvmtNo_Ref) then
						strMessage = strMessage & strRet(Index1,C_RefMvmtNo_Ref) & ";"
						intIflg=False
						Exit for
					End if 
				Next
			
			End If

			.vspdData.Row = Index1 + 1

			If IntIflg <> False then
				.vspdData.MaxRows = CLng(TempRow) + CLng(index1) + 1
				iInsRow = CLng(TempRow) + CLng(index1) + 1
	
				Call .vspdData.SetText(0		,	iInsRow, ggoSpread.InsertFlag)
				Call .vspdData.SetText(C_TrackingNo,	iInsRow, "*")
				Call .vspdData.SetText(C_LotNo,	iInsRow, "*")
	
'				Call SetState("C",iInsRow)

				Call .vspdData.SetText(C_PlantCd	,	iInsRow, strRet(index1,C_PlantCd_Ref))
				Call vspdData_Change(C_PlantCd, iInsRow)
				Call .vspdData.SetText(C_PlantNm	,	iInsRow, strRet(index1,C_PlantNm_Ref))
				Call .vspdData.SetText(C_itemCd		,	iInsRow, strRet(index1,C_ItemCd_Ref))
				Call vspdData_Change(C_itemCd, iInsRow)
				Call .vspdData.SetText(C_itemNm		,	iInsRow, strRet(index1,C_ItemNm_Ref))
				Call .vspdData.SetText(C_Spec		,	iInsRow, strRet(index1,C_SpplSpec_Ref))
				Call .vspdData.SetText(C_Unit		,	iInsRow, strRet(index1,C_MvmtUnit_Ref))
				Call .vspdData.SetText(C_GrQty		,	iInsRow, strRet(index1,C_MvmtQty_Ref))

				.vspdData.Row = iInsRow 	 
				.vspdData.Col = C_ProcurType
				If Trim(.vspdData.Text) <> "P" Then
					Call .vspdData.SetText(C_WorkPrc	,	iInsRow, strRet(index1,C_MvmtPrc_Ref))
				Else
					Call .vspdData.SetText(C_MvmtPrc	,	iInsRow, strRet(index1,C_MvmtPrc_Ref))
				End If

				
				Call .vspdData.SetText(C_SLCd		,	iInsRow, strRet(index1,C_SLCd_Ref))
				Call .vspdData.SetText(C_SLNm		,	iInsRow, strRet(index1,C_SLNm_Ref))
				Call .vspdData.SetText(C_LotNo		,	iInsRow, strRet(index1,C_Lot_No_Ref))
				Call .vspdData.SetText(C_LotNoSeq	,	iInsRow, strRet(index1,C_Lot_Seq_Ref))
				Call .vspdData.SetText(C_TrackingNo	,	iInsRow, strRet(index1,C_TrackingNo_Ref))
				Call .vspdData.SetText(C_RefMvmtNo	,	iInsRow, strRet(index1,C_RefMvmtNo_Ref))
					
				' 반품출고수량 - 재입고수량 
				IF strRet(index1,C_TotRetQty_ref) <> "" Then
					temp= UNICDbl(strRet(index1,C_MvmtQty_ref)) - UNICDbl(strRet(index1,C_TotRetQty_ref))
					Call .vspdData.SetText(C_GrQty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
				End If
				IF strRet(index1,C_IvQty_Ref) <> "" Then
					temp = UNICDbl(GetSpreadText(.vspdData,C_GrQty,iInsRow,"X","X")) - UNICDbl(strRet(index1,C_IvQty_Ref))
					Call .vspdData.SetText(C_GrQty,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
				End If
				Call vspdData_Change(C_GrQty, .vspddata.Row)
			Else
				IntIFlg=True
			End if 
		    
		Next
	
		intEndRow = .vspdData.MaxRows
	
		Call SetSpreadColorRef(intStartRow,intEndRow)
		
		if strMessage<>"" then
			Call DisplayMsgBox("17a005","X",strmessage,"입출고번호")
			.vspdData.ReDraw = True
			Exit Function
		End if

		.vspdData.ReDraw = True
	
	End with

	ggoOper.SetReqAttr	frm1.txtMvmtType, "Q"
	ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"

	Call SetToolBar("1110111100001111")
	
End Function

'================================== 2.2.5 SetSpreadColorRef() ==================================================
' Function Name : SetSpreadColorRef
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColorRef(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
	    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
	    ggoSpread.spreadlock		C_PlantCd	, pvStartRow, C_UnitPop,  pvEndRow
	    ggoSpread.SSSetProtected	C_Unit		, pvStartRow, pvEndRow
	    ggoSpread.spreadunlock		C_GrQty		, pvStartRow, C_GrQtyPop,  pvEndRow
	    ggoSpread.SSSetRequired		C_GrQty 	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_StockQty	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_Cur		, pvStartRow, pvEndRow
	    ggoSpread.spreadlock		C_CurPop	, pvStartRow, C_CurPop,  pvEndRow
	    ggoSpread.SSSetProtected	C_MvmtPrc	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_DocAmt	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_WorkPrc	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_WorkLocAmt, pvStartRow, pvEndRow
	    ggoSpread.spreadunlock		C_SlCd		, pvStartRow, C_SlcdPop,  pvEndRow
	    ggoSpread.SSSetRequired		C_SlCd 		, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_SlNm		, pvStartRow, pvEndRow
	    ggoSpread.spreadlock		C_RetTypeNm	, pvStartRow, C_InspResultNo,  pvEndRow
	    ggoSpread.SSSetProtected	C_RetTypeNm	, pvStartRow, pvEndRow
	    ggoSpread.spreadunlock		C_RemarkDtl , pvStartRow, C_RemarkDtl,  pvEndRow
	    ggoSpread.SSSetProtected	C_TrackingNo, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_GRNo      , pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_GRSeqNo   , pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_InspFlg   , pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_InspSts   , pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_GRMeth    , pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_InspReqNo , pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_InspResultNo, pvStartRow, pvEndRow
	
    End With
End Sub

'==============================================================================================================================
Function OpenMvmtNo()
	
		Dim strRet
		Dim arrParam(3)
		Dim iCalledAspName
		Dim IntRetCD
	
		If lblnWinEvent = True Or UCase(frm1.txtGrNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
		lblnWinEvent = True

		arrParam(0) = ""	' Trim(frm1.hdnSupplierCd.Value)
		arrParam(1) = ""	' Trim(frm1.hdnGroupCd.Value)
		arrParam(2) = ""	' Trim(frm1.hdnMvmtType.Value)		
		arrParam(3) = ""	' Rcpt flg , which must be "Y" or "N" or ""
				
		iCalledAspName = AskPRAspName("M4132PA1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4141PA1", "X")
			lblnWinEvent = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		lblnWinEvent = False
		If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
		If strRet(0) = "" Then
			frm1.txtGrNo.focus	
			Set gActiveElement = document.activeElement
			Exit Function
		Else
			frm1.txtGrNo.value = strRet(0)
			frm1.txtGrNo.focus	
			Set gActiveElement = document.activeElement
		End If	
		
End Function
'==============================================================================================================================
Function OpenGrType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "입출고유형"	
	arrParam(1) = "M_Mvmt_type A, (SELECT DISTINCT RCPT_TYPE, SUBCONTRA_FLG FROM M_CONFIG_PROCESS WHERE STO_FLG = 'N') B "
	
	arrParam(2) = Trim(frm1.txtMvmtType.Value)
	
	'arrParam(4) = "SUBCONTRA_FLG = " & FilterVar("N", "''", "S") & "AND RET_FLG = " & FilterVar("Y", "''", "S") & "AND USAGE_FLG = " & FilterVar("Y", "''", "S") & "  "
	arrParam(4) =  " A.IO_TYPE_CD = B.RCPT_TYPE AND A.SUBCONTRA_FLG = " & FilterVar("N", "''", "S")
	arrParam(4) = arrParam(4) &  " AND A.USAGE_FLG = " & FilterVar("Y", "''", "S") 
	arrParam(5) = "입출고유형"			
	
    arrField(0) = "ED12" & Chr(11) & "A.IO_TYPE_CD"
    arrField(1) = "ED20" & Chr(11) & "A.IO_TYPE_NM"
    arrField(2) = "ED10" & Chr(11) & "A.RET_FLG"
    arrField(3) = "ED12" & Chr(11) & "A.SUBCONTRA2_FLG"
    arrField(4) = "ED10" & Chr(11) & "A.RCPT_FLG"
    arrField(5) = "ED10" & Chr(11) & "B.SUBCONTRA_FLG"
    
    
    arrHeader(0) = "입출고유형"		
    arrHeader(1) = "입출고유형명"
    arrHeader(2) = "반품여부"		
    arrHeader(3) = "자품목정산여부"
    arrHeader(4) = "입고여부"
    arrHeader(5) = "외주가공여부"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else 
		frm1.txtMvmtType.Value	= arrRet(0)		
		frm1.txtMvmtTypeNm.Value= arrRet(1)
		Call changeMvmtType()
		lgBlnFlgChgValue = True
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function
'==============================================================================================================================
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSupplierCd.className)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"			
	arrParam(1) = "B_Biz_Partner"
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
	arrParam(3) = ""								
	
	'arrParam(4) = "Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "		'사외거래처만"	
	arrParam(4) = "Bp_Type <> " & FilterVar("C", "''", "S") & " AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "		'사외거래처만"	
	arrParam(5) = "공급처"			
	
    arrField(0) = "BP_CD"				
    arrField(1) = "BP_NM"
    arrField(2) = "GR_INSP_TYPE"

	arrHeader(0) = "공급처"		
	arrHeader(1) = "공급처명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		frm1.hdnGrInspType.Value = arrRet(2)
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
	
	arrParam(4) = "B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  "
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
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
	
	iCurRow = frm1.vspdData.ActiveRow

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,iCurRow,"X","X")))
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_PlantCd,	iCurRow, arrRet(0))		
		Call frm1.vspdData.SetText(C_PlantNm,	iCurRow, arrRet(1))		
		
		frm1.vspdData.Col = C_ItemCd
		
		If Trim(frm1.vspdData.Text) <> "" Then
			Call checkItemCd(arrRet(0), Trim(frm1.vspdData.Text), iCurRow)
		End If
	End If	
End Function
'==============================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(11)
	Dim iCalledAspName
	
	Dim strPlantCd, strItemCd, strQty

	If IsOpenPop = True Then Exit Function
	
	frm1.vspdData.Col = C_PlantCd	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 	 
	If Trim(frm1.vspdData.Text) = "" Then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Call SheetFocus(frm1.vspdData.ActiveRow, C_PlantCd)
		Exit Function
	End If

	IsOpenPop = True
	
	frm1.vspdData.Col = C_PlantCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	arrParam(0) = Trim(frm1.vspdData.Text)
	
	frm1.vspdData.Col = C_ItemCd
	arrParam(1) = Trim(frm1.vspdData.Text)

	
	If Trim(frm1.hdnSubcontra2flg.value) = "Y" Or Trim(frm1.hdnSubcontraflg.value) = "Y" Then
		arrParam(2) = "12!MO"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
		arrParam(3) = "20!O"
	Else
		arrParam(2) = "36!PP"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
		arrParam(3) = "30!P"
	End If
	
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명	
	arrField(2) = 3 ' -- 규격 
	arrField(3) = 4 ' -- 단위 
	arrField(7) = 9 ' -- 조달구분 
	arrField(8) = 14 ' -- Lot관리 
	arrField(9) = 15 ' -- 창고코드 
	arrField(10) = 25 ' -- Tracking관리 
	arrField(11) = 43 ' -- 검사품 여부 
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
		frm1.vspdData.Col = C_ItemCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_ItemNm
		frm1.vspdData.Text = arrRet(1)
		frm1.vspdData.Col = C_Spec
		frm1.vspdData.Text = arrRet(2)
		frm1.vspdData.Col = C_Unit
		frm1.vspdData.Text = arrRet(3)
		frm1.vspdData.Col = C_SlCd
		frm1.vspdData.Text = arrRet(9)
		frm1.vspdData.Col = C_Cur
		frm1.vspdData.Text = parent.gCurrency
		frm1.vspdData.Col = C_ProcurType
		frm1.vspdData.Text = arrRet(7)

		frm1.vspdData.Col = C_PlantCd
		strPlantCd = frm1.vspdData.Text
		frm1.vspdData.Col = C_ItemCd
		strItemCd = frm1.vspdData.Text

		Call changeByItem(strPlantCd, strItemCd, frm1.vspdData.ActiveRow)
	End If	
End Function
'==============================================================================================================================
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위"					
	arrParam(1) = "B_Unit_OF_MEASURE"		
	
	frm1.vspdData.Col = C_Unit
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
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
		frm1.vspdData.Col  = C_Unit
		frm1.vspdData.text = arrRet(0)	
		'Call ChangeReturnCost()
	End If	
End Function
'==============================================================================================================================
Function OpenGrQty()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5),arrField(6)
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = C_PlantCd
	If Trim(frm1.vspdData.text) = "" then 
		Call DisplayMsgBox("169901", "X", "X", "X")
		Call SheetFocus(frm1.vspdData.ActiveRow, C_PlantCd)
		Exit Function
	End If
	
	frm1.vspdData.Col = C_SlCd
	If Trim(frm1.vspdData.text) = "" then 
		Call DisplayMsgBox("169902", "X", "X", "X")
		Call SheetFocus(frm1.vspdData.ActiveRow, C_SlCd)
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	frm1.vspdData.Col = C_PlantCd
	arrParam(0) = Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_PlantNm
	arrParam(1) = Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_SlCd
	arrParam(2) = Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_SlNm
	arrParam(3) = Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_ItemCd
	arrParam(4) = Trim(frm1.vspdData.text)
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
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1211PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	Set gActiveElement = document.activeElement   

End Function
'==============================================================================================================================
Function OpenCurr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	arrParam(0) = "화폐"
	arrParam(1) = "B_Currency"

	frm1.vspdData.Col = C_Cur
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	arrParam(2) = Trim(frm1.vspdData.text)

	arrParam(4) = ""
	arrParam(5) = "화폐"
	arrField(0) = "Currency"
	arrField(1) = "Currency_Desc"
		    
	arrHeader(0) = "화폐"
	arrHeader(1) = "화폐명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col  = C_Cur
		frm1.vspdData.text = arrRet(0)	
	End If

End Function
'==============================================================================================================================
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
	
	iCurRow = frm1.vspdData.ActiveRow		
	
	If IsOpenPop = True Then Exit Function

	frm1.vspdData.Col = C_PlantCd	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 	 
	If Trim(frm1.vspdData.Text) = "" Then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Call SheetFocus(frm1.vspdData.ActiveRow, C_PlantCd)
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "창고"
	arrParam(1) = "B_STORAGE_LOCATION"
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_SLCd,iCurRow,"X","X")))
	arrParam(4) = "PLANT_CD= " & FilterVar(UCase(GetSpreadText(frm1.vspdData,C_PlantCd,iCurRow,"X","X")), "''", "S") & " "
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
		Call frm1.vspdData.SetText(C_SLCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_SLNm,	iCurRow, arrRet(1))
	End If	
End Function
'==============================================================================================================================
Function OpenLotNo()
	Dim arrRet
	Dim arrParam(5),arrField(6)
	Dim IntRetCD
	Dim iCalledAspName

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	frm1.vspdData.Col = C_PlantCd
	arrParam(0) = frm1.vspdData.Text
	
	frm1.vspdData.Col = C_PlantNm
	arrParam(1) = frm1.vspdData.Text
	
	If arrParam(0) = "" Then	
		Call DisplayMsgBox("169901","X", "X", "X")   <% '공장정보가 필요합니다 %>
		Call SheetFocus(frm1.vspdData.ActiveRow, C_PlantCd)
		Exit Function
	End If		
		
    frm1.vspdData.Col = C_SLCd
	arrParam(2) = frm1.vspdData.Text
		
	frm1.vspdData.Col = C_SLNm
	arrParam(3) = frm1.vspdData.Text
	
	frm1.vspdData.Col = C_ItemCd
	arrParam(4) = frm1.vspdData.Text
	
	If arrParam(4) = "" Then
		Call DisplayMsgBox("169915","X", "X", "X")   <% '품목코드를 입력하십시오 %>
		Call SheetFocus(frm1.vspdData.ActiveRow, C_ItemCd)
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
			Call .SetText(C_LotNo, .ActiveRow, arrRet(5))
			Call .SetText(C_LotNoSeq, .ActiveRow, arrRet(6))
		End With
	End If	
End Function
'==============================================================================================================================
Function OpenRet()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function 

	IsOpenPop = True
 
    frm1.vspdData.Col=C_RetType 
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 

	arrParam(0) = "반품유형"				
	arrParam(1) = "B_MINOR"	
	
	arrParam(2) = Trim(frm1.vspdData.Text)		
		
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9017", "''", "S") & " "	
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
			.vspdData.Col = C_RetType 
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_RetTypeNm
			.vspdData.Text = arrRet(1)
	End With
	Call vspdData_Change(C_RetType , frm1.vspdData.ActiveRow)   
	End If	
End Function
'==============================================================================================================================
Function OpenTrackingNo()

	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	frm1.vspdData.Col = C_PlantCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	If Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		IsOpenPop = False
		Exit Function
	End if
    
    arrParam(0) = ""
    arrParam(1) = ""
    arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	
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
		frm1.vspdData.Col = C_TrackingNo
		frm1.vspdData.Text = arrRet
	End If	

End Function
'==============================================================================================================================
Sub Form_Load()
	    
    Call LoadInfTB19029                                                  
    Call ggoOper.LockField(Document, "N")                                
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                 
    Call SetDefaultVal
    Call InitVariables
	
End Sub
'==============================================================================================================================
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
'==============================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then Exit Sub
    
    If frm1.vspdData.MaxRows = 0 Then Exit Sub
End Sub
'==============================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'==============================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==============================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'==============================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'==============================================================================================================================
Sub PopRestoreSpreadColumnInf()
    
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
    Call ChangeTag(True)
End Sub
'==============================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    
    frm1.vspdData.Row = Row
	frm1.vspdData.Col = Col
    
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)
    
    Select Case	Col
    	Case C_PlantCd
    		Call changePlantCd(Row, Trim(frm1.vspdData.Text))
    	Case C_ItemCd
    		Call changeItemCd(Row, Trim(frm1.vspdData.Text))
    	Case C_GrQty
    		Call changeQty(Row, Trim(frm1.vspdData.Text))
    	Case C_MvmtPrc, C_WorkPrc
    		Call changeAmt(Row, Col, Trim(frm1.vspdData.Text))
    	Case C_SlCd
    		Call changeSlCd(Row, Trim(frm1.vspdData.Text))
    End Select
    
End Sub
'==============================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData

		Select Case Col
			Case C_PlantPop			' 공장 
				Call OpenPlant()
			Case C_ItemPop			' 품목 
				Call OpenItem()
			Case C_UnitPop			' 단위 
				Call OpenUnit()
			Case C_GrQtyPop			' 수량 
				Call OpenGrQty()
			Case C_CurPop			' 화폐 
				Call OpenCurr()
			Case C_SlCdPop			' 창고 
				Call OpenSL()
			Case C_LotNoPop			' Lot No.
				Call OpenLotNo()
			Case C_RetTypePop
			    Call OpenRet()	
			Case C_TrackingNoPop	'Tracking NO
				Call OpenTrackingNo()
		End Select   
    End With
End Sub
'==============================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
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
'==============================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    Err.Clear                                               
	On Error Resume Next                                   
	
	ggoSpread.Source = frm1.vspdData
	
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    
    If Not chkField(Document, "1") Then	Exit Function
    If DbQuery = False Then Exit Function
       
    FncQuery = True											
    Set gActiveElement = document.ActiveElement   
    
End Function
'==============================================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    On Error Resume Next                                   
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ChangeTag(False)
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call SetDefaultVal
    Call InitVariables
    
    FncNew = True                                           
	Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncDelete() 
    
	Dim IntRetCD
	
	On Error Resume Next 
    Err.Clear                                               
    
    FncDelete = False
    
    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function
    														
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then            
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    FncDelete = True                            
    Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
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
    If Not chkField(Document, "2") Then  Exit Function
 
    
    ggoSpread.Source = frm1.vspdData                 
    If Not ggoSpread.SSDefaultCheck Then  Exit Function
    If frm1.vspdData.Maxrows < 1 then Exit Function

    For intIndex = 1 to frm1.vspdData.MaxCols 
		frm1.vspdData.SetColItemData intindex,0
	Next
	
    If DbSave = False Then Exit Function
    
    FncSave = True                                                       
    Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function FncCopy() 
	On Error Resume Next                                                          '☜: If process fails 
End Function
'==============================================================================================================================
Function FncCancel()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear    

	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo

    If frm1.vspdData.Maxrows < 1 then
		Call SetToolBar("1110010000001111")
		'Call ggoOper.LockField(Document, "N")
		ggoOper.SetReqAttr	frm1.txtMvmtType, "N"
		ggoOper.SetReqAttr	frm1.txtSupplierCd, "N"
    End If
    
    Set gActiveElement = document.ActiveElement                                                  
End Function
'==============================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow

	On Error Resume Next
	Err.Clear

	' 상단입력체크 
	If Not chkField(Document, "2") Then  Exit Function
	
	'Call ggoOper.LockField(Document, "Q")
	ggoOper.SetReqAttr	frm1.txtMvmtType, "Q"
	ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"
	
	FncInsertRow = False
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End IF
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow, imRow
		'Call SetSpreadLock()
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
	End With

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_PlantCd
	frm1.vspdData.Text = Parent.gPlant
	frm1.vspdData.Col = C_PlantNm
	frm1.vspdData.Text = Parent.gPlantNm

	Call SetToolBar("1110111100001111")	
	
    If Err.number = 0 Then FncInsertRow = True
	Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear  
    
    ggoSpread.Source = frm1.vspdData
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    	.focus
		ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncPrint()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear  
    
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncExcel()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear
	
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_SINGLEMULTI)	
    
    Set gActiveElement = document.ActiveElement							
End Function
'==============================================================================================================================
Function FncFind() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
    
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False) 
    Set gActiveElement = document.ActiveElement                                 
End Function
'==============================================================================================================================
Function FncExit()
	Dim IntRetCD
	
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
	
	FncExit = False
	ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")           
		If IntRetCD = vbNo Then Exit Function
	End If
    
    FncExit = True
    Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function DbDelete() 
    Dim strVal
    
    On Error Resume Next       
    Err.Clear                                                           
	
	DbDelete = False													
    frm1.txtMode.value = Parent.UID_M0003
    
    If LayerShowHide(1) = False Then Exit Function
    Call ExecMyBizASP(frm1, BIZ_PGM_ID)							
    DbDelete = True                                             
	Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function DbDeleteOk()											
	lgBlnFlgChgValue = False
	Call MainNew()
End Function
'==============================================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey     
    Dim strVal
        
    DbQuery = False
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
    
    If LayerShowHide(1) = False Then Exit Function
    
	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtGrNo=" & .hdnGrNo.value
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtGrNo=" & Trim(.txtGrNo.value)
		End if
  
		Call RunMyBizASP(MyBizASP, strVal)							
    End With
    
    DbQuery = True
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function DbQueryOk()											
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear
    
    lgIntFlgMode = Parent.OPMD_UMODE									
    
    Call ggoOper.LockField(Document, "Q")						
	Call ChangeTag(True)
	lgBlnFlgChgValue = False	
	Call SetToolBar("11101011000111")
	
	Call RemovedivTextArea
	
	if interface_Account = "N" then		
		frm1.btnGlSel.disabled = true
	Else 
		frm1.btnGlSel.disabled = False		
	End if
	frm1.vspdData.focus

End Function
'==============================================================================================================================
Function DbSave() 
	'On Error Resume Next                                                          '☜: If process fails
    Err.Clear  
    Dim lRow        
    Dim strVal, strDel
	Dim iColSep, iRowSep
	
	Dim strCUTotalvalLen 	'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  	'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim objTEXTAREA 		'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
	
    DbSave = False                                                          
    
	Call DisableToolBar(Parent.TBC_SAVE)                                          '☜: Disable Save Button Of ToolBar

    If LayerShowHide(1) = False Then
		Exit Function
	End If 

    iColSep = Parent.gColSep													
	iRowSep = Parent.gRowSep													
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]
	
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	
	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode
	
	strVal = ""
	strDel = ""

	With frm1
		For lRow = 1 To .vspdData.MaxRows
 
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
  
		        Case ggoSpread.InsertFlag
					If Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"))) = "" Or Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"))) = "0" then
						Call DisplayMsgBox("970021","X","수량","X")
						Call RemovedivTextArea
						Call LayerShowHide(0)
						Call SheetFocus(lRow, C_GrQty)
						.vspdData.EditMode = True
						Exit Function
					End if

					strVal = "C" & iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Unit,lRow, "X","X"))					& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Cur,lRow, "X","X"))					& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_MvmtPrc,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_DocAmt,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_WorkPrc,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_WorkLocAmt,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SlCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotNoSeq,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MakerLotNo,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MakerLotSeqNo,lRow, "X","X"))			& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RetType,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_TrackingNo,lRow, "X","X"))				& iColSep 
					'strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_InspFlg,lRow, "X","X"))				& iColSep
					If Trim(GetSpreadText(frm1.vspdData,C_InspFlg,lRow, "X","X")) = "0" Then
						strVal = strVal & "N"	& iColSep
					Else
						strVal = strVal & "Y"	& iColSep
					End If
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GRMeth,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_InspReqNo,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_InspResultNo,lRow, "X","X"))			& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RefMvmtNo,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProcurType,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(frm1.hdnGrInspType.value)											& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RemarkDtl,lRow, "X","X"))				& iColSep 
					strVal = strVal & lRow & iRowSep
				Case ggoSpread.DeleteFlag
					strDel = "D" & iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X"))					& iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_RefMvmtNo,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"))					& iColSep 
					strDel = strDel & lRow & iRowSep
		   	End Select 
		
			.vspdData.Row = lRow
			.vspdData.Col = 0
			Select Case .vspdData.Text
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
	End With
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA   = document.createElement("TEXTAREA")
	   objTEXTAREA.name  = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA   = document.createElement("TEXTAREA")
	   objTEXTAREA.name  = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If

	'------ Developer Coding part (End ) -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	If Err.number = 0 Then	 
	   DbSave = True                                                             '☜: Processing is OK
	End If

	Set gActiveElement = document.ActiveElement         
End Function
'==============================================================================================================================
Function DbSaveOk()												

	Call InitVariables
	Call ChangeTag(true)
	Call MainQuery()
	
End Function
'==============================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'==============================================================================================================================
Function RemovedivTextArea()
	Dim ii
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function
'==============================================================================================================================
Function changeMvmtType()
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                            
  	If CheckRunningBizProcess = True Then
		Exit Function
	End If
	
	If Trim(frm1.txtMvmtType.Value) = "" Then
		frm1.txtMvmtTypeNm.Value = ""
		Exit Function
	End If
	
    changeMvmtType = False               
    
    If LayerShowHide(1) = False Then
         Exit Function
    End If  
    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeMvmtType"
    strVal = strVal & "&txtMvmtType=" & Trim(frm1.txtMvmtType.Value)

    Call RunMyBizASP(MyBizASP, strVal)
	
	lgBlnFlgChgValue = true

    changeMvmtType = True                
End Function
'==============================================================================================================================
Function changeSpplCd()
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                        
    If CheckRunningBizProcess = True Then
		Exit Function
	End If 
    changeSpplCd = False           
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeSpplCd"
    strVal = strVal & "&txtSupplierCd=" & FilterVar(Trim(frm1.txtSupplierCd.Value),"","SNM")
    
    Call RunMyBizASP(MyBizASP, strVal)
	
	lgBlnFlgChgValue = true
    
    changeSpplCd = True            

End Function
'==============================================================================================================================
Function changeGroupCd()
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                        
    If CheckRunningBizProcess = True Then
		Exit Function
	End If 
    changeGroupCd = False           

	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeGroupCd"
    strVal = strVal & "&txtGroupCd=" & FilterVar(Trim(frm1.txtGroupCd.Value),"","SNM")
    
    Call RunMyBizASP(MyBizASP, strVal)
	
	lgBlnFlgChgValue = true
    
    changeGroupCd = True            

End Function
'==============================================================================================================================
Function changePlantCd(Row, strPlantCd)
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_PlantNm

	If 	CommonQueryRs(" PLANT_NM ", " B_PLANT ", _
							 " PLANT_CD = " & FilterVar(strPlantCd, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("125000","X",strPlantCd,"X")
		Call frm1.vspdData.SetText(C_PlantCd, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_PlantNm, frm1.vspdData.Row, "")
		Exit Function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.vspdData.Text = lgF0(0)

	frm1.vspdData.Col = C_ItemCd
	
	If Trim(frm1.vspdData.Text) <> "" Then
		If checkItemCd(strPlantCd, Trim(frm1.vspdData.Text), Row) Then
			Exit Function
		End If
	End If
End Function
'==============================================================================================================================
Function checkItemCd(strPlantCd, strItemCd, Row)
	checkItemCd = True

	frm1.vspdData.Row = Row

	If 	CommonQueryRs(" ITEM_CD ", " B_ITEM_BY_PLANT ", _
							 " PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") , _
							 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("122729","X","X","X")
		Call frm1.vspdData.SetText(C_ItemCd, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_ItemNm, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_Spec, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_Unit, frm1.vspdData.Row, "")
		Exit Function
	End If

	checkItemCd = False
End Function
'==============================================================================================================================
Function changeItemCd(Row, strItemCd)
	Dim strPlantCd

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_PlantCd
	If Trim(frm1.vspdData.Text) = "" Then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Call frm1.vspdData.SetText(C_ItemCd, frm1.vspdData.Row, "")
		Exit Function
	Else
		strPlantCd = Trim(frm1.vspdData.Text)
	End If

	If checkItemCd(strPlantCd, strItemCd, Row) Then
		Exit Function
	End If

	If 	CommonQueryRs(" B.ITEM_NM, B.SPEC, B.BASIC_UNIT, A.MAJOR_SL_CD, A.PROCUR_TYPE ", " B_ITEM_BY_PLANT A, B_ITEM B ", _
							 " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND A.ITEM_CD = " & FilterVar(strItemCd, "''", "S") , _
							 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("122729","X","X","X")
		'Call frm1.vspdData.SetText(C_ItemCd, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_ItemNm, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_Spec, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_Unit, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_SlCd, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_ProcurType, frm1.vspdData.Row, "")
		Exit Function
	End If
	lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1, Chr(11))
	lgF2 = Split(lgF2, Chr(11))
	lgF3 = Split(lgF3, Chr(11))
	lgF4 = Split(lgF4, Chr(11))

	Call frm1.vspdData.SetText(C_ItemNm, frm1.vspdData.Row, lgF0(0))
	Call frm1.vspdData.SetText(C_Spec, frm1.vspdData.Row, lgF1(0))
	Call frm1.vspdData.SetText(C_Unit, frm1.vspdData.Row, lgF2(0))
	Call frm1.vspdData.SetText(C_Cur, frm1.vspdData.Row, parent.gCurrency)
	Call frm1.vspdData.SetText(C_SlCd, frm1.vspdData.Row, lgF3(0))
	Call frm1.vspdData.SetText(C_ProcurType, frm1.vspdData.Row, lgF4(0))

	Call changeByItem(strPlantcd, strItemCd, Row)

End Function

'==========================================   ChangeItemPlant2()  ======================================
'	Name : ChangeItemPlant2()
'	[2005/09/16 Sim Hae Young Add Sub]
'=========================================================================================================
Sub ChangeItemPlant2(lRow)

	Dim lgF2By2
	Dim arrVal1
	Dim arrVal2

	Dim iStrSelect
	Dim iStrSql

	Dim iOrderUnitArr
	Dim iOrderUnitArr2
	Dim iOrderUnitArr3
	Dim iSLCdArr
	Dim iSLNmArr
	Dim iItemNmArr
	Dim iSpecArr
	Dim iHSCdArr
	Dim iHSNmArr
	Dim iPlantNmArr
	Dim iProcur_type
	Dim iTracking_Flg
	Dim iUnder_Tol
	Dim iOver_Tol

	Err.Clear

	iStrSelect = ""
	iStrSelect = " B.PUR_UNIT, A.ORDER_UNIT_PUR, C.BASIC_UNIT "

	iStrSql =""
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT S.ITEM_CD,S.ORDER_UNIT_PUR "
	iStrSql = iStrSql & " 	FROM B_ITEM_BY_PLANT S LEFT OUTER JOIN B_STORAGE_LOCATION T ON(S.MAJOR_SL_CD=T.SL_CD AND S.PLANT_CD=T.PLANT_CD) "
	iStrSql = iStrSql & " WHERE S.PLANT_CD=" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND S.ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND S.VALID_FROM_DT <= GETDATE() AND S.VALID_TO_DT >= GETDATE() "
	iStrSql = iStrSql & " ) A "
	iStrSql = iStrSql & " LEFT OUTER JOIN  "
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT ITEM_CD,PUR_UNIT "
	iStrSql = iStrSql & " 	FROM M_SUPPLIER_ITEM_BY_PLANT "
	iStrSql = iStrSql & " WHERE PLANT_CD=" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND BP_CD  =" & FilterVar(Trim(frm1.txtSupplierCd.value), "''" , "S") 
	iStrSql = iStrSql & " 	AND VALID_FR_DT <= GETDATE() AND VALID_TO_DT >= GETDATE() "
	iStrSql = iStrSql & " ) B "
	iStrSql = iStrSql & " ON(A.ITEM_CD=B.ITEM_CD)  "
	iStrSql = iStrSql & " LEFT OUTER JOIN  "
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT ITEM_CD,BASIC_UNIT "
	iStrSql = iStrSql & " FROM B_ITEM  "
	iStrSql = iStrSql & " WHERE ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " ) C "
	iStrSql = iStrSql & " ON(A.ITEM_CD=C.ITEM_CD)  "
   
	If CommonQueryRs2by2(iStrSelect, iStrSql, , lgF2By2)= False Then
		Call DisplayMsgBox("122700","X","X","X")
		Err.Clear

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_itemCd
		frm1.vspdData.text = ""
		Exit Sub
	End If

	arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))

	arrVal2 = Split(arrVal1(0), chr(11))

	iOrderUnitArr  	= Trim(arrVal2(1))
	iOrderUnitArr2	= Trim(arrVal2(2))
	iOrderUnitArr3	= Trim(arrVal2(3))

	With frm1.vspdData
		.Row = lRow

		.Col = C_Unit
		If Trim(iOrderUnitArr)<>"" Then
			.text = Trim(iOrderUnitArr)
		Else
			If Trim(iOrderUnitArr2)<>"" Then
				.text = Trim(iOrderUnitArr2)
			Else
				.text = Trim(iOrderUnitArr3)
			End If
		End If

	End With

End Sub

'==========================================   lookupPrice()  ======================================
'	Name : lookupPrice()
'	Description :
'==================================================================================================
Function lookupPrice(ByVal Row)

    Err.Clear

    If CheckRunningBizProcess = True Then
		Exit Function
	End If

    Dim strVal

	lgBlnFlgChgValue = true

	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	' === 2005.07.15 단가관련 수정 =================
	frm1.vspdData.Col = C_ItemCd
	If Trim(frm1.vspdData.text) = "" Then
		Call DisplayMsgBox("169915","X","X","X")
		Call LayerShowHide(0)
		Exit Function
	End If
	' === 2005.07.15 단가관련 수정 =================
    

    strVal = BIZ_PGM_ID & "?txtMode=" & "lookupPrice"
    strVal = strVal & "&txtStampDt=" & Trim(frm1.txtGmDt.text)
    strVal = strVal & "&txtBpCd=" & Trim(frm1.txtSupplierCd.Value)
	frm1.vspdData.Col = C_itemCd
    strVal = strVal & "&txtItemCd=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_PlantCd
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_Unit
    strVal = strVal & "&txtUnit=" & Trim(frm1.vspdData.text)
    frm1.vspdData.Col = C_Cur
    strVal = strVal & "&txtCurrency=" & Trim(frm1.vspdData.text)
    strVal = strVal & "&txtRow=" & Row

    If LayerShowHide(1) = False Then Exit Function
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'==============================================================================================================================
Function changeByItem(strPlantCd, strItemCd, Row)
	Dim Procur_Type
	Dim Lot_Flg
	Dim Major_Sl_Cd
	Dim Tracking_Flg
	Dim Recv_Inspec_flg
	Dim Rcpt_flg
	Dim Ret_flg
	Dim Item_acct
	Dim Item_acct_grp
	
	Dim str_prc_ctrl_indctr,mvmt_prc,strStdFlg

	frm1.vspdData.Row = Row

	If 	CommonQueryRs(" A.PROCUR_TYPE, A.LOT_FLG, A.MAJOR_SL_CD, A.TRACKING_FLG, A.RECV_INSPEC_FLG, A.ITEM_ACCT, dbo.ufn_GetItemAcctGrp(A.ITEM_ACCT) ITEM_ACCT_GRP ", _
							" B_ITEM_BY_PLANT A, B_ITEM_ACCT_INF B", _
							" A.ITEM_ACCT = B.ITEM_ACCT AND A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND A.ITEM_CD = " & FilterVar(strItemCd, "''", "S") , _
							lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("122729","X","X","X")
	End If

	lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1, Chr(11))
	lgF2 = Split(lgF2, Chr(11))
	lgF3 = Split(lgF3, Chr(11))
	lgF4 = Split(lgF4, Chr(11))
	lgF5 = Split(lgF5, Chr(11))
	lgF6 = Split(lgF6, Chr(11))

	Procur_Type		= Trim(lgF0(0))	' 조달구분 
	Lot_Flg			= Trim(lgF1(0))	' Lot구분 
	Major_Sl_Cd		= Trim(lgF2(0))	' 창고 
	Tracking_Flg	= Trim(lgF3(0))	' Tracking 구분 
	Recv_Inspec_flg	= Trim(lgF4(0))	' 검사품 
	Item_acct		= Trim(lgF5(0))	' 품목계정 
	Item_acct_grp	= Trim(lgF6(0))	' 품목계정 그룹 
	
	
	If Trim(frm1.hdnSubcontra2flg.Value) = "Y" Or Trim(frm1.hdnSubcontraflg.value) = "Y" Then
		If  (Trim(Procur_Type) < "M" Or Trim(Procur_Type) > "O") Or (Trim(Item_acct_grp) < "1" Or Trim(Item_acct_grp) > "3") Then
			
			Call DisplayMsgBox("179019","X","X","X")
			frm1.vspddata.Col = C_itemCd
			frm1.vspddata.text = ""
			frm1.vspddata.Col = C_ItemNm
			frm1.vspddata.text = ""
			frm1.vspddata.Col = C_Spec
			frm1.vspddata.text = ""
			Exit Function
		End If
	Else
		If Trim(Procur_Type) <> "P" Or (Trim(Item_acct_grp) < "3" Or Trim(Item_acct_grp) > "7") Then
			
			Call DisplayMsgBox("179019","X","X","X")
			frm1.vspddata.Col = C_itemCd
			frm1.vspddata.text = ""
			frm1.vspddata.Col = C_ItemNm
			frm1.vspddata.text = ""
			frm1.vspddata.Col = C_Spec
			frm1.vspddata.text = ""
			Exit Function
		End If
	End If


	' 품목단가 
	'=====================================================
	'단가정책에 따른 표준단가 반영 
	'2007.05.17 Modified by KSJ
	'=====================================================
	If 	CommonQueryRs(" PRC_CTRL_INDCTR, STD_PRC ", " I_MATERIAL_VALUATION ", _
							 " PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") _
							 ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	Else

		lgF0 = Split(lgF0, Chr(11))  
		lgF1 = Split(lgF1, Chr(11)) 
		
		str_prc_ctrl_indctr = Trim(lgF0(0))    '단가구분 
		mvmt_prc            = Trim(lgF1(0))    '표준단가 
		
		If UCase(Trim(str_prc_ctrl_indctr)) = "S" Then   '단가구분이 표준단가이면 
		
				If 	CommonQueryRs(" TOP 1 A.MINOR_CD ", " B_MINOR A, B_CONFIGURATION B ", " A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD  AND A.MAJOR_CD = " & FilterVar("M4105", "''", "S") & " AND B.SEQ_NO = 1 AND B.REFERENCE = " & FilterVar("Y", "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				    '데이타가 없으면 표준단가반영 
				    frm1.vspdData.Col = C_MvmtPrc
				    frm1.vspdData.Text = mvmt_prc
				Else

				    '진단가 최신단가 구분하여 가져오도록 변경 
					Call ChangeItemPlant2(frm1.vspdData.Row)
					Call lookupPrice(frm1.vspdData.Row)

				end if
							 
		else
			'진단가 최신단가 구분하여 가져오도록 변경 
			Call ChangeItemPlant2(frm1.vspdData.Row)
			Call lookupPrice(frm1.vspdData.Row)
		end if
		
	End If
	
	'If CommonQueryRs(" CASE MOVING_AVG_PRC WHEN 0 THEN STD_PRC ELSE MOVING_AVG_PRC END ", " I_MATERIAL_VALUATION ", _
	'						 " PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") _
	'						 ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	'Else
	'	lgF0 = Split(lgF0, Chr(11))
	'	frm1.vspdData.Col = C_MvmtPrc
	'	frm1.vspdData.Text = lgF0(0)
	'End If

	frm1.vspdData.Col = C_MvmtPrc
	'=====================================================
	
	' 품목금액 
	Call changeAmt(frm1.vspdData.Row, C_MvmtPrc, frm1.vspdData.Text)

	' 창고명 
	frm1.vspdData.Col = C_SlCd
	frm1.vspdData.Text = Major_Sl_Cd
	Call changeSlCd(frm1.vspdData.Row, Major_Sl_Cd)

	' 조달구분 
	If UCase(Trim(Procur_Type)) = "P" Then
		frm1.vspdData.Col = C_WorkPrc
		frm1.vspdData.Text = 0
		frm1.vspdData.Col = C_WorkLocAmt
		frm1.vspdData.Text = 0
		
		ggoSpread.SSSetProtected	C_WorkPrc, Row, Row
	Else
		'구매품이 아닐경우에는 단가 입력 protect
		'2007.06.19 Modified by KSJ
		ggoSpread.SSSetProtected	C_MvmtPrc, Row, Row
		
		ggoSpread.SpreadUnLock 		C_WorkPrc, Row, C_WorkPrc, Row
		ggoSpread.SSSetRequired		C_WorkPrc, Row, Row
	End If

	' Lot
	If UCase(Trim(Lot_Flg)) = "N" Then
		frm1.vspdData.Col = C_LotNo
		frm1.vspdData.Text = "*"
		frm1.vspdData.Col = C_LotNoSeq
		frm1.vspdData.Text = 0

		ggoSpread.SSSetProtected	C_LotNo, Row, Row
		ggoSpread.SSSetProtected	C_LotNoPop, Row, Row
		ggoSpread.SSSetProtected	C_LotNoSeq, Row, Row
	Else
		If 	CommonQueryRs(" RCPT_FLG, RET_FLG ", " M_MVMT_TYPE ", _
								 " IO_TYPE_CD = " & FilterVar(Trim(frm1.txtMvmtType.Value), "''", "S") _
								 ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Else
			lgF0 = Split(lgF0, Chr(11))
			Rcpt_flg = lgF0(0)
			lgF1 = Split(lgF1, Chr(11))
			Ret_flg = lgF1(0)
		End If

		If Trim(Rcpt_flg) = "N" And Trim(Ret_flg) = "Y" Then
			ggoSpread.SpreadUnLock 		C_LotNo, Row, C_LotNoSeq, Row
			ggoSpread.SSSetRequired		C_LotNo, Row, Row
			ggoSpread.SSSetRequired		C_LotNoSeq, Row, Row
		Else
			frm1.vspdData.Col = C_LotNo
			frm1.vspdData.Text = ""
			frm1.vspdData.Col = C_LotNoSeq
			frm1.vspdData.Text = 0
	
			' 자동채번,수동채번 코드 
			If 	CommonQueryRs(" LOT_GEN_MTHD ", " B_LOT_CONTROL ", _
									 " PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") _
									 ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Else
				lgF0 = Split(lgF0, Chr(11))
				
				If Trim(lgF0(0)) = "A" Then
					ggoSpread.SSSetProtected	C_LotNo, Row, Row
					ggoSpread.SSSetProtected	C_LotNoPop, Row, Row
					ggoSpread.SSSetProtected	C_LotNoSeq, Row, Row
				Else
					ggoSpread.SpreadUnLock 		C_LotNo, Row, C_LotNoSeq, Row
					ggoSpread.SSSetRequired		C_LotNo, Row, Row
					ggoSpread.SSSetRequired		C_LotNoSeq, Row, Row
				End If
			End If
		End If
	End If
	
	' Tracking
	If UCase(Trim(Tracking_Flg)) = "Y" Then
		ggoSpread.spreadUnlock 		C_TrackingNo, Row, C_TrackingNoPop, Row
		ggoSpread.SSSetRequired		C_TrackingNo, Row, Row	
	Else
		frm1.vspdData.Col = C_TrackingNo
		frm1.vspdData.Text = "*"
		ggoSpread.SSSetProtected	C_TrackingNo, Row, Row
		ggoSpread.SSSetProtected	C_TrackingNoPop, Row, Row	
	End If

	' 검사품 여부 
	If UCase(Trim(Recv_Inspec_flg)) = "Y" Then
		Call frm1.vspdData.SetText(C_InspFlg, Row, "1")
		'If 	CommonQueryRs(" MINOR_NM ", " B_MINOR ", _
		'						 " MAJOR_CD = " & FilterVar("B9016", "''", "S") & " AND MINOR_CD = " & FilterVar(frm1.hdnGrInspType.Value, "''", "S") _
		'						 ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		'Else
		'	lgF0 = Split(lgF0, Chr(11))
			'frm1.vspdData.Col = C_GRMeth
			'frm1.vspdData.Text = Trim(lgF0(0)) 히든 필드이므로 명을 보여줄필요 없음 
			 
		'End If
		
			frm1.vspdData.Col = C_GRMeth
			frm1.vspdData.Text = Trim(frm1.hdnGrInspType.Value)
	Else
		Call frm1.vspdData.SetText(C_InspFlg, Row, "0")
	End If

End Function
'==============================================================================================================================
Function changeQty(Row, strQty)
	Dim strAmt

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_MvmtPrc
	strAmt = Trim(frm1.vspdData.Text)
	
	frm1.vspdData.Col = C_DocAmt
	frm1.vspdData.Text = CDbl(strQty) * CDbl(strAmt)

	frm1.vspdData.Col = C_WorkPrc
	strAmt = Trim(frm1.vspdData.Text)
	
	frm1.vspdData.Col = C_WorkLocAmt
	frm1.vspdData.Text = CDbl(strQty) * CDbl(strAmt)
	
End Function
'==============================================================================================================================
Function changeAmt(Row, Col, strAmt)
	Dim strQty

	frm1.vspdData.Row = Row
	
	frm1.vspdData.Col = C_GrQty	

	strQty = Trim(frm1.vspdData.Text)
	
	frm1.vspdData.Col = Col + 1
	frm1.vspdData.Text = CDbl(strQty) * CDbl(strAmt)
	
End Function
'==============================================================================================================================
Function changeSlCd(Row, strSlCd)
	Dim strPlantCd

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_PlantCd	
	If Trim(frm1.vspdData.Text) = "" Then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Exit Function
	Else
		strPlantCd = Trim(frm1.vspdData.Text)
	End If

	If 	CommonQueryRs(" SL_NM ", " B_STORAGE_LOCATION ", _
							 " PLANT_CD = " & FilterVar(strPlantCd, "''", "S") & " AND SL_CD = " & FilterVar(strSlCd, "''", "S") , _
							 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("125710","X","X","X")
		Call frm1.vspdData.SetText(C_SlCd, frm1.vspdData.Row, "")
		Call frm1.vspdData.SetText(C_SlNm, frm1.vspdData.Row, "")
		Exit Function
	End If

	frm1.vspdData.Col = C_SlNm
	lgF0 = Split(lgF0, Chr(11))
	frm1.vspdData.Text = lgF0(0)

End Function
'==============================================================================================================================
Sub txtGmDt_DblClick(Button)
	If Button = 1 then
		frm1.txtGmDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtGmDt.focus
	End if
End Sub
'==============================================================================================================================
Sub txtGmDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'==============================================================================================================================

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function
</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	

</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>예외입고/반품등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenRetRef">예외반품출고참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>입출고번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입출고번호" NAME="txtGrNo" MAXLENGTH=18 SIZE=32 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGmNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()"></TD>
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
								<TD CLASS="TD5" NOWRAP>입출고유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입출고유형" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="23NXXU" OnChange="VBScript:changeMvmtType()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGrType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="입출고유형" NAME="txtMvmtTypeNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>입출고일</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=입출고일 NAME="txtGmDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="23N1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="23XXXU" OnChange="VBScript:changeSpplCd()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="24X"></TD>	
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="23XXXU" OnChange="VBScript:changeGroupCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupNm" SIZE=20 tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>입출고번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입출고번호" NAME="txtGrNo1" SIZE=34 MAXLENGTH=18 tag="21XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT ALT="비고" NAME="txtRemark" MAXLENGTH=120 SIZE=34 tag="21XXXU"></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"  TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGrNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGrInspType" tag="24" TabIndex="-1">

<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnSubcontra2flg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24" TabIndex="-1">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
