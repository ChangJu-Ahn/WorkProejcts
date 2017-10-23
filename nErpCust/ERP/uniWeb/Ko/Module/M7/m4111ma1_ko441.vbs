Option Explicit					
' 코딩시 주의할점 
' hdnSubcontraflg 히든필드는 외주가공여부를 나타내는 필드이다.
' 2005년 10월 패치에서 M_MVMT_TYPE 테이플의 SUBCONTRA_FLG 필드는 사급여부를 바꾸었음 
' SUBCONTRA2_FLG 필드를 만들어 외주가공여부로 사용한다.

'==============================================================================================================================
Function ChangeTag(Byval Changeflg)
	
	frm1.vspdData.ReDraw = false

	If Changeflg = true then
		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "Q"
		ggoOper.SetReqAttr	frm1.txtRemark, "Q"
		Call SetSpreadLock 
	Else
		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "N"
'		Call ggoOper.LockField(Document, "N")
		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "D"
		Call SetSpreadLock 
	End if 

	frm1.vspdData.ReDraw = true
	
End Function 
'==============================================================================================================================
Sub InitSpreadPosVariables()
	C_PlantCd		        = 1
	C_PlantNm		        = 2
	C_ItemCd		        = 3
	C_ItemNm		        = 4
	C_Spec			        = 5 
	C_TrackingNo	        = 6
	C_InspFlg		        = 7
	C_MvmtRcptLookUpBtn     = 8
	C_GrQty			= 9
	C_StockQty		= 10
	C_GRUnit		= 11
	C_Cur		    = 12
	C_MvmtPrc	    = 13
	C_DocAmt		= 14
	C_LocAmt		= 15
	C_SlCd			= 16
	C_SlCdPop		= 17
	C_SlNm			= 18
	C_InspSts		= 19
	C_GRMeth		= 20
	C_LotNo			= 21
	C_LotSeqNo		= 22
	C_MakerLotNo	= 23
	C_MakerLotSeqNo	= 24
	C_RemarkDtl		= 25
	C_GRNo			= 26
	C_GRSeqNo		= 27
	C_InspReqNo		= 28
	C_InspResultNo	= 29
	C_PoNo			= 30
	C_PoSeqNo		= 31
	C_CCNo			= 32
	C_CCSeqNo		= 33
	C_LLCNo			= 34
	C_LLCSeqNo		= 35
	C_Stateflg		= 36
	C_InspMethCd	= 37
	C_MvmtNo		= 38
	C_IvNo			= 39
	C_IvSeqNo		= 40
	C_RetOrdQty		= 41
	'2008-03-29 7:43오후 :: hanc
	C_TransTime     = 42
    C_MainLot       = 43
    C_ImportTime    = 44
    C_CreateType    = 45

	
End Sub

Sub SetDefaultVal()
	
	frm1.txtGmDt.Text = EndDate
	frm1.txtGroupCd.Value = Parent.gPurGrp
    Call SetToolBar("1110000000001111")
    frm1.txtMvmtNo.focus 
    Set gActiveElement = document.activeElement    
    interface_Account = GetSetupMod(Parent.gSetupMod, "a")
	frm1.btnGlSel.disabled = true
	frm1.btnGlSel_1.disabled = true    
End Sub

Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20060403",,Parent.gAllowDragDropSpread  
		
		.ReDraw = false
		
		.MaxCols = C_CreateType + 1
    	.MaxRows = 0

		Call AppendNumberPlace("6", "5", "0")
    	Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 		C_PlantCd,	"공장", 10
		ggoSpread.SSSetEdit 		C_PlantNm,	"공장명", 20
		ggoSpread.SSSetEdit 		C_ItemCd,	"품목", 10
		ggoSpread.SSSetEdit 		C_ItemNm,	"품목명", 20 
		ggoSpread.SSSetEdit 		C_Spec,	    "품목규격", 20 	
		ggoSpread.SSSetEdit 		C_TrackingNo,	"Tracking No.", 15 	
		ggoSpread.SSSetCheck 		C_InspFlg,	"검사품여부",10,,,true
		ggoSpread.SSSetButton 		C_MvmtRcptLookUpBtn	
		SetSpreadFloatLocal 		C_GrQty,	"입고수량",15,1, 3
		SetSpreadFloatLocal 		C_StockQty,	"재고처리수량",15,1, 3
		ggoSpread.SSSetEdit 		C_GRUnit,	"단위", 10
		ggoSpread.SSSetEdit 		C_Cur,	    "화폐", 10
		
		ggoSpread.SSSetFloat		C_MvmtPrc,	"입고단가"		, 15	,"C" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat 		C_DocAmt,	"입고금액"		, 15	,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat 		C_LocAmt,	"입고자국금액"	, 15	,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec		    
		
		ggoSpread.SSSetEdit			C_SlCd,		"창고", 10
		ggoSpread.SSSetButton 		C_SlCdPop
		ggoSpread.SSSetEdit 		C_SlNm,		"창고명", 20	    
		ggoSpread.SSSetEdit 		C_InspSts,	"검사상태", 10
		ggoSpread.SSSetEdit 		C_GRMeth,	"납입시검사방법", 20
		ggoSpread.SSSetEdit 		C_LotNo,	"Lot No.", 20, , , 25, 2    
		SetSpreadFloatLocal 		C_LotSeqNo, "LOT NO 순번", 20,1,6
		ggoSpread.SSSetEdit 		C_MakerLotNo,	"MAKER LOT NO.", 20,,,12,2    
		SetSpreadFloatLocal 		C_MakerLotSeqNo,"Maker Lot 순번", 20,1,6
		
		ggoSpread.SSSetEdit 		C_RemarkDtl,	"비고", 20
		
		ggoSpread.SSSetEdit 		C_GRNo,		"재고처리번호", 20
		ggoSpread.SSSetFloat 		C_GRSeqNo,	"재고처리순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_InspReqNo,"검사요청번호", 20
		ggoSpread.SSSetEdit 		C_InspResultNo,"검사결과등록번호", 20
		ggoSpread.SSSetEdit 		C_PoNo,		"발주번호", 20
		ggoSpread.SSSetFloat 		C_PoSeqNo,	"발주순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_CCNo,		"통관번호", 20
		ggoSpread.SSSetFloat 		C_CCSeqNo,	"통관순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0    
		ggoSpread.SSSetEdit 		C_LLCNo,	"LOCAL L/C번호", 20
		ggoSpread.SSSetFloat 		C_LLCSeqNo,	"LOCAL L/C순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		
		ggoSpread.SSSetEdit 		C_Stateflg,	 "State Flag"	, 10
		ggoSpread.SSSetEdit 		C_InspMethCd,"Insp.Meth."	, 10
		ggoSpread.SSSetEdit 		C_MvmtNo,	 "Movmt.No."	, 10
		ggoSpread.SSSetEdit 		C_IvNo,	     "매입번호", 20
		ggoSpread.SSSetFloat 		C_IvSeqNo,	"매입순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		'ggoSpread.SSSetEdit 		C_IvNo,		 "IV No."		, 10
		'ggoSpread.SSSetEdit 		C_IvSeqNo,	 "IV Seq. No."	, 10
		SetSpreadFloatLocal 		C_RetOrdQty, "반품수량",15,1, 3

        '2008-03-29 7:31오후 :: hanc
		ggoSpread.SSSetEdit 		C_TransTime  ,	"TransTime", 20 
		ggoSpread.SSSetEdit 		C_MainLot    ,	"MainLot", 25
		ggoSpread.SSSetEdit 		C_ImportTime ,	"ImportTime", 20
		ggoSpread.SSSetEdit 		C_CreateType ,	"CreateType", 20


		Call ggoSpread.MakePairsColumn(C_SlCd,C_SlCdPop)
		Call ggoSpread.SSSetColHidden(C_Stateflg,C_MvmtNo,True)
		Call ggoSpread.SSSetColHidden(C_MvmtRcptLookUpBtn, C_MvmtRcptLookUpBtn,True)
		Call ggoSpread.SSSetColHidden(C_RetOrdQty,C_RetOrdQty,True)
		Call ggoSpread.SSSetColHidden(C_TransTime,C_TransTime,True)
		Call ggoSpread.SSSetColHidden(C_MainLot,C_MainLot,True)
		Call ggoSpread.SSSetColHidden(C_ImportTime,C_ImportTime,True)
		Call ggoSpread.SSSetColHidden(C_CreateType,C_CreateType,True)
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)		
		
		Call SetSpreadLock()
		
		.ReDraw = true
		
    End With
End Sub
'==============================================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLock -1, -1
	ggoSpread.SpreadUnlock C_MvmtRcptLookUpBtn,	-1,	C_MvmtRcptLookUpBtn,-1
End Sub
'==============================================================================================================================
Sub SetSpreadLock_mes()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLock -1, -1
	ggoSpread.SpreadUnlock C_MvmtRcptLookUpBtn,	-1,	C_MvmtRcptLookUpBtn,-1
End Sub
'==============================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLock	-1, pvStartRow, -1, pvEndRow
	ggoSpread.SpreadUnlock C_MakerLotNo,		pvStartRow,		C_MvmtRcptLookUpBtn,	pvEndRow
	ggoSpread.SpreadUnlock C_MakerLotSeqNo,		pvStartRow,		C_MvmtRcptLookUpBtn,	pvEndRow
	ggoSpread.SpreadUnlock C_MvmtRcptLookUpBtn,	pvStartRow,		C_MvmtRcptLookUpBtn,	pvEndRow
	ggoSpread.SpreadUnlock C_SlCd,				pvStartRow,		C_SlCdPop,				pvEndRow
	ggoSpread.SpreadUnlock C_GrQty,				pvStartRow,		C_GrQty,				pvEndRow
	ggoSpread.SpreadUnlock C_RemarkDtl,			pvStartRow,		C_RemarkDtl,				pvEndRow
	ggoSpread.SSSetRequired 	C_GrQty,	pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_SlCd ,	pvStartRow, pvEndRow
End Sub
'==============================================================================================================================
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
			C_MvmtRcptLookUpBtn = iCurColumnPos(8)
			C_GrQty			= iCurColumnPos(9)
			C_StockQty		= iCurColumnPos(10)
			C_GRUnit		= iCurColumnPos(11)
			C_Cur		    = iCurColumnPos(12)
			C_MvmtPrc	    = iCurColumnPos(13)
			C_DocAmt		= iCurColumnPos(14)
			C_LocAmt		= iCurColumnPos(15)
			C_SlCd			= iCurColumnPos(16)
			C_SlCdPop		= iCurColumnPos(17)
			C_SlNm			= iCurColumnPos(18)
			C_InspSts		= iCurColumnPos(19)
			C_GRMeth		= iCurColumnPos(20)
			C_LotNo			= iCurColumnPos(21)
			C_LotSeqNo		= iCurColumnPos(22)
			C_MakerLotNo	= iCurColumnPos(23)
			C_MakerLotSeqNo	= iCurColumnPos(24)
			C_RemarkDtl		= iCurColumnPos(25)
			C_GRNo			= iCurColumnPos(26)
			C_GRSeqNo		= iCurColumnPos(27)
			C_InspReqNo		= iCurColumnPos(28)
			C_InspResultNo  = iCurColumnPos(29)
			C_PoNo			= iCurColumnPos(30)
			C_PoSeqNo		= iCurColumnPos(31)
			C_CCNo			= iCurColumnPos(32)
			C_CCSeqNo		= iCurColumnPos(33)
			C_LLCNo			= iCurColumnPos(34)
			C_LLCSeqNo		= iCurColumnPos(35)
			C_Stateflg		= iCurColumnPos(36)
			C_InspMethCd	= iCurColumnPos(37)
			C_MvmtNo		= iCurColumnPos(38)
			C_IvNo			= iCurColumnPos(39)
			C_IvSeqNo		= iCurColumnPos(40)
			C_RetOrdQty		= iCurColumnPos(41)
            '2008-03-29 7:32오후 :: hanc
			C_TransTime  	= iCurColumnPos(42)
			C_MainLot    	= iCurColumnPos(43)
			C_ImportTime 	= iCurColumnPos(44)
			C_CreateType 	= iCurColumnPos(45)

	End Select

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
'==============================================================================================================================
Function OpenPoRef()

	Dim strRet
	Dim arrParam(12)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","발주내역참조" )
		Exit Function
	End if 

	if Trim(frm1.txtMvmtType.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "입고형태","X")
		frm1.txtMvmtType.focus
		Set gActiveElement = document.activeElement
		Exit Function	
	elseif Trim(frm1.txtSupplierCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공급처","X")
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if

	if (UCase(Trim(frm1.hdnImportflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="Y") or _
	   (UCase(Trim(frm1.hdnSubcontraflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="N") then
		Call DisplayMsgBox("17A012", "X","입고형태" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","발주내역참조" )
		'입고형태 는(은) 발주참조를 할수 없습니다."
		Exit Function
	End if

	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = Trim(frm1.txtGroupNm.value)
	arrParam(4) = "N"		'Clsflg
	arrParam(5) = "Y"		'Releaseflg
	arrParam(8) = "GR"		'Rcptflg
	arrParam(9) = Trim(frm1.txtMvmtType.Value)
	arrParam(10)= ""
	arrParam(11)= ""
	arrParam(12)= ""

	iCalledAspName = AskPRAspName("M3112RA5")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA5", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.	
	If strRet(0,0) = "" Then
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetPoRef(strRet)
	End If	
		
End Function

'==============================================================================================================================
'20071211::hanc
Function OpenPoRef1()

	Dim strRet
	Dim arrParam(12)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","발주내역참조" )
		Exit Function
	End if 

'20071211::hanc입고형태, 공급처를 필수로 넘겨받지 않음.
'20071211::hanc	if Trim(frm1.txtMvmtType.Value) = "" then
'20071211::hanc		Call DisplayMsgBox("17A002","X" , "입고형태","X")
'20071211::hanc		frm1.txtMvmtType.focus
'20071211::hanc		Set gActiveElement = document.activeElement
'20071211::hanc		Exit Function	
'20071211::hanc	elseif Trim(frm1.txtSupplierCd.Value) = "" then
'20071211::hanc		Call DisplayMsgBox("17A002","X" , "공급처","X")
'20071211::hanc		frm1.txtSupplierCd.focus
'20071211::hanc		Set gActiveElement = document.activeElement
'20071211::hanc		Exit Function
'20071211::hanc	End if

	if (UCase(Trim(frm1.hdnImportflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="Y") or _
	   (UCase(Trim(frm1.hdnSubcontraflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="N") then
		Call DisplayMsgBox("17A012", "X","입고형태" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","발주내역참조" )
		'입고형태 는(은) 발주참조를 할수 없습니다."
		Exit Function
	End if

	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = Trim(frm1.txtGroupNm.value)
	arrParam(4) = "N"		'Clsflg
	arrParam(5) = "Y"		'Releaseflg
	arrParam(8) = "GR"		'Rcptflg
	arrParam(9) = Trim(frm1.txtMvmtType.Value)
	arrParam(10)= ""
	arrParam(11)= ""
	arrParam(12)= ""

	iCalledAspName = AskPRAspName("M3112RA5_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA5", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.	
	If strRet(0,0) = "" Then
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetPoRef1(strRet)
	End If	
		
End Function
'==============================================================================================================================
'20080211::hanc
Function OpenPoRef2()

	Dim strRet
	Dim arrParam(12)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","MES입고참조" )
		Exit Function
	End if 

	if (UCase(Trim(frm1.hdnImportflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="Y") or _
	   (UCase(Trim(frm1.hdnSubcontraflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="N") then
		Call DisplayMsgBox("17A012", "X","입고형태" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","발주내역참조" )
		'입고형태 는(은) 발주참조를 할수 없습니다."
		Exit Function
	End if

	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = Trim(frm1.txtGroupNm.value)
	arrParam(4) = "N"		'Clsflg
	arrParam(5) = "Y"		'Releaseflg
	arrParam(8) = "GR"		'Rcptflg
	arrParam(9) = Trim(frm1.txtMvmtType.Value)
	arrParam(10)= ""
	arrParam(11)= ""
	arrParam(12)= ""

	iCalledAspName = AskPRAspName("M3112RA6_KO441")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA5", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.	
	If strRet(0,0) = "" Then
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetPoRef2(strRet)
	End If	
    Call	SetSpreadLock_mes()     '20080214::hanc::mes입고 참조 후 수량을 수정할 수 없다.
		
End Function
'==============================================================================================================================
Function SetPoRef(strRet)

	Dim Index1, Count1, Row1
	Dim temp
			
	Const C_PoNo_Ref		= 0
	Const C_PoSeq_Ref		= 1
	Const C_PlantCd_Ref		= 2
	Const C_SLCd_Ref		= 3
	Const C_ItemCd_Ref		= 4
	Const C_ItemNm_Ref		= 5
	Const C_Spec_Ref		= 6	
	Const C_TrackingNo_Ref	= 7
		
	Const C_POQty_Ref		= 8
	Const C_Unit_Ref		= 9
	Const C_POPrc_Ref		= 10
	Const C_POAmt_Ref		= 11
	Const C_POCur_Ref		= 12
	Const C_PODlvyDt_Ref	= 13

	Const C_GRQty_Ref		= 14
	Const C_LCQty_Ref		= 15
	Const C_PreIvQty_Ref	= 16
    Const C_InspectQty_Ref	= 17
    Const C_IvQty_Ref		= 18
    
    Const C_InspFlg_Ref		= 19
	Const C_InspMeth_Ref	= 20
	Const C_InspMethCd_Ref	= 21
	Const C_PlantNm_Ref		= 22
	Const C_SLNm_Ref		= 23
	Const C_Pur_Grp_Ref		= 24
	Const C_LCRCPTQTY_Ref	= 25
    Const C_Lot_flg			= 26
    Const C_Lot_gen_mtd		= 27
    
    
	Count1 = Ubound(strRet,1)
	
	With frm1.vspdData
		
		.Redraw = False
	
		Call fncinsertrow(Count1 + 1)
		
		For index1 = 0 to Count1
			
			Row1 = .ActiveRow + Index1
			
			'C_LCRCPTQTY_Ref 는 히든필드(HH)이므로 
			temp = UNICDbl(strRet(index1,C_POQty_Ref)) - (UNICDbl(strRet(index1,C_GRQty_Ref)) + UNICDbl(strRet(index1,C_PreIvQty_Ref)) _
				 + UNICDbl(strRet(index1, C_InspectQty_Ref)) + UNICDbl(strRet(index1, C_LCQty_Ref)) - CDbl(strRet(index1, C_LCRCPTQTY_Ref))  )

			Call .SetText(C_PlantCd,	Row1, strRet(index1,C_PlantCd_Ref))
			Call .SetText(C_PlantNm,	Row1, strRet(index1,C_PlantNm_Ref))
			Call .SetText(C_itemCd,		Row1, strRet(index1,C_ItemCd_Ref))
			Call .SetText(C_itemNm,		Row1, strRet(index1,C_ItemNm_Ref))
			Call .SetText(C_Spec,		Row1, strRet(index1,C_Spec_Ref))
			Call .SetText(C_TrackingNo, Row1, strRet(index1,C_TrackingNo_Ref))
			Call .SetText(C_GRQty,		Row1, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			Call .SetText(C_GRUnit,		Row1, strRet(index1,C_Unit_Ref))
			Call .SetText(C_Cur,		Row1, strRet(index1,C_POCur_Ref))
			Call .SetText(C_MvmtPrc,	Row1, 0)
			Call .SetText(C_DocAmt,		Row1, 0)
			Call .SetText(C_LocAmt,		Row1, 0)
			Call .SetText(C_SLCd,		Row1, strRet(index1,C_SLCd_Ref))
			Call .SetText(C_SLNm,		Row1, strRet(index1,C_SLNm_Ref))
			Call .SetText(C_PoNo,		Row1, strRet(index1,C_PoNo_Ref))
			Call .SetText(C_PoSeqNo,	Row1, strRet(index1,C_PoSeq_Ref))
			Call .SetText(C_InspMethCd, Row1, strRet(index1,C_InspMethCd_Ref))
			Call .SetText(C_Stateflg,	Row1, "PO")

			.Row = Row1	
			.Col = C_InspFlg										
			If strRet(index1,C_InspFlg_Ref) = "Y" Then
				Call .SetText(C_InspFlg,	Row1, "1")
				Call .SetText(C_GRMeth,		Row1, strRet(index1,C_InspMeth_Ref))
			Else
				Call .SetText(C_InspFlg,	Row1, "0")
				Call .SetText(C_GRMeth,		Row1, "")
			End if
	
			If UCase(Trim(strRet(index1,C_Lot_flg))) = "Y" and UCase(Trim(strRet(index1,C_Lot_gen_mtd))) = "M" Then
				ggoSpread.spreadUnlock	C_LotNo, Row1, C_LotSeqNo, Row1
				ggoSpread.SSSetRequired	C_LotNo, Row1, Row1
			ElseIf UCase(Trim(strRet(index1,C_Lot_flg))) = "N" Then
				Call .SetText(C_LotNo,	Row1, "*")
			End If
		
			'C_Pur_Grp_Ref
			IF index1 = 0 Then
				frm1.txtGroupCd.value = strRet(index1,C_Pur_Grp_Ref)		
				Call changeGroupCd()                                            '20071211::hanc
			End IF

		Next
	
		Call LocalReFormatSpreadCellByCellByCurrency()
		Call setReference()
		
		.ReDraw = True
		
	End with
	
End Function
'==============================================================================================================================
'20080213::hanc
Function SetPoRef2(strRet)
	Dim Index1, Count1, Row1
	Dim temp
			
    Const C_supplier		= 3        '20071211::hanc
    Const C_MvmtType		= 2        '20071211::hanc
	Const C_PoNo_Ref		= 4
	Const C_PoSeq_Ref		= 5
	Const C_PlantCd_Ref		= 6
	Const C_SLCd_Ref		= 7
	Const C_ItemCd_Ref		= 8
	Const C_ItemNm_Ref		= 9
	Const C_Spec_Ref		= 10	
	Const C_TrackingNo_Ref	= 11
		
	Const C_POQty_Ref		= 12
	Const C_Unit_Ref		= 13
	Const C_POPrc_Ref		= 14
	Const C_POAmt_Ref		= 15
	Const C_POCur_Ref		= 16
	Const C_PODlvyDt_Ref	= 17

	Const C_GRQty_Ref		= 18
	Const C_LCQty_Ref		= 19
	Const C_PreIvQty_Ref	= 20
    Const C_InspectQty_Ref	= 21
    Const C_IvQty_Ref		= 22
    
    Const C_InspFlg_Ref		= 23
	Const C_InspMeth_Ref	= 24
	Const C_InspMethCd_Ref	= 25
	Const C_PlantNm_Ref		= 26
	Const C_SLNm_Ref		= 27
	Const C_Pur_Grp_Ref		= 28
	Const C_LCRCPTQTY_Ref	= 29
    Const C_Lot_flg			= 30
    Const C_Lot_gen_mtd		= 31
    Const C_MVMT_NO_REF		= 32
    
    
	Count1 = Ubound(strRet,1)
	
	With frm1.vspdData
		
		.Redraw = False
	
		Call fncinsertrow(Count1 + 1)
		
		For index1 = 0 to Count1
			
			Row1 = .ActiveRow + Index1
			
			'C_LCRCPTQTY_Ref 는 히든필드(HH)이므로 
'			temp = UNICDbl(strRet(index1,C_POQty_Ref)) - (UNICDbl(strRet(index1,C_GRQty_Ref)) + UNICDbl(strRet(index1,C_PreIvQty_Ref)) _
'				 + UNICDbl(strRet(index1, C_InspectQty_Ref)) + UNICDbl(strRet(index1, C_LCQty_Ref)) - CDbl(strRet(index1, C_LCRCPTQTY_Ref))  )
            temp = UNICDbl(strRet(index1,C_GRQty_Ref))
            
			Call .SetText(C_PlantCd,	Row1, strRet(index1,C_PlantCd_Ref))
			Call .SetText(C_PlantNm,	Row1, strRet(index1,C_PlantNm_Ref))
			Call .SetText(C_itemCd,		Row1, strRet(index1,C_ItemCd_Ref))
			Call .SetText(C_itemNm,		Row1, strRet(index1,C_ItemNm_Ref))
			Call .SetText(C_Spec,		Row1, strRet(index1,C_Spec_Ref))
			Call .SetText(C_TrackingNo, Row1, strRet(index1,C_TrackingNo_Ref))
			Call .SetText(C_GRQty,		Row1, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			Call .SetText(C_GRUnit,		Row1, strRet(index1,C_Unit_Ref))
			Call .SetText(C_Cur,		Row1, strRet(index1,C_POCur_Ref))
			Call .SetText(C_MvmtPrc,	Row1, 0)
			Call .SetText(C_DocAmt,		Row1, 0)
			Call .SetText(C_LocAmt,		Row1, 0)
			Call .SetText(C_SLCd,		Row1, strRet(index1,C_SLCd_Ref))
			Call .SetText(C_SLNm,		Row1, strRet(index1,C_SLNm_Ref))
			Call .SetText(C_PoNo,		Row1, strRet(index1,C_PoNo_Ref))
			Call .SetText(C_PoSeqNo,	Row1, strRet(index1,C_PoSeq_Ref))
			Call .SetText(C_InspMethCd, Row1, strRet(index1,C_InspMethCd_Ref))
			Call .SetText(C_Stateflg,	Row1, "PO")
            
            Call .SetText(C_RemarkDtl,  Row1, strRet(index1,C_MVMT_NO_REF))         '20080215::hanc
            Call .SetText(C_MakerLotNo,  Row1, strRet(index1,C_MVMT_NO_REF))        '20080215::hanc

            Call .SetText(C_TransTime ,  Row1, strRet(index1,33))        '20080215::hanc
            Call .SetText(C_MainLot   ,  Row1, strRet(index1,34))        '20080215::hanc
            Call .SetText(C_ImportTime,  Row1, strRet(index1,35))        '20080215::hanc
            Call .SetText(C_CreateType,  Row1, strRet(index1,36))        '20080215::hanc
			            
			.Row = Row1	
			.Col = C_InspFlg										
			If strRet(index1,C_InspFlg_Ref) = "Y" Then
				Call .SetText(C_InspFlg,	Row1, "1")
				Call .SetText(C_GRMeth,		Row1, strRet(index1,C_InspMeth_Ref))
			Else
				Call .SetText(C_InspFlg,	Row1, "0")
				Call .SetText(C_GRMeth,		Row1, "")
			End if
	
			If UCase(Trim(strRet(index1,C_Lot_flg))) = "Y" and UCase(Trim(strRet(index1,C_Lot_gen_mtd))) = "M" Then
				ggoSpread.spreadUnlock	C_LotNo, Row1, C_LotSeqNo, Row1
				ggoSpread.SSSetRequired	C_LotNo, Row1, Row1
			ElseIf UCase(Trim(strRet(index1,C_Lot_flg))) = "N" Then
				Call .SetText(C_LotNo,	Row1, "*")
			End If
		
			'C_Pur_Grp_Ref
			IF index1 = 0 Then
				frm1.txtGroupCd.value       = strRet(index1,C_Pur_Grp_Ref)
				frm1.txtSupplierCd.value    = strRet(index1,C_supplier)         '20071211::hanc
				frm1.txtMvmtType.value      = strRet(index1,C_MvmtType)         '20071211::hanc
'20080213::hanc				Call changeGroupCd()                                            '20071211::hanc
				Call changeMvmtType()                                           '20071211::hanc
				Call changeSpplCd()                                             '20071211::hanc
			End IF

		Next
	
		Call LocalReFormatSpreadCellByCellByCurrency()
		Call setReference()
		
		.ReDraw = True
		
	End with

End Function

'==============================================================================================================================
'20071211::hanc
Function SetPoRef1(strRet)
	Dim Index1, Count1, Row1
	Dim temp
			
    Const C_supplier		= 3        '20071211::hanc
    Const C_MvmtType		= 2        '20071211::hanc
	Const C_PoNo_Ref		= 4
	Const C_PoSeq_Ref		= 5
	Const C_PlantCd_Ref		= 6
	Const C_SLCd_Ref		= 7
	Const C_ItemCd_Ref		= 8
	Const C_ItemNm_Ref		= 9
	Const C_Spec_Ref		= 10	
	Const C_TrackingNo_Ref	= 11
		
	Const C_POQty_Ref		= 12
	Const C_Unit_Ref		= 13
	Const C_POPrc_Ref		= 14
	Const C_POAmt_Ref		= 15
	Const C_POCur_Ref		= 16
	Const C_PODlvyDt_Ref	= 17

	Const C_GRQty_Ref		= 18
	Const C_LCQty_Ref		= 19
	Const C_PreIvQty_Ref	= 20
    Const C_InspectQty_Ref	= 21
    Const C_IvQty_Ref		= 22
    
    Const C_InspFlg_Ref		= 23
	Const C_InspMeth_Ref	= 24
	Const C_InspMethCd_Ref	= 25
	Const C_PlantNm_Ref		= 26
	Const C_SLNm_Ref		= 27
	Const C_Pur_Grp_Ref		= 28
	Const C_LCRCPTQTY_Ref	= 29
    Const C_Lot_flg			= 30
    Const C_Lot_gen_mtd		= 31
    
    
	Count1 = Ubound(strRet,1)
	
	With frm1.vspdData
		
		.Redraw = False
	
		Call fncinsertrow(Count1 + 1)
		
		For index1 = 0 to Count1
			
			Row1 = .ActiveRow + Index1
			
			'C_LCRCPTQTY_Ref 는 히든필드(HH)이므로 
			temp = UNICDbl(strRet(index1,C_POQty_Ref)) - (UNICDbl(strRet(index1,C_GRQty_Ref)) + UNICDbl(strRet(index1,C_PreIvQty_Ref)) _
				 + UNICDbl(strRet(index1, C_InspectQty_Ref)) + UNICDbl(strRet(index1, C_LCQty_Ref)) - CDbl(strRet(index1, C_LCRCPTQTY_Ref))  )

			Call .SetText(C_PlantCd,	Row1, strRet(index1,C_PlantCd_Ref))
			Call .SetText(C_PlantNm,	Row1, strRet(index1,C_PlantNm_Ref))
			Call .SetText(C_itemCd,		Row1, strRet(index1,C_ItemCd_Ref))
			Call .SetText(C_itemNm,		Row1, strRet(index1,C_ItemNm_Ref))
			Call .SetText(C_Spec,		Row1, strRet(index1,C_Spec_Ref))
			Call .SetText(C_TrackingNo, Row1, strRet(index1,C_TrackingNo_Ref))
			Call .SetText(C_GRQty,		Row1, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			Call .SetText(C_GRUnit,		Row1, strRet(index1,C_Unit_Ref))
			Call .SetText(C_Cur,		Row1, strRet(index1,C_POCur_Ref))
			Call .SetText(C_MvmtPrc,	Row1, 0)
			Call .SetText(C_DocAmt,		Row1, 0)
			Call .SetText(C_LocAmt,		Row1, 0)
			Call .SetText(C_SLCd,		Row1, strRet(index1,C_SLCd_Ref))
			Call .SetText(C_SLNm,		Row1, strRet(index1,C_SLNm_Ref))
			Call .SetText(C_PoNo,		Row1, strRet(index1,C_PoNo_Ref))
			Call .SetText(C_PoSeqNo,	Row1, strRet(index1,C_PoSeq_Ref))
			Call .SetText(C_InspMethCd, Row1, strRet(index1,C_InspMethCd_Ref))
			Call .SetText(C_Stateflg,	Row1, "PO")
'			Call .SetText(C_StockCd,	Row1, "ETC")        '2008-05-16 2:24오후 :: hanc

			.Row = Row1	
			.Col = C_InspFlg										
			If strRet(index1,C_InspFlg_Ref) = "Y" Then
				Call .SetText(C_InspFlg,	Row1, "1")
				Call .SetText(C_GRMeth,		Row1, strRet(index1,C_InspMeth_Ref))
			Else
				Call .SetText(C_InspFlg,	Row1, "0")
				Call .SetText(C_GRMeth,		Row1, "")
			End if
	
			If UCase(Trim(strRet(index1,C_Lot_flg))) = "Y" and UCase(Trim(strRet(index1,C_Lot_gen_mtd))) = "M" Then
				ggoSpread.spreadUnlock	C_LotNo, Row1, C_LotSeqNo, Row1
				ggoSpread.SSSetRequired	C_LotNo, Row1, Row1
			ElseIf UCase(Trim(strRet(index1,C_Lot_flg))) = "N" Then
				Call .SetText(C_LotNo,	Row1, "*")
			End If
		
			'C_Pur_Grp_Ref
			IF index1 = 0 Then
				frm1.txtGroupCd.value       = strRet(index1,C_Pur_Grp_Ref)
				frm1.txtSupplierCd.value    = strRet(index1,C_supplier)         '20071211::hanc
				frm1.txtMvmtType.value      = strRet(index1,C_MvmtType)         '20071211::hanc
				Call changeGroupCd()                                            '20071211::hanc
				Call changeMvmtType()                                           '20071211::hanc
				Call changeSpplCd()                                             '20071211::hanc
			End IF

		Next
	
		Call LocalReFormatSpreadCellByCellByCurrency()
		Call setReference()
		
		.ReDraw = True
		
	End with
	
End Function
'==============================================================================================================================
Function OpenCcRef()

	Dim strRet
	Dim arrParam(7)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","통관참조" )
		Exit Function
	End if 
	
	if Trim(frm1.txtMvmtType.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "입고형태","X")
		frm1.txtMvmtType.focus
		Set gActiveElement = document.activeElement
		Exit Function	
	elseif Trim(frm1.txtSupplierCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공급처","X")
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	
	if Not(UCase(Trim(frm1.hdnImportflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="Y") then
		Call DisplayMsgBox("17A012", "X","입고형태" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","통관참조" )
		'입고형태 는 통관참조를 할수 없습니다."
		Exit Function
	End if

	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = Trim(frm1.txtGroupNm.value)
	'arrParam(5) = Trim(frm1.hdnSubcontraflg.value)
	arrParam(5) = Trim(frm1.txtMvmtType.value)
	
	iCalledAspName = AskPRAspName("M4212RA1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4212RA1_KO441", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If strRet(0,0) = "" Then
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else	
		Call SetCcRef(strRet)
	End If	
		
End Function
'==============================================================================================================================
Function SetCcRef(strRet)

	Dim Index1, Count1, Row1
    Dim temp
	
	Const C_CcNo_Ref		= 0
	Const C_CcSeq_Ref		= 1
	Const C_PlantCd_Ref		= 2
	Const C_SLCd_Ref		= 3
	Const C_ItemCd_Ref		= 4
	Const C_ItemNm_Ref		= 5
	Const C_Spec_Ref		= 6
	Const C_TrackingNo_Ref	= 7	
	Const C_Qty_Ref			= 8
	Const C_Unit_Ref		= 9
	Const C_CcDt_Ref		= 10
	Const C_PoNo_Ref		= 11
	Const C_PoSeq_Ref		= 12
	Const C_BLNo_Ref		= 13
	Const C_BLSeq_Ref		= 14
	Const C_RcptQty_Ref		= 15
	Const C_PlantNm_Ref		= 16
	Const C_SLNm_Ref		= 17
	Const C_Price_Ref		= 18
	Const C_Cur_Ref			= 19
	Const C_InspFlg_Ref		= 20
	Const C_InspMeth_Ref	= 21
	Const C_InspMethCd_Ref	= 22
	Const C_Pur_Grp_Ref		= 23
	Const C_InspectQty_Ref  = 24
    Const C_Lot_flg			= 25
    Const C_Lot_gen_mtd		= 26
	
	Count1 = Ubound(strRet,1)
	
	With frm1.vspdData
		
		.Redraw = False
	
		Call fncinsertrow(Count1 + 1)
		
		For index1 = 0 to Count1
			
			Row1 = .ActiveRow + Index1
			temp = unicdbl(strRet(index1,C_Qty_Ref)) - (unicdbl(strRet(index1,C_RcptQty_Ref)) + unicdbl(strRet(index1, C_InspectQty_Ref)))
				
			Call .SetText(C_PlantCd,	Row1, strRet(index1,C_PlantCd_Ref))
			Call .SetText(C_PlantNm,	Row1, strRet(index1,C_PlantNm_Ref))
			Call .SetText(C_itemCd,		Row1, strRet(index1,C_ItemCd_Ref))
			Call .SetText(C_itemNm,		Row1, strRet(index1,C_ItemNm_Ref))
			Call .SetText(C_Spec,		Row1, strRet(index1,C_Spec_Ref))
			Call .SetText(C_TrackingNo, Row1, strRet(index1,C_TrackingNo_Ref))
			Call .SetText(C_GRQty,		Row1, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			Call .SetText(C_GRUnit,		Row1, strRet(index1,C_Unit_Ref))
			Call .SetText(C_Cur,		Row1, strRet(index1,C_Cur_Ref))
			Call .SetText(C_MvmtPrc,	Row1, 0)
			Call .SetText(C_DocAmt,		Row1, 0)
			Call .SetText(C_LocAmt,		Row1, 0)
			Call .SetText(C_SLCd,		Row1, strRet(index1,C_SLCd_Ref))
			Call .SetText(C_SLNm,		Row1, strRet(index1,C_SLNm_Ref))
			Call .SetText(C_PoNo,		Row1, strRet(index1,C_PoNo_Ref))
			Call .SetText(C_PoSeqNo,	Row1, strRet(index1,C_PoSeq_Ref))
			Call .SetText(C_InspMethCd, Row1, strRet(index1,C_InspMethCd_Ref))
			Call .SetText(C_Stateflg,	Row1, "CC")
			Call .SetText(C_CcNo, Row1, strRet(index1,C_CcNo_Ref))
			Call .SetText(C_CcSeqNo, Row1, strRet(index1,C_CcSeq_Ref))
				
			.Row = Row1	
			.Col = C_InspFlg					
			If strRet(index1,C_InspFlg_Ref) = "Y" Then
				Call .SetText(C_InspFlg,	Row1, "1")
				Call .SetText(C_GRMeth,		Row1, strRet(index1,C_InspMeth_Ref))
				
			Else
				Call .SetText(C_InspFlg,	Row1, "0")
				Call .SetText(C_GRMeth,		Row1, "")
			End if
				
			If UCase(Trim(strRet(index1,C_Lot_flg))) = "Y" and UCase(Trim(strRet(index1,C_Lot_gen_mtd))) = "M" Then
				ggoSpread.spreadUnlock	C_LotNo, Row1, C_LotSeqNo, Row1
				ggoSpread.SSSetRequired	C_LotNo, Row1, Row1
			ElseIf UCase(Trim(strRet(index1,C_Lot_flg))) = "N" Then
				Call .SetText(C_LotNo,	Row1, "*")
			End If
							'C_Pur_Grp_Ref
			IF index1 = 0 Then
				frm1.txtGroupCd.value = strRet(index1,C_Pur_Grp_Ref)	
    			Call changeGroupCd()                                            '20071211::hanc
	
			End IF
			
		Next
		
		Call LocalReFormatSpreadCellByCellByCurrency()
		Call setReference()
		
		.ReDraw = True
		
	End with
	
End Function
'==============================================================================================================================
Function OpenLLCRef()

	Dim strRet
	Dim arrParam(7)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","LOCAL L/C참조" )
		Exit Function
	End if 
	
	if Trim(frm1.txtMvmtType.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "입고형태","X")
		frm1.txtMvmtType.focus
		Set gActiveElement = document.activeElement
		Exit Function	
	elseif Trim(frm1.txtSupplierCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공급처","X")
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if

	if Not (UCase(Trim(frm1.hdnImportflg.Value))<>"Y" and UCase(Trim(frm1.hdnRcptflg.Value))="Y" and UCase(Trim(frm1.hdnRetflg.Value))<>"Y") then
		Call DisplayMsgBox("17A012", "X","입고형태" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","LOCAL L/C참조" )
		'입고형태 는(은) Local L/C참조를 할수 없습니다."
		Exit Function
	End if
	
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = Trim(frm1.txtGroupNm.value)

	iCalledAspName = AskPRAspName("M3212RA4_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3212RA4_KO441", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam,document), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If strRet(0,0) = "" Then
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetLLCRef(strRet)
	End If	
		
End Function
'==============================================================================================================================
Function SetLLCRef(strRet)

    Dim Index1,Count1,Row1
    Dim temp
	
	Const C_LcNo_Ref		= 0
	Const C_LCSeq_Ref		= 1
	Const C_PlantCd_Ref		= 2
	Const C_SLCd_Ref		= 3
	Const C_ItemCd_Ref		= 4
	Const C_ItemNm_Ref		= 5
	Const C_Spec_Ref		= 6
	Const C_TrackingNo_Ref	= 7	
	Const C_Qty_Ref			= 8
	Const C_Unit_Ref		= 9
	Const C_OpenDt_Ref		= 10
	Const C_PoNo_Ref		= 11
	Const C_PoSeq_Ref		= 12
	Const C_SLNm_Ref		= 13
	Const C_RcptQty_Ref		= 14
	Const C_Price_Ref		= 15
	Const C_Cur_Ref			= 16
	Const C_PlantNm_Ref		= 17
	Const C_InspFlg_Ref		= 18
	Const C_InspMeth_Ref	= 19
	Const C_InspMethCd_Ref	= 20
	Const C_Pur_Grp_Ref		= 21
	Const C_PreIvQty_Ref    = 22
    Const C_Lot_flg			= 23
    Const C_Lot_gen_mtd		= 24
    
	Count1 = Ubound(strRet,1)
	
	With frm1.vspdData
		
		.Redraw = False
	
		Call fncinsertrow(Count1 + 1)
		
		For index1 = 0 to Count1
				
			Row1 = .ActiveRow + Index1
				
			temp = unicdbl(strRet(index1,C_Qty_Ref)) - (unicdbl(strRet(index1,C_RcptQty_Ref)) + unicdbl(strRet(index1,C_PreIvQty_Ref)))
				
			Call .SetText(C_PlantCd,	Row1, strRet(index1,C_PlantCd_Ref))
			Call .SetText(C_PlantNm,	Row1, strRet(index1,C_PlantNm_Ref))
			Call .SetText(C_itemCd,		Row1, strRet(index1,C_ItemCd_Ref))
			Call .SetText(C_itemNm,		Row1, strRet(index1,C_ItemNm_Ref))
			Call .SetText(C_Spec,		Row1, strRet(index1,C_Spec_Ref))
			Call .SetText(C_TrackingNo, Row1, strRet(index1,C_TrackingNo_Ref))
			Call .SetText(C_GRQty,		Row1, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			Call .SetText(C_GRUnit,		Row1, strRet(index1,C_Unit_Ref))
			Call .SetText(C_Cur,		Row1, strRet(index1,C_Cur_Ref))
			Call .SetText(C_MvmtPrc,	Row1, 0)
			Call .SetText(C_DocAmt,		Row1, 0)
			Call .SetText(C_LocAmt,		Row1, 0)
			Call .SetText(C_SLCd,		Row1, strRet(index1,C_SLCd_Ref))
			Call .SetText(C_SLNm,		Row1, strRet(index1,C_SLNm_Ref))
			Call .SetText(C_PoNo,		Row1, strRet(index1,C_PoNo_Ref))
			Call .SetText(C_PoSeqNo,	Row1, strRet(index1,C_PoSeq_Ref))
			Call .SetText(C_InspMethCd, Row1, strRet(index1,C_InspMethCd_Ref))
			Call .SetText(C_Stateflg,	Row1, "LC")
			Call .SetText(C_LLCNo, Row1, strRet(index1,C_LcNo_Ref))
			Call .SetText(C_LLCSeqNo, Row1, strRet(index1,C_LCSeq_Ref))

			.Row = Row1	
			.Col = C_InspFlg										
			If strRet(index1,C_InspFlg_Ref) = "Y" Then
				Call .SetText(C_InspFlg,	Row1, "1")
				Call .SetText(C_GRMeth,		Row1, strRet(index1,C_InspMeth_Ref))
			Else
				Call .SetText(C_InspFlg,	Row1, "0")
				Call .SetText(C_GRMeth,		Row1, "")
			End if
			
			If UCase(Trim(strRet(index1,C_Lot_flg))) = "Y" and UCase(Trim(strRet(index1,C_Lot_gen_mtd))) = "M" Then
				ggoSpread.spreadUnlock	C_LotNo, Row1, C_LotSeqNo, Row1
				ggoSpread.SSSetRequired	C_LotNo, Row1, Row1
			ElseIf UCase(Trim(strRet(index1,C_Lot_flg))) = "N" Then
				Call .SetText(C_LotNo,	Row1, "*")
			End If

				'C_Pur_Grp_Ref
			IF index1 = 0 Then
				frm1.txtGroupCd.value = strRet(index1,C_Pur_Grp_Ref)
				Call changeGroupCd()                                            '20071211::hanc
		
			End IF
			
		Next
		
		Call LocalReFormatSpreadCellByCellByCurrency()
		Call setReference()
		
		.ReDraw = True
		
	End with
	
End Function
'==============================================================================================================================
Function OpenIvRef()

	Dim strRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD

	If lblnWinEvent = True Then Exit Function
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","매입내역참조" )
		Exit Function
	End if 

	if Trim(frm1.txtMvmtType.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "입고형태","X")
		frm1.txtMvmtType.focus
		Set gActiveElement = document.activeElement
		Exit Function	
	elseif Trim(frm1.txtSupplierCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공급처","X")
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	
	if (UCase(Trim(frm1.hdnImportflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="Y") or _
	   (UCase(Trim(frm1.hdnSubcontraflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="N") then
		Call DisplayMsgBox("17A012", "X","입고형태" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","매입내역참조" )
		Exit Function
	End if
		
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = Trim(frm1.txtGroupNm.value)
	
	iCalledAspName = AskPRAspName("M5112RA1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5112RA1_KO441", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.	
	If strRet(0,0) = "" Then
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else	
		Call SetIvRef(strRet)
	End If	
		
End Function
'==============================================================================================================================
Function SetIvRef(strRet)

	Dim Index1, Count1, Row1
    Dim temp
    
	Const C_IvNo_Ref		= 0
	Const C_IvSeqNo_Ref		= 1
	Const C_PlantCd_Ref		= 2
	Const C_PlantNm_Ref		= 3	
	Const C_ItemCd_Ref		= 4
	Const C_ItemNm_Ref		= 5
	Const C_Spec_Ref		= 6		
	Const C_IvQty_Ref		= 7	
	Const C_MvmtQty_Ref		= 8
	Const C_IvUnit_Ref		= 9	
	Const C_IvPrc_Ref		= 10
	Const C_IvDocAmt_Ref	= 11
	Const C_IvCur_Ref		= 12
	Const C_PONo_Ref		= 13
	Const C_POSeqNo_Ref		= 14
	Const C_TrackingNo_Ref	= 15
	Const C_SlCd_Ref		= 16
	Const C_SLNm_Ref		= 17
	Const C_InspMethCd_Ref	= 18
	Const C_InspMethNm_Ref	= 19
	Const C_InspFlg_Ref		= 20
	Const C_Pur_Grp_Ref		= 21
	Const C_InspectQty_Ref  = 22
    Const C_Lot_flg			= 23
    Const C_Lot_gen_mtd		= 24
	
	Count1 = Ubound(strRet,1)
	
	With frm1.vspdData
		
		.Redraw = False
	
		Call fncinsertrow(Count1 + 1)
		
		For index1 = 0 to Count1
				
			Row1 = .ActiveRow + Index1

			temp = UNICDbl(strRet(index1,C_IvQty_Ref)) - (UNICDbl(strRet(index1,C_MvmtQty_Ref)) + unicdbl(strRet(index1,C_InspectQty_Ref)))
				
			Call .SetText(C_PlantCd,	Row1, strRet(index1,C_PlantCd_Ref))
			Call .SetText(C_PlantNm,	Row1, strRet(index1,C_PlantNm_Ref))
			Call .SetText(C_itemCd,		Row1, strRet(index1,C_ItemCd_Ref))
			Call .SetText(C_itemNm,		Row1, strRet(index1,C_ItemNm_Ref))
			Call .SetText(C_Spec,		Row1, strRet(index1,C_Spec_Ref))
			Call .SetText(C_TrackingNo, Row1, strRet(index1,C_TrackingNo_Ref))
			Call .SetText(C_GRQty,		Row1, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			Call .SetText(C_GRUnit,		Row1, strRet(index1,C_IvUnit_Ref))
			Call .SetText(C_Cur,		Row1, strRet(index1,C_IvCur_Ref))
			Call .SetText(C_MvmtPrc,	Row1, 0)
			Call .SetText(C_DocAmt,		Row1, 0)
			Call .SetText(C_LocAmt,		Row1, 0)
			Call .SetText(C_SLCd,		Row1, strRet(index1,C_SLCd_Ref))
			Call .SetText(C_SLNm,		Row1, strRet(index1,C_SLNm_Ref))
			Call .SetText(C_PoNo,		Row1, strRet(index1,C_PoNo_Ref))
			Call .SetText(C_PoSeqNo,	Row1, strRet(index1,C_POSeqNo_Ref))
			Call .SetText(C_InspMethCd, Row1, strRet(index1,C_InspMethCd_Ref))
			Call .SetText(C_Stateflg,	Row1, "IV")
			Call .SetText(C_IvNo, Row1, strRet(index1,C_IvNo_Ref))
			Call .SetText(C_IvSeqNo, Row1, strRet(index1,C_IvSeqNo_Ref))
			
			.Row = Row1	
			.Col = C_InspFlg										
			If strRet(index1,C_InspFlg_Ref) = "Y" Then
				Call .SetText(C_InspFlg,	Row1, "1")
				Call .SetText(C_GRMeth,		Row1, strRet(index1,C_InspMethNm_Ref))
			Else
				Call .SetText(C_InspFlg,	Row1, "0")
				Call .SetText(C_GRMeth,		Row1, "")
			End if

			If UCase(Trim(strRet(index1,C_Lot_flg))) = "Y" and UCase(Trim(strRet(index1,C_Lot_gen_mtd))) = "M" Then
				ggoSpread.spreadUnlock	C_LotNo, Row1, C_LotSeqNo, Row1
				ggoSpread.SSSetRequired	C_LotNo, Row1, Row1
			ElseIf UCase(Trim(strRet(index1,C_Lot_flg))) = "N" Then
				Call .SetText(C_LotNo,	Row1, "*")
			End If
			
				'C_Pur_Grp_Ref
			IF index1 = 0 Then
				frm1.txtGroupCd.value = strRet(index1,C_Pur_Grp_Ref)		
				Call changeGroupCd()                                            '20071211::hanc
			End IF
		Next
	
		Call LocalReFormatSpreadCellByCellByCurrency()
		Call setReference()
		
		.ReDraw = True
		
	End with
	
End Function
'==============================================================================================================================
Function OpenSlCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	Dim iCurRow
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCurRow = frm1.vspdData.ActiveRow
	
	arrParam(0) = "창고"						
	arrParam(1) = "B_STORAGE_LOCATION"			
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_SlCd,iCurRow,"X","X"))
'	arrParam(3) = Trim(frm1.txtSlNm.Value)
	arrParam(4) = "PLANT_CD= " & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,iCurRow,"X","X")), " " , "S") & " "
	arrParam(5) = "창고"						
	
    arrField(0) = "SL_CD"						
    arrField(1) = "SL_NM"						
    
    arrHeader(0) = "창고"					
    arrHeader(1) = "창고명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) <> "" Then
		Call frm1.vspdData.SetText(C_SlCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_SlNm,	iCurRow, arrRet(1))
	End If	
	
End Function


'2008-05-16 3:13오후 :: hanc
Function FncCopy() 
    Dim SumTotal,tmpGrossAmt
    if frm1.vspdData.Maxrows < 1	then exit function
    ggoSpread.Source = frm1.vspdData	
    
    ggoSpread.CopyRow
    
    frm1.vspdData.ReDraw = False
    
    Call SetSpreadColor(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)
    
    frm1.vspdData.ReDraw = True
    
        
'  	frm1.vspdData.Row = frm1.vspdData.ActiveRow
'    frm1.vspdData.Col = C_TrackingNo
'	
'	
' 	if Trim(frm1.vspdData.Text) = "*" then
'		ggoSpread.spreadlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
'	else
'	    ggoSpread.spreadUnlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
'		ggoSpread.sssetrequired C_TrackingNo, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
'	end if
	    
    frm1.vspdData.ReDraw = True
End Function

'==============================================================================================================================
Function OpenMvmtType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtMvmtType.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "입고형태"	
	arrParam(1) = "( select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b where a.rcpt_type = b.io_type_cd    and a.sto_flg = " & FilterVar("N", "''", "S") & "  AND a.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and ((b.RCPT_FLG=" & FilterVar("Y", "''", "S") & "  AND b.RET_FLG=" & FilterVar("N", "''", "S") & " ) or (b.RET_FLG=" & FilterVar("N", "''", "S") & "  And b.SUBCONTRA_FLG=" & FilterVar("N", "''", "S") & " )) ) c"
	arrParam(2) = Trim(frm1.txtMvmtType.Value)
	'arrParam(4) = "((RCPT_FLG='Y' AND RET_FLG='N') or (RET_FLG='N' And SUBCONTRA_FLG='N')) AND USAGE_FLG='Y' "
	arrParam(5) = "입고형태"			
	
    arrField(0) = "IO_Type_Cd"
    arrField(1) = "IO_Type_NM"
    
    arrHeader(0) = "입고형태"		
    arrHeader(1) = "입고형태명"
    
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
Function OpenMvmtNo()
	
		Dim strRet
		Dim arrParam(3)
		Dim iCalledAspName
		Dim IntRetCD
	
		If IsOpenPop = True Or UCase(frm1.txtMvmtNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
		IsOpenPop = True

		arrParam(0) = ""'Trim(frm1.hdnSupplierCd.Value)
		arrParam(1) = ""'Trim(frm1.hdnGroupCd.Value)
		arrParam(2) = ""'Trim(frm1.hdnMvmtType.Value)		
		arrParam(3) = ""'This is for Inspection check, must be nothing.
		
		iCalledAspName = AskPRAspName("M4111PA3_ko441")   '20080214::hanc
	    
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4111PA3_ko441", "X")
			IsOpenPop = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
		IsOpenPop = False
		
		If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
		
		If strRet(0) = "" Then
			frm1.txtMvmtNo.focus	
			Set gActiveElement = document.activeElement
			Exit Function
		Else
			frm1.txtMvmtNo.value = strRet(0)
			frm1.txtMvmtNo.focus	
			frm1.hPoNo.value = strRet(1)        'HANC
			Set gActiveElement = document.activeElement
		End If	
		
End Function
'==============================================================================================================================
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
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
Function OpenMvmtRcpt(lgCurRow)
	
	Dim strRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	
	arrParam(0) = Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lgCurRow,"X","X"))
	arrParam(1) = Trim(GetSpreadText(frm1.vspdData,C_PlantNm,lgCurRow,"X","X"))
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lgCurRow,"X","X"))
	arrParam(3) = Trim(GetSpreadText(frm1.vspdData,C_ItemNm,lgCurRow,"X","X"))
	arrParam(4) = UCase(Trim(frm1.txtSupplierCd.value))
	arrParam(5) = UCase(Trim(frm1.txtSupplierNm.value))
	
	iCalledAspName = AskPRAspName("M4111PA5")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4111PA5", "X")
		IsOpenPop = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If strRet(0) <> "" Then
		'Call SetMvmtNo(strRet)
	End If	
		
End Function
'------------------------------------------  OpenBackFlushRef()  -----------------------------------------
'	Name : OpenBackFlushRef()
'	Description : BackFlush Simmulation Reference
'---------------------------------------------------------------------------------------------------------
Function OpenBackFlushRef()
	
	Dim arrRet,arrParam_1(1)
	Dim strRet
	Dim IntRows
	Dim strVal
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	strVal = ""
	
	With frm1.vspdData
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = C_GrQty		' Produced Qty
			If UNICDbl(.Text) > CDbl(0) Then
				.Col = C_PlantCd	
				strVal = strVal & UCase(Trim(.Text)) & parent.gColSep
				.Col = C_PoNo			
				strVal = strVal & UCase(Trim(.Text)) & parent.gColSep
				.Col = C_PoSeqNo		
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_GrQty
				strVal = strVal & UniConvNum(.Text,0) & parent.gRowSep
			End If
		Next
	End With

	' 2005.06.15 추가 입고내역이 없을 경우 조회 안되도록 수정 
	If CDbl(frm1.vspdData.MaxRows) = 0 Then
		Call DisplayMsgBox("174100","X","X","X")		
		Exit Function
	End If

	arrParam_1(0) = strVal
	arrParam_1(1) = UCase(Trim(frm1.txtSupplierCd.Value))

	iCalledAspName = AskPRAspName("m4111ra5")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m4111ra5", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam_1), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'==============================================================================================================================
Sub CurFormatNumSprSheet()

	With frm1
		ggoSpread.Source = frm1.vspdData
		'입고금액(화폐단위가 있을시에만 포텟팅)
		
		if .hdnMvmtCur.value <> "" Then
		ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .hdnMvmtCur.Value,  Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		End if
		'입고자국금액 
		ggoSpread.SSSetFloatByCellOfCur C_LocAmt,-1, parent.gCurrency,  Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
	End With
End Sub
'==============================================================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
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
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6"				  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","99999"
    End Select
         
End Sub
'==============================================================================================================================
Function setReference()

	ggoOper.SetReqAttr	frm1.txtMvmtType, "Q"
	ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"
	Call SetToolBar("11101001001111")
End Function
'==============================================================================================================================
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
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
   
	ggoSpread.Source = frm1.vspdData
    
	If Col = C_SlCdPop Then
    	Call OpenSlCd()
    ElseIf Col = C_MvmtRcptLookUpBtn Then
		Call OpenMvmtRcpt(frm1.vspdData.ActiveRow)
	End If	
    
End Sub
'==============================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==============================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
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
    
    Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData, -1, -1 ,C_Cur			,C_MvmtPrc	,"C" ,"I","X","X")
    Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData, -1, -1 ,C_Cur			,C_DocAmt	,"A" ,"I","X","X")
    Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData, -1, -1 ,parent.gCurrency	,C_LocAmt	,"A" ,"I","X","X")
    Call ChangeTag(True)
    
End Sub
'==============================================================================================================================
Sub txtGmDt_DblClick(Button)
	if Button = 1 then
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
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    
    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
    
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)
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
    
    On Error Resume Next                                                 
    Err.Clear                                               
    
    If Trim(frm1.txtMvmtNo.value) = "" Then
        Call DisplayMsgBox("ZZ0004","X","X","X")
       Exit function
    End If

	ggoSpread.Source = frm1.vspdData
	
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
    ggoSpread.ClearSpreadData        
    Call InitVariables

    If Not chkFieldByCell(frm1.txtMvmtNo, "A",1) Then Exit Function
   
    If 	CommonQueryRs(" IsNull(DLVY_ORD_FLG, ''), B.ret_flg "," M_PUR_GOODS_MVMT A inner join M_MVMT_TYPE B ON (A.io_type_cd = B.io_type_cd) ", " MVMT_RCPT_NO = " & FilterVar(frm1.txtMvmtNo.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
		Call DisplayMsgBox("174100","X","X","X")
		Call SetDefaultVal
		lgBlnFlgChgValue = false
		'frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Call ggoOper.ClearField(Document, "2") 'Clear field only when error rise.
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1, Chr(11))

	'-- Modify for Isuue 8867 by Byun Jee Hyun
	If Trim(lgF0(0)) = "Y" Then
		Call DisplayMsgBox("17a014", "X","납입지시","조회" )
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Call ggoOper.ClearField(Document, "2") 'Clear field only when error rise.
		Exit function
	End If 

	If Trim(lgF1(0)) = "Y" Then
		Call DisplayMsgBox("17a014", "X","반품","조회" )
		frm1.txtMvmtNo.focus
		Set gActiveElement = document.activeElement
		Call ggoOper.ClearField(Document, "2") 'Clear field only when error rise.
		Exit function
	End If
	'-- End of Isuue 8867


	If Trim(frm1.hPoNo.value) = "" Then
        If DbQuery_e = False Then Exit Function     '20080214::hanc
	Else
        If DbQuery = False Then Exit Function     '20080214::hanc
	End If
        
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
    
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.ClearSpreadData                  
    Call SetDefaultVal
    Call InitVariables
        
    FncNew = True                     
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

    'If Not chkField(Document, "2") Then	Exit Function
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkFieldByCell(frm1.txtMvmtType, "A",1)	then
       Exit Function
    End If
    
    If Not chkFieldByCell(frm1.txtGmDt, "A",1)	then
       Exit Function
    End If
    
    If Not chkFieldByCell(frm1.txtSupplierCd, "A",1)	then
       Exit Function
    End If
    
	With frm1
		if CompareDateByFormat(.txtGmDt.text,EndDate,.txtGmDt.Alt,"현재일", _
                   "970025",.txtGmDt.UserDefinedFormat,Parent.gComDateType,True) = False Then	
			Exit Function
		End if   

	    '------------------------------------------------------------------------
		'SR 번호 : 10198
	    '내   용 : 입고시 입고번호 수동일 경우 기 입고번호가 있으면 에러 메세지 처리 
	    'date    : 2005/08/17
	    'Modifier: Kim Duk Hyun 
	    '------------------------------------------------------------------------
		.vspdData.Col = 0  
		If .vspdData.Text <> ggoSpread.DeleteFlag Then
			If .txtMvmtNo1.Value <> "" Then
				If 	CommonQueryRs(" MVMT_RCPT_NO "," M_PUR_GOODS_MVMT ", " MVMT_RCPT_NO = " & FilterVar(.txtMvmtNo1.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
							
					Call DisplayMsgBox("174112","X",.txtMvmtNo1.Value,"X")
					.txtMvmtNo1.Value = ""
					Call SetFocusToDocument("M") 
					.txtMvmtNo1.focus
					Exit function
				End If
			End If
		End If
                      
		ggoSpread.Source = .vspdData									
		If Not ggoSpread.SSDefaultCheck Then Exit Function
    
		If .vspdData.Maxrows < 1 then Exit Function
		
		For intIndex = 1 to .vspdData.MaxCols 
			.vspdData.SetColItemData intindex,0	
		Next
	End with
	    
    frm1.vspdData.Col = C_PoNo 
    
    'dbsave_e 와 같이 _e를 붙인 이유 : po_no가 없는 것은 프로세스가 다르다. 따라서 이화면(구매입고)에 없는 프로세스인 po_no가 없는 프로세스를 추가하여야 하였다.

	If frm1.vspdData.Text = "" Then
        If DbSave_e = False Then Exit Function      '20080213::hanc
	Else
        If DbSave   = False Then Exit Function      '20080213::hanc
	End If

    
    FncSave = True                                                      
    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncCancel()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    	
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
    
    Set gActiveElement = document.ActiveElement                                                   
End Function
'==============================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
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
	
	With frm1

		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow , imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
    End With
	
	If Err.number = 0 Then FncInsertRow = True
	Set gActiveElement = document.ActiveElement
	
End Function
'==============================================================================================================================
Function LocalReFormatSpreadCellByCellByCurrency() 
	On Error Resume Next
	Err.Clear
	
	With frm1
		
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,-1, -1,C_Cur,C_MvmtPrc,		"C" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,-1, -1,C_Cur,C_DocAmt,		"A" ,"I","X","X") 
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,-1, -1,parent.gCurrency,C_LocAmt,	"A" ,"I","X","X") 
		
    End With
	
	Set gActiveElement = document.ActiveElement
	
End Function
'==============================================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    ggoSpread.Source = frm1.vspdData
    if frm1.vspdData.Maxrows < 1	then exit function
    
    frm1.vspdData .focus
	ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow
	
    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncPrint()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncExcel()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_SINGLEMULTI)		
    Set gActiveElement = document.ActiveElement   						
End Function
'==============================================================================================================================
Function FncFind() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False)  
    Set gActiveElement = document.ActiveElement                                 
End Function
'==============================================================================================================================
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
    Set gActiveElement = document.ActiveElement   
End Function
'20080214::hanc==============================================================================================================================
Function DbQuery_e() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey     
    Dim strVal
        
    DbQuery_e = False
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
    
    If LayerShowHide(1) = False Then Exit Function
    
	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID1 & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtGrNo=" & .hdnMvmtNo.value
		else
		    strVal = BIZ_PGM_ID1 & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtGrNo=" & Trim(.txtMvmtNo.value)
		End if
  
		Call RunMyBizASP(MyBizASP, strVal)							
    End With
    
    DbQuery_e = True
	Set gActiveElement = document.ActiveElement   
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
    Err.Clear                                                                     '☜: Clear error status
  
    If LayerShowHide(1) = False Then Exit Function
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtRcptNo=" & .hdnRcptNo.value
		    strVal = strVal & "&txtMvmtNo=" & .hdnMvmtNo.value
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtMvmtNo=" & Trim(.txtMvmtNo.value)
		End if
    
		Call RunMyBizASP(MyBizASP, strVal)									
    End With
    
    DbQuery = True
End Function
'==============================================================================================================================
Function DbQueryOk()													
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	lgIntFlgMode = Parent.OPMD_UMODE											
'    Call ggoOper.LockField(Document, "Q")								
	lgBlnFlgChgValue = False	
	
	Call SetToolBar("11101011000111")
	Call RemovedivTextArea
	Call ChangeTag(True)
	
	if interface_Account = "N" then		
		frm1.btnGlSel.disabled = true
	Else 
		frm1.btnGlSel.disabled = False		
	End if
	frm1.btnGlSel_1.disabled = true
	
	frm1.vspdData.focus
End Function

'20080213::hanc==============================================================================================================================
Function DbSave_e()
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
	
    DbSave_e = False                                                          
    
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
'					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Unit,lRow, "X","X"))					& iColSep 
					strVal = strVal & "EA"					& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Cur,lRow, "X","X"))					& iColSep 
'단가 : 확인후 꼭 넣도록 한다.					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_MvmtPrc,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & 10		& iColSep 
'금액 : 확인 후 꼭 넣도록 한다.					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_DocAmt,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & 500		& iColSep 
'가공비단가					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_WorkPrc,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & 10		& iColSep 
'가공비금액					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_WorkLocAmt,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & 0		& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SlCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotSeqNo,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MakerLotNo,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MakerLotSeqNo,lRow, "X","X"))			& iColSep 
'반품유형					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RetType,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim("")				& iColSep 
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
'출고번호					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RefMvmtNo,lRow, "X","X"))				& iColSep 
					strVal = strVal & ""				& iColSep 
'조달구분					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProcurType,lRow, "X","X"))				& iColSep 
					strVal = strVal & "P"				& iColSep 
'					strVal = strVal & Trim(frm1.hdnGrInspType.value)											& iColSep
					strVal = strVal & "A"											& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RemarkDtl,lRow, "X","X"))				& iColSep 
					strVal = strVal & lRow & iColSep
                    '2008-03-29 7:37오후 :: hanc
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_TransTime  ,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MainLot    ,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ImportTime ,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_CreateType ,lRow, "X","X"))				& iRowSep 
				Case ggoSpread.DeleteFlag
					strDel = "D" & iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X"))					& iColSep
'출고번호					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_RefMvmtNo,lRow, "X","X"))				& iColSep 
					strDel = strDel & ""				& iColSep 
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
	Call ExecMyBizASP(frm1, BIZ_PGM_ID1)

	If Err.number = 0 Then	 
	   DbSave = True                                                             '☜: Processing is OK
	End If

	Set gActiveElement = document.ActiveElement         
End Function

'==============================================================================================================================
Function DbSave() 

    Dim lRow        
    Dim strVal, strDel
	Dim iColSep, iRowSep
	
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

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
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
	
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
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					If Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"))) = "" Or Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"))) = "0" then 
						Call DisplayMsgBox("970021","X","입고수량","X")
						Call RemovedivTextArea
						Call LayerShowHide(0)
						Exit Function
					End if

					strVal = "C"																					& iColSep				
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow, "X","X"))					& iColSep & iColSep
        			strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow, "X","X"))						& iColSep & iColSep & iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_TrackingNo,lRow, "X","X"))					& iColSep
					If Trim(GetSpreadText(frm1.vspdData,C_InspFlg,lRow, "X","X")) = "0" Then
						strVal = strVal & "N"																		& iColSep
					Else
						strVal = strVal & "Y"																		& iColSep
					End If
					If Trim(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X")) = "" Then
						strVal = strVal & "0"																		& iColSep & iColSep
					Else
						strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"),0)			& iColSep & iColSep
					End If
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GRUnit,lRow, "X","X"))						& iColSep & iColSep & iColSep & iColSep & iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SlCd,lRow, "X","X"))						& iColSep & iColSep & iColSep & iColSep & iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotNo,lRow, "X","X"))						& iColSep
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(frm1.vspdData,C_LotSeqNo,lRow, "X","X")), 0)	& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MakerLotNo,lRow, "X","X"))					& iColSep
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(frm1.vspdData,C_MakerLotSeqNo,lRow, "X","X")), 0)		& iColSep & iColSep & iColSep & iColSep & iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PoNo,lRow, "X","X"))						& iColSep
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(frm1.vspdData,C_PoSeqNo,lRow, "X","X")), 0)		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_CCNo,lRow, "X","X"))						& iColSep
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(frm1.vspdData,C_CCSeqNo,lRow, "X","X")), 0)		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LLCNo,lRow, "X","X"))						& iColSep
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(frm1.vspdData,C_LLCSeqNo,lRow, "X","X")), 0)	& iColSep & iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_InspMethCd,lRow, "X","X"))					& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X"))						& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_IvNo,lRow, "X","X"))						& iColSep
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(frm1.vspdData,C_IvSeqNo,lRow, "X","X")), 0)		& iColSep
					strVal = strVal & lRow & iColSep

					'2005.12.16 Remark_Dtl 추가
					strVal = strVal & iColSep & iColSep & iColSep & iColSep & iColSep
					strVal = strVal & iColSep & iColSep & iColSep & iColSep & iColSep
					strVal = strVal & iColSep & iColSep & iColSep
					'2007.01.16 검사결과등록을 위해 인자값 추가 
					If Trim(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X")) = "" Then
						strVal = strVal & "0"																		& iColSep & iColSep
					Else
						strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"),0)			& iColSep & iColSep
					End If
					'===========================================
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RemarkDtl,lRow, "X","X"))					& iColSep
					'2008-03-29 7:33오후 :: hanc
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_TransTime  ,lRow, "X","X"))						& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MainLot    ,lRow, "X","X"))						& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ImportTime ,lRow, "X","X"))						& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_CreateType ,lRow, "X","X"))						& iRowSep

				Case ggoSpread.DeleteFlag
					strDel = "D" & iColSep   '0
					strDel = strDel & iColSep & iColSep & iColSep & iColSep & iColSep  ' 1 - 5
					strDel = strDel & iColSep & iColSep & iColSep & iColSep & iColSep  ' 6 - 10
					strDel = strDel & iColSep & iColSep & iColSep & iColSep & iColSep  ' 11 - 15
					strDel = strDel & iColSep & iColSep & iColSep & iColSep & iColSep  ' 16 - 20
					strDel = strDel & iColSep & iColSep & iColSep & iColSep & iColSep  ' 21 - 25
					strDel = strDel & iColSep & iColSep & iColSep & iColSep & iColSep  ' 26 - 30
					strDel = strDel & iColSep & iColSep & iColSep & iColSep & iColSep  ' 31 - 35
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X")) & iColSep & iColSep & iColSep
					If Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_RetOrdQty,lRow, "X","X"))) >  "0" then 
						Call DisplayMsgBox("172126","X",lRow & "행","X")
						Call RemovedivTextArea
						Call LayerShowHide(0)
						Exit Function
					End if
					strDel = strDel & lRow & iRowSep	'39
					
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
	'------ Developer Coding part (End ) -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	If Err.number = 0 Then	 
	   DbSave = True                                                             '☜: Processing is OK
	End If

	Set gActiveElement = document.ActiveElement 
End Function

'==============================================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
	Call InitVariables
	Call MainQuery()
End Function
'==============================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'==============================================================================================================================
Function changeMvmtType()

    changeMvmtType = False                 

	With frm1
		If 	CommonQueryRs(" A.IO_TYPE_NM, A.RCPT_FLG, A.IMPORT_FLG, A.RET_FLG, B.SUBCONTRA_FLG ", _
					" M_MVMT_TYPE A, M_CONFIG_PROCESS B ", _
					" A.IO_TYPE_CD = B.RCPT_TYPE AND B.STO_FLG = " & FilterVar("N", "''", "S") & "  AND B.USAGE_FLG= " & FilterVar("Y", "''", "S") & "  AND (A.RET_FLG = " & FilterVar("N", "''", "S") & "   AND (A.RCPT_FLG = " & FilterVar("Y", "''", "S") & "  OR A.SUBCONTRA_FLG = " & FilterVar("N", "''", "S") & " )) AND A.USAGE_FLG = " & FilterVar("Y", "''", "S") & "  AND A.IO_TYPE_CD = " & FilterVar(.txtMvmtType.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("171900","X","X","X")
			.txtMvmtTypeNm.Value = ""
			Call SetFocusToDocument("M") 
			.btnGlSel_1.disabled = true
			.txtMvmtType.focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		lgF3 = Split(lgF3, Chr(11))
		lgF4 = Split(lgF4, Chr(11))
		
		.txtMvmtTypeNm.Value	= lgF0(0)
		.hdnRcptflg.Value 		= lgF1(0)
		.hdnImportflg.Value		= lgF2(0)
		.hdnRetflg.Value 		= lgF3(0)
		.hdnSubcontraflg.Value  = lgF4(0)
		
		If .hdnSubcontraflg.value = "Y" Then
			.btnGlSel_1.disabled = false
		Else
			.btnGlSel_1.disabled = true
		End If

	End With

	lgBlnFlgChgValue = true
    
    changeMvmtType = True                  

End Function
'==============================================================================================================================
Function changeSpplCd()

	With frm1
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
'==============================================================================================================================
Function RemovedivTextArea()
	Dim ii
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function
'==============================================================================================================================
Function changeGroupCd()

	changeGroupCd = False
	
	With frm1
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
