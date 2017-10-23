Const BIZ_PGM_QRY_ID  = "i1611mb1_ko366.asp"                                                                                           
Const C_MaxKey        = 10

Dim C_ItemCd
Dim C_ItemNm
Dim C_ItemAcct
Dim C_SlCd
Dim C_CCCd
Dim C_DocumentDt
Dim C_OrderUnit
Dim C_BaseUnit
Dim C_Qty
Dim C_Price
Dim C_Amount
Dim C_CostOfDevy
Dim C_TrackingNo

Dim C_ProjectNm
Dim C_StructNm

Dim C_LotNo
Dim C_LotSubNo
Dim C_TrnsType
Dim C_MovType
Dim C_DocumentNo
Dim C_SeqNo
Dim C_SoNo
Dim C_SoSeq
Dim C_PoNo
Dim C_PoSeq
Dim C_ProdtNo
Dim C_BpCd
Dim C_BpNm
Dim C_TrnsSlCd
Dim C_WcCd
Dim C_WcNm
Dim C_DocumentText
Dim C_Spec
Dim C_DebitCreditFlag
Dim C_SubSeqNo
Dim C_TempGlNo
Dim C_GlNo



Dim lgIsOpenPop
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4

 '==========================================  2.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    
    lgBlnFlgChgValue = False                             
    lgStrPrevKey = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""
    lgLngCurRows = 0                                
    lgSortKey = 1
End Sub



'==========================================  2.2 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtFrSLCd.focus
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
        	frm1.txtPlantCd.value = lgPLCd
	End If

	frm1.txtTrnsFrDt.Text = FromDate
	frm1.txtTrnsToDt.Text = StartDate                           
End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()
    With frm1
    
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20050101", , Parent.gAllowDragDropSpread
    
		.vspdData.ReDraw = false
		.vspdData.MaxCols = C_GlNo + 1
		.vspdData.MaxRows = 0
    
		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "3", "0")
    
		ggoSpread.SSSetEdit C_ItemCd, "품목", 18,,,,2
		ggoSpread.SSSetEdit C_ItemNm, "품목명", 25
'20071221::hanc		ggoSpread.SSSetEdit C_ItemAcct, "품목계정", 8
		ggoSpread.SSSetEdit C_SlCd, "창고", 6
'20071221::hanc		ggoSpread.SSSetEdit C_CCCd, "Cost CD", 6
		ggoSpread.SSSetDate C_DocumentDt, "수불일자", 10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit C_OrderUnit, "Order 단위", 10
		ggoSpread.SSSetEdit C_BaseUnit, "단위", 5
		ggoSpread.SSSetFloat C_Qty, "수량", 15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_Price, "단가", 15,Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_Amount, "금액", 15,Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_CostOfDevy, "부대비", 15,Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No.", 25
		
'20071221::hanc		ggoSpread.SSSetEdit C_ProjectNm, "현장명", 10
'20071221::hanc		ggoSpread.SSSetEdit C_StructNm, "건설사명", 10
		
		ggoSpread.SSSetEdit C_LotNo, "Lot No.", 18
		ggoSpread.SSSetFloat C_LotSubNo, "Lot 순번", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit C_TrnsType, "수불구분", 10
		ggoSpread.SSSetEdit C_MovType, "이동유형", 10
		ggoSpread.SSSetEdit C_DocumentNo, "수불번호", 10
		ggoSpread.SSSetFloat C_SeqNo, "수불상세", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit C_SoNo, "수주번호", 10
		ggoSpread.SSSetFloat C_SoSeq, "수주상세", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit C_PoNo, "구매번호", 10
		ggoSpread.SSSetFloat C_PoSeq, "구매상세", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit C_ProdtNo, "제조오더번호", 15
		ggoSpread.SSSetEdit C_Spec, "규격", 25
		ggoSpread.SSSetEdit C_DebitCreditFlag, "재고증감구분", 15
		ggoSpread.SSSetEdit C_BpCd, "거래처", 10
		ggoSpread.SSSetEdit C_BpNm, "거래처명", 10
		ggoSpread.SSSetEdit C_TrnsSlCd, "이동창고", 10
		ggoSpread.SSSetEdit C_WcCd, "작업장", 10
		ggoSpread.SSSetEdit C_WcNm, "작업장명", 10
		ggoSpread.SSSetEdit C_DocumentText, "비고", 10
		ggoSpread.SSSetEdit C_SubSeqNo, "수불상세 SubNo.", 10

		ggoSpread.SSSetEdit C_TempGlNo, "결의전표번호", 15
		ggoSpread.SSSetEdit C_GlNo, "회계전표번호", 15
		
		
	
		
			Call ggoSpread.SSSetColHidden(C_BpCd,C_WcNm, True)
			Call ggoSpread.SSSetColHidden(C_SubSeqNo,C_SubSeqNo, True)
			Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
			.vspdData.ReDraw = true
		ggoSpread.SSSetSplit2(2)
		
   		ggoSpread.Source = .vspdData
   		
   	End With 
	
	Call SetSpreadLock()
	
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	
'20071221::hanc	C_ItemCd = 1
'20071221::hanc	C_ItemNm = 2
'20071221::hanc	C_ItemAcct = 3
'20071221::hanc	C_SlCd = 4
'20071221::hanc	C_CCCd = 5
'20071221::hanc	C_DocumentDt = 6
'20071221::hanc	C_OrderUnit = 7
'20071221::hanc	C_BaseUnit = 8
'20071221::hanc	C_Qty = 9
'20071221::hanc	C_Price = 10
'20071221::hanc	C_Amount = 11
'20071221::hanc	C_CostOfDevy = 12
'20071221::hanc	C_TrackingNo = 13
'20071221::hanc	C_ProjectNm = 14
'20071221::hanc	C_StructNm = 15
'20071221::hanc	C_LotNo = 16
'20071221::hanc	C_LotSubNo = 17
'20071221::hanc	C_TrnsType = 18
'20071221::hanc	C_MovType = 19
'20071221::hanc	C_DocumentNo = 20
'20071221::hanc	C_SeqNo = 21
'20071221::hanc	C_SoNo = 22
'20071221::hanc	C_SoSeq = 23
'20071221::hanc	C_PoNo = 24
'20071221::hanc	C_PoSeq = 25
'20071221::hanc	C_ProdtNo = 26
'20071221::hanc	C_Spec = 27
'20071221::hanc	C_DebitCreditFlag = 28
'20071221::hanc	C_BpCd = 29
'20071221::hanc	C_BpNm = 30
'20071221::hanc	C_TrnsSlCd = 31
'20071221::hanc	C_WcCd = 32
'20071221::hanc	C_WcNm = 33
'20071221::hanc	C_DocumentText = 34
'20071221::hanc	C_SubSeqNo = 35
'20071221::hanc	C_TempGlNo = 36
'20071221::hanc	C_GlNo = 37
	
	C_ItemCd = 1
	C_ItemNm = 2
	C_SlCd = 3
	C_DocumentDt = 4
	C_OrderUnit = 5
	C_BaseUnit = 6
	C_Qty = 7
	C_Price = 8
	C_Amount = 9
	C_CostOfDevy = 10
	C_TrackingNo = 11
	C_LotNo = 12
	C_LotSubNo = 13
	C_TrnsType = 14
	C_MovType = 15
	C_DocumentNo = 16
	C_SeqNo = 17
	C_SoNo = 18
	C_SoSeq = 19
	C_PoNo = 20
	C_PoSeq = 21
	C_ProdtNo = 22
	C_Spec = 23
	C_DebitCreditFlag = 24
	C_BpCd = 25
	C_BpNm = 26
	C_TrnsSlCd = 27
	C_WcCd = 28
	C_WcNm = 29
	C_DocumentText = 30
	C_SubSeqNo = 31
	C_TempGlNo = 32
	C_GlNo = 33				

'20071221::hanc	C_ItemAcct = 3
'20071221::hanc	C_CCCd = 4
'20071221::hanc	C_ProjectNm = 14
'20071221::hanc	C_StructNm = 15
	

End Sub


'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		
 		'C_ItemCd = iCurColumnPos(1)
		'C_ItemNm = iCurColumnPos(2)
		'C_ItemAcct = iCurColumnPos(3) 
		'C_SlCd = iCurColumnPos(4)
		'C_CCCd = iCurColumnPos(5)
		'C_DocumentDt = iCurColumnPos(6)
		'C_OrderUnit = iCurColumnPos(7)
		'C_BaseUnit = iCurColumnPos(8)
		'C_Qty = iCurColumnPos(9)
		'C_Price = iCurColumnPos(10)
		'C_Amount = iCurColumnPos(11)
		'C_CostOfDevy = iCurColumnPos(12)
		'C_TrackingNo = iCurColumnPos(13)
		'C_LotNo = iCurColumnPos(14)
		'C_LotSubNo = iCurColumnPos(15)
		'C_TrnsType = iCurColumnPos(16)
		'C_MovType = iCurColumnPos(17)
		'C_DocumentNo = iCurColumnPos(18)
		'C_SeqNo = iCurColumnPos(19)
		'C_SoNo = iCurColumnPos(20)
		'C_SoSeq = iCurColumnPos(21)
		'C_PoNo = iCurColumnPos(22)
		'C_PoSeq = iCurColumnPos(23)
		'C_ProdtNo = iCurColumnPos(24)
		'C_Spec = iCurColumnPos(25)
		'C_DebitCreditFlag = iCurColumnPos(26)
		'C_BpCd = iCurColumnPos(27)
		'C_BpNm = iCurColumnPos(28)
		'C_TrnsSlCd = iCurColumnPos(29)
		'C_WcCd = iCurColumnPos(30)
		'C_WcNm = iCurColumnPos(31)
		'C_DocumentText = iCurColumnPos(32)
		'C_SubSeqNo = iCurColumnPos(33)
		'C_TempGlNo = iCurColumnPos(34)
		'C_GlNo = iCurColumnPos(35)
		'
		'C_ProjectNm = iCurColumnPos(36)
		'C_StructNm = iCurColumnPos(37)
 		
'20071221::hanc 	C_ItemCd = iCurColumnPos(1)
'20071221::hanc		C_ItemNm = iCurColumnPos(2)
'20071221::hanc		C_ItemAcct = iCurColumnPos(3) 
'20071221::hanc		C_SlCd = iCurColumnPos(4)
'20071221::hanc		C_CCCd = iCurColumnPos(5)
'20071221::hanc		C_DocumentDt = iCurColumnPos(6)
'20071221::hanc		C_OrderUnit = iCurColumnPos(7)
'20071221::hanc		C_BaseUnit = iCurColumnPos(8)
'20071221::hanc		C_Qty = iCurColumnPos(9)
'20071221::hanc		C_Price = iCurColumnPos(10)
'20071221::hanc		C_Amount = iCurColumnPos(11)
'20071221::hanc		C_CostOfDevy = iCurColumnPos(12)
'20071221::hanc		C_TrackingNo = iCurColumnPos(13)
'20071221::hanc		C_ProjectNm = iCurColumnPos(14)
'20071221::hanc		C_StructNm = iCurColumnPos(15)
'20071221::hanc		C_LotNo = iCurColumnPos(16)
'20071221::hanc		C_LotSubNo = iCurColumnPos(17)
'20071221::hanc		C_TrnsType = iCurColumnPos(18)
'20071221::hanc		C_MovType = iCurColumnPos(19)
'20071221::hanc		C_DocumentNo = iCurColumnPos(20)
'20071221::hanc		C_SeqNo = iCurColumnPos(21)
'20071221::hanc		C_SoNo = iCurColumnPos(22)
'20071221::hanc		C_SoSeq = iCurColumnPos(23)
'20071221::hanc		C_PoNo = iCurColumnPos(24)
'20071221::hanc		C_PoSeq = iCurColumnPos(25)
'20071221::hanc		C_ProdtNo = iCurColumnPos(26)
'20071221::hanc		C_Spec = iCurColumnPos(27)
'20071221::hanc		C_DebitCreditFlag = iCurColumnPos(28)
'20071221::hanc		C_BpCd = iCurColumnPos(29)
'20071221::hanc		C_BpNm = iCurColumnPos(30)
'20071221::hanc		C_TrnsSlCd = iCurColumnPos(31)
'20071221::hanc		C_WcCd = iCurColumnPos(32)
'20071221::hanc		C_WcNm = iCurColumnPos(33)
'20071221::hanc		C_DocumentText = iCurColumnPos(34)
'20071221::hanc		C_SubSeqNo = iCurColumnPos(35)
'20071221::hanc		C_TempGlNo = iCurColumnPos(36)
'20071221::hanc		C_GlNo = iCurColumnPos(37)
		
        C_ItemCd = iCurColumnPos(1)
        C_ItemNm = iCurColumnPos(2)
        C_SlCd = iCurColumnPos(3)
        C_DocumentDt = iCurColumnPos(4)
        C_OrderUnit = iCurColumnPos(5)
        C_BaseUnit = iCurColumnPos(6)
        C_Qty = iCurColumnPos(7)
        C_Price = iCurColumnPos(8)
        C_Amount = iCurColumnPos(9)
        C_CostOfDevy = iCurColumnPos(10)
        C_TrackingNo = iCurColumnPos(11)
        C_LotNo = iCurColumnPos(12)
        C_LotSubNo = iCurColumnPos(13)
        C_TrnsType = iCurColumnPos(14)
        C_MovType = iCurColumnPos(15)
        C_DocumentNo = iCurColumnPos(16)
        C_SeqNo = iCurColumnPos(17)
        C_SoNo = iCurColumnPos(18)
        C_SoSeq = iCurColumnPos(19)
        C_PoNo = iCurColumnPos(20)
        C_PoSeq = iCurColumnPos(21)
        C_ProdtNo = iCurColumnPos(22)
        C_Spec = iCurColumnPos(23)
        C_DebitCreditFlag = iCurColumnPos(24)
        C_BpCd = iCurColumnPos(25)
        C_BpNm = iCurColumnPos(26)
        C_TrnsSlCd = iCurColumnPos(27)
        C_WcCd = iCurColumnPos(28)
        C_WcNm = iCurColumnPos(29)
        C_DocumentText = iCurColumnPos(30)
        C_SubSeqNo = iCurColumnPos(31)
        C_TempGlNo = iCurColumnPos(32)
        C_GlNo = iCurColumnPos(33)	
		
		
 	End Select
 
End Sub


'-----------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus   
		Exit Function
	End If
 
	If Plant_SLCd_Check(0) = False Then Exit Function

	If lgIsOpenPop = True Then Exit Function
 
	lgIsOpenPop = True

	arrParam(0) = "품목"     
	arrParam(1) = "B_Item_By_Plant,B_Item"
	arrParam(2) = Trim(frm1.txtItemCd.Value)
	arrParam(3) = ""
	arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	arrParam(5) = "품목"      

	arrField(0) = "B_Item_By_Plant.Item_Cd"  
	arrField(1) = "B_Item.Item_NM"   
	    
	arrHeader(0) = "품목"      
	arrHeader(1) = "품목명"    
    
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If 
	Set gActiveElement = document.activeElement
End Function

 '-----------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	If frm1.txtPlantCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장팝업"        
	arrParam(1) = "B_PLANT"          
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""           
	arrParam(4) = ""           
	arrParam(5) = "공장"
	 
	arrField(0) = "B_PLANT.PLANT_CD"
	arrField(1) = "B_PLANT.PLANT_NM"
	    
	arrHeader(0) = "공장"       
	arrHeader(1) = "공장명"     
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	lgIsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If 
	Set gActiveElement = document.activeElement  
End Function

'------------------------------------------  OpenSL1()  -------------------------------------------------
Function OpenSL1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus    
		Exit Function
	End If

	If Plant_SLCd_Check(0) = False Then Exit Function

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "창고팝업"                                                                    
	arrParam(1) = "B_STORAGE_LOCATION"                           
	arrParam(2) = Trim(frm1.txtFrSlCd.Value)                     
	arrParam(3) = ""
	arrParam(4) = "Plant_cd= " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "창고"
	 
	arrField(0) = "Sl_Cd" 
	arrField(1) = "Sl_Nm" 
	 
	arrHeader(0) = "창고"  
	arrHeader(1) = "창고명"  

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	  
	lgIsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtFrSlCd.Value = arrRet(0)
		frm1.txtFrSlNm.Value = arrRet(1)
		frm1.txtFrSlCd.focus
	End If 
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  OpenSL2()  -------------------------------------------------
Function OpenSL2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus   
		Exit Function
	End If

	If Plant_SLCd_Check(0) = False Then Exit Function

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "창고팝업" 
	arrParam(1) = "B_STORAGE_LOCATION"
	arrParam(2) = Trim(frm1.txtToSlCd.Value)
	arrParam(3) = ""
	arrParam(4) = "Plant_cd= " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "창고"
	 
	arrField(0) = "Sl_Cd" 
	arrField(1) = "Sl_Nm" 
	 
	arrHeader(0) = "창고"  
	arrHeader(1) = "창고명"  

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	  
	lgIsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtToSlCd.focus
		Exit Function
	Else
		frm1.txtToSlCd.Value = arrRet(0)
		frm1.txtToSlNm.Value = arrRet(1)
		frm1.txtToSlCd.focus
	End If 
	Set gActiveElement = document.activeElement
End Function

 '------------------------------------------  OpenItemAcct()  -------------------------------------------------
Function OpenItemAcct()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If lgIsOpenPop = True Then Exit Function
	 
	lgIsOpenPop = True

	arrParam(0) = "품목계정 팝업"                 
	arrParam(1) = "B_MINOR"     
	arrParam(2) = Trim(frm1.txtItemAcct.Value)   
	arrParam(3) = ""       
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & ""    
	arrParam(5) = "품목계정"     
	 
	arrField(0) = "MINOR_CD"      
	arrField(1) = "MINOR_NM"      
	 
	arrHeader(0) = "품목계정"      
	arrHeader(1) = "계정명"      
	 
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	lgIsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtItemAcct.focus
		Exit Function
	Else
		frm1.txtItemAcct.Value = arrRet(0)
		frm1.txtItemAcctNm.Value = arrRet(1)
		frm1.txtItemAcct.focus
	End If
	Set gActiveElement = document.activeElement 
End Function

 '------------------------------------------  OpenMovType()  -------------------------------------------------
Function OpenMovType()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If lgIsOpenPop = True Or UCase(frm1.txtMovType.ClassName)="PROTECTED" Then Exit Function 

	lgIsOpenPop = True

	arrParam(0) = "수불유형 팝업"     
	arrParam(1) = "B_MINOR"      
	arrParam(2) = Trim(frm1.txtMovType.Value)
	arrParam(3) = ""
	arrParam(4) = "MAJOR_CD = " & FilterVar("I0001", "''", "S") & ""
	arrParam(5) = "수불유형"
	 
	arrField(0) = "MINOR_CD" 
	arrField(1) = "MINOR_NM" 
	 
	arrHeader(0) = "수불유형"  
	arrHeader(1) = "수불유형명"  
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	lgIsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtMovType.focus
		Exit Function
	Else
		frm1.txtMovType.Value   = arrRet(0)
		frm1.txtMovTypeNm.Value = arrRet(1)
		frm1.txtMovType.focus
	End If
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  OpenWC()  -------------------------------------------------
Function OpenWCCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus    
		Exit Function
	End If

	If Plant_SLCd_Check(0) = False Then Exit Function

	If lgIsOpenPop = True Then Exit Function
	 
	If UCase(frm1.txtWCCd.ClassName) = UCase(Parent.UCN_PROTECTED) Then	Exit Function

	lgIsOpenPop = True

	arrParam(0) = "작업장팝업" 
	arrParam(1) = "P_WORK_CENTER"    
	arrParam(2) = Trim(frm1.txtWCCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "작업장"   
	 
	arrField(0) = "WC_CD" 
	arrField(1) = "WC_NM" 
	 
	arrHeader(0) = "작업장"  
	arrHeader(1) = "작업장명"  
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	lgIsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtWCCd.focus
		Exit Function
	Else
		frm1.txtWcCd.Value   = arrRet(0)
		frm1.txtWcNm.Value   = arrRet(1)
		frm1.txtWcCd.Focus
	End If 
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  OpenTrackingNo()  --------------------------------------------------
' Name : OpenTrackingNo()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingNo()
	Dim iCalledAspName, IntRetCD

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If
		
End Function

'------------------------------------------  SetTrackingNo()  --------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetTrackingNo(Byval arrRet)

    With frm1
			.txtTrackingNo.Value = arrRet(0)
			.txtTrackingNo.focus
			Set gActiveElement = document.activeElement
	End With
	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'=======================================================================================================
'   Function Name : ShowHiddenData
'=======================================================================================================
Sub ShowHiddenData(ByVal Row)
	With frm1
		.vspdData.Row    = Row
		.vspdData.Col    = C_DocumentText
		.Remark.Value    = .vspdData.Text
		
		.vspdData.Row    = Row
		.vspdData.Col    = C_WcNm
		.WcNm.Value		 = .vspdData.Text
			  
		.vspdData.Col    = C_WcCd
		.WcCd.Value      = .vspdData.Text
			  
		.vspdData.Col    = C_TrnsSlCd
		.TrnsSlCd.Value  = .vspdData.Text
			  
		.vspdData.Col    = C_BpNm
		.txtBpNm.Value   = .vspdData.Text
			  
		.vspdData.Col    = C_BpCd
		.BpCd.Value      = .vspdData.Text
	End With
End Sub


'==========================================================================================
'   Event Name : txtTrnsFrDt
'==========================================================================================
Sub txtTrnsFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtTrnsFrDt.Action = 7
		Call SetFocusToDocument("M")        
		frm1.txtTrnsFrDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtTrnsToDt
'==========================================================================================
Sub txtTrnsToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtTrnsToDt.Action = 7
		Call SetFocusToDocument("M")        
		frm1.txtTrnsToDt.Focus
	End if
End Sub

Function  txtTrnsFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then Call MainQuery()
End Function

Function txtTrnsToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then Call MainQuery()
End Function


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet
	Call ggoSpread.ReOrderingSpreadData
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then Exit Sub
    If CheckRunningBizProcess = True Then Exit Sub
    
    If	frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and _
		lgStrPrevKey <> "" and lgStrPrevKey2 <> "" and lgStrPrevKey3 <> "" and lgStrPrevKey4 <> "" Then
		Call DisableToolBar(Parent.TBC_QUERY)
		If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData

	Call SetPopupMenuItemInf("0000111111")
         
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Sub
		
		If Row < 1 Then
			ggoSpread.Source = frm1.vspdData
			 
 			If lgSortKey = 1 Then
 				ggoSpread.SSSort Col					'Sort in Ascending
 				lgSortKey = 2
 			Else
 				ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 				lgSortKey = 1
			End If 
					
		End If
    End With
    
    Call ShowHiddenData(frm1.vspdData.ActiveRow)
	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If NewRow > 0 And Row <> NewRow Then
		Call ShowHiddenData(NewRow)
	End If
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

 '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery() 
    FncQuery = False                                                        
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
   
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkFieldByCell(frm1.txtPlantCd, "A",1) Then Exit Function
	If Not chkFieldByCell(frm1.txtTrnsFrDt, "A",1) Then Exit Function
	If Not chkFieldByCell(frm1.txtTrnsToDt, "A",1) Then Exit Function
	
	If ValidDateCheck(frm1.txtTrnsFrDt, frm1.txtTrnsToDt) = False Then Exit Function
    
    '-----------------------
    'Erase contents area
    '-----------------------
	Call ggoSpread.ClearSpreadData      

	Call InitVariables               
    
    If Plant_SLCd_Check(1) = False Then Exit Function
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery() = False Then Exit Function

    FncQuery = True 
    
    Call SetToolbar("11000000000111")    

End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
   FncExit = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
 Dim strVal
    DbQuery = False
    
    Err.Clear
	Call LayerShowHide(1)
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID &	"?txtMode="				& Parent.UID_M0001 & _
										"&lgIntFlgMode="		& lgIntFlgMode & _
										"&txtPlantCd="			& Trim(.hPlantCd.value) & _
										"&txtItemCd="			& Trim(.hItemCd.value) & _
										"&txtTrnsFrDt="			& Trim(.hTrnsFrDt.value) & _
										"&txtTrnsToDt="			& Trim(.hTrnsToDt.value) & _
										"&txtFrSlCd="			& Trim(.hFrSlCd.value) & _
										"&txtToSlCd="			& Trim(.hToSlCd.value) & _
										"&cboItemAcct="			& Trim(.hItemAcct.value) & _
										"&txtWcCd="				& Trim(.hWcCd.value) & _
										"&txtMovType="			& Trim(.hMovType.value) & _
										"&cboTrnsType="			& Trim(.hTrnsType.value) & _
										"&txtRefNo="			& Trim(.hRefNo.value) & _
										"&txtTrackingNo="		& Trim(.hTrackingNo.value) & _
										"&lgStrPrevKey="		& lgStrPrevKey & _
										"&lgStrPrevKey2="		& lgStrPrevKey2 & _
										"&lgStrPrevKey3="		& lgStrPrevKey3 & _
										"&lgStrPrevKey4="		& lgStrPrevKey4 & _
										"&txtMaxRows="			& .vspdData.MaxRows
		Else
    
			strVal = BIZ_PGM_QRY_ID &	"?txtMode="				& Parent.UID_M0001 & _
										"&lgIntFlgMode="		& lgIntFlgMode & _
										"&txtPlantCd="			& Trim(.txtPlantCd.Value) & _
										"&txtItemCd="			& Trim(.txtItemCd.value) & _
										"&txtTrnsFrDt="			& Trim(.txtTrnsFrDt.Text) & _
										"&txtTrnsToDt="			& Trim(.txtTrnsToDt.Text) & _
										"&txtFrSlCd="			& Trim(.txtFrSlCd.value) & _
										"&txtToSlCd="			& Trim(.txtToSlCd.value) & _
										"&cboItemAcct="			& Trim(.cboItemAcct.value) & _
										"&txtWcCd="				& Trim(.txtWcCd.value) & _
										"&txtMovType="			& Trim(.txtMovType.value) & _
										"&cboTrnsType="			& Trim(.cboTrnsType.value) & _
										"&txtRefNo="			& Trim(.txtRefNo.value) & _
										"&txtTrackingNo="		& Trim(.txtTrackingNo.value) & _
										"&lgStrPrevKey="		& lgStrPrevKey & _
										"&lgStrPrevKey2="		& lgStrPrevKey2 & _
										"&lgStrPrevKey3="		& lgStrPrevKey3 & _
										"&lgStrPrevKey4="		& lgStrPrevKey4 & _
										"&txtMaxRows="			& .vspdData.MaxRows
		End if
    End With

    Call RunMyBizASP(MyBizASP, strVal)         
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
	lgIntFlgMode = parent.OPMD_UMODE
    
     Call ShowHiddenData(frm1.vspdData.ActiveRow)
     frm1.vspdData.Focus
End Function


'========================================================================================
' Function Name : Plant_SLCd_Check
'========================================================================================
Function Plant_SLCd_Check(ByVal ChkIndex)
	
	Plant_SLCd_Check = False
 
 '-----------------------
 'Check Plant CODE  
 '-----------------------
    If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit function
    End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)

If ChkIndex = 1 Then
 '-----------------------
 'Check SLCd CODE  
 '-----------------------
	If frm1.txtFrSlCd.value <> "" Then	
		
		If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND SL_CD = " & FilterVar(frm1.txtFrSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtFrSLCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
				Call DisplayMsgBox("125700","X","X","X")
				frm1.txtFrSlNm.Value = ""
				frm1.txtFrSlCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			Else
				lgF0 = Split(lgF0, Chr(11))
				frm1.txtFrSlNm.Value = lgF0(0)
				Call DisplayMsgBox("169922","X","X","X")
				frm1.txtFrSlCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtFrSlNm.Value = lgF0(0)
	Else
		frm1.txtFrSlNm.Value = ""
	End if
	
	If frm1.txtToSlCd.value <> "" Then
		If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND SL_CD = " & FilterVar(frm1.txtToSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtToSLCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
				Call DisplayMsgBox("125700","X","X","X")
				frm1.txtToSlNm.Value = ""
				frm1.txtToSlCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			Else
				lgF0 = Split(lgF0, Chr(11))
				frm1.txtToSlNm.Value = lgF0(0)
				Call DisplayMsgBox("169922","X","X","X")
				frm1.txtToSlCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtToSlNm.Value = lgF0(0)
	Else
		frm1.txtToSlNm.Value = ""
	End if 
	
	If frm1.txtItemCd.value <> "" Then
		If  CommonQueryRs(" B.ITEM_NM "," B_ITEM_BY_PLANT A, B_ITEM B ", " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
				Call DisplayMsgBox("122600","X","X","X")
				frm1.txtItemNm.Value = ""
				frm1.txtItemCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			Else
				lgF0 = Split(lgF0, Chr(11))
				frm1.txtItemNm.Value = lgF0(0)
				Call DisplayMsgBox("122700","X","X","X")
				frm1.txtItemCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtItemNm.Value = lgF0(0)
	Else
		frm1.txtItemNm.Value = ""
	End if 
	 
	If frm1.txtWcCd.value <> "" Then
		If  CommonQueryRs(" WC_NM "," P_WORK_CENTER ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND WC_CD = " & FilterVar(frm1.txtWcCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("182100","X","X","X")
			frm1.txtWcNm.Value = ""
			frm1.txtWcCd.Focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtWcNm.Value = lgF0(0)
	Else
		frm1.txtWcNm.Value = ""
	End if 

	If frm1.txtMovType.value <> "" Then
		If  CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND MINOR_CD= " & FilterVar(frm1.txtMovType.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("169948","X","X","X")
			frm1.txtMovTypeNm.Value = ""
			frm1.txtMovType.Focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtMovTypeNm.Value = lgF0(0)
	Else
		frm1.txtMovTypeNm.Value = ""
	End if 
	
 End If 
  
 Plant_SLCd_Check = True
End Function

'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim intCnt
	Dim var1, var2
    Dim StrEbrFile
    Dim intRetCd
	Dim ObjName
	
	StrEbrFile = "I1311OA1_KO412"

	frm1.vspddata.Row  = frm1.vspddata.ActiveRow
	
	frm1.vspddata.col  = C_DocumentNo 'GetKeyPos("A",9)
	var1 = Trim(frm1.vspddata.text)
    Err.Clear
    
	frm1.vspddata.col  = C_DocumentDt 'GetKeyPos("A",8)
	var2 = Trim(Left(UniConvDate(frm1.vspddata.Text),4))

	StrUrl = StrUrl & "ITEM_DOCUMENT_NO|" & Var1
	StrUrl = StrUrl & "|DOCUMENT_YEAR|" & Var2

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	if	left(Var1,2)	=	"ST"	then
		Call FncEBRPrint(EBAction,ObjName,StrUrl)
	Else
		Call DisplayMsgBox("I10000","X", "X", "X")   
		Exit Function
	End if
End Function

'========================================================================================
Function FncBtnPreview()
	Dim strUrl
	Dim intCnt
	Dim var1, var2
        Dim StrEbrFile
        Dim intRetCd
	Dim ObjName
	
	StrEbrFile = "I1311OA1_KO412"

	frm1.vspddata.Row  = frm1.vspddata.ActiveRow
	
	frm1.vspddata.col  = C_DocumentNo 'GetKeyPos("A",9)
	var1 = Trim(frm1.vspddata.text)
        
        Err.Clear
    
	frm1.vspddata.col  = C_DocumentDt 'GetKeyPos("A",8)
	var2 = Trim(Left(UniConvDate(frm1.vspddata.Text),4))

	StrUrl = StrUrl & "ITEM_DOCUMENT_NO|"	& Var1
	StrUrl = StrUrl & "|DOCUMENT_YEAR|"		& Var2

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	
	If	left(Var1,2)	=	"ST"	then
		Call FncEBRPreview(ObjName,StrUrl)
	Else
		Call DisplayMsgBox("I10000","X", "X", "X")   
		Exit Function
	End if
End Function
'========================================================================================