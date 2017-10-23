Option Explicit

Const DNCheck = "DNCheck"

Const BIZ_PGM_ID = "s4311mb1.asp"            '☆: Head Query 비지니스 로직 ASP명 
Const BIZ_OnLine_ID = "s3111ab1.asp"           '☆: OnLine ADO 비지니스 로직 ASP명 

Const TAB1 = 1                 '☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

Dim C_SlCd
Dim C_SlCdPopup
Dim C_ItemCd
Dim C_ItemPopup
Dim C_ItemNm
Dim C_Spec
Dim C_DnUnit
Dim C_DnUnitPopup
Dim C_DnQty
Dim C_Price
Dim C_TotalAmt
Dim C_NetAmt
Dim C_LotNo
Dim C_LotNoPopup
Dim C_LotSeq
Dim C_CartonNo
Dim C_VatIncFlag
Dim C_VatIncFlagNm
Dim C_VatType
Dim C_VatTypePopup
Dim C_VatTypeNm
Dim C_VatRate
Dim C_VatAmt
Dim C_RetType
Dim C_RetTypePopup
Dim C_RetTypeNm
Dim C_ReservedPrice
Dim C_Remark
Dim C_LotMgmtFlag
Dim C_DnSeq
Dim C_OldNetAmt             ' 변경전 금액 
Dim C_OldVatAmt
Dim C_DBNetAmt              ' DB의 금액 
Dim C_DbVatAmt

'=========================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgBlnFlgChgValue1           '2002.12.18 납품처상세정보 변경 여부 
Dim lgBlnFlgChgValue2           '2002.12.18 운송정보 변경 여부 

Dim lgBlnFlgDtlChgValue        ' Variable is for Dtl Dirty flag

Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgStrPrevKey
Dim lgBlnFlgHdrMode
Dim lgBlnFlgDtlMode
Dim lgSortKey

'=========================================
Dim IsOpenPop      ' Popup
Dim lsClickCfmYes
Dim lsClickCfmNo
Dim PrevRadioFlag     '☜: Radio Button의 이전값 
Dim PrevRadioType     '☜: Radio Button의 이전값 
Dim PrevRadioDnParcel
Dim arrCollectVatType
Dim lgArrVATTypeInfo    ' VAT Type 정보(VAT_TYPE, VAT_TYPE_NAME, VAT_RATE)
Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6, i '@@@CommonQueryRs 를 위한 변수 

'=========================================
Sub FormatField()
    With frm1
        ' 날짜 OCX Foramt 설정 
        Call FormatDATEField(.txtDlvyDt)
        Call FormatDATEField(.txtPlannedGIDt)
        Call FormatDATEField(.txtArriv_dt)
        Call FormatDATEField(.txtGI_Dt)
        ' 수량 OCX Format 설정 
        Call FormatDoubleSingleField(.txtVat_rate)
        Call FormatDoubleSingleField(.txtNet_amt)
        Call FormatDoubleSingleField(.txtVat_amt)
        Call FormatDoubleSingleField(.txtTot_amt)
        Call FormatDoubleSingleField(.txtTotal_Amt)
        Call FormatDoubleSingleField(.txtCol_amt)
    End With
End Sub

'=========================================
Sub LockFieldInit(ByVal pvFlag)
    With frm1
        ' 날짜 OCX
        Call LockObjectField(.txtDlvyDt, "R")
        Call LockObjectField(.txtPlannedGIDt, "R")
        Call LockObjectField(.txtArriv_dt, "O")
        Call LockObjectField(.txtGI_Dt, "P")
        ' 수량 OCX
        Call LockObjectField(.txtVat_rate, "P")
        Call LockObjectField(.txtNet_amt, "P")
        Call LockObjectField(.txtVat_amt, "P")
        Call LockObjectField(.txtTot_amt, "P")
        Call LockObjectField(.txtTotal_Amt, "P")
        Call LockObjectField(.txtCol_amt, "P")
        
        If pvFlag = "N" Then
			Call LockHTMLField(.txtDnNo, "O")	
        End If
    End With

End Sub

'=========================================
Sub LockFieldQuery()
    With frm1
		Call LockHTMLField(.txtDnNo, "P")	
		Call LockHTMLField(.txtDn_Type, "P")	
    End With
End Sub

'=========================================
Sub initSpreadPosVariables()
        
    C_SlCd = 1             '창고 
    C_SlCdPopup = 2        '창고Popup
    C_ItemCd = 3           '품목 
    C_ItemPopup = 4
    C_ItemNm = 5           '품목명 
    C_Spec = 6             '품목규격 
    C_DnUnit = 7           '단위 
    C_DnUnitPopup = 8
    C_DnQty = 9            '출고요청수량 
    C_Price = 10            '단가 
    C_TotalAmt = 11
    C_NetAmt = 12          '금액 
    C_LotNo = 13           'LOT No
    C_LotNoPopup = 14      'LOT NoPopup
    C_LotSeq = 15          'LOT No 순번 
    C_CartonNo = 16
    C_VatIncFlag = 17
    C_VatIncFlagNm = 18
    C_VatType = 19
    C_VatTypePopup = 20
    C_VatTypeNm = 21
    C_VatRate = 22
    C_VatAmt = 23         '부가세금액 
    C_RetType = 24
    C_RetTypePopup = 25
    C_RetTypeNm = 26
    C_ReservedPrice = 27
    C_Remark = 28
    C_LotMgmtFlag = 29
    C_DnSeq = 30           '출하순번 
    C_OldNetAmt = 31
    C_OldVatAmt = 32
    C_DBNetAmt = 33
    C_DbVatAmt = 34
End Sub

'=========================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgBlnFlgDtlChgValue = False
    '2002.12.20 SMJ
    lgBlnFlgChgValue1 = False
    lgBlnFlgChgValue2 = False
    
    lgIntGrpCount = 0                           'initializes Group View Size
    IsOpenPop = False
    lsClickCfmYes = False
    lsClickCfmNo = False

End Sub

'=========================================
Sub SetDefaultVal()

    On Error Resume Next

    With frm1
        .txtSales_Grp.Value = Parent.gSalesGrp
        .txtCurrency.Value = Parent.gCurrency
        .txtPlant.Value = Parent.gPlant
        .txtPlantNm.Value = Parent.gPlantNm
        '2002-09-26 SMJ 추가 (납기일, 출고예정일)
        .txtPlannedGIDt.Text = EndDate
        .txtDlvyDt.Text = EndDate
        .rdoVat_Calc_Type1.Checked = True
              
    End With

    If Err.Number <> 0 Then
        MsgBox Err.Description
        Err.Clear
    End If

    lgBlnFlgChgValue = False
    '2002.12.20 SMJ
    lgBlnFlgChgValue1 = False
    lgBlnFlgChgValue2 = False
     
    frm1.btnPosting.disabled = True
    frm1.btnPostCancel.disabled = True
    frm1.btnPosting.Value = "출고처리"
    frm1.btnPostCancel.Value = "출고처리취소"
    frm1.chkArFlag.Checked = False
    frm1.chkVatFlag.Checked = False

End Sub

'=========================================
Sub SetRowDefaultVal(ByVal pvRowCnt, ByVal pvInsFlag)

    With frm1.vspdData
    
        .Row = pvRowCnt

        .Col = C_DnSeq:         .Text = 0
        .Col = C_LotNo:         .Text = ""
        .Col = C_LotSeq:        .Text = 0
        
        .Col = C_DnQty:         .Text = 0
        .Col = C_Price:         .Text = 0
        .Col = C_TotalAmt:      .Text = 0
        .Col = C_NetAmt:        .Text = 0
        .Col = C_VatAmt:        .Text = 0
        
        .Col = C_OldNetAmt:     .Text = 0
        .Col = C_OldVatAmt:     .Text = 0
        .Col = C_DBNetAmt:      .Text = 0
        .Col = C_DbVatAmt:      .Text = 0

        ' 입력인 경우 
        If pvInsFlag = "Y" Then
            ' 창고 
            .Col = C_SlCd:          .Text = Trim(frm1.txtSlCd.Value)
             
            '부가세 포함여부 Default값 설정 
            If frm1.rdoVat_Inc_flag1.Checked Then
                .Col = C_VatIncFlag:        .Text = "1"
                .Col = C_VatIncFlagNm:      .Text = "별도"
            Else
                .Col = C_VatIncFlag:        .Text = "2"
                .Col = C_VatIncFlagNm:      .Text = "포함"
            End If
              
            ' Header에 부가세 유형이 등록되어 있는 경우 해당 값을 Default값을 설정한다.
            If Trim(frm1.txtVat_Type.Value) <> "" Then
                .Col = C_VatType:       .Text = frm1.txtVat_Type.Value
                .Col = C_VatTypeNm:     .Text = frm1.txtVatTypeNm.Value
                .Col = C_VatRate:       .Text = frm1.txtVat_rate.Text
            End If
        End If
    End With

End Sub

'=========================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables
    
    With ggoSpread
        .Source = frm1.vspdData
        
        .Spreadinit "V20030901", , Parent.gAllowDragDropSpread
 
        frm1.vspdData.MaxCols = C_DbVatAmt + 1            '☜: 최대 Columns의 항상 1개 증가시킴 
        
        frm1.vspdData.MaxRows = 0

        frm1.vspdData.ReDraw = False

        Call GetSpreadColumnPos("A")

        .SSSetEdit C_SlCd, "창고", 8, , , 7, 2
        .SSSetButton C_SlCdPopup
        .SSSetEdit C_ItemCd, "품목", 18, , , 18, 2
        .SSSetButton C_ItemPopup
        .SSSetEdit C_ItemNm, "품목명", 30
        .SSSetEdit C_Spec, "규격", 30
        .SSSetEdit C_DnUnit, "단위", 8, , , 5, 2
        .SSSetButton C_DnUnitPopup
        .SSSetFloat C_DnQty, "수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetFloat C_Price, "단가", 15, Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetFloat C_TotalAmt, "금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetFloat C_NetAmt, "순금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
        .SSSetEdit C_LotNo, "Lot No.", 12, , , 25, 2
        .SSSetButton C_LotNoPopup
        .SSSetEdit C_CartonNo, "Carton No", 15, , , 10, 2
        Call AppendNumberPlace("7", "5", "0")
        .SSSetFloat C_LotSeq, "LOT NO 순번", 15, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetCombo C_VatIncFlagNm, "VAT포함구분명", 15, 2
        .SetCombo "별도" & vbTab & "포함", C_VatIncFlagNm
        .SSSetEdit C_VatIncFlag, "VAT포함구분", 5, 2
        .SetCombo "1" & vbTab & "2", C_VatIncFlag
        .SSSetEdit C_VatType, "VAT유형", 10, , , 5, 2
        .SSSetButton C_VatTypePopup
        .SSSetEdit C_VatTypeNm, "VAT유형명", 20
        .SSSetFloat C_VatRate, "VAT율", 10, Parent.ggExchRateNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetFloat C_VatAmt, "VAT금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetEdit C_RetType, "반품유형", 10, , , 5, 2
        .SSSetButton C_RetTypePopup
        .SSSetEdit C_RetTypeNm, "반품유형명", 20
        .SSSetFloat C_ReservedPrice, "적립금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetEdit C_Remark, "비고", 50, , , 120
        .SSSetEdit C_LotMgmtFlag, "LOT관리여부", 2
        .SSSetEdit C_DnSeq, "출하순번", 3

        .SSSetFloat C_OldNetAmt, "순금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetFloat C_OldVatAmt, "VAT금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetFloat C_DBNetAmt, "순금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        .SSSetFloat C_DbVatAmt, "VAT금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"

        Call .MakePairsColumn(C_SlCd, C_SlCdPopup)
        Call .MakePairsColumn(C_ItemCd, C_ItemPopup)
        Call .MakePairsColumn(C_DnUnit, C_DnUnitPopup)
        Call .MakePairsColumn(C_LotNo, C_LotNoPopup)
        Call .MakePairsColumn(C_VatType, C_VatTypePopup)
        Call .MakePairsColumn(C_RetType, C_RetTypePopup)

        Call .SSSetColHidden(C_VatIncFlag, C_VatIncFlag, True)
        Call .SSSetColHidden(C_LotMgmtFlag, C_LotMgmtFlag, True)
        Call .SSSetColHidden(C_NetAmt, C_NetAmt, True)
        Call .SSSetColHidden(C_DnSeq, frm1.vspdData.MaxCols, True)
        
        frm1.vspdData.ReDraw = True
  
    End With
End Sub

'=========================================
Sub SetSpreadLock()
End Sub

'=========================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
 
    With ggoSpread
        .SpreadUnLock C_ItemCd, pvStartRow, C_ItemCd, pvEndRow
        .SpreadUnLock C_DnUnit, pvStartRow, C_DnUnit, pvEndRow
        .SpreadUnLock C_DnQty, pvStartRow, C_DnQty, pvEndRow
        .SpreadUnLock C_Price, pvStartRow, C_Price, pvEndRow
        .SpreadUnLock C_TotalAmt, pvStartRow, C_TotalAmt, pvEndRow
        .SpreadUnLock C_SlCd, pvStartRow, C_SlCd, pvEndRow
        .SpreadUnLock C_VatIncFlagNm, pvStartRow, C_VatIncFlagNm, pvEndRow
        .SpreadUnLock C_VatType, pvStartRow, C_VatType, pvEndRow
        .SpreadUnLock C_RetType, pvStartRow, C_RetType, pvEndRow
        .SpreadUnLock C_Remark, pvStartRow, C_Remark, pvEndRow
        .SpreadUnLock C_SlCdPopup, pvStartRow, C_SlCdPopup, pvEndRow
        .SpreadUnLock C_ItemPopup, pvStartRow, C_ItemPopup, pvEndRow
        .SpreadUnLock C_DnUnitPopup, pvStartRow, C_DnUnitPopup, pvEndRow
        .SpreadUnLock C_LotNoPopup, pvStartRow, C_LotNoPopup, pvEndRow
        .SpreadUnLock C_VatTypePopup, pvStartRow, C_VatTypePopup, pvEndRow
        .SpreadUnLock C_RetTypePopup, pvStartRow, C_RetTypePopup, pvEndRow

        .SSSetRequired C_ItemCd, pvStartRow, pvEndRow
        .SSSetRequired C_DnUnit, pvStartRow, pvEndRow
        .SSSetRequired C_DnQty, pvStartRow, pvEndRow
        .SSSetRequired C_Price, pvStartRow, pvEndRow
        .SSSetRequired C_TotalAmt, pvStartRow, pvEndRow
        .SSSetRequired C_SlCd, pvStartRow, pvEndRow
        .SSSetRequired C_VatIncFlagNm, pvStartRow, pvEndRow
        .SSSetRequired C_VatType, pvStartRow, pvEndRow
        .SSSetRequired C_RetType, pvStartRow, pvEndRow

        .SSSetProtected C_VatTypeNm, pvStartRow, pvEndRow
        .SSSetProtected C_RetTypeNm, pvStartRow, pvEndRow
        .SSSetProtected C_ItemNm, pvStartRow, pvEndRow
        .SSSetProtected C_Spec, pvStartRow, pvEndRow
        .SSSetProtected C_VatRate, pvStartRow, pvEndRow
        .SSSetProtected C_VatAmt, pvStartRow, pvEndRow
        .SSSetProtected C_ReservedPrice, pvStartRow, pvEndRow
        .SSSetProtected C_LotMgmtFlag, pvStartRow, pvEndRow
        .SSSetProtected C_DnSeq, pvStartRow, pvEndRow
    End With

    Call ChangeLotRetField(pvStartRow, pvEndRow)
    Call ChangeReturnItemRetField
End Sub

'=========================================
Sub SetSpreadColorConfirmed(ByVal lRow)
   ggoSpread.SSSetProtected lRow, lRow
End Sub

'========================================
Function OpenConDnNo()
    Dim iCalledAspName
    Dim strRet

    If IsOpenPop = True Then Exit Function
   
    IsOpenPop = True
  
    iCalledAspName = AskPRAspName("S4111PA2")
            
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "S4111PA2", "x")
        IsOpenPop = False
        Exit Function
    End If

    frm1.txtConDn_no.focus
    
    strRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=Y", Array(window.Parent), _
        "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
  
    If strRet <> "" Then
         frm1.txtConDn_no.Value = strRet
    End If

End Function

'========================================
Function OpenRequried(ByVal iRequried)
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 OpenRequried = False
 
 If IsOpenPop = True Then Exit Function

 If lsClickCfmYes = True Then Exit Function

 IsOpenPop = True

 Select Case iRequried
 Case 0

 Case 1
    If lsClickCfmNo = True Then
        IsOpenPop = False
        Exit Function
    End If

    If UCase(frm1.txtSales_Grp.className) = Parent.UCN_PROTECTED Then
        IsOpenPop = False
        Exit Function
    End If

    arrParam(0) = "영업그룹"
    arrParam(1) = "B_SALES_GRP"
    arrParam(2) = Trim(frm1.txtSales_Grp.Value)
    arrParam(3) = ""
    arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
    arrParam(5) = "영업그룹"
          
    arrField(0) = "SALES_GRP"
    arrField(1) = "SALES_GRP_NM"
             
    arrHeader(0) = "영업그룹"
    arrHeader(1) = "영업그룹명"
    
    frm1.txtSales_Grp.focus

 Case 2
    If frm1.txtCol_Type.ReadOnly = True Then
        IsOpenPop = False
        Exit Function
    End If
    
    arrParam(0) = "수금유형"
    arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"
    arrParam(2) = Trim(frm1.txtCol_Type.Value)
    arrParam(3) = ""
    arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("A1006", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 4 And Config.Reference in (" & FilterVar("CS", "''", "S") & ") "
    arrParam(5) = "수금유형"
             
    arrField(0) = "Minor.MINOR_CD"
    arrField(1) = "Minor.MINOR_NM"
                
    arrHeader(0) = "수금유형"
    arrHeader(1) = "수금유형명"

    frm1.txtCol_Type.focus
    
 Case 4
    If frm1.txtTaxBizAreaCd.ReadOnly = True Then
        IsOpenPop = False
        Exit Function
    End If
    
    arrParam(0) = "세금신고사업장"
    arrParam(1) = "B_TAX_BIZ_AREA"
    arrParam(2) = Trim(frm1.txtTaxBizAreaCd.Value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "세금신고사업장"

    arrField(0) = "TAX_BIZ_AREA_CD"
    arrField(1) = "TAX_BIZ_AREA_NM"

    arrHeader(0) = "세금신고사업장"
    arrHeader(1) = "세금신고사업장명"

    frm1.txtTaxBizAreaCd.focus
    
 End Select

 arrParam(3) = ""   '☜: [Condition Name Delete]
    
 If iRequried = 0 Then
  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
   "dialogWidth=570px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 Else
  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
   "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 End If
 
 IsOpenPop = False

 If arrRet(0) <> "" Then
  Call SetRequried(arrRet, iRequried)
  OpenRequried = True
 End If
End Function

'========================================
Function OpenBp(ByVal iRequried)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If lsClickCfmYes = True Then Exit Function

    Select Case iRequried
        Case 0
            If frm1.txtSold_to_party.ReadOnly = True Then Exit Function

            IsOpenPop = True
            arrParam(0) = "주문처"
            arrParam(1) = "B_BIZ_PARTNER"
            arrParam(2) = Trim(frm1.txtSold_to_party.Value)
            arrParam(3) = ""   '☜: [Condition Name Delete]
            arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag = " & FilterVar("Y", "''", "S") & " "
            arrParam(5) = "주문처"
             
            arrField(0) = "BP_CD"
            arrField(1) = "BP_NM"
            arrField(2) = "BP_RGST_NO"
                
            If lsClickCfmNo = True Then
                IsOpenPop = False
                Exit Function
            End If

            If UCase(frm1.txtSold_to_party.className) = Parent.UCN_PROTECTED Then
                IsOpenPop = False
                Exit Function
            End If
             
            arrHeader(0) = "주문처"
            arrHeader(1) = "주문처명"
            arrHeader(2) = "사업자등록번호"
            
            frm1.txtSold_to_party.focus
              
        Case 1
            If frm1.txtShip_to_party.ReadOnly = True Then
                IsOpenPop = False
                Exit Function
            End If

            If Trim(frm1.txtSold_to_party.Value) = "" Then
                Call DisplayMsgBox("203150", "X", "X", "X")
                'MsgBox "주문처를 먼저 입력하세요!"
                frm1.txtSold_to_party.focus
                IsOpenPop = False
                Exit Function
            End If

            arrParam(0) = "납품처"
            arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"
            arrParam(2) = Trim(frm1.txtShip_to_party.Value)
            arrParam(3) = Trim(frm1.txtShip_to_partyNm.Value)
            arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SSH", "''", "S") & " " _
                & "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE <= " & FilterVar("CS", "''", "S") & " " _
                & "AND PARTNER_FTN.BP_CD= " & FilterVar(frm1.txtSold_to_party.Value, "''", "S") & " "
            arrParam(5) = "납품처"
  

            arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"
            arrField(1) = "PARTNER.BP_NM"
            arrField(2) = "PARTNER_FTN.BP_CD"
            arrField(3) = "PARTNER.BP_RGST_NO"
            arrField(4) = "PARTNER_FTN.PARTNER_FTN"
            arrField(5) = "PARTNER.CONTRY_CD"
     
            arrHeader(0) = "납품처"
            arrHeader(1) = "납품처명"
            arrHeader(2) = "거래처코드"
            arrHeader(3) = "사업자등록번호"
            arrHeader(4) = "거래처타입"
            arrHeader(5) = "국가코드"

            frm1.txtShip_to_party.focus
  
    End Select


    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetBp(arrRet, iRequried)
        Call GetTaxBizArea("BP")
    End If

End Function

'========================================
Function OpenOption(ByVal iOption)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    Select Case iOption
        Case 3
            If frm1.txtDeal_Type.ReadOnly = True Then
                IsOpenPop = False
                Exit Function
            End If

            If lsClickCfmYes = True Or lsClickCfmNo = True Then
                IsOpenPop = False
                Exit Function
            End If

            arrParam(0) = "판매유형"
            arrParam(1) = "B_MINOR"
            arrParam(2) = Trim(frm1.txtDeal_Type.Value)
            arrParam(3) = Trim(frm1.txtDeal_Type_nm.Value)
            arrParam(4) = "MAJOR_CD=" & FilterVar("S0001", "''", "S") & ""
            arrParam(5) = "판매유형"
              
            arrField(0) = "MINOR_CD"
            arrField(1) = "MINOR_NM"
                 
            arrHeader(0) = "판매유형"
            arrHeader(1) = "판매유형명"

            frm1.txtDeal_Type.focus
            
        Case 4
            If frm1.txtTrans_Meth.ReadOnly = True Then
                IsOpenPop = False
                Exit Function
            End If

            arrParam(0) = "운송방법"
            arrParam(1) = "B_MINOR"
            arrParam(2) = Trim(frm1.txtTrans_Meth.Value)
            arrParam(3) = Trim(frm1.txtTrans_Meth_nm.Value)
            arrParam(4) = "MAJOR_CD=" & FilterVar("B9009", "''", "S") & ""
            arrParam(5) = "운송방법"
              
            arrField(0) = "MINOR_CD"
            arrField(1) = "MINOR_NM"
                 
            arrHeader(0) = "운송방법"
            arrHeader(1) = "운송방법명"
            
            frm1.txtTrans_Meth.focus

        Case 5
            If frm1.txtPay_terms.ReadOnly = True Then
                IsOpenPop = False
                Exit Function
            End If

            If lsClickCfmYes = True Or lsClickCfmNo = True Then
                IsOpenPop = False
                Exit Function
            End If

            arrParam(0) = "결제방법"
            arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"
            arrParam(2) = Trim(frm1.txtPay_terms.Value)
            arrParam(3) = ""
            arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9004", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
                            & " And Config.MINOR_CD = Minor.MINOR_CD" _
                            & " And Config.SEQ_NO = 1" _
                            & " And Config.Reference = " & FilterVar("N", "''", "S") & " "

            arrParam(5) = "결제방법"
            arrField(0) = "Minor.MINOR_CD"
            arrField(1) = "Minor.MINOR_NM"
                 
            arrHeader(0) = "결제방법"
            arrHeader(1) = "결제방법명"

            frm1.txtPay_terms.focus
            
        Case 6
            If frm1.txtVat_Type.ReadOnly = True Then
                IsOpenPop = False
                Exit Function
            End If

            If lsClickCfmYes = True Or lsClickCfmNo = True Then
                IsOpenPop = False
                Exit Function
            End If

            arrParam(0) = "VAT유형"
            arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"
            arrParam(2) = Trim(frm1.txtVat_Type.Value)
            arrParam(3) = ""
            arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
            & " And Config.MINOR_CD = Minor.MINOR_CD" _
            & " And Config.SEQ_NO = 1"
            arrParam(5) = "VAT유형"
              
            arrField(0) = "Minor.MINOR_CD"
            arrField(1) = "Minor.MINOR_NM"
            arrField(2) = "Config.REFERENCE"
                      
            arrHeader(0) = "VAT유형"
            arrHeader(1) = "VAT유형명"
            arrHeader(2) = "VAT율"
            
            frm1.txtVat_Type.focus
                 
    End Select

    arrParam(3) = ""   '☜: [Condition Name Delete]

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetOption(arrRet, iOption)
    End If
End Function

'========================================
Function OpenDNType()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If lsClickCfmYes = True Then Exit Function

    IsOpenPop = True

    If frm1.txtDn_Type.ReadOnly = True Then
        IsOpenPop = False
        Exit Function
    End If

    arrParam(0) = "출하형태"
    arrParam(1) = "b_minor A, I_MOVETYPE_CONFIGURATION B, S_SO_TYPE_CONFIG C"
    arrParam(2) = Trim(frm1.txtDn_Type.Value)
    arrParam(3) = ""
    arrParam(4) = "A.minor_cd = B.mov_type And B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " " _
                    & "and A.major_cd = " & FilterVar("I0001", "''", "S") & " and b.mov_type = c.mov_type and c.so_mgmt_flag = " & FilterVar("N", "''", "S") & "  and c.EXPORT_FLAG = " & FilterVar("N", "''", "S") & "  and c.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  and c.REL_BILL_FLAG = " & FilterVar("Y", "''", "S") & " "
    arrParam(5) = "출하형태"

    arrField(0) = "A.MINOR_CD"
    arrField(1) = "A.MINOR_NM"
    arrField(2) = "C.SO_TYPE"
    arrField(3) = "C.RET_ITEM_FLAG"
    arrField(4) = "C.REL_BILL_FLAG"

    arrHeader(0) = "출하형태"
    arrHeader(1) = "출하형태명"
    arrHeader(2) = "수주형태"
    arrHeader(3) = "반품여부"
    arrHeader(4) = "매출여부"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=650px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
     
    IsOpenPop = False

    frm1.txtDn_Type.focus
    
    If arrRet(0) <> "" Then
        Call SetDNType(arrRet)
    End If
End Function

'========================================
Function OpenPlant()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If frm1.txtPlant.ReadOnly = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "공장"
    arrParam(1) = "B_PLANT"
    arrParam(2) = Trim(frm1.txtPlant.Value)
    arrParam(4) = ""
    arrParam(5) = "공장"
     
    arrField(0) = "PLANT_CD"
    arrField(1) = "PLANT_NM"
        
    arrHeader(0) = "공장"
    arrHeader(1) = "공장명"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
                                    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    frm1.txtPlant.focus
    
    If arrRet(0) <> "" Then
        Call SetPlant(arrRet)
    End If
End Function

'========================================
Function OpenSL()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If frm1.txtSlCd.ReadOnly = True Then Exit Function

    IsOpenPop = True
    
    arrParam(0) = "창고"
    arrParam(1) = "b_storage_location"
    arrParam(2) = Trim(frm1.txtSlCd.Value)
    arrParam(3) = ""
     
    If Len(frm1.txtPlant.Value) Then
         arrParam(4) = "plant_cd = " + FilterVar(frm1.txtPlant.Value, " ", "S")
    Else
         arrParam(4) = ""
    End If
     
    arrParam(5) = "창고"
     
    arrField(0) = "sl_cd"
    arrField(1) = "sl_nm"
        
    arrHeader(0) = "창고"
    arrHeader(1) = "창고명"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    frm1.txtSlCd.focus
    
    If arrRet(0) <> "" Then
        Call SetSL(arrRet)
    End If
End Function

' 재고담당자 Popup
'========================================
Function OpenInvMgrPopUp()

    Dim iArrRet
    Dim iArrParam(5), iArrField(6), iArrHeader(6)

    If IsOpenPop Then Exit Function

    IsOpenPop = True

    With frm1
        '재고담당 
        If .txtInvMgr.ReadOnly Then
            IsOpenPop = False
            Exit Function
        End If

        iArrParam(1) = "dbo.B_MINOR"
        iArrParam(2) = Trim(.txtInvMgr.Value)
        iArrParam(3) = ""
        iArrParam(4) = "MAJOR_CD = " & FilterVar("I0004", "''", "S") & ""
                
        iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
        iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
                            
        iArrHeader(0) = .txtInvMgr.ALT
        iArrHeader(1) = .txtInvMgrNm.ALT

        .txtInvMgr.focus

        iArrParam(0) = iArrHeader(0)                            ' 팝업 Title
        iArrParam(5) = iArrHeader(0)                            ' 조회조건 명칭 

        iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
            "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

        IsOpenPop = False

        If iArrRet(0) <> "" Then
            .txtInvMgr.Value = iArrRet(0)
            .txtInvMgrNm.Value = iArrRet(1)
            lgBlnFlgChgValue1 = True            ' Header 정보 변경여부 설정 
        End If
    End With
    
End Function

'========================================
Function OpenItem(ByVal strCode)
    Dim iCalledAspName
    Dim arrParam(2)
    Dim strRet

    arrParam(0) = strCode

    If Len(frm1.txtPlant.Value) Then
        arrParam(1) = frm1.txtPlant.Value
    Else
        Call DisplayMsgBox("189220", "X", "X", "X")
        Exit Function
    End If
  
    ggoSpread.Source = frm1.vspdData
    frm1.vspdData.Col = C_SlCd
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
 
    If Len(frm1.vspdData.Text) Then
        arrParam(2) = frm1.vspdData.Text
    Else
        Call DisplayMsgBox("17A002", "X", "창고", "X")
        Exit Function
    End If

    If IsOpenPop = True Then Exit Function
   
    iCalledAspName = AskPRAspName("S3112PA3")
        
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "S3112PA3", "x")
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True
  
    strRet = window.showModalDialog(iCalledAspName, Array(window.Parent, arrParam), _
        "dialogWidth=820px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If strRet(0) <> "" Then
        frm1.vspdData.Col = C_ItemCd
        frm1.vspdData.Text = strRet(0)
        frm1.vspdData.Col = C_ItemNm
        frm1.vspdData.Text = strRet(1)
        frm1.vspdData.Col = C_LotMgmtFlag
        frm1.vspdData.Text = strRet(2)

        Call vspdData_Change(C_ItemCd, frm1.vspdData.Row)
     End If

End Function

'========================================
Function OpenDnDtl(ByVal strCode, ByVal iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    Dim OriginCol, TempCd

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    Select Case iWhere
        Case 1 '품목 
            arrParam(1) = "b_item item, b_plant plant, b_item_by_plant item_plant"
            arrParam(2) = strCode
            '  arrParam(4) = "item.item_cd=item_plant.item_cd and plant.plant_cd=item_plant.plant_cd and (item_plant.item_acct = '10' or item_plant.item_acct = '20' or item_plant.item_acct = '60')"
            arrParam(4) = "item.item_cd=item_plant.item_cd and plant.plant_cd=item_plant.plant_cd"
            arrParam(5) = "품목"
             
            arrField(0) = "item.item_cd"
            arrField(1) = "item.item_nm"
            arrField(2) = "plant.plant_cd"
            arrField(3) = "plant.plant_nm"
                
            arrHeader(0) = "품목"
            arrHeader(1) = "품목명"
            arrHeader(2) = "공장"
            arrHeader(3) = "공장명"

        Case 2 '단위 
            arrParam(1) = "b_unit_of_measure"
            arrParam(2) = strCode
            arrParam(4) = ""
            arrParam(5) = "단위"
             
            arrField(0) = "unit"
            arrField(1) = "unit_nm"
                
            arrHeader(0) = "단위"
            arrHeader(1) = "단위명"


        Case 3
            arrParam(0) = "VAT포함구분"
            arrParam(1) = "B_MINOR"
            arrParam(2) = strCode
            arrParam(4) = "MAJOR_CD=" & FilterVar("S4035", "''", "S") & ""
            arrParam(5) = "VAT포함구분"
              
            arrField(0) = "MINOR_CD"
            arrField(1) = "MINOR_NM"
                      
            arrHeader(0) = "VAT포함구분"
            arrHeader(1) = "VAT포함구분명"

        Case 4 'VAT유형 
            arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"
            arrParam(2) = strCode
            arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
                            & " And Config.MINOR_CD = Minor.MINOR_CD" _
                            & " And Config.SEQ_NO = 1"
            arrParam(5) = "VAT유형"

            arrField(0) = "Minor.MINOR_CD"
            arrField(1) = "Minor.MINOR_NM"
            arrField(2) = "Config.REFERENCE"
                      
            arrHeader(0) = "VAT유형"
            arrHeader(1) = "VAT유형명"
            arrHeader(2) = "VAT율"

        Case 5 '창고 
            arrParam(1) = "b_storage_location"
            arrParam(2) = strCode
              
            If Len(frm1.txtPlant.Value) Then
                arrParam(4) = "plant_cd = " & FilterVar(frm1.txtPlant.Value, "''", "S") & " "
            Else
                Call DisplayMsgBox("189220", "X", "X", "X")
                IsOpenPop = False
                Exit Function
            End If
             
            arrParam(5) = "창고"
             
            arrField(0) = "sl_cd"
            arrField(1) = "sl_nm"
                
            arrHeader(0) = "창고"
            arrHeader(1) = "창고명"

        Case 6
            arrParam(1) = "B_MINOR"
            arrParam(2) = strCode
            arrParam(4) = "MAJOR_CD=" & FilterVar("B9017", "''", "S") & ""
            arrParam(5) = "반품유형"
             
            arrField(0) = "Minor_cd"
            arrField(1) = "Minor_nm"
                
            arrHeader(0) = "반품유형"
            arrHeader(1) = "반품유형명"

    End Select

    arrParam(0) = arrParam(5)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetDNDtl(arrRet, iWhere)
    End If
     
End Function

'========================================
Function OpenLotNoPopup(ByVal iWhere)
    Dim iCalledAspName
    Dim arrRet
    Dim Param1
    Dim Param2
    Dim Param3
    Dim Param4
    Dim Param5
    Dim Param6
    Dim Param7
    Dim Param8, Param9

    If IsOpenPop = True Then Exit Function

    frm1.vspdData.Col = C_ItemCd
     
    If frm1.vspdData.Text = "" Then
        Call DisplayMsgBox("17A002", "X", "품목코드", "X")
        Exit Function
    End If
  
    frm1.vspdData.Col = C_SlCd
 
    If frm1.vspdData.Text = "" Then
        Call DisplayMsgBox("17A002", "X", "창고", "X")
        Exit Function
    End If

    iCalledAspName = AskPRAspName("I2212RA1")
        
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "I2212RA1", "x")
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True
 
    With frm1.vspdData
        .Row = iWhere

        .Col = C_SlCd:          Param1 = .Text
        .Col = C_ItemCd:        Param2 = .Text
        Param3 = "*" 'Tracking No
        Param4 = frm1.txtPlant.Value
        Param5 = "J"
        .Col = C_LotNo:         Param6 = .Text
        Param7 = ""
        .Col = C_ItemNm:        Param8 = .Text
        .Col = C_DnUnit:        Param9 = .Text
           
        arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent, Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9), _
                                        "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
         
        IsOpenPop = False

        If Trim(arrRet(0)) <> "" Then
            .Col = C_LotNo: .Text = arrRet(3)
            .Col = C_LotSeq: .Text = arrRet(4)
            Call vspdData_Change(.Col, .Row)
            lgBlnFlgDtlChgValue = True
        End If
    End With
 
End Function

'========================================
Function OpenZip()
    Dim arrRet
    Dim arrParam(2)

    If IsOpenPop = True Then Exit Function
    
    If Trim(frm1.txtShip_to_party.Value) = "" Then
        MsgBox "납품처를 먼저 입력하세요", vbInformation, Parent.gLogoName
        Call changeTabs(TAB1) '첫번째 Tab
        frm1.txtShip_to_party.focus
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True
    
    arrParam(0) = Trim(frm1.txtZIP_cd.Value)
    arrParam(1) = ""
    arrParam(2) = Trim(frm1.txtHCntryCd.Value)

    arrRet = window.showModalDialog("../../comasp/ZipPopup.asp", Array(window.Parent, arrParam), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
        
    IsOpenPop = False
    
    frm1.txtZIP_cd.focus
    
    If arrRet(0) <> "" Then
        Call SetZip(arrRet)
    End If
            
End Function

'2003.01.11 SMJ
'========================================
Function OpenTransCo()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "운송회사"
    arrParam(1) = "B_MAJOR A , B_MINOR B"
    arrParam(2) = ""
    arrParam(3) = ""
    arrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9031", "''", "S") & " "
    arrParam(5) = "운송회사"

    arrField(0) = "B.MINOR_CD"
    arrField(1) = "B.MINOR_NM"

    arrHeader(0) = "운송회사"
    arrHeader(1) = "운송회사명"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
            "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    frm1.txtTransCo.focus
    
    If arrRet(0) <> "" Then
        Call SetTransCo(arrRet)
    End If
End Function

'========================================
Function OpenVehicleNo()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function
        
    IsOpenPop = True

    arrParam(0) = "차량번호"
    arrParam(1) = "B_MAJOR A , B_MINOR B"
    arrParam(2) = ""
    arrParam(3) = ""
    arrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9032", "''", "S") & " "
    arrParam(5) = "차량번호"

    arrField(0) = "B.MINOR_CD"
    arrField(1) = "B.MINOR_NM"

    arrHeader(0) = "차량관리번호"
    arrHeader(1) = "차량번호"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
                                    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    frm1.txtVehicleNo.focus
    
    If arrRet(0) <> "" Then
        Call SetVehicleNo(arrRet)
    End If
    
End Function

'========================================
Function GetTaxBizArea(ByVal pvStrFlag)

 Dim iStrSelectList, iStrFromList, iStrWhereList
 Dim iiStrSoldToParty, iStrSalesGrp, iStrTaxBizArea
 Dim iStrRs
 Dim iArrTaxBizArea(1), iArrTemp

 GetTaxBizArea = False
 
 '세금신고 사업장 Edting시 유효값 Check 및 사업장 명 Fetch
 If pvStrFlag = "NM" Then
    iStrTaxBizArea = frm1.txtTaxBizAreaCd.Value
 Else
    iiStrSoldToParty = frm1.txtSold_to_party.Value
    iStrSalesGrp = frm1.txtSales_Grp.Value
    '발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 
    If Len(iiStrSoldToParty) > 0 And Len(iStrSalesGrp) > 0 Then pvStrFlag = "*"
 End If
 
 iStrSelectList = " * "
 iStrFromList = " dbo.ufn_s_GetTaxBizArea ( " & FilterVar(iiStrSoldToParty, "''", "S") & ",  " & FilterVar(iStrSalesGrp, "''", "S") & ",  " & FilterVar(iStrTaxBizArea, "''", "S") & ",  " & FilterVar(pvStrFlag, "''", "S") & ") "
 iStrWhereList = ""
 
 Err.Clear
    
 If CommonQueryRs2by2(iStrSelectList, iStrFromList, iStrWhereList, iStrRs) Then
	If iStrRs <> "" Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrTaxBizArea(0) = iArrTemp(1)
		iArrTaxBizArea(1) = iArrTemp(2)

		Call SetRequried(iArrTaxBizArea, 4)
		GetTaxBizArea = True
	End If
 Else
    If Err.Number <> 0 Then Err.Clear

    ' 세금 신고 사업장을 Editing한 경우 
    If pvStrFlag = "NM" Then
        If Not OpenRequried(4) Then
            frm1.txtTaxBizAreaCd.Value = ""
            frm1.txtTaxBizAreaNm.Value = ""
        Else
            GetTaxBizArea = True
        End If
    End If
 End If
End Function

'========================================
Sub GetContryCd()
    Dim iContryCd
    
    Call CommonQueryRs("A.Contry_Cd", " B_BIZ_PARTNER A", "A.BP_CD =  " & FilterVar(frm1.txtShip_to_party.Value, "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    If lgF0 = "" Then
        Call DisplayMsgBox("229938", "X", "X", "X")
        frm1.txtHCntryCd.Value = ""
        Exit Sub
    End If

    iContryCd = Split(lgF0, Chr(11))
    
    frm1.txtHCntryCd.Value = iContryCd(0)
    
End Sub

'========================================
Function SetRequried(ByVal arrRet, ByVal iRequried)

    If arrRet(0) <> "" Then

        Select Case iRequried
            Case 1            ' 영업그룹 
                frm1.txtSales_Grp.Value = arrRet(0)
                frm1.txtSales_GrpNm.Value = arrRet(1)
                Call GetTaxBizArea("BA")
            Case 2
                frm1.txtCol_Type.Value = arrRet(0)
                frm1.txtCol_Type_nm.Value = arrRet(1)
                Call txtCol_Type_OnChange
            Case 4
                frm1.txtTaxBizAreaCd.Value = arrRet(0)
                frm1.txtTaxBizAreaNm.Value = arrRet(1)
        End Select

        lgBlnFlgChgValue = True

    End If

End Function

'========================================
Function SetBp(ByVal arrRet, ByVal iRequried)

    If arrRet(0) <> "" Then

        Select Case iRequried
            Case 0            ' 주문처 
                frm1.txtSold_to_party.Value = arrRet(0)
                frm1.txtSold_to_partyNm.Value = arrRet(1)
                Call SoldToPartyLookUp
            Case 1
                frm1.txtShip_to_party.Value = arrRet(0)
                frm1.txtShip_to_partyNm.Value = arrRet(1)
                frm1.txtHCntryCd.Value = arrRet(5)
        End Select

        lgBlnFlgChgValue = True

    End If

End Function

'========================================
Function SetOption(ByVal arrRet, ByVal iOption)
    If arrRet(0) <> "" Then

        Select Case iOption
            Case 3
                frm1.txtDeal_Type.Value = arrRet(0)
                frm1.txtDeal_Type_nm.Value = arrRet(1)
            Case 4  ' 운송방법 
                frm1.txtTrans_Meth.Value = arrRet(0)
                frm1.txtTrans_Meth_nm.Value = arrRet(1)
            Case 5
                frm1.txtPay_terms.Value = arrRet(0)
                frm1.txtPay_terms_nm.Value = arrRet(1)
            Case 6
                frm1.txtVat_Type.Value = arrRet(0)
                frm1.txtVatTypeNm.Value = arrRet(1)
                frm1.txtVat_rate.Text = arrRet(2)
                Call ChangeGridVatRate
            Case 12
                frm1.txtVat_Inc_Flag.Value = arrRet(0)
                frm1.txtVat_Inc_Flag_Nm.Value = arrRet(1)

        End Select

        lgBlnFlgChgValue = True

    End If
End Function

'========================================
Function SetDNType(arrRet)
    frm1.txtDn_Type.Value = arrRet(0)
    frm1.txtDn_TypeNm.Value = arrRet(1)
    frm1.txtSO_TYPE.Value = arrRet(2)
    frm1.txtRetItemFlag.Value = arrRet(3)
    frm1.txtRetBillFlag.Value = arrRet(4)

    lgBlnFlgChgValue = True

    Call ChangeReturnItemRetField
End Function

'========================================
Function SetPlant(arrRet)
    frm1.txtPlant.Value = arrRet(0)
    frm1.txtPlantNm.Value = arrRet(1)
    lgBlnFlgChgValue = True

    Call txtPlant_OnChange
End Function

'========================================
Function SetSL(arrRet)
    frm1.txtSlCd.Value = arrRet(0)
    frm1.txtSlNm.Value = arrRet(1)
    lgBlnFlgChgValue = True
End Function

'========================================
Function SetDNDtl(ByVal arrRet, ByVal iWhere)
    With frm1

    Select Case iWhere
        Case 1 '품목 
            .vspdData.Col = C_ItemCd
            .vspdData.Text = arrRet(0)
            .vspdData.Col = C_ItemName
            .vspdData.Text = arrRet(1)
            .vspdData.Col = C_PlantCd
            .vspdData.Text = arrRet(2)

            Call vspdData_Change(C_ItemCd, frm1.vspdData.Row)
        Case 2 '단위 
            .vspdData.Col = C_DnUnit
            .vspdData.Text = arrRet(0)

            Call vspdData_Change(C_DnUnit, frm1.vspdData.Row)
        Case 4 'VAT유형 
            .vspdData.Col = C_VatType
            .vspdData.Text = arrRet(0)
            .vspdData.Col = C_VatTypeNm
            .vspdData.Text = arrRet(1)
            .vspdData.Col = C_VatRate
            .vspdData.Text = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
            Call SetVATTypeForSpread(frm1.vspdData.ActiveRow)

            Call vspdData_Change(C_VatType, frm1.vspdData.Row)
        Case 5 '창고 
            .vspdData.Col = C_SlCd
            .vspdData.Text = arrRet(0)

            Call vspdData_Change(C_SlCd, frm1.vspdData.Row)
        Case 6
            .vspdData.Col = C_RetType
            .vspdData.Text = arrRet(0)
            .vspdData.Col = C_RetTypeNm
            .vspdData.Text = arrRet(1)
            Call vspdData_Change(C_RetType, .vspdData.Row)
        Case Else
            Exit Function
        End Select
      
    End With

    lgBlnFlgDtlChgValue = True
 
End Function

'========================================
Function SetShipToPlceRef(ByVal arrRet)
    On Error Resume Next
    
    frm1.txtSTP_Inf_No.Value = arrRet(0)            '납품처상세정보번호 
    frm1.txtZIP_cd.Value = arrRet(1)                '우편번호 
    frm1.txtADDR1_Dlv.Value = arrRet(2)             '납품주소1
    frm1.txtADDR2_Dlv.Value = arrRet(3)             '납품주소2
    frm1.txtADDR3_Dlv.Value = arrRet(4)             '납품주소3
    frm1.txtReceiver.Value = arrRet(5)              '인수자명 
    frm1.txtTel_No1.Value = arrRet(6)               '전화번호1
    frm1.txtTel_No2.Value = arrRet(7)               '전화번호2
    lgBlnFlgChgValue1 = True

End Function

'========================================
Function SetTrnsMethRef(ByVal arrRet)
        
    frm1.txtTrnsp_Inf_No.Value = arrRet(0)          '운송정보번호 
    frm1.txtTransCo.Value = arrRet(1)               '운송회사 
    frm1.txtDriver.Value = arrRet(2)                '운전자명 
    frm1.txtVehicleNo.Value = arrRet(3)             '차량번호 
    frm1.txtSender.Value = arrRet(4)                '인계자명 
    lgBlnFlgChgValue2 = True

End Function

'========================================
Sub SoldToPartyLookUp()

    Err.Clear
    
    If LayerShowHide(1) = False Then
        Exit Sub
    End If
     
    Dim strVal
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUp"
    strVal = strVal & "&txtSold_to_party=" & Trim(frm1.txtSold_to_party.Value)
    
    Call RunMyBizASP(MyBizASP, strVal)
 
End Sub

'========================================
' 박정순 추가 (2006-05-26) 영업그룹명/결제방법명이 default 처리 안되는 것 수정.
'========================================
Sub SoldToPartyLookUpOK()
    Dim aSold_to_partyNm 

    Call CommonQueryRs("sales_grp_nm", " B_SALES_GRP", " sales_grp = " & FilterVar(frm1.txtSales_Grp.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    If lgF0 = "" Then
	    frm1.txtSales_GrpNm.Value = ""
        Exit Sub
    End If

    aSold_to_partyNm = Split(lgF0, Chr(11))

    frm1.txtSales_GrpNm.Value = aSold_to_partyNm(0)


    Call CommonQueryRs("minor_nm", " b_minor", " major_cd  = " & FilterVar("B9004", "''", "S") &  " AND  minor_cd = " & FilterVar(frm1.txtPay_terms.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    If lgF0 = "" Then
	    frm1.txtPay_terms_nm.Value = ""
        Exit Sub
    End If

    aSold_to_partyNm = Split(lgF0, Chr(11))

    frm1.txtPay_terms_nm.Value = aSold_to_partyNm(0)


End Sub

'2002.12.20 SMJ
'========================================
Sub SetZip(arrRet)
    With frm1
        .txtZIP_cd.Value = arrRet(0)
        .txtADDR1_Dlv.Value = arrRet(1)
         .txtSTP_Inf_No.Value = ""
        lgBlnFlgChgValue1 = True
    End With
End Sub

'2003.01.11 SMJ
'========================================
Function SetTransCo(arrRet)
    frm1.txtTransCo.Value = arrRet(1)
    If frm1.txtTrnsp_Inf_No.Value <> "" Then
        frm1.txtTrnsp_Inf_No.Value = ""
    End If
        
    lgBlnFlgChgValue2 = True
End Function

'========================================
Function SetVehicleNo(arrRet)
    frm1.txtVehicleNo.Value = arrRet(1)
    If frm1.txtTrnsp_Inf_No.Value <> "" Then
        frm1.txtTrnsp_Inf_No.Value = ""
    End If
        
    lgBlnFlgChgValue2 = True
End Function

'========================================
Sub DNTypeLookUp()
    Dim iDNTypeNm, iSOType, iRetItemFlag, iRelBillFlag

    Err.Clear
    
    Call CommonQueryRs("A.MINOR_NM, C.SO_TYPE, C.RET_ITEM_FLAG, C.REL_BILL_FLAG", " B_MINOR A, I_MOVETYPE_CONFIGURATION B, S_SO_TYPE_CONFIG C", " A.minor_cd = B.mov_type And B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & "  and A.major_cd = " & FilterVar("I0001", "''", "S") & " and b.mov_type = c.mov_type and c.so_mgmt_flag = " & FilterVar("N", "''", "S") & "  and c.EXPORT_FLAG = " & FilterVar("N", "''", "S") & " and c.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  and c.REL_BILL_FLAG = " & FilterVar("Y", "''", "S") & "  And A.Minor_cd = " & FilterVar(frm1.txtDn_Type.Value, "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    If lgF0 = "" Then
        Call DisplayMsgBox("204255", "X", "X", "X")
        frm1.txtDn_Type.Value = ""
        frm1.txtDn_TypeNm.Value = ""
        frm1.txtSO_TYPE.Value = ""
        frm1.txtRetItemFlag.Value = ""
        frm1.txtRetBillFlag.Value = ""
        Exit Sub
    End If

    iDNTypeNm = Split(lgF0, Chr(11))
    iSOType = Split(lgF1, Chr(11))
    iRetItemFlag = Split(lgF2, Chr(11))
    iRelBillFlag = Split(lgF3, Chr(11))

    frm1.txtDn_TypeNm.Value = iDNTypeNm(0)
    frm1.txtSO_TYPE.Value = iSOType(0)
    frm1.txtRetItemFlag.Value = iRetItemFlag(0)
    frm1.txtRetBillFlag.Value = iRelBillFlag(0)
     
    Call ChangeReturnItemRetField
End Sub

'========================================
Sub UnLockColor_CfmNo()
    With ggoOper
        Call .SetReqAttr(frm1.txtSold_to_party, "N")
        Call .SetReqAttr(frm1.txtShip_to_party, "N")
        Call .SetReqAttr(frm1.txtDn_Type, "N")
        Call .SetReqAttr(frm1.txtDeal_Type, "N")
        Call .SetReqAttr(frm1.txtPlant, "N")
        Call .SetReqAttr(frm1.txtSales_Grp, "N")
        Call .SetReqAttr(frm1.txtDlvyDt, "N")
        Call .SetReqAttr(frm1.txtRemark, "D")
        Call .SetReqAttr(frm1.txtPlannedGIDt, "N")
        Call .SetReqAttr(frm1.txtPay_terms, "N")
        Call .SetReqAttr(frm1.txtTaxBizAreaCd, "N")
        Call .SetReqAttr(frm1.txt_Payterms_txt, "D")
        Call .SetReqAttr(frm1.txtVat_Type, "D")
        Call .SetReqAttr(frm1.rdoVat_Inc_flag1, "D")
        Call .SetReqAttr(frm1.rdoVat_Inc_flag2, "D")
        Call .SetReqAttr(frm1.rdoVat_Calc_Type1, "D")
        Call .SetReqAttr(frm1.rdoVat_Calc_Type2, "D")
        Call .SetReqAttr(frm1.txtTrans_Meth, "D")
        Call .SetReqAttr(frm1.txtDlvyPlace, "D")
        Call .SetReqAttr(frm1.chkArFlag, "D")
        Call .SetReqAttr(frm1.chkVatFlag, "D")
        Call .SetReqAttr(frm1.txtCol_Type, "D")
        Call .SetReqAttr(frm1.txtSlCd, "D")
        Call .SetReqAttr(frm1.txtInvMgr, "D")
    End With
End Sub

'========================================
Sub LockColor_CfmYes()
    With ggoOper
        Call .SetReqAttr(frm1.txtSold_to_party, "Q")
        Call .SetReqAttr(frm1.txtShip_to_party, "Q")
        Call .SetReqAttr(frm1.txtDn_Type, "Q")
        Call .SetReqAttr(frm1.txtDeal_Type, "Q")
        Call .SetReqAttr(frm1.txtPlant, "Q")
        Call .SetReqAttr(frm1.txtSales_Grp, "Q")
        Call .SetReqAttr(frm1.txtGI_Dt, "Q")
        Call .SetReqAttr(frm1.txtDlvyDt, "Q")
        Call .SetReqAttr(frm1.txtRemark, "Q")
        Call .SetReqAttr(frm1.txtPlannedGIDt, "Q")
        Call .SetReqAttr(frm1.txtPay_terms, "Q")
        Call .SetReqAttr(frm1.txtTaxBizAreaCd, "Q")
        Call .SetReqAttr(frm1.txt_Payterms_txt, "Q")
        Call .SetReqAttr(frm1.txtVat_Type, "Q")
        Call .SetReqAttr(frm1.rdoVat_Inc_flag1, "Q")
        Call .SetReqAttr(frm1.rdoVat_Inc_flag2, "Q")
        Call .SetReqAttr(frm1.rdoVat_Calc_Type1, "Q")
        Call .SetReqAttr(frm1.rdoVat_Calc_Type2, "Q")
        Call .SetReqAttr(frm1.txtTrans_Meth, "Q")
        Call .SetReqAttr(frm1.txtDlvyPlace, "Q")
        Call .SetReqAttr(frm1.chkArFlag, "Q")
        Call .SetReqAttr(frm1.chkVatFlag, "Q")
        Call .SetReqAttr(frm1.txtCol_Type, "Q")
        Call .SetReqAttr(frm1.txtSlCd, "Q")
        Call .SetReqAttr(frm1.txtInvMgr, "Q")
        '2003.01.22 SMJ
        Call .SetReqAttr(frm1.txtArriv_dt, "Q")
        Call .SetReqAttr(frm1.txtArriv_Tm, "Q")
    End With
End Sub

'========================================
Sub CurFormatNumericOCX()
 With frm1

  '개설금액 
  ggoOper.FormatFieldByObjectOfCur .txtNet_amt, .txtDoc_cur.Value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
  
 End With
End Sub

'========================================
Sub ChangeVATIncFlagRetField()

 If frm1.rdoVat_Calc_Type2.Checked = True And frm1.txtVat_Type.Value = "" Then
  Call DisplayMsgBox("17A002", "X", "VAT유형", "X")
  frm1.rdoVat_Calc_Type1.Checked = True
  Exit Sub
 End If

 If frm1.txtGINo.Value = "" Then
  ggoSpread.Source = frm1.vspdData

  If frm1.rdoVat_Calc_Type1.Checked = True Then
   Call ggoOper.SetReqAttr(frm1.txtVat_Type, "D")

   If frm1.vspdData.MaxRows > 0 Then
    ggoSpread.SpreadUnLock C_VatType, -1, C_VatType
    ggoSpread.SpreadUnLock C_VatIncFlagNm, -1, C_VatIncFlagNm
    ggoSpread.SpreadUnLock C_VatTypePopup, -1, C_VatTypePopup
'    ggoSpread.SSSetRequired C_VatIncFlag, -1
    ggoSpread.SSSetRequired C_VatIncFlagNm, -1
    ggoSpread.SSSetRequired C_VatType, -1
   End If
  Else
   Call ggoOper.SetReqAttr(frm1.txtVat_Type, "Q")
   
   If frm1.vspdData.MaxRows > 0 Then
   
    ggoSpread.SSSetProtected C_VatIncFlag, -1
    ggoSpread.SSSetProtected C_VatIncFlagNm, -1
    ggoSpread.SSSetProtected C_VatType, -1
    ggoSpread.SSSetProtected C_VatTypePopup, -1, C_VatTypePopup

    Dim i
    For i = 1 To frm1.vspdData.MaxRows
     frm1.vspdData.Row = i
     frm1.vspdData.Col = C_VatType
     frm1.vspdData.Text = frm1.txtVat_Type.Value
     frm1.vspdData.Col = C_VatTypeNm
     frm1.vspdData.Text = frm1.txtVatTypeNm.Value
     
     If frm1.rdoVat_Inc_flag1.Checked = True Then
      frm1.vspdData.Col = C_VatIncFlag
      frm1.vspdData.Text = "1"
      frm1.vspdData.Col = C_VatIncFlagNm
      frm1.vspdData.Text = "별도"
     Else
      frm1.vspdData.Col = C_VatIncFlag
      frm1.vspdData.Text = "2"
      frm1.vspdData.Col = C_VatIncFlagNm
      frm1.vspdData.Text = "포함"
     End If
    Next
   End If
  End If
 End If
End Sub

'=========================================
Sub ChangeLotRetField(ByVal pvStartRow, ByVal pvEndRow)
    Dim iRow
    
    
    For iRow = pvStartRow To pvEndRow
        frm1.vspdData.Row = iRow
        frm1.vspdData.Col = C_LotMgmtFlag
 
        If frm1.vspdData.Text = "Y" Then
            ggoSpread.SpreadUnLock C_LotNo, iRow, C_LotNo, iRow
            ggoSpread.SpreadUnLock C_LotSeq, iRow, C_LotSeq, iRow
            ggoSpread.SpreadUnLock C_LotNoPopup, iRow, C_LotNoPopup, iRow
            ggoSpread.SSSetRequired C_LotNo, iRow, iRow
            ggoSpread.SSSetRequired C_LotSeq, iRow, iRow
        Else
            ggoSpread.SSSetProtected C_LotNo, iRow, iRow
            ggoSpread.SSSetProtected C_LotSeq, iRow, iRow
            ggoSpread.SSSetProtected C_LotNoPopup, iRow, iRow
        End If
    Next
End Sub

'========================================
Function CookiePage(ByVal Kubun)

 On Error Resume Next

 Const CookieSplit = 4877      'Cookie Split String : CookiePage Function Use
 Dim strTemp, arrVal

 If Kubun = 1 Then

  WriteCookie CookieSplit, frm1.txtConDn_no.Value & Parent.gRowSep & frm1.txtRadioType.Value

 ElseIf Kubun = 0 Then

  strTemp = ReadCookie(CookieSplit)
   
  If strTemp = "" Then Exit Function
   
  arrVal = Split(strTemp, Parent.gRowSep)

  If arrVal(0) = "" Then Exit Function
  
  frm1.txtConDn_no.Value = arrVal(0)
  
'  frm1.txtConDN_no.value =  strTemp

  If Err.Number <> 0 Then
   Err.Clear
   WriteCookie CookieSplit, ""
   Exit Function
  End If
  
  Call MainQuery
     
  WriteCookie CookieSplit, ""
  
 End If

End Function

'========================================
Function JumpChgCheck()

 Dim IntRetCD

 If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then Exit Function
 End If

 Call CookiePage(1)
 Call PgmJump(BIZ_PGM_JUMP_ID)

End Function

'========================================
Sub InitVATTypeInfo()
    On Error Resume Next

    Dim iStrSelectList, iStrFromList, iStrWhereList
    Dim iArrVATType, iArrVATTypeNm, iArrVATRate
    Dim iIntIndex
    
    Err.Clear
    
    iStrSelectList = " Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE "
    iStrFromList = " B_MINOR Minor,B_CONFIGURATION Config "
    iStrWhereList = " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 "
    
    If CommonQueryRs(iStrSelectList, iStrFromList, iStrWhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
        iArrVATType = Split(lgF0, Parent.gColSep)
        iArrVATTypeNm = Split(lgF1, Parent.gColSep)
        iArrVATRate = Split(lgF2, Parent.gColSep)
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description
            Err.Clear
        End If
        Exit Sub
    End If

    ReDim lgArrVATTypeInfo(UBound(iArrVATType) - 1, 2)

    For iIntIndex = 0 To UBound(iArrVATType) - 1
        lgArrVATTypeInfo(iIntIndex, 0) = iArrVATType(iIntIndex)
        lgArrVATTypeInfo(iIntIndex, 1) = iArrVATTypeNm(iIntIndex)
        lgArrVATTypeInfo(iIntIndex, 2) = iArrVATRate(iIntIndex)
    Next
End Sub

'========================================
Sub GetVATType(ByVal pvStrVATType, ByRef prStrVATTypeNm, ByRef prStrVATRate)
    Dim iIntIndex

    For iIntIndex = 0 To UBound(lgArrVATTypeInfo, 1)

'	 품목정보의 vat type 이 space 가 포함되어 있는 경우 처리. (박정순 수정 (2006-05-26) ) 
'        If UCase(lgArrVATTypeInfo(iIntIndex, 0)) = UCase(pvStrVATType) Then  
         If Trim(UCase(lgArrVATTypeInfo(iIntIndex, 0))) = Trim(UCase(pvStrVATType)) Then
            prStrVATTypeNm = lgArrVATTypeInfo(iIntIndex, 1)
            prStrVATRate = lgArrVATTypeInfo(iIntIndex, 2)
            Exit Sub
        End If
    Next

    prStrVATTypeNm = ""
    prStrVATRate = "0"
End Sub

'========================================
Sub SetVATTypeForSpread(ByVal pvIntRow)
    Dim iStrVATType, iStrVATTypeNm, iStrVATRate
    
    With frm1.vspdData
        .Row = pvIntRow
        .Col = C_VatType: iStrVATType = .Text
         
        '
        Call GetVATType(iStrVATType, iStrVATTypeNm, iStrVATRate)
         
        .Col = C_VatTypeNm: .Text = iStrVATTypeNm
        .Col = C_VatRate: .Text = iStrVATRate
    End With
End Sub

'=====================================================
Sub SetVatTypeForHdr()
    Dim iStrVATType, iStrVATTypeNm, iStrVATRate

    With frm1
        iStrVATType = .txtVat_Type.Value
         
        Call GetVATType(iStrVATType, iStrVATTypeNm, iStrVATRate)

        .txtVatTypeNm.Value = iStrVATTypeNm
        .txtVat_rate.Text = iStrVATRate 'UNIFormatNumber(UNICDbl(iStrVATRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
    End With
End Sub

'==================================================
Function LotNoChange(CRow)

    Dim strCon, iStrItemCd, strSlCd, strLotNo
    
    If Trim(frm1.txtPlant.Value) = "" Then
        Call DisplayMsgBox("189220", "X", "X", "X")
        Exit Function
    End If

    frm1.vspdData.Row = CRow
    frm1.vspdData.Col = C_ItemCd
    If Trim(frm1.vspdData.Text) = "" Then
        Exit Function
    Else
        iStrItemCd = frm1.vspdData.Text
    End If
    frm1.vspdData.Col = C_SlCd
    If Trim(frm1.vspdData.Text) = "" Then
        Exit Function
    Else
        strSlCd = frm1.vspdData.Text
    End If
    
    frm1.vspdData.Col = C_LotNo
    If Trim(frm1.vspdData.Text) = "" Then
        Exit Function
    Else
        strLotNo = frm1.vspdData.Text
    End If
        
    strCon = " PLANT_CD = " & FilterVar(frm1.txtPlant.Value, "''", "S") & "  AND ITEM_CD =  " & FilterVar(iStrItemCd, "''", "S") & "  "
    strCon = strCon & " AND SL_CD =  " & FilterVar(strSlCd, "''", "S") & "  AND TRACKING_NO = " & FilterVar("*", "''", "S") & "  AND LOT_NO =  " & FilterVar(strLotNo, "''", "S") & "  "

    Call CommonQueryRs("LOT_NO", "I_ONHAND_STOCK_DETAIL", strCon, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    If lgF0 = "" Then
        Call DisplayMsgBox("161101", "X", "X", "X")
        frm1.vspdData.Text = ""
        Exit Function
    End If

End Function

'========================================
Function GetItemPrice(iRow)

    Dim iStrItemCd, iStrDnUnit, iStrSoldToParty, iStrPayMeth, iStrCurrency, iStrDealType, iStrValidDt
    Dim iStrSelectList, iStrFromList, iStrWhereList
    Dim iStrRs, iStrItemInfo

    With frm1
        .vspdData.Row = iRow
        
        .vspdData.Col = C_ItemCd
        iStrItemCd = .vspdData.Text
        
        .vspdData.Col = C_DnUnit
        iStrDnUnit = .vspdData.Text

        iStrSoldToParty = .txtSold_to_party.Value
        iStrPayMeth = .txtPay_terms.Value
        iStrCurrency = .txtCurrency.Value
        iStrDealType = .txtDeal_Type.Value
        iStrValidDt = UniConvDateToYYYYMMDD(.txtPlannedGIDt.Text, Parent.gDateFormat, "")
    End With

    If Len(Trim(iStrItemCd)) = 0 Or Len(Trim(iStrDnUnit)) = 0 Then Exit Function
    
    iStrSelectList = " dbo.ufn_s_GetItemSalesPrice( " & FilterVar(iStrSoldToParty, "''", "S") & ",  " & FilterVar(iStrItemCd, "''", "S") & ", " & _
                                                  " " & FilterVar(iStrDealType, "''", "S") & ",  " & FilterVar(iStrPayMeth, "''", "S") & ", " & _
                                                  " " & FilterVar(iStrDnUnit, "''", "S") & ",  " & FilterVar(iStrCurrency, "''", "S") & ", " & _
                                                  " " & FilterVar(iStrValidDt, "''", "S") & ")"
    iStrFromList = ""
    iStrWhereList = ""

    Err.Clear
    
    If CommonQueryRs2by2(iStrSelectList, iStrFromList, iStrWhereList, iStrRs) Then
    
        iStrItemInfo = Split(iStrRs, Chr(11))

        frm1.vspdData.Col = C_Price
        frm1.vspdData.Text = UNIFormatNumber(iStrItemInfo(1), ggUnitCost.DecPoint, -2, 0, ggUnitCost.RndPolicy, ggUnitCost.RndUnit)
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description
            Err.Clear
        End If
    End If

End Function

'========================================
Function ItemByHScodeChange(CRow)

    Dim strVal
    Dim strCon
    Dim arrLot
     
    If Trim(frm1.txtPlant.Value) = "" Then
        Call DisplayMsgBox("189220", "X", "X", "X")
        Exit Function
    End If

    frm1.vspdData.Row = CRow
    frm1.vspdData.Col = C_ItemCd
    If Trim(frm1.vspdData.Text) = "" Then Exit Function

    strCon = "ITEM_CD =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "  AND PLANT_CD = " & FilterVar(frm1.txtPlant.Value, "''", "S") & " "

    Call CommonQueryRs("LOT_FLG", "B_ITEM_BY_PLANT", strCon, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    frm1.vspdData.Row = CRow
    If lgF0 = "" Then
        Call DisplayMsgBox("122700", "X", "X", "X")
		frm1.vspdData.Col = C_ItemCd
		frm1.vspdData.Text = ""
        Exit Function
    Else
        arrLot = Split(lgF0, Chr(11))
        frm1.vspdData.Col = C_LotMgmtFlag
        frm1.vspdData.Text = arrLot(0)
    End If

    Call ChangeLotRetField(CRow, CRow)
     
    If LayerShowHide(1) = False Then
        Exit Function
    End If

    strVal = ""
    strVal = BIZ_PGM_ID & "?txtMode=" & "ItemByHsCode"
    '품목 
    frm1.vspdData.Col = C_ItemCd
    strVal = strVal & "&ItemCd=" & Trim(frm1.vspdData.Text)
    '현재 ROW 위치 
    strVal = strVal & "&CRow=" & CRow

    Call RunMyBizASP(MyBizASP, strVal)
     
End Function

'========================================
Sub CheckCreditlimit()

    Err.Clear

    If LayerShowHide(1) = False Then Exit Sub

    Dim iStrVal
    
    iStrVal = BIZ_PGM_ID & "?txtMode=" & "ChkGiCreditLimit"
    iStrVal = iStrVal & "&txtHDnNo=" & Trim(frm1.txtHDNNo.Value)
    
    Call RunMyBizASP(MyBizASP, iStrVal)
 
End Sub

'========================================
Function BatchButton(ByVal iKubun)
    Dim iStrVal
    Dim iIntRetVal
    Err.Clear
     With frm1
        Select Case iKubun
            ' 출고처리 
            Case 2
                '변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 %>
                If lgBlnFlgChgValue = True Or lgBlnFlgChgValue1 = True Or lgBlnFlgChgValue2 = True Then
                    iIntRetVal = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
                    If iIntRetVal = vbNo Then Exit Function
                End If

                .txtBatch.Value = "Posting"
        
                If LayerShowHide(1) = False Then
                    Exit Function
                End If
                
            ' 출고처리 취소 
            Case 3
                .txtBatch.Value = "PostCancel"
                If LayerShowHide(1) = False Then
                    Exit Function
                End If
            Case Else
                Exit Function
        End Select

        If .chkArFlag.Checked And .txtArFlag.Value <> "Y" Then .txtArFlag.Value = "Y"
        If UCase(Trim(.txtArFlag.Value)) = "Y" And .chkVatFlag.Checked Then .txtVatFlag.Value = "Y"
        
        iStrVal = BIZ_PGM_ID & "?txtMode=" & "ARPOST" _
                            & "&txtHDnNo=" & Trim(.txtHDNNo.Value) _
                            & "&txtActualGIDt=" & Trim(.txtGI_Dt.Text) _
                            & "&txtARFlag=" & Trim(.txtArFlag.Value) _
                            & "&txtVatFlag=" & Trim(.txtVatFlag.Value) _
                            & "&txtGINo=" & Trim(.txtGINo.Value)
     End With

     Call RunMyBizASP(MyBizASP, iStrVal)
     
     Exit Function
     
End Function

'========================================
Sub BatchSetting()
    If frm1.vspdData.MaxRows < 1 Then
        frm1.btnPosting.disabled = True
        frm1.btnPostCancel.disabled = True
    Else
        If Len(Trim(frm1.txtGINo.Value)) Then
            frm1.btnPosting.disabled = True
            frm1.btnPostCancel.disabled = False
         Else
            frm1.btnPosting.disabled = False
            frm1.btnPostCancel.disabled = True
        End If
    End If
End Sub

'========================================
Sub ChangeReturnItemRetField()
    ggoSpread.Source = frm1.vspdData
    If frm1.txtRetItemFlag.Value = "Y" Then
        frm1.vspdData.Col = C_RetType: frm1.vspdData.ColHidden = False
        frm1.vspdData.Col = C_RetTypePopup: frm1.vspdData.ColHidden = False
        frm1.vspdData.Col = C_RetTypeNm: frm1.vspdData.ColHidden = False

        ggoSpread.SpreadUnLock C_RetType, -1, C_RetType
        ggoSpread.SpreadUnLock C_RetTypePopup, -1, C_RetTypePopup
        ggoSpread.SSSetRequired C_RetType, -1, -1
    Else
        frm1.vspdData.Col = C_RetType: frm1.vspdData.ColHidden = True
        frm1.vspdData.Col = C_RetTypePopup: frm1.vspdData.ColHidden = True
        frm1.vspdData.Col = C_RetTypeNm: frm1.vspdData.ColHidden = True

        ggoSpread.SSSetProtected C_RetType, -1, -1
          
        Dim ii
          
        For ii = 1 To frm1.vspdData.MaxRows
            frm1.vspdData.Col = C_RetType
            frm1.vspdData.Row = ii
            frm1.vspdData.Text = ""
        Next
    End If
End Sub

'========================================
Sub SetTransLT()
 Dim strValue, LT_Dt

 strValue = "BP_CD = " & FilterVar(frm1.txtShip_to_party.Value, "''", "S") & ""
  

 Call CommonQueryRs(" TRANS_LT ", " B_BIZ_PARTNER ", strValue, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

 If lgF0 = "" Then Exit Sub
 
 LT_Dt = Split(lgF0, Chr(11))

    frm1.txtHTransit_LT.Value = LT_Dt(0)

' Call SetPlannedGIDt()
End Sub

'========================================
Sub SetPlannedGIDt()
	If Len(frm1.txtHTransit_LT.Value) Then
		frm1.txtPlannedGIDt.Text = UNIDateAdd("d", -UNICDbl(frm1.txtHTransit_LT.Value), frm1.txtDlvyDt.Text, Parent.gDateFormat)
	Else
		frm1.txtPlannedGIDt.Text = frm1.txtDlvyDt.Text
	End If
End Sub

'========================================
Function BtnSpreadCheck()
    Dim IntRetCD
 
    BtnSpreadCheck = False

    If Trim(frm1.txtGI_Dt.Text) = "" Then
        MsgBox "실제출고일을 입력하세요", vbInformation, Parent.gLogoName

        If gPageNo <> TAB2 Then Call changeTabs(TAB2)
        frm1.txtGI_Dt.focus
        Exit Function
    End If
        
    '==================================================
    ' 2002.2.4 SMJ
    ' 실제출고일이 현재일보다 작게입력되도록 수정 
    '==================================================
    If UniConvDateToYYYYMMDD(frm1.txtGI_Dt.Text, Parent.gDateFormat, "") > UniConvDateToYYYYMMDD(EndDate, Parent.gDateFormat, "") Then
        IntRetCD = DisplayMsgBox("970024", "X", frm1.txtGI_Dt.ALT, "현재일")
        Call SetFocusToDocument("M")
        
        If gPageNo <> TAB2 Then Call changeTabs(TAB2)
        frm1.txtGI_Dt.focus
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData

    '변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
        If IntRetCD = vbNo Then Exit Function
    End If

    '변경이 없을때 작업진행여부 체크 
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")
        If IntRetCD = vbNo Then Exit Function
    End If

     BtnSpreadCheck = True

End Function

'========================================
Sub ChangeGridVatIncType()
    Dim i
 
    If frm1.vspdData.MaxRows < 1 Then Exit Sub
    
    If frm1.rdoVat_Calc_Type2.Checked = True Then
        If frm1.rdoVat_Inc_flag1.Checked = True Then
            For i = 1 To frm1.vspdData.MaxRows
                frm1.vspdData.Row = i

                frm1.vspdData.Col = C_VatIncFlag
                frm1.vspdData.Text = 1

                frm1.vspdData.Col = C_VatIncFlagNm
                frm1.vspdData.Text = "별도"

                Call vspdData_Change(C_VatIncFlag, i)
            Next
        ElseIf frm1.rdoVat_Inc_flag2.Checked = True Then
            For i = 1 To frm1.vspdData.MaxRows
                frm1.vspdData.Row = i

                frm1.vspdData.Col = C_VatIncFlag
                frm1.vspdData.Text = 2

                frm1.vspdData.Col = C_VatIncFlagNm
                frm1.vspdData.Text = "포함"

                Call vspdData_Change(C_VatIncFlag, i)
            Next
        End If
    End If
End Sub

'========================================
Sub ChangeGridVatRate()
    Dim i
 
    If frm1.vspdData.MaxRows < 1 Then Exit Sub
    
    If frm1.rdoVat_Calc_Type2.Checked = True Then
        For i = 1 To frm1.vspdData.MaxRows
            frm1.vspdData.Row = i
            frm1.vspdData.Col = C_VatRate
            frm1.vspdData.Text = frm1.txtVat_rate.Text
            Call vspdData_Change(C_VatType, i)
        Next
    End If
End Sub

'========================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To frm1.vspdData.MaxCols - 1
           frm1.vspdData.Col = iDx
           frm1.vspdData.Row = iRow

           If Not frm1.vspdData.ColHidden Then
              Call SetActiveCell(frm1.vspdData, iDx, iRow, "M", "X", "X")
              Exit For
           End If
           
       Next
          
    End If
End Sub

'========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
                        
            C_SlCd = iCurColumnPos(1)
            C_SlCdPopup = iCurColumnPos(2)
            C_ItemCd = iCurColumnPos(3)
            C_ItemPopup = iCurColumnPos(4)
            C_ItemNm = iCurColumnPos(5)
            C_Spec = iCurColumnPos(6)
            C_DnUnit = iCurColumnPos(7)
            C_DnUnitPopup = iCurColumnPos(8)
            C_DnQty = iCurColumnPos(9)
            C_Price = iCurColumnPos(10)
            C_TotalAmt = iCurColumnPos(11)
            C_NetAmt = iCurColumnPos(12)
            C_LotNo = iCurColumnPos(13)
            C_LotNoPopup = iCurColumnPos(14)
            C_LotSeq = iCurColumnPos(15)
            C_CartonNo = iCurColumnPos(16)
            C_VatIncFlag = iCurColumnPos(17)
            C_VatIncFlagNm = iCurColumnPos(18)
            C_VatType = iCurColumnPos(19)
            C_VatTypePopup = iCurColumnPos(20)
            C_VatTypeNm = iCurColumnPos(21)
            C_VatRate = iCurColumnPos(22)
            C_VatAmt = iCurColumnPos(23)
            C_RetType = iCurColumnPos(24)
            C_RetTypePopup = iCurColumnPos(25)
            C_RetTypeNm = iCurColumnPos(26)
            C_ReservedPrice = iCurColumnPos(27)
            C_Remark = iCurColumnPos(28)
            C_LotMgmtFlag = iCurColumnPos(29)
            C_DnSeq = iCurColumnPos(30)
            C_OldNetAmt = iCurColumnPos(31)
            C_OldVatAmt = iCurColumnPos(32)
            C_DBNetAmt = iCurColumnPos(33)
            C_DbVatAmt = iCurColumnPos(34)
    End Select
End Sub

'========================================
Sub Form_Load()
	Call LoadInfTB19029              '⊙: Load table , B_numeric_format
	Call FormatField()
	Call LockFieldInit("L")                                   '⊙: Lock  Suitable  Field
	
	Call InitSpreadSheet
	Call SetDefaultVal
	Call InitVariables                                                      
	Call InitVATTypeInfo()		' 부가세 유형정보를 초기화 한다.

    Call SetToolbar("11101101001011")          '⊙: 버튼 툴바 제어 
	Call CookiePage(0)
	Call ChangeTabs(TAB1)             'Because Textbox OCX Formatfield Display
 
 	frm1.txtConDn_no.focus

	gIsTab     = "Y" : gTabMaxCnt = 3

End Sub

'=======================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function btnShipToPlceRef_OnClick()
    Dim iCalledAspName
    Dim arrRet
    Dim ShipToPartyCd
    
    On Error Resume Next
    
    If Trim(frm1.txtShip_to_party.Value) = "" Then
        Call DisplayMsgBox("204256", "X", "X", "X")  '☜ 바뀐부분 
        Exit Function
    End If

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
    
    iCalledAspName = AskPRAspName("S4111RA1")
        
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "S4111RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    ShipToPartyCd = Trim(frm1.txtShip_to_party.Value)
    
    arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent, ShipToPartyCd), "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        If Err.Number <> 0 Then
            Err.Clear
        End If
        Exit Function
    Else
        Call SetShipToPlceRef(arrRet)
        Call SetToolbar("11101000000111")                   '⊙: 버튼 툴바 제어 
    End If

End Function

'========================================
Function btnTrnsMethRef_OnClick()
    Dim iCalledAspName
    Dim arrRet
    
    On Error Resume Next

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("S4111RA2")
        
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "S4111RA2", "x")
        IsOpenPop = False
        Exit Function
    End If

    arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent, ""), "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        If Err.Number <> 0 Then
            Err.Clear
        End If
        Exit Function
    Else
        Call SetTrnsMethRef(arrRet)
        Call SetToolbar("11101000000111")                   '⊙: 버튼 툴바 제어 
    End If

End Function

'=======================================================
Sub rdoVat_Calc_Type1_OnClick()
    Call ChangeVATIncFlagRetField
    lgBlnFlgChgValue = True
End Sub

'=======================================================
Sub rdoVat_Calc_Type2_OnClick()
    Call ChangeVATIncFlagRetField
    Call ChangeGridVatRate
    lgBlnFlgChgValue = True
End Sub

'=======================================================
Sub rdoVat_Inc_flag1_OnClick()
    lgBlnFlgChgValue = True
    Call ChangeGridVatIncType
End Sub

'=======================================================
Sub rdoVat_Inc_flag2_OnClick()
    lgBlnFlgChgValue = True
    Call ChangeGridVatIncType
End Sub

'=======================================================
Sub txtDlvyDt_Change()
    lgBlnFlgChgValue = True
    Call SetPlannedGIDt
End Sub

'=======================================================
Sub txtPlannedGIDt_Change()
     lgBlnFlgChgValue = True
End Sub

'2002. 12.20 SMJ
'실제납품일 
'=======================================================
Sub txtArriv_dt_Change()
    lgBlnFlgChgValue = True
End Sub

' 부가세 유형 
'=======================================================
Sub txtVat_type_OnChange()
    lgBlnFlgChgValue = True
    Call SetVatTypeForHdr
    Call ChangeGridVatRate
End Sub

'2002.12.18 SMJ
'납품처 상세정보, 운송정보 변경여부 
'=======================================================
Sub txtZip_cd_Onchange()
    lgBlnFlgChgValue1 = True
    If frm1.txtSTP_Inf_No.Value <> "" Then
         frm1.txtSTP_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtReceiver_Onchange()
    lgBlnFlgChgValue1 = True
    If frm1.txtSTP_Inf_No.Value <> "" Then
         frm1.txtSTP_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtADDR1_Dlv_Onchange()
    lgBlnFlgChgValue1 = True
    If frm1.txtSTP_Inf_No.Value <> "" Then
         frm1.txtSTP_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtADDR2_Dlv_Onchange()
    lgBlnFlgChgValue1 = True
    If frm1.txtSTP_Inf_No.Value <> "" Then
         frm1.txtSTP_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtADDR3_Dlv_Onchange()
    lgBlnFlgChgValue1 = True
    If frm1.txtSTP_Inf_No.Value <> "" Then
         frm1.txtSTP_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtTel_No1_Onchange()
    lgBlnFlgChgValue1 = True
    If frm1.txtSTP_Inf_No.Value <> "" Then
         frm1.txtSTP_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtTel_No2_Onchange()
    lgBlnFlgChgValue1 = True
    If frm1.txtSTP_Inf_No.Value <> "" Then
         frm1.txtSTP_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtTransCo_Onchange()
    lgBlnFlgChgValue2 = True
    If frm1.txtTrnsp_Inf_No.Value <> "" Then
         frm1.txtTrnsp_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtSender_Onchange()
    lgBlnFlgChgValue2 = True
    If frm1.txtTrnsp_Inf_No.Value <> "" Then
         frm1.txtTrnsp_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtVehicleNo_Onchange()
    lgBlnFlgChgValue2 = True
    If frm1.txtTrnsp_Inf_No.Value <> "" Then
         frm1.txtTrnsp_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtDriver_Onchange()
    lgBlnFlgChgValue2 = True
    If frm1.txtTrnsp_Inf_No.Value <> "" Then
         frm1.txtTrnsp_Inf_No.Value = ""
    End If
End Sub

'=======================================================
Sub txtGI_Dt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGI_Dt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtGI_Dt.focus
    End If
End Sub

'=======================================================
Sub txtDlvyDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDlvyDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDlvyDt.focus
    End If
End Sub

'=======================================================
Sub txtPlannedGIDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlannedGIDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtPlannedGIDt.focus
    End If
End Sub

'2002.12.20 SMJ
'=======================================================
Sub txtArriv_dt_DblClick(Button)
    If Button = 1 Then
        frm1.txtArriv_dt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtArriv_dt.focus
    End If
End Sub

'=======================================================
Sub txtSold_to_party_OnChange()
    If Trim(frm1.txtSold_to_party.Value) <> "" Then Call SoldToPartyLookUp
End Sub

'=======================================================
Sub txtDn_Type_OnChange()
 If Trim(frm1.txtDn_Type.Value) <> "" Then Call DNTypeLookUp
End Sub

'=======================================================
Sub txtSales_Grp_OnChange()
 If Trim(frm1.txtSales_Grp.Value) = "" Then
  frm1.txtSales_GrpNm.Value = ""
 Else
  '영업그룹과 관련된 세금신고사업장을 Fetch한다.
  Call GetTaxBizArea("BA")
 End If
End Sub

'=======================================================
Function txtTaxBizAreaCd_OnChange()
    If Trim(frm1.txtTaxBizAreaCd.Value) = "" Then
         frm1.txtTaxBizAreaNm.Value = ""
    Else
        If Not GetTaxBizArea("NM") Then
            frm1.txtTaxBizAreaCd.focus
            txtTaxBizAreaCd_OnChange = False
        End If
    End If
End Function

'=======================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

    If Row <= 0 Then Exit Sub

    ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
        .Row = Row
        .Col = Col
    
        Select Case Col
            Case C_ItemPopup
                Call OpenItem(.Text)
                
            Case C_DnUnitPopup
                Call OpenDnDtl(.Text, 2)
                
            Case C_LotNoPopup
                Call OpenLotNoPopup(Row)

            Case C_VatTypePopup
                Call OpenDnDtl(.Text, 4)
                
            Case C_SlCdPopup
                Call OpenDnDtl(.Text, 5)
                
            Case C_RetTypePopup
                Call OpenDnDtl(.Text, 6)
        End Select
    End With
    
    Call SetActiveCell(frm1.vspdData, Col - 1, Row, "M", "X", "X")

End Sub

'=======================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    If Trim(frm1.txtGINo.Value) = "" Then
        Call SetPopupMenuItemInf("1101111111")
    Else
        Call SetPopupMenuItemInf("0000111111")
    End If

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
       
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col                'Sort in Ascending
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey     'Sort in Descending
            lgSortKey = 1
        End If
        
        Exit Sub
    End If

End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=======================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex
 
    With frm1.vspdData
        .Row = Row
        .Col = Col
        intIndex = .Value
     
        .Col = C_VatIncFlag
        .Value = intIndex + 1
    End With
    
    Call vspdData_Change(C_VatIncFlag, Row)
    
    ggoSpread.UpdateRow Row
End Sub

'=======================================================
Sub vspdData_MouseDown(Button, Shift, x, y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'=======================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
    ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
        .Row = Row:     .Col = 0
        If .Text = ggoSpread.DeleteFlag Then Exit Sub
        
        ggoSpread.UpdateRow Row
        lgBlnFlgDtlChgValue = True

        Select Case Col
            Case C_DnQty
                Call CalcAmtLoc(Row)

            Case C_Price
                Call CalcAmtLoc(Row)

            Case C_ItemCd
                .Col = C_SlCd
                If Len(.Text) Then
                    Call ItemByHScodeChange(Row)
                Else
                     Call DisplayMsgBox("17A002", "X", "창고", "X")
                     .Col = C_ItemCd:       .Text = ""
                End If

            Case C_DnUnit
                Call GetItemPrice(Row)
                Call CalcAmtLoc(Row)
         
            Case C_VatType
                Call SetVATTypeForSpread(Row)
                Call CalcVatAmt(Row)

            Case C_TotalAmt
                Call CalcNetAmt(Row)
                Call CalcTotal("U", Row)

            Case C_VatIncFlag
                Call CalcNetAmt(Row)
                Call CalcTotal("U", Row)

            ' 2003.02.07 SMJ
            ' lotno onchange추가 
            Case C_LotNo
                Call LotNoChange(Row)
                
        End Select
    End With

End Sub

'========================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    If OldLeft <> NewLeft Then Exit Sub

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) And lgStrPrevKey <> "" Then
        If CheckRunningBizProcess Then Exit Sub
      
        Call DisableToolBar(Parent.TBC_QUERY)
        Call DbQuery
    End If
End Sub

'=======================================================
Sub btnPosting_OnClick()
    Dim IntRetCD
 
    If BtnSpreadCheck = False Then Exit Sub
    ' 여신 Check후 출고처리한다.
    Call CheckCreditlimit
End Sub

'=======================================================
Sub btnPostCancel_OnClick()

    If BtnSpreadCheck = False Then Exit Sub
    Call BatchButton(3)

End Sub

' 후속작업여부 - 매출채권 
'=======================================================
Sub chkArFlag_OnClick()

    On Error Resume Next

    Select Case frm1.chkArFlag.Checked
        Case True
            lblArFlag.disabled = False
        Case False
            lblArFlag.disabled = True
            lblVatFlag.disabled = True
            frm1.chkVatFlag.Checked = False
     End Select

    ' 수금유형이 등록된 경우 AR Flag가 변경된 경우 저장여부를 check한다.
    If frm1.txtArFlag.Value = "Y" Then
        lgBlnFlgChgValue = True
    End If
End Sub

' 후속작업여부 - 세금계산서 
'=======================================================
Sub chkVatFlag_OnClick()

    On Error Resume Next

    Select Case frm1.chkVatFlag.Checked
        Case True
            lblVatFlag.disabled = False
            lblArFlag.disabled = False
            frm1.chkArFlag.Checked = True
        Case False
            lblVatFlag.disabled = True
    End Select

'   lgBlnFlgChgValue = True
End Sub

' 후속작업여부 - 매출채권 
'=======================================================
Sub chkArFlag_OnPropertyChange()
    With frm1
        If .chkArFlag.Checked = True Then
            If UNICDbl(frm1.txtCol_amt.Text) > 0 Then
                Call ggoOper.SetReqAttr(.txtCol_Type, "N")
            Else
                Call ggoOper.SetReqAttr(.txtCol_Type, "D")
            End If

            If Len(.txtCol_Type.Value) Then
                Call ggoOper.SetReqAttr(.txtCol_amt, "N")
            Else
                Call ggoOper.SetReqAttr(.txtCol_amt, "D")
            End If
        Else
            Call ggoOper.SetReqAttr(.txtCol_Type, "Q")
            Call ggoOper.SetReqAttr(.txtCol_amt, "Q")
            .txtCol_Type.Value = ""
            .txtCol_Type_nm.Value = ""
            .txtCol_amt.Text = "0"
        End If
    End With
End Sub

'=======================================================
Sub txtCol_Type_OnChange()
    If Len(frm1.txtCol_Type.Value) Then
        If frm1.chkArFlag.Checked = True Then
            If Len(Trim(frm1.txtGINo.Value)) Then
                Call ggoOper.SetReqAttr(frm1.txtCol_amt, "Q")
            Else
                Call ggoOper.SetReqAttr(frm1.txtCol_amt, "N")
            End If
        Else
            Call ggoOper.SetReqAttr(frm1.txtCol_amt, "Q")
        End If
    Else
        If frm1.chkArFlag.Checked = True Then
            Call ggoOper.SetReqAttr(frm1.txtCol_amt, "D")
        Else
            Call ggoOper.SetReqAttr(frm1.txtCol_amt, "Q")
        End If
    End If
    lgBlnFlgChgValue = True
End Sub

'=======================================================
Sub txtCol_amt_Change()
    If UNICDbl(frm1.txtCol_amt.Text) > 0 Then
        If frm1.chkArFlag.Checked = True Then
            If Len(Trim(frm1.txtGINo.Value)) Then
                Call ggoOper.SetReqAttr(frm1.txtCol_Type, "Q")
            Else
                Call ggoOper.SetReqAttr(frm1.txtCol_Type, "N")
            End If
        Else
            Call ggoOper.SetReqAttr(frm1.txtCol_Type, "Q")
        End If
    Else
        If frm1.chkArFlag.Checked = True Then
            Call ggoOper.SetReqAttr(frm1.txtCol_Type, "D")
        Else
            Call ggoOper.SetReqAttr(frm1.txtCol_Type, "Q")
        End If
    End If
    lgBlnFlgChgValue = True
End Sub

'=======================================================
Sub txtPlant_OnChange()
 Dim i
 
 If frm1.vspdData.MaxRows > 0 And lgIntFlgMode = Parent.OPMD_UMODE Then
  For i = 1 To frm1.vspdData.MaxRows
   frm1.vspdData.Col = 0
   frm1.vspdData.Row = i
  
   If frm1.vspdData.Text <> ggoSpread.InsertFlag Then
       ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow i
   End If
  Next
 End If
End Sub

'=======================================================
Function FncQuery()
    Dim IntRetCD
    
    FncQuery = False
    
    Err.Clear

    If lgBlnFlgChgValue = True Or lgBlnFlgChgValue1 = True Or lgBlnFlgChgValue2 = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    
    If Not chkFieldByCell(frm1.txtConDn_no, "A", TAB1) Then Exit Function

    Call ggoOper.ClearField(document, "2")
    Call InitVariables
    
    Call LockFieldInit("N")
    Call UnLockColor_CfmNo

    Call DbQuery
    
    FncQuery = True
        
End Function

'=======================================================
Function FncNew()
    Dim IntRetCD
    
    FncNew = False
    
    If lgBlnFlgChgValue = True Or lgBlnFlgChgValue1 = True Or lgBlnFlgChgValue2 = True Or ggoSpread.SSCheckChange Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    
    Call ggoOper.ClearField(document, "A")
    Call LockFieldInit("N")
    Call UnLockColor_CfmNo
    Call SetDefaultVal

    Call InitVariables
    
    Call SetToolbar("11101111001011")          '⊙: 버튼 툴바 제어 
	
	Call ChangeTabs(TAB1)
	frm1.txtDnNo.focus
	
	Set gActiveElement = document.ActiveElement
	
    FncNew = True

End Function

'=====================================================
Function FncDelete()
    
    Dim IntRetCD
    
    FncDelete = False
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
    Call DbDelete
    
    FncDelete = True
    
End Function

'=====================================================
Function FncSave()
    On Error Resume Next
     
    Dim IntRetCD
    
    FncSave = False
    
    Err.Clear

    If lgBlnFlgChgValue = False And lgBlnFlgDtlChgValue = False And lgBlnFlgChgValue1 = False And lgBlnFlgChgValue2 = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If

    With frm1
        ' 입력필수 항목 Check
        If Not chkFieldByCell(.txtDn_Type, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtSold_to_party, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtDeal_Type, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtShip_to_party, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtSales_Grp, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtTaxBizAreaCd, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtPay_terms, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtDlvyDt, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtPlannedGIDt, "A", "1") Then Exit Function
        If Not chkFieldByCell(.txtPlant, "A", "2") Then Exit Function
        
	' 필드길이 Check
	If Not ChkFieldLengthByCell(.txtRemark, "A", 1) Then Exit Function
	If Not ChkFieldLengthByCell(.txt_Payterms_txt, "A", 1) Then Exit Function
	If Not ChkFieldLengthByCell(.txtArriv_Tm, "A", 1) Then Exit Function
	If Not ChkFieldLengthByCell(.txtReceiver, "A", 3) Then Exit Function
	If Not ChkFieldLengthByCell(.txtADDR1_Dlv, "A", 3) Then Exit Function
	If Not ChkFieldLengthByCell(.txtADDR2_Dlv, "A", 3) Then Exit Function
	If Not ChkFieldLengthByCell(.txtADDR3_Dlv, "A", 3) Then Exit Function
	If Not ChkFieldLengthByCell(.txtDlvyPlace, "A", 3) Then Exit Function
	If Not ChkFieldLengthByCell(.txtTel_No1, "A", 3) Then Exit Function
	If Not ChkFieldLengthByCell(.txtTel_No2, "A", 3) Then Exit Function
	If Not ChkFieldLengthByCell(.txtSender, "A", 3) Then Exit Function
	If Not ChkFieldLengthByCell(.txtDriver, "A", 3) Then Exit Function

        '2002.08.29 SMJ
        '수금액을 입력하고 수금유형이 없으면 에러 
        If .txtCol_Type.Value <> "" And UNICDbl(.txtCol_amt.Text) = 0 Then
            Call DisplayMsgBox("970029", "X", .txtCol_amt.ALT, "X")
            
            If gPageNo <> TAB2 Then
                Call ChangeTabs(TAB2)
            End If

            .txtCol_amt.focus()
		    Set gActiveElement = document.ActiveElement
            Exit Function
        End If
        
        If .txtCol_Type.Value = "" And UNICDbl(.txtCol_amt.Text) > 0 Then
            Call DisplayMsgBox("970029", "X", .txtCol_Type.ALT, "X")
            If gPageNo <> TAB2 Then
                Call ChangeTabs(TAB2)
            End If
            
            .txtCol_Type.focus()
		    Set gActiveElement = document.ActiveElement
            Exit Function
        End If
        
        ggoSpread.Source = .vspdData

        If Not ggoSpread.SSDefaultCheck Then Exit Function

        If lgBlnFlgChgValue = False And lgBlnFlgChgValue1 = False And lgBlnFlgChgValue2 = False Then
            .txtHdrStateFlg.Value = "1"
        Else
            .txtHdrStateFlg.Value = "2"
        End If
    End With
    
    Call DbSave                                                    '☜: Save db data

    FncSave = True

End Function

'=====================================================
Function FncCopy()
    With frm1.vspdData
        If .MaxRows < 1 Then Exit Function
        
	    If gPageNo <> TAB2 Then Call ChangeTabs(TAB2)
        
        .ReDraw = False
         
        ggoSpread.Source = frm1.vspdData
        ggoSpread.CopyRow
        SetSpreadColor .ActiveRow, .ActiveRow
        
        Call SetRowDefaultVal(.ActiveRow, "N")
        
        Call SubSetErrPos(.ActiveRow)

        .ReDraw = True

        lgBlnFlgDtlChgValue = True

    End With
    
    Set gActiveElement = document.ActiveElement
End Function

'========================================
Function FncCancel()
    If frm1.vspdData.MaxRows < 1 Then Exit Function

	If gPageNo <> TAB2 Then Call ChangeTabs(TAB2)
    
    Call CalcTotal("C", frm1.vspdData.ActiveRow)

    ggoSpread.Source = frm1.vspdData

    ggoSpread.EditUndo
End Function

'=======================================================
Function FncInsertRow(pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next
    Err.Clear
    
    FncInsertRow = False
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If
    
    If gPageNo <> TAB2 Then Call ChangeTabs(TAB2)
    
    With frm1.vspdData
        .focus
        ggoSpread.Source = frm1.vspdData

        .ReDraw = False

        ggoSpread.InsertRow .ActiveRow, imRow

        SetSpreadColor .ActiveRow, .ActiveRow + imRow - 1
        
        For i = .ActiveRow To .ActiveRow + imRow - 1
            Call SetRowDefaultVal(i, "Y")
        Next
        
        Call SubSetErrPos(.ActiveRow)
        
        .ReDraw = True

        lgBlnFlgDtlChgValue = True

    End With
    
    If Err.Number = 0 Then
       FncInsertRow = True
    End If
    
    Set gActiveElement = document.ActiveElement
    
End Function

'========================================
Function FncDeleteRow()
    If frm1.vspdData.MaxRows < 1 Then Exit Function

	If gPageNo <> TAB2 Then Call ChangeTabs(TAB2)

    frm1.vspdData.focus
    Set gActiveElement = document.ActiveElement
        
    ggoSpread.Source = frm1.vspdData
                
    Call CalcTotal("D", 0)
        
    ggoSpread.DeleteRow

    lgBlnFlgDtlChgValue = True
End Function

'========================================
Function FncPrint()
    Call Parent.FncPrint
End Function

'========================================
Function FncExcel()
 Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================
Function FncFind()
 Call Parent.FncFind(Parent.C_SINGLE, True)
End Function

'=====================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit (gActiveSpdSheet.ActiveCol)
    
End Sub

'========================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf
End Sub

'========================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf
    Call InitSpreadSheet
    
    Call ggoSpread.ReOrderingSpreadData
    If Trim(frm1.txtGINo.Value) = "" Then
        SetSpreadColor -1, -1
    Else
        SetSpreadColorConfirmed -1
    End If

End Sub

'=====================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	If lgBlnFlgChgValue = True Or lgBlnFlgChgValue1 = True Or lgBlnFlgChgValue2 = True OR ggoSpread.SSCheckChange  Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If

	FncExit = True
End Function

'=====================================================
Function DbDelete()
    Err.Clear
    
    DbDelete = False

    If LayerShowHide(1) = False Then
        Exit Function
    End If

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
    strVal = strVal & "&txtDnNo=" & Trim(frm1.txtDnNo.Value)    '☜: 삭제 조건 데이타 
    
    Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True

End Function

'========================================
Function DbDeleteOk()
 lgBlnFlgChgValue = False
 Call MainNew
End Function

'========================================
Function DbQuery()
    
    Err.Clear
    
    DbQuery = False

    If LayerShowHide(1) = False Then
        Exit Function
    End If
     
    Dim iStrVal
    
    iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        iStrVal = iStrVal & "&txtConDN_no=" & Trim(frm1.txtHDNNo.Value)
    Else
        iStrVal = iStrVal & "&txtConDN_no=" & Trim(frm1.txtConDn_no.Value)
    End If
    iStrVal = iStrVal & "&lgStrPrevKey=" & lgStrPrevKey
    iStrVal = iStrVal & "&txtLastRow=" & frm1.vspdData.MaxRows
 
    Call RunMyBizASP(MyBizASP, iStrVal)
 
    DbQuery = True

End Function

'========================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE

	Call LockFieldQuery

    If Len(Trim(frm1.txtGINo.Value)) Then
        Call SetToolbar("11101000000111")
        Call LockColor_CfmYes
    Else
        Call SetToolbar("11111111001111")
        If gPageNo = TAB2 Then
            frm1.vspdData.focus
        Else
            frm1.txtConDn_no.focus
        End If
        Call chkArFlag_OnPropertyChange
    End If

    Call ChangeVATIncFlagRetField
    Call BatchSetting
 
    '2002-09-16 SMJ
    If frm1.vspdData.MaxRows > 0 And Trim(frm1.txtGINo.Value) = "" Then
        Call ggoOper.SetReqAttr(frm1.txtGI_Dt, "D")
    Else
        Call ggoOper.SetReqAttr(frm1.txtGI_Dt, "Q")
    End If
     
    lgBlnFlgChgValue = False
    lgBlnFlgChgValue1 = False
    lgBlnFlgChgValue2 = False
    
End Function

'========================================
Function DbHdrQueryOk()
 
    lgIntFlgMode = Parent.OPMD_UMODE

	Call LockFieldQuery
    Call SetToolbar("11111111001111")
    
    Call ChangeVATIncFlagRetField
    Call chkArFlag_OnPropertyChange
    Call BatchSetting
    
	frm1.btnPosting.disabled = True
	frm1.btnPostCancel.disabled = True
    frm1.txtConDn_no.focus
    
    lgBlnFlgChgValue = False
End Function

'========================================
Function DbSave()
    On Error Resume Next
    
    Err.Clear
 
    DbSave = False

    If LayerShowHide(1) = False Then Exit Function

    Dim iLngRow
    Dim iStrIns, iStrUpd, iStrDel
    Dim iArrData
    
    Dim iLngCTotalvalLen        '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장 
    Dim iLngUTotalvalLen
    Dim iLngDTotalValLen

    Dim iTmpCBuffer             '현재의 버퍼 [수정,신규]
    Dim iTmpCBufferCount        '현재의 버퍼 Position
    Dim iTmpCBufferMaxCount     '현재의 버퍼 Chunk Size

    Dim iTmpUBuffer             '현재의 버퍼 [수정,신규]
    Dim iTmpUBufferCount        '현재의 버퍼 Position
    Dim iTmpUBufferMaxCount     '현재의 버퍼 Chunk Size

    Dim iTmpDBuffer             '현재의 버퍼 [삭제]
    Dim iTmpDBufferCount        '현재의 버퍼 Position
    Dim iTmpDBufferMaxCount     '현재의 버퍼 Chunk Size
 
    iTmpCBufferMaxCount = Parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
    iTmpUBufferMaxCount = Parent.C_CHUNK_ARRAY_COUNT
    iTmpDBufferMaxCount = Parent.C_CHUNK_ARRAY_COUNT

    iTmpCBufferCount = -1
    iTmpUBufferCount = -1
    iTmpDBufferCount = -1

    iLngCTotalvalLen = 0
    iLngUTotalvalLen = 0
    iLngDTotalValLen = 0

    ReDim iArrData(40)
    ReDim iTmpCBuffer(iTmpCBufferMaxCount)  '최기 버퍼의 설정[신규]
    ReDim iTmpUBuffer(iTmpUBufferMaxCount)  '최기 버퍼의 설정[수정]
    ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[삭제]

    ' 변경되지 않는 Default 값.
    iArrData(5) = "0"       ' 출고요청덤수량 
    iArrData(7) = "0"       ' Picking덤수량 
    iArrData(10) = "0"      ' 과부족허용량(+)
    iArrData(11) = "0"      ' 과부족허용량(-)
    iArrData(14) = "0"      ' 통관량 
    iArrData(15) = ""       ' 수주번호 
    iArrData(16) = "0"      ' 수주순번 
    iArrData(17) = "0"      ' 납품순번 
    iArrData(18) = ""       ' L/C번호 
    iArrData(19) = "0"      ' L/C순번 
    iArrData(21) = ""       ' 검사구분 
    iArrData(22) = "0"      ' ext1_qty
    iArrData(23) = "0"      ' ext1_amt
    iArrData(24) = ""       ' ext1_cd
    iArrData(25) = "0"      ' ext2_qty
    iArrData(26) = "0"      ' ext2_amt
    iArrData(27) = ""       ' ext2_cd
    iArrData(28) = "0"      ' ext3_qty
    iArrData(29) = "0"      ' ext3_amt
    iArrData(30) = ""       ' ext3_cd

    With frm1.vspdData
        For iLngRow = 1 To .MaxRows
            .Row = iLngRow
            .Col = 0

            '삭제인 경우 
            If .Text = ggoSpread.DeleteFlag Then
                ' Row 번호, 출하순번 
                .Col = C_DnSeq
                iStrDel = CStr(iLngRow) & Parent.gColSep & Trim(.Text) & Parent.gRowSep
                
                If iLngDTotalValLen + Len(iStrDel) > Parent.C_FORM_LIMIT_BYTE Then    '한개의 form element에 넣을 한개치가 넘으면 
                    Call MakeTextArea("txtDSpread", iTmpDBuffer)
                                      
                   iTmpDBufferMaxCount = Parent.C_CHUNK_ARRAY_COUNT
                   ReDim iTmpDBuffer(iTmpDBufferMaxCount)
                   iTmpDBufferCount = -1
                   iLngDTotalValLen = 0
                End If
                                   
                iTmpDBufferCount = iTmpDBufferCount + 1

                If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
                   iTmpDBufferMaxCount = iTmpDBufferMaxCount + Parent.C_CHUNK_ARRAY_COUNT
                   ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
                End If

                iTmpDBuffer(iTmpDBufferCount) = iStrDel
                iLngDTotalValLen = iLngDTotalValLen + Len(iStrDel)

            ' 입력, 수정인 경우 
            ElseIf .Text <> "" Then
                iArrData(0) = iLngRow               ' Row번호 
                .Col = C_DnSeq:             iArrData(1) = Trim(.Text)                   ' 출하순번 
                .Col = C_ItemCd:            iArrData(2) = Trim(.Text)
                .Col = C_DnQty:             iArrData(3) = UNIConvNum(Trim(.Text), 0)    ' 출고요청수량 
                .Col = C_DnUnit:            iArrData(4) = Trim(.Text)                   ' 단위 
                .Col = C_DnQty:             iArrData(6) = UNIConvNum(Trim(.Text), 0)    ' Picking수량 
                iArrData(8) = UCase(Trim(frm1.txtPlant.Value))                          ' 공장 
                .Col = C_SlCd:              iArrData(9) = Trim(.Text)                   ' 창고 
                
                ' LOT No                
                .Col = C_LotNo
                If Trim(.Text) = "" Then
					iArrData(12) = "*"
				Else
					iArrData(12) = Trim(.Text)
				End If
				
				' LOT No 순번 
                .Col = C_LotSeq:            iArrData(13) = UNIConvNum(Trim(.Text), 0)
                .Col = C_Remark:            iArrData(20) = Trim(.Text)                  ' 비고 
                .Col = C_CartonNo:          iArrData(31) = Trim(.Text)                  ' Carton 번호 
                .Col = C_Price:             iArrData(32) = UNIConvNum(Trim(.Text), 0)   ' 단가 
                .Col = C_NetAmt:            iArrData(33) = UNIConvNum(Trim(.Text), 0)   ' 출고금액 
                .Col = C_NetAmt:            iArrData(34) = UNIConvNum(Trim(.Text), 0)   ' 출고자국금액 
                .Col = C_VatIncFlag:        iArrData(35) = Trim(.Text)                  ' 부가세 포함 여부 
                .Col = C_VatType:           iArrData(36) = Trim(.Text)                  ' 부가세 유형 
                .Col = C_VatRate:           iArrData(37) = UNIConvNum(Trim(.Text), 0)   ' 부가세 율 
                .Col = C_VatAmt:            iArrData(38) = UNIConvNum(Trim(.Text), 0)   ' 부가세 액 
                .Col = C_VatAmt:            iArrData(39) = UNIConvNum(Trim(.Text), 0)   ' 부가세 자국액 
                .Col = C_RetType:           iArrData(40) = Trim(.Text)                  ' 반품유형 

                .Col = 0
                ' 입력 
                If .Text = ggoSpread.InsertFlag Then
                    iStrIns = Join(iArrData, Parent.gColSep) & Parent.gRowSep

                    If iLngCTotalvalLen + Len(iStrIns) > Parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 Data 한개치가 넘으면 
                        Call MakeTextArea("txtCSpread", iTmpCBuffer)
                        
                       iTmpCBufferMaxCount = Parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
                       ReDim iTmpCBuffer(iTmpCBufferMaxCount)
                       iTmpCBufferCount = -1
                       iLngCTotalvalLen = 0
                    End If
                       
                    iTmpCBufferCount = iTmpCBufferCount + 1
                      
                    If iTmpCBufferCount > iTmpCBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
                       iTmpCBufferMaxCount = iTmpCBufferMaxCount + Parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
                       ReDim Preserve iTmpCBuffer(iTmpCBufferMaxCount)
                    End If
                    iTmpCBuffer(iTmpCBufferCount) = iStrIns
                    iLngCTotalvalLen = iLngCTotalvalLen + Len(iStrIns)

                ' 수정 
                ElseIf .Text = ggoSpread.UpdateFlag Then
                    iStrUpd = Join(iArrData, Parent.gColSep) & Parent.gRowSep

                    If iLngUTotalvalLen + Len(iStrUpd) > Parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 Data 한개치가 넘으면 
                        Call MakeTextArea("txtUSpread", iTmpUBuffer)
                        
                       iTmpUBufferMaxCount = Parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
                       ReDim iTmpUBuffer(iTmpUBufferMaxCount)
                       iTmpUBufferCount = -1
                       iLngUTotalvalLen = 0
                    End If
                       
                    iTmpUBufferCount = iTmpUBufferCount + 1
                      
                    If iTmpUBufferCount > iTmpUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
                       iTmpUBufferMaxCount = iTmpUBufferMaxCount + Parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
                       ReDim Preserve iTmpUBuffer(iTmpUBufferMaxCount)
                    End If
                    iTmpUBuffer(iTmpUBufferCount) = iStrUpd
                    iLngUTotalvalLen = iLngUTotalvalLen + Len(iStrUpd)
                End If
            End If
        Next
    End With

   ' 나머지 데이터 처리 
    If iTmpCBufferCount > -1 Then Call MakeTextArea("txtCSpread", iTmpCBuffer)
    If iTmpUBufferCount > -1 Then Call MakeTextArea("txtUSpread", iTmpUBuffer)
    If iTmpDBufferCount > -1 Then Call MakeTextArea("txtDSpread", iTmpDBuffer)

    With frm1
        .txtMode.Value = Parent.UID_M0002           '☜: 비지니스 처리 ASP 의 상태 
        .txtFlgMode.Value = lgIntFlgMode
        .txtlgBlnChgValue1.Value = lgBlnFlgChgValue1
        .txtlgBlnChgValue2.Value = lgBlnFlgChgValue2
    End With

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)
 
    DbSave = True
    
End Function

'=====================================================
Function DbSaveOk()
    Call InitVariables
    Call RemovedivTextArea
    Call MainQuery
End Function

'========================================
Sub MakeTextArea(ByVal pvStrName, ByRef prArrData)
    Dim iObjTEXTAREA        '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

    Set iObjTEXTAREA = document.createElement("TEXTAREA")            '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
    iObjTEXTAREA.Name = pvStrName
    iObjTEXTAREA.Value = Join(prArrData, "")
    divTextArea.appendChild (iObjTEXTAREA)
End Sub

'========================================
Function RemovedivTextArea()
    Dim iIntIndex
    
    For iIntIndex = 1 To divTextArea.children.length
        divTextArea.removeChild (divTextArea.children(0))
    Next
End Function

' 자국금액 / Vat금액 / VAT 자국금액 계산 
Sub CalcAmtLoc(ByVal pvIntRow)
    Dim iDblTotalAmt, iDblDnQty, iDblPrice
    Dim iStrVatIncFlag

    With frm1.vspdData
        .Row = pvIntRow
        
        ggoSpread.Source = frm1.vspdData
        .Col = 0
        If .Text = ggoSpread.DeleteFlag Then Exit Sub

        .Col = C_DnQty: iDblDnQty = UNICDbl(.Text)
        .Col = C_Price: iDblPrice = UNICDbl(.Text)
        
        iDblTotalAmt = iDblDnQty * iDblPrice
        .Col = C_TotalAmt
        If iDblTotalAmt <> 0 Then
            ' 자국금액계산(부가세 포함금액)
            .Text = UNIConvNumPCToCompanyByCurrency(iDblTotalAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
            
            ' Net 금액 계산 
            Call CalcNetAmt(pvIntRow)
        Else
            .Col = C_TotalAmt: .Text = "0"
            .Col = C_NetAmt: .Text = "0"
            .Col = C_VatAmt: .Text = "0"
        End If
    End With
    
    ' 총금액 계산 
    Call CalcTotal("U", pvIntRow)
End Sub

' Net 금액 계산 
Sub CalcNetAmt(ByVal pvIntRow)
    Dim iDblVatRate, iDblTotalAmt, iDblVatAmt
    Dim iStrVatIncFlag, iStrNetAmt
    
    With frm1.vspdData
        .Row = pvIntRow:        .Col = C_TotalAmt
        iStrNetAmt = .Text:     iDblTotalAmt = UNICDbl(.Text)
                            
        ' VAT 금액 계산 
        .Col = C_VatIncFlag: iStrVatIncFlag = .Text
                
        .Col = C_VatRate: iDblVatRate = UNICDbl(.Text)
        .Col = C_VatAmt: .Text = FncCalcVatAmt(iDblTotalAmt, iStrVatIncFlag, iDblVatRate, Parent.gCurrency)
        iDblVatAmt = UNICDbl(.Text)
                
        ' Net 금액 계산 
        If iStrVatIncFlag = "1" Then
            .Col = C_NetAmt: .Text = iStrNetAmt
        Else
            .Col = C_NetAmt: .Text = UNIConvNumPCToCompanyByCurrency(iDblTotalAmt - iDblVatAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
        End If
    End With
End Sub

' VAT 금액 계산 
Sub CalcVatAmt(ByVal pvIntRow)
    Dim iDblTotalAmt, iDblVatAmt, iDblNetAmt, iDblVatRate, iDblDnQty, iDblPrice
    Dim iStrVatIncFlag, iStrNetAmt

    With frm1.vspdData
        .Row = pvIntRow
        .Col = C_TotalAmt: iDblTotalAmt = UNICDbl(.Text)
        iStrNetAmt = .Text
                            
        .Col = C_VatIncFlag: iStrVatIncFlag = .Text
        .Col = C_VatRate: iDblVatRate = UNICDbl(.Text)
                            
        .Col = C_VatAmt: .Text = FncCalcVatAmt(iDblTotalAmt, iStrVatIncFlag, iDblVatRate, Parent.gCurrency)
        iDblVatAmt = UNICDbl(.Text)
            
        ' Net 금액 계산 
        If iStrVatIncFlag = "1" Then
            .Col = C_NetAmt: .Text = iStrNetAmt
        Else
            .Col = C_NetAmt: .Text = UNIConvNumPCToCompanyByCurrency(iDblTotalAmt - iDblVatAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
        End If
    End With

    ' 총금액 계산 
    Call CalcTotal("U", pvIntRow)
End Sub

' 총합계금액을 재계산한다.
Sub CalcTotal(ByVal pvStrFlag, ByVal pvIntRow)
    On Error Resume Next
    
    Dim iLngRow, iLngFirstRow, iLngLastRow
    Dim iDblNetAmt, iDblVatAmt, iDblOldNetAmt, iDblOldVatAmt, iDblDiffNetAmt, iDblDiffVatAmt
    Dim iStrNetAmt, iStrVatAmt
    
    iDblDiffNetAmt = 0
    iDblDiffVatAmt = 0
    
    With frm1.vspdData
        Select Case pvStrFlag
            ' 추가/수정 
            Case "U"
                .Row = pvIntRow
                .Col = C_OldNetAmt: iDblOldNetAmt = UNICDbl(.Text)
                .Col = C_OldVatAmt: iDblOldVatAmt = UNICDbl(.Text)
                
                .Col = C_NetAmt: iStrNetAmt = .Text:            iDblDiffNetAmt = UNICDbl(.Text) - iDblOldNetAmt
                .Col = C_VatAmt: iStrVatAmt = .Text:            iDblDiffVatAmt = UNICDbl(.Text) - iDblOldVatAmt

                ' 변경후 값 설정 
                .Col = C_OldNetAmt:         .Text = iStrNetAmt
                .Col = C_OldVatAmt:         .Text = iStrVatAmt

            ' 취소 
            Case "C"
                ggoSpread.Source = frm1.vspdData
    
                .Row = pvIntRow
                .Col = C_OldNetAmt: iDblOldNetAmt = UNICDbl(.Text)
                .Col = C_OldVatAmt: iDblOldVatAmt = UNICDbl(.Text)
                .Col = 0
                Select Case .Text
                    Case ggoSpread.InsertFlag
                        iDblDiffNetAmt = -iDblOldNetAmt
                        iDblDiffVatAmt = -iDblOldVatAmt

                    Case ggoSpread.UpdateFlag
                        .Col = C_DBNetAmt: iDblDiffVatAmt = UNICDbl(.Text) - iDblOldNetAmt
                        .Col = C_DbVatAmt: iDblDiffVatAmt = UNICDbl(.Text) - iDblOldVatAmt

                    Case ggoSpread.DeleteFlag
                        .Col = C_DBNetAmt: iDblDiffNetAmt = UNICDbl(.Text)
                        .Col = C_DbVatAmt: iDblDiffVatAmt = UNICDbl(.Text)
                End Select

            ' 삭제 
            Case "D"
                ggoSpread.Source = frm1.vspdData
                iLngFirstRow = .SelBlockRow
                If iLngFirstRow = -1 Then
                    iLngFirstRow = 1
                    iLngLastRow = .MaxRows
                Else
                    iLngLastRow = .SelBlockRow2
                End If
                        
                For iLngRow = iLngFirstRow To iLngLastRow
                    .Row = iLngRow
                    .Col = 0
                    If .Text <> ggoSpread.DeleteFlag And .Text <> ggoSpread.InsertFlag Then
                        .Col = C_NetAmt: iDblDiffNetAmt = iDblDiffNetAmt - UNICDbl(.Text)
                        .Col = C_VatAmt: iDblDiffVatAmt = iDblDiffVatAmt - UNICDbl(.Text)
                    End If
                Next
                
        End Select
    End With

    Call SetTotal(iDblDiffNetAmt, iDblDiffVatAmt)
End Sub

' 총금액 설정 
Sub SetTotal(ByVal pvDblNetAmt, ByVal pvDblVatAmt)
    Dim iDblTotalAmt
    With frm1
        iDblTotalAmt = UNICDbl(.txtTotal_Amt.Text) + pvDblNetAmt + pvDblVatAmt
        pvDblNetAmt = UNICDbl(.txtNet_amt.Text) + pvDblNetAmt
        pvDblVatAmt = UNICDbl(.txtVat_amt.Text) + pvDblVatAmt

        .txtNet_amt.Text = UNIConvNumPCToCompanyByCurrency(pvDblNetAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
        .txtVat_amt.Text = UNIConvNumPCToCompanyByCurrency(pvDblVatAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo, "X")
        .txtTot_amt.Text = UNIConvNumPCToCompanyByCurrency(iDblTotalAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
        .txtTotal_Amt.Text = .txtTot_amt.Text
        
    End With
End Sub

' 부가세 금액을 계산한다.
Function FncCalcVatAmt(ByVal pvDblAmt, _
                    ByVal pvStrVatIncFlag, _
                    ByVal pvDblVatRate, _
                    ByVal pvStrCurrency)
    Dim iDblVatAmt
    
    ' 부가세가 별도인 경우 
    If pvStrVatIncFlag = "1" Then
        iDblVatAmt = pvDblAmt * pvDblVatRate / 100
    Else
        iDblVatAmt = pvDblAmt * pvDblVatRate / (100 + pvDblVatRate)
    End If
    
    FncCalcVatAmt = UNIConvNumPCToCompanyByCurrency(iDblVatAmt, pvStrCurrency, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo, "X")

End Function

'==============================================================================================================================
' 박정순 추가 (2006-04-27) 
'==============================================================================================================================
Function OpenGLRef() 

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
		
	Dim iCnt, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6 , a
	Dim rtn

'	msgbox frm1.txtConDn_no.Value 

'	rtn = CommonQueryRs("GOODS_MV_NO", "S_DN_HDR", " dn_no   = '" & frm1.txtConDn_no.Value & "'" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	rtn = CommonQueryRs("TOP 1 A.document_year", "i_goods_movement_header A , i_goods_movement_detail B ", " a.item_document_no = b.item_document_no and a.item_document_no   = '" & frm1.txtGINo.value & "' and B.dn_no   = '" & frm1.txtConDn_no.Value & "'" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

        If rtn = True Then
		
		iCnt = replace(lgF0, Chr(11),"")	

		If iCnt  <> "" Then

		   rtn = CommonQueryRs("temp_gl_no, gl_no", "a_batch", " ref_no    = '" & frm1.txtGINo.value & "-" & iCnt & "'"  , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

'		   MSGBOX frm1.txtGINo.value & "-" & iCnt

		   if replace(lgF0, Chr(11),"") <> ""  Then 
			a = "T"
			arrParam(0) = replace(lgF0, Chr(11),"")
		   End if

		   if replace(lgF1, Chr(11),"") <> ""  Then 
			a = "A"
			arrParam(0) = replace(lgF1, Chr(11),"")
		   End if

		End if
	ELSE
		a = "B"
	End If


	arrParam(1) = ""

   If a = "A" Then               '회계전표팝업 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			lblnWinEvent = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif a = "T" Then          '결의전표팝업 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			lblnWinEvent = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif a= "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다. 
    End if

	lblnWinEvent = False
	
End Function
'==============================================================================================================================
