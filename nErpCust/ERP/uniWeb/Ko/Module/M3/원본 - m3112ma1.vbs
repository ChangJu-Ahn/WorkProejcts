Option Explicit		

'******************************************  1.2 Global 변수/상수 선언  ***********************************

Const BIZ_PGM_ID 					= "m3112mb1.asp"		
Const BIZ_PGM_JUMP_ID 				= "M3111MA1"
Const BIZ_PGM_JUMP_ID_PUR_CHARGE	= "M6111MA2"
Const BIZ_PGM_JUMPORDERRUN_ID		= "m3110ma1.asp"

'==========================================  1.2.1 Global 상수 선언  ======================================
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
Dim C_Check
Dim C_CostCon
Dim C_CostConCd	
Dim C_OrderAmt
Dim C_NetAmt
Dim C_OrgNetAmt
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
Dim C_TrackingNo
Dim C_TrackingNoPop
Dim C_Lot_No 
Dim C_Lot_Seq 
Dim C_RetCd 
Dim C_Popup8
Dim C_RetNm 
Dim C_Over
Dim C_Under	
Dim C_Bal_Qty	
Dim C_Bal_Doc_Amt
Dim C_Bal_Loc_Amt
Dim C_ExRate
Dim C_SeqNo
Dim C_PrNo	
Dim C_MvmtNo	
Dim C_PoNo
Dim C_PoSeqNo
Dim C_MaintSeq
Dim C_SoNo	
Dim C_SoSeqNo
Dim C_OrgNetAmt1  
Dim C_reference 
Dim C_Stateflg
Dim C_Remrk
'==========================================  1.2.2 Global 변수 선언  =====================================


Dim lblnWinEvent
Dim releaseFlg
Dim arrCollectVatType

'==========================================  1.2.3 Global Variable값 정의  ===============================

Dim IsOpenPop          

'==========================================   Release()  ======================================
Sub Release()

    Err.Clear
    
    If CheckRunningBizProcess = True Then	
		Exit Sub
	End If                
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Trim(frm1.hdnMode.Value)	
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.Value)
    strVal = strVal & "&txtUpdtUserId=" & Parent.gUsrID   
    
    If LayerShowHide(1) = False Then Exit Sub
	Call RunMyBizASP(MyBizASP, strVal)								
	
End Sub
'==========================================   Cfm()  ======================================
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
			frm1.btnCfm.disabled = False	'200308          
			Exit Sub
		Else 
			frm1.btnCfm.disabled = True		'200308   
		End If
		frm1.hdnMode.Value = "Release"
					                                                
	elseif Trim(frm1.hdnReleaseflg.Value) = "Y" then
			
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			frm1.btnCfm.disabled = False	'200308    
			Exit Sub
		Else 		
			frm1.btnCfm.disabled = True		'200308 	
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
'=========================================================================================================
Sub ChangeItemPlant(byVal intStartRow ,byVal IntEndRow)
    Err.Clear                                                       
	
	Dim strVal
    Dim intIndex
    Dim lGrpCnt
	Dim igColSep,igRowSep
	
	igColSep = Parent.gColSep
	igRowSep = Parent.gRowSep
	
	If Trim(frm1.txtHMaintNo.Value) <> "" Then Exit Sub
		
    frm1.txtMode.Value = "LookUpItemPlant"			
	lGrpCnt = 1
	strVal = ""
	    
	For intIndex = intStartRow To intEndRow
		strVal = strVal & CStr(intIndex) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SLCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OrderUnit,intIndex,"X","X")) & igRowSep
			
		lGrpCnt = lGrpCnt + 1

		Call frm1.vspdData.SetText(C_Cost	,	intIndex, "")
		Call frm1.vspdData.SetText(C_Over	,	intIndex, "")
		Call frm1.vspdData.SetText(C_Under	,	intIndex, "")
	Next
		
	frm1.txtMaxRows.value = lGrpCnt-1
	frm1.txtSpread.value = strVal
	
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
	iStrSelect = " B.PUR_UNIT, A.ORDER_UNIT_PUR, C.BASIC_UNIT, A.MAJOR_SL_CD, A.SL_NM, C.ITEM_NM, C.SPEC, C.HS_CD,C.HS_NM, D.PLANT_NM, A.PROCUR_TYPE, A.TRACKING_FLG, "
	iStrSelect = iStrSelect & " B.UNDER_TOL, ISNULL(B.OVER_TOL, A.OVER_TOL) OVER_TOL  "

	iStrSql =""
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT S.ITEM_CD,S.ORDER_UNIT_PUR, S.MAJOR_SL_CD, S.PROCUR_TYPE, T.SL_NM, S.TRACKING_FLG, "
	iStrSql = iStrSql & " 			CASE WHEN S.OVER_RCPT_FLG = 'Y' THEN S.OVER_RCPT_RATE ELSE 0 END OVER_TOL "
	iStrSql = iStrSql & " 	FROM B_ITEM_BY_PLANT S LEFT OUTER JOIN B_STORAGE_LOCATION T ON(S.MAJOR_SL_CD=T.SL_CD AND S.PLANT_CD=T.PLANT_CD) "
	iStrSql = iStrSql & " WHERE S.PLANT_CD=" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND S.ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND S.VALID_FROM_DT <= GETDATE() AND S.VALID_TO_DT >= GETDATE() "
	iStrSql = iStrSql & " )A "
	iStrSql = iStrSql & " LEFT OUTER JOIN  "
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT ITEM_CD,PUR_UNIT,UNDER_TOL,OVER_TOL "
	iStrSql = iStrSql & " 	FROM M_SUPPLIER_ITEM_BY_PLANT "
	iStrSql = iStrSql & " WHERE PLANT_CD=" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND BP_CD IN(SELECT BP_CD FROM M_PUR_ORD_HDR WHERE PO_NO=" & FilterVar(Trim(frm1.txtPoNo.value), "''" , "S") & ") "
	iStrSql = iStrSql & " 	AND VALID_FR_DT <= GETDATE() AND VALID_TO_DT >= GETDATE() "
	iStrSql = iStrSql & " )B "
	iStrSql = iStrSql & " ON(A.ITEM_CD=B.ITEM_CD)  "
	iStrSql = iStrSql & " LEFT OUTER JOIN  "
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT S.ITEM_CD,S.BASIC_UNIT,S.ITEM_NM, S.SPEC, S.HS_CD, T.HS_NM "
	iStrSql = iStrSql & " FROM B_ITEM S LEFT OUTER JOIN B_HS_CODE T ON(S.HS_CD=T.HS_CD) "
	iStrSql = iStrSql & " WHERE S.ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " )C "
	iStrSql = iStrSql & " ON(A.ITEM_CD=C.ITEM_CD),  "
	iStrSql = iStrSql & " (SELECT PLANT_NM FROM B_PLANT WHERE PLANT_CD=" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X")), "''" , "S") & ")D "


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
	iSLCdArr		= Trim(arrVal2(4))
	iSLNmArr		= Trim(arrVal2(5))
	iItemNmArr		= Trim(arrVal2(6))
	iSpecArr		= Trim(arrVal2(7))
	iHSCdArr		= Trim(arrVal2(8))
	iHSNmArr		= Trim(arrVal2(9))
	iPlantNmArr		= Trim(arrVal2(10))
	iProcur_type	= Trim(arrVal2(11))
	iTracking_Flg	= Trim(arrVal2(12))
	iUnder_Tol      = Trim(arrVal2(13))
    iOver_Tol       = Trim(arrVal2(14))

	With frm1.vspdData
		.Row = lRow

		.Col = C_OrderUnit
		If Trim(iOrderUnitArr)<>"" Then
			.text = Trim(iOrderUnitArr)
		Else
			If Trim(iOrderUnitArr2)<>"" Then
				.text = Trim(iOrderUnitArr2)
			Else
				.text = Trim(iOrderUnitArr3)
			End If
		End If

		'=============================
		'품목의 조달구분 체크 
		'=============================
		If (Trim(iProcur_type)="P") And (Trim(frm1.hdnSubcontraflg.Value) = "Y") then
			Call DisplayMsgBox("179019","X","X","X")
			.Col = C_itemCd
			.text = ""
			Exit Sub
		End If
		If (Trim(iProcur_type)<>"P") And (Trim(frm1.hdnSubcontraflg.Value) = "N") then
			Call DisplayMsgBox("179019","X","X","X")
			.Col = C_itemCd
			.text = ""
			Exit Sub
		End If


		.Col = C_SLCd
		.text = Trim(iSLCdArr)

		.Col = C_SLNm
		.text = Trim(iSLNmArr)

		.Col = C_itemNm
		.text = Trim(iItemNmArr)

		.Col = C_SpplSpec
		.text = Trim(iSpecArr)

		.Col = C_HSCd
		.text = Trim(iHSCdArr)

		.Col = C_HSNm
		.text = Trim(iHSNmArr)

		.Col = C_PlantNm
		.text = Trim(iPlantNmArr)
		
		.Col = C_PrNo

		If .text = "" Then		
			If iTracking_Flg <> "Y" Then
				ggoSpread.spreadlock C_TrackingNo, .Row, C_TrackingNoPop, .Row
				.Col = C_TrackingNo
				.text = "*"
			Else
				ggoSpread.spreadUnlock C_TrackingNo, .Row, C_TrackingNoPop, .Row
				ggoSpread.sssetrequired C_TrackingNo, .Row, .Row
				.Col = C_TrackingNo
				.text = ""
			End If
		End If
		
		'2006.12.8 Modified by KSJ
		.Col = C_Over
		.text = iOver_Tol
		
		.Col = C_Under
		.text = iUnder_Tol

	End With

End Sub

'==========================================   ChangeItemPlantForUnit()  ======================================
'	Name : ChangeItemPlantForUnit()
'	Description : 단위변경시 
'=========================================================================================================

Sub ChangeItemPlantForUnit(byVal intStartRow ,byVal IntEndRow)

    Err.Clear                                       

    Dim strVal
    Dim intIndex
    Dim lGrpCnt
	Dim igColSep,igRowSep
	
	igColSep = Parent.gColSep
	igRowSep = Parent.gRowSep

	If Trim(frm1.txtHMaintNo.Value) <> "" Then Exit Sub
		
    frm1.txtMode.Value = "LookUpItemPlantForUnit"	
	lGrpCnt = 1
	strVal = ""
	For intIndex = intStartRow To intEndRow
		strVal = strVal & CStr(intIndex) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OrderUnit,intIndex,"X","X")) & igRowSep
		
		lGrpCnt = lGrpCnt + 1
	Next

	frm1.txtMaxRows.value = lGrpCnt-1
	frm1.txtSpread.value = strVal
	
    If LayerShowHide(1) = False Then Exit Sub
    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)					
	
End Sub

'==========================================   ChangeItemPlantForUnit2()  ======================================
'	Name : ChangeItemPlantForUnit2()
'	Description : 단위변경시 
'=========================================================================================================

Sub ChangeItemPlantForUnit2(byVal lRow)

	Dim strsstemp1,strsstemp2,strsstemp3
	Dim strWhere, strPriceType
    
    ggoSpread.Source = frm1.vspdData

    with frm1.vspdData 		
		.Row = lRow

		.Col 		= C_ItemCd
		strssTemp1 	= Trim(.Text)
		.Col 		= C_PlantCd
		strssTemp2 	= Trim(.Text)
		.Col 		= C_OrderUnit
		strssTemp3 	= Trim(.Text)
		
		If strssTemp1 = "" Or strssTemp2 = "" Or strssTemp3 = "" Then
			Exit Sub
		End if

		' 단가type 의 유무를 조사 
		Call CommonQueryRs(" MINOR_CD ", " B_CONFIGURATION ", " MAJOR_CD = " & FilterVar("M0001", "''", "S") & " AND REFERENCE = " & FilterVar("Y", "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Err.number <> 0 Then
			MsgBox Err.description, VbInformation, parent.gLogoName
			Err.Clear 
			Exit Sub
		End If
		
		If Len(lgF0) > 0 Then
			lgF0 = Split(lgF0, Chr(11))
			strPriceType = lgF0(0)
		Else
			Call DisplayMsgBox("171214", "X", "X", "X")
			Exit Sub
		End If
	
		strWhere = " PLANT_CD = " & FilterVar(strssTemp2, "''", "S")
		strWhere = strWhere & " AND ITEM_CD = " & FilterVar(strssTemp1, "''", "S")
		strWhere = strWhere & " AND BP_CD = " & FilterVar(Trim(frm1.txtSupplierCd.value), "''", "S")
		strWhere = strWhere & " AND PUR_UNIT = " & FilterVar(strssTemp3, "''", "S")
		strWhere = strWhere & " AND PUR_CUR = " & FilterVar(Trim(frm1.txtCurr.value), "''", "S")
		strWhere = strWhere & " AND VALID_FR_DT <= " & FilterVar(Trim(frm1.txtPoDt.text), "''", "S")
		If Trim(strPriceType) = "T" Then
			strWhere = strWhere & " AND PRC_FLG =  'T' "
		End If
		strWhere = strWhere & " ORDER BY VALID_FR_DT DESC "
	
		Call CommonQueryRs(" PUR_PRC, PRC_FLG ", " M_SUPPLIER_ITEM_PRICE ", strwhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Err.number <> 0 Then
			MsgBox Err.description, VbInformation, parent.gLogoName
			Err.Clear 
			Exit Sub
		End If
	
		If Len(lgF0) > 0 Then
			lgF0 = Split(lgF0, Chr(11))
			lgF1 = Split(lgF1, Chr(11))
			.Col = C_Cost
			.Text = lgF0(0)
			.Col = C_CostConCd
			.Text = lgF1(0)
		Else
			strWhere = " PLANT_CD = " & FilterVar(strssTemp2, "''", "S")
			strWhere = strWhere & " AND ITEM_CD = " & FilterVar(strssTemp1, "''", "S")
			strWhere = strWhere & " AND PUR_UNIT = " & FilterVar(strssTemp3, "''", "S")
			strWhere = strWhere & " AND PUR_CUR = " & FilterVar(Trim(frm1.txtCurr.value), "''", "S")
			strWhere = strWhere & " AND VALID_FR_DT <= " & FilterVar(Trim(frm1.txtPoDt.text), "''", "S")
			If Trim(strPriceType) = "T" Then
				strWhere = strWhere & " AND PRC_FLG =  'T' "
			End If
			strWhere = strWhere & " ORDER BY VALID_FR_DT DESC "
	
			Call CommonQueryRs(" PUR_PRC, PRC_FLG ", " M_ITEM_PUR_PRICE ", strwhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			If Err.number <> 0 Then
				MsgBox Err.description, VbInformation, parent.gLogoName
				Err.Clear 
				Exit Sub
			End If
	
			If Len(lgF0) > 0 Then
				lgF0 = Split(lgF0, Chr(11))
				lgF1 = Split(lgF1, Chr(11))
				.Col = C_Cost
				.Text = lgF0(0)
				.Col = C_CostConCd
				.Text = lgF1(0)
			Else
				.Col = C_Cost
				.Text = 0
			End If
		End If
		
		Call vspdData_Change(C_Cost, lRow)
		Call vspdData_Change(C_CostConCd, lRow)
	End With

End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
    Call SetToolbar("1110000000001111")
    frm1.btnCfmSel.disabled = true
    frm1.btnCfm.value = "확정"
    ' === 2005.07.15 단가 일괄불러오기 =============
    frm1.btnCallPrice.disabled = False
    ' === 2005.07.15 단가 일괄불러오기 =============    
    frm1.txtPoNo.focus 
	Set gActiveElement = document.activeElement
End Sub

'========================================  2.2.3 InitSpreadPosVariables() ========================================
Sub InitSpreadPosVariables()
	C_PlantCd 		= 1 
	C_Popup1		= 2 
	C_PlantNm 		= 3 
	C_itemCd 		= 4 
	C_Popup2 		= 5 
	C_itemNm 		= 6 
	C_SpplSpec      = 7   
	C_OrderQty		= 8 
	C_OrderUnit		= 9 
	C_Popup3		= 10
	C_Cost			= 11
	C_Check			= 12
	C_CostCon		= 13
	C_CostConCd		= 14
	C_OrderAmt		= 15
	C_NetAmt        = 16  
	C_OrgNetAmt     = 17  
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
	C_TrackingNo	= 32
	C_TrackingNoPop	= 33
	C_Lot_No        = 34  
	C_Lot_Seq       = 35  
	C_RetCd         = 36
	C_Popup8        = 37
	C_RetNm         = 38      
	C_Over			= 39
	C_Under			= 40
	C_Bal_Qty		= 41
	C_Bal_Doc_Amt	= 42
	C_Bal_Loc_Amt	= 43
	C_ExRate		= 44
	C_SeqNo 		= 45
	C_PrNo			= 46
	C_MvmtNo		= 47
	C_PoNo			= 48
	C_PoSeqNo		= 49
	C_MaintSeq		= 50
	C_SoNo			= 51
	C_SoSeqNo		= 52
	C_OrgNetAmt1    = 53  
	C_reference     = 54
	C_Stateflg		= 55
	C_Remrk			= 56
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables
	
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20051201",,Parent.gAllowDragDropSpread  
	
	.ReDraw = false

    .MaxCols = C_Remrk+1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0
	
    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit 	C_PlantCd, "공장", 7,,,4,2
    ggoSpread.SSSetButton 	C_Popup1
    ggoSpread.SSSetEdit 	C_PlantNm, "공장명", 20
    ggoSpread.SSSetEdit 	C_ItemCd, "품목", 18,,,18,2
    ggoSpread.SSSetButton 	C_Popup2
    ggoSpread.SSSetEdit 	C_ItemNm, "품목명", 20    
    ggoSpread.SSSetEdit		C_SpplSpec, "품목규격", 20        '품목규격 추가 
    SetSpreadFloatLocal		C_OrderQty, "발주수량",15,1,3
    ggoSpread.SSSetEdit 	C_OrderUnit, "단위", 6,,,3,2
    ggoSpread.sssetButton 	C_Popup3
    SetSpreadFloatLocal		C_Cost, "단가",15,1,4
    ggoSpread.sssetButton	C_Check
    ggoSpread.SSSetCombo 	C_CostCon, "단가구분", 10,0,False
    ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
    ggoSpread.SSSetCombo 	C_CostConCd, "단가구분코드", 10,0,False
    ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
    SetSpreadFloatLocal		C_OrderAmt, "금액",15,1,2
    SetSpreadFloatLocal		C_NetAmt, "발주순금액",15,1,2
    SetSpreadFloatLocal		C_OrgNetAmt, "C_OrgNetAmt",15,1,2
    SetSpreadFloatLocal		C_OrgNetAmt1, "C_OrgNetAmt1",15,1,2
    ggoSpread.SSSetDate 	C_DlvyDt, "납기일", 10, 2, Parent.gDateFormat
    ggoSpread.SSSetEdit 	C_HSCd, "HS부호", 15,,,20,2
    ggoSpread.sssetButton 	C_Popup5
    ggoSpread.SSSetEdit 	C_HSNm, "HS명", 20
    ggoSpread.SSSetEdit 	C_SLCd, "창고", 10,,,7,2
    ggoSpread.SSSetButton 	C_Popup6
    ggoSpread.SSSetEdit 	C_SLNm, "창고명", 20
    ggoSpread.SSSetEdit 	C_TrackingNo, "Tracking No.",  15,,,25,2
    ggoSpread.SSSetButton 	C_TrackingNoPop
    ggoSpread.SSSetEdit 	C_Lot_No, "Lot No.",  15,,,9,2           '13 차 추가 
    ggoSpread.SSSetEdit 	C_Lot_Seq, "Lot No.순번",  15,,,15,2      '13 차 추가  
    SetSpreadFloatLocal 	C_Over, "과부족허용율(+)(%)",20,1,6
    SetSpreadFloatLocal 	C_Under,"과부족허용율(-)(%)",20,1,6
    ggoSpread.SSSetCombo	C_IOFlg,"VAT포함여부", 15,2,False               '13 차 추가 
    ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
    ggoSpread.SSSetCombo 	C_IOFlgCd, "VAT포함여부코드", 15,2,False
    ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd
    ggoSpread.SSSetEdit 	C_VatType, "VAT", 7,,,4,2
    ggoSpread.SSSetButton 	C_Popup7
    ggoSpread.SSSetEdit 	C_VatNm, "VAT명", 20 
    SetSpreadFloatLocal		C_VatRate, "VAT율(%)",15,1,5
    SetSpreadFloatLocal		C_VatAmt, "VAT금액",15,1,2
    ggoSpread.SSSetEdit 	C_RetCd , "반품유형", 10,,,4,2
    ggoSpread.SSSetButton 	C_Popup8
    ggoSpread.SSSetEdit 	C_RetNm , "반품유형명", 20 
    SetSpreadFloatLocal		C_Bal_Qty, "Bal. Qty.",15,1,3  
    SetSpreadFloatLocal		C_Bal_Doc_Amt, "Bal. Doc. Amt.",15,1,2  
    SetSpreadFloatLocal		C_Bal_Loc_Amt, "Bal. Loc. Amt.",15,1,2  
    SetSpreadFloatLocal		C_ExRate, "Xch. Rate",15,1,5  
    ggoSpread.SSSetEdit 	C_SeqNo, "순번", 10
    ggoSpread.SSSetEdit 	C_PrNo, "구매요청번호", 20
    ggoSpread.SSSetEdit 	C_MvmtNo, "구매입고번호", 20
    ggoSpread.SSSetEdit 	C_PoNo, "발주번호", 20
    ggoSpread.SSSetEdit 	C_PoSeqNo, "발주SEQNO", 20
    ggoSpread.SSSetEdit 	C_MaintSeq, "maintseq", 10
	ggoSpread.SSSetEdit 	C_SoNo, "", 10
	ggoSpread.SSSetEdit 	C_SoSeqNo, "", 10
    ggoSpread.SSSetEdit 	C_Stateflg, "stateflg", 10
    ggoSpread.SSSetEdit 	C_reference, "reference", 10
    ggoSpread.SSSetEdit 	C_Remrk, "비고", 20,,,120,2
 
    
	Call ggoSpread.MakePairsColumn(C_PlantCd,C_Popup1)
	Call ggoSpread.MakePairsColumn(C_ItemCd,C_Popup2)
	Call ggoSpread.MakePairsColumn(C_OrderUnit,C_Popup3)
	Call ggoSpread.MakePairsColumn(C_Cost,C_Check)
	Call ggoSpread.MakePairsColumn(C_HSCd,C_Popup5)
	Call ggoSpread.MakePairsColumn(C_SLCd,C_Popup6)
	Call ggoSpread.MakePairsColumn(C_TrackingNo,C_TrackingNoPop)
	Call ggoSpread.MakePairsColumn(C_VatType,C_Popup7)
	Call ggoSpread.MakePairsColumn(C_RetCd,C_Popup8)

	Call ggoSpread.SSSetColHidden(C_SeqNo,C_SeqNo,True)	
	Call ggoSpread.SSSetColHidden(C_Lot_Seq,C_Lot_Seq,True)	
	Call ggoSpread.SSSetColHidden(C_Lot_No,C_Lot_No,True)	
	Call ggoSpread.SSSetColHidden(C_IOFlgCd,C_IOFlgCd,True)	
	Call ggoSpread.SSSetColHidden(C_Bal_Qty,C_Bal_Qty,True)	
	Call ggoSpread.SSSetColHidden(C_Bal_Doc_Amt,C_Bal_Doc_Amt,True)	
	Call ggoSpread.SSSetColHidden(C_Bal_Loc_Amt,C_Bal_Loc_Amt,True)	
	Call ggoSpread.SSSetColHidden(C_ExRate,C_ExRate,True)	
	Call ggoSpread.SSSetColHidden(C_CostConCd,C_CostConCd,True)	
	'Call ggoSpread.SSSetColHidden(C_PrNo,C_PrNo,True)	
	Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,True)	
	Call ggoSpread.SSSetColHidden(C_PoNo,C_PoNo,True)	
	Call ggoSpread.SSSetColHidden(C_PoSeqNo,C_PoSeqNo,True)	
	Call ggoSpread.SSSetColHidden(C_MaintSeq,C_MaintSeq,True)	
	Call ggoSpread.SSSetColHidden(C_SoNo,C_SoNo,True)	
	Call ggoSpread.SSSetColHidden(C_SoSeqNo,C_SoSeqNo,True)	
	Call ggoSpread.SSSetColHidden(C_Stateflg,C_Stateflg,True)	
	Call ggoSpread.SSSetColHidden(C_RetCd,C_RetCd,True)	
	Call ggoSpread.SSSetColHidden(C_Popup8,C_Popup8,True)	
	Call ggoSpread.SSSetColHidden(C_RetNm,C_RetNm,True)	
	Call ggoSpread.SSSetColHidden(C_OrgNetAmt,C_OrgNetAmt,True)	
	Call ggoSpread.SSSetColHidden(C_OrgNetAmt1,C_OrgNetAmt1,True)	
	Call ggoSpread.SSSetColHidden(C_reference,C_reference,True)	
        
    ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
    ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
    ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
    ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd
    
    Call SetSpreadLock
    
	.ReDraw = true
	
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
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
    ggoSpread.sssetrequired C_CostCon, -1
    If Trim(.hdnreference.value) = "N" then
        ggoSpread.spreadlock C_OrderAmt, -1
    else 
        ggoSpread.SpreadUnLock C_OrderAmt, -1
        ggoSpread.sssetrequired C_OrderAmt, -1
    end if
    
    ggoSpread.spreadlock C_NetAmt, -1
    ggoSpread.SpreadUnLock C_DlvyDT, -1
    ggoSpread.sssetrequired C_DlvyDT, -1
    ggoSpread.spreadlock C_HSCd, -1
    ggoSpread.spreadlock C_Popup5, -1
    ggoSpread.spreadlock C_HSNm, -1
    ggoSpread.SpreadUnLock C_SLCd , -1
    ggoSpread.sssetrequired C_SLCd, -1
    ggoSpread.SpreadUnLock C_Popup6 , -1
    ggoSpread.spreadlock C_SLNm, -1
    
    ggoSpread.SpreadUnLock C_VatType , -1
    ggoSpread.SpreadUnLock C_Popup7 , -1
    ggoSpread.SpreadUnLock C_VatNm , -1
    ggoSpread.SpreadUnLock C_VatRate , -1
    ggoSpread.SpreadUnLock C_VatAmt , -1
    
    ggoSpread.SpreadUnLock C_Popup8 , -1
    ggoSpread.spreadlock C_RetNm , -1
    ggoSpread.spreadUnLock C_IOFlg , -1    '13차추가 
    ggoSpread.sssetrequired C_IOFlg, -1
    ggoSpread.SpreadLock C_IOFlgCd, -1
    ggoSpread.spreadlock C_Lot_No , -1     '13차추가 
    ggoSpread.spreadlock C_Lot_Seq , -1    '13차추가 
    ggoSpread.spreadlock C_TrackingNo , -1  
	ggoSpread.spreadUnlock C_Under, -1
	ggoSpread.spreadUnlock C_Over, -1
	ggoSpread.spreadlock C_PrNo, -1       '2006-09
    
    End With
End Sub
'================================== SetSpreadLockAfterQuery() ======================================
Sub SetSpreadLockAfterQuery()

	Dim index,Count,index1 

    With frm1
    
    .vspdData.ReDraw = False
    
    if .vspdData.MaxRows < 1 then
		Exit sub
	end if
	
    if .txtRelease.Value = "Y" then
		For index = C_PlantCd to C_Stateflg
			ggoSpread.SpreadLock index , -1
		Next
	Else
		For index1 = Cint(.hdnmaxrow.value) + 1 to .vspdData.MaxRows
		    ggoSpread.SpreadLock frm1.vspddata.maxcols, index1, frm1.vspddata.maxcols, index1
			ggoSpread.SpreadLock C_SeqNo , index1,C_SeqNo,index1
			ggoSpread.SpreadLock C_PlantCd ,index1,C_PlantCd,index1
			ggoSpread.SpreadLock C_Popup1 , index1,C_Popup1,index1
			ggoSpread.spreadlock C_PlantNm , index1,C_PlantNm,index1
			ggoSpread.SpreadLock C_ItemCd, index1,C_ItemCd,index1
			ggoSpread.SpreadLock C_Popup2 , index1,C_Popup2,index1
			ggoSpread.spreadlock C_ItemNm , index1,C_ItemNm,index1
			ggoSpread.spreadlock C_SpplSpec,index1,C_SpplSpec,index1         '품목규격 추가 
			ggoSpread.SpreadUnLock C_OrderQty,index1,C_OrderQty,index1
			ggoSpread.sssetrequired C_OrderQty, index1,index1
			
			if UCase(frm1.hdnRetflg.Value) = "N" then
				ggoSpread.SpreadUnLock C_OrderUnit , index1,C_OrderUnit,index1
				ggoSpread.sssetrequired C_OrderUnit, index1,index1
				ggoSpread.SpreadUnLock C_Popup3 , index1,C_Popup3,index1
				ggoSpread.SpreadUnLock C_Cost , index1,C_Cost,index1
				ggoSpread.sssetrequired C_Cost, index1,index1
			else
				ggoSpread.SpreadLock C_OrderUnit , index1,C_OrderUnit,index1
				ggoSpread.SpreadLock C_Popup3 , index1,C_Popup3,index1
				ggoSpread.SpreadLock C_Cost , index1,C_Cost,index1
			end if		

			ggoSpread.SpreadUnLock C_CostCon, index1,C_CostCon,index1
			ggoSpread.sssetrequired C_CostCon, index1,index1
			ggoSpread.spreadlock C_NetAmt, index1,C_NetAmt,index1		

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
			ggoSpread.SpreadUnLock C_IOFlg, index1,C_IOFlgCd,index1 
			ggoSpread.SSSetRequired	C_IOFlg, index1,index1 
			ggoSpread.SSSetRequired	C_IOFlgCd, index1,index1
		    
			ggoSpread.SpreadUnLock C_VatType , index1,C_VatType,index1
			ggoSpread.SpreadUnLock C_Popup7 , index1,C_Popup7,index1
			ggoSpread.spreadlock C_VatNm , index1,C_VatNm,index1
			ggoSpread.spreadlock C_VatRate , index1,C_VatRate,index1
			ggoSpread.spreadlock C_VatAmt , index1,C_VatAmt,index1
		'******************************************
		  '13차추가]
			if .hdnRetflg.Value = "Y" then
				ggoSpread.spreadUnLock C_RetCd , index1,C_RetCd,index1
				ggoSpread.SpreadUnLock C_Popup8 , index1,C_Popup8,index1
				ggoSpread.spreadlock   C_RetNm , index1,C_RetNm,index1
				ggoSpread.spreadUnLock C_Lot_No , index1,C_Lot_No,index1       
				ggoSpread.spreadUnLock C_Lot_Seq , index1,C_Lot_Seq,index1 
			else
				ggoSpread.spreadlock C_RetCd , index1,C_RetCd,index1		
				ggoSpread.spreadlock C_Popup8 , index1,C_Popup8,index1		
				ggoSpread.spreadlock C_RetNm , index1,C_RetNm,index1		
		        ggoSpread.spreadlock C_Lot_No , index1,C_Lot_No,index1        
		        ggoSpread.spreadlock C_Lot_Seq , index1,C_Lot_Seq,index1      
		    end if        
		'******************************************
		    ggoSpread.SpreadUnLock C_SLCd , index1,C_SLCd,index1
		    ggoSpread.sssetrequired C_SLCd, index1,index1
		    ggoSpread.SpreadUnLock C_Popup6 , index1,C_Popup6,index1
		    ggoSpread.spreadlock C_SLNm, index1,C_SLNm,index1
			
            .vspdData.Row = index1
			.vspdData.Col = C_TrackingNo
			if Trim(.vspdData.Text) = "*" then
				ggoSpread.spreadlock C_TrackingNo, index1, C_TrackingNoPop, index1
			else
				ggoSpread.spreadUnlock C_TrackingNo, index1, C_TrackingNoPop, index1
				ggoSpread.sssetrequired C_TrackingNo, index1, index1
			end if

			'************************************************ 13차	

			frm1.vspdData.Row = index1
		    frm1.vspdData.Col = C_PrNo
			if Trim(.vspdData.Text) <> "" then
				ggoSpread.spreadlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.spreadlock C_Popup3 , index1, C_Popup3, index1
		        ggoSpread.spreadlock C_DlvyDT, index1,C_DlvyDT, index1
			else
				ggoSpread.spreadUnlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.sssetrequired C_OrderUnit, index1, index1
				ggoSpread.SpreadUnLock C_DlvyDT, index1,C_DlvyDT, index1
			    ggoSpread.sssetrequired C_DlvyDT, index1, index1
			end if
		    ggoSpread.spreadUnlock C_Under,index1,C_Under,index1
		    ggoSpread.spreadUnlock C_Over,index1,C_Over,index1
		    ggoSpread.spreadlock C_PrNo, index1, C_PrNo, index1
	    next
	End if
	
    .vspdData.ReDraw = True
    
    End With
End Sub
'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SeqNo		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_PlantCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ItemCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SpplSpec	, pvStartRow, pvEndRow '품목규격 추가 
    ggoSpread.SSSetRequired		C_OrderQty	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_OrderUnit	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_Cost		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_CostCon	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_CostConCd	, pvStartRow, pvEndRow
    
    If Trim(.hdnreference.value) = "N" then
        ggoSpread.SSSetProtected	C_OrderAmt, pvStartRow, pvEndRow
    else 
        ggoSpread.SSSetRequired  C_OrderAmt, pvStartRow, pvEndRow
    end if
    
    ggoSpread.SSSetProtected	C_NetAmt, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_DlvyDt, pvStartRow, pvEndRow
    
    if Trim(frm1.hdnImportflg.value) <> "Y" then
	    ggoSpread.SSSetProtected	C_HSCd	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_Popup5, pvStartRow, pvEndRow
	else
		ggoSpread.spreadUnlock	C_HSCd	, pvStartRow, C_HSCd, pvEndRow
		ggoSpread.sssetrequired	C_HSCd	, pvStartRow, pvEndRow
		ggoSpread.spreadUnlock	C_Popup5, pvStartRow, C_Popup5, pvEndRow
	end if
	
	ggoSpread.SSSetProtected		C_TrackingNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_TrackingNoPop, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_HSNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_SLCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_SLNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatRate, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatAmt , pvStartRow, pvEndRow
    '******************************************
	ggoSpread.SSSetRequired		C_IOFlg	 , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected		C_IOFlgCd, pvStartRow, pvEndRow  '13차추가 
	if .hdnRetflg.Value <> "Y" then
		ggoSpread.SSSetProtected C_RetCd	, pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected C_Popup8, pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected C_RetNm	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Lot_No, pvStartRow, pvEndRow        
		ggoSpread.SSSetProtected C_Lot_Seq, pvStartRow, pvEndRow      
	end if        
	'******************************************
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColorRef(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SeqNo		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_PlantCd	, pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_PlantCd	, pvStartRow, C_PlantCd, pvEndRow
	ggoSpread.spreadlock		C_Popup1	, pvStartRow, C_Popup1,  pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm	, pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_ItemCd	, pvStartRow, C_ItemCd, pvEndRow
	ggoSpread.spreadlock		C_Popup2	, pvStartRow, C_Popup2,  pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SpplSpec	, pvStartRow, pvEndRow '품목규격 추가 
    ggoSpread.spreadUnlock		C_OrderQty	, pvStartRow, C_OrderQty,	pvEndRow '20040609추가 
    ggoSpread.SSSetRequired		C_OrderQty	, pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_OrderUnit	, pvStartRow, C_OrderUnit, pvEndRow
    ggoSpread.spreadlock		C_Popup3	, pvStartRow, C_Popup3, pvEndRow	'20040524 팝업속성수정 
    ggoSpread.spreadUnlock		C_Cost		, pvStartRow, C_Cost,	pvEndRow	'20040609추가 
    ggoSpread.SSSetRequired		C_Cost		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_CostCon	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_CostConCd	, pvStartRow, pvEndRow
    
    If Trim(.hdnreference.value) = "N" then
        ggoSpread.SSSetProtected	C_OrderAmt, pvStartRow, pvEndRow
    else 
        ggoSpread.SSSetRequired  C_OrderAmt, pvStartRow, pvEndRow
    end if
    
    ggoSpread.SSSetProtected	C_NetAmt, pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_DlvyDt	, pvStartRow, C_DlvyDt, pvEndRow
    if Trim(frm1.hdnImportflg.value) <> "Y" then
	    ggoSpread.SSSetProtected	C_HSCd	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_Popup5, pvStartRow, pvEndRow
	else
		ggoSpread.spreadUnlock	C_HSCd	, pvStartRow, C_HSCd, pvEndRow
		ggoSpread.sssetrequired	C_HSCd	, pvStartRow, pvEndRow
		ggoSpread.spreadUnlock	C_Popup5, pvStartRow, C_Popup5, pvEndRow
	end if
	
	ggoSpread.SSSetProtected		C_TrackingNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_TrackingNoPop, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_HSNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired			C_SLCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_SLNm	, pvStartRow, pvEndRow
    ggoSpread.spreadUnlock			C_Popup6, pvStartRow, C_Popup6, pvEndRow '20040524 팝업속성수정 
    ggoSpread.SSSetProtected		C_VatNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatRate, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatAmt , pvStartRow, pvEndRow
    '******************************************
	ggoSpread.SSSetRequired		C_IOFlg	 , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected		C_IOFlgCd, pvStartRow, pvEndRow  '13차추가 
	if .hdnRetflg.Value <> "Y" then
		ggoSpread.SSSetProtected C_RetCd	, pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected C_Popup8, pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected C_RetNm	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Lot_No, pvStartRow, pvEndRow        
		ggoSpread.SSSetProtected C_Lot_Seq, pvStartRow, pvEndRow      
	end if        
	'******************************************
    End With
End Sub
'==================================== GetSpreadColumnPos() ====================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd 		= iCurColumnPos(1)
			C_Popup1		= iCurColumnPos(2)
			C_PlantNm 		= iCurColumnPos(3)
			C_itemCd 		= iCurColumnPos(4)
			C_Popup2 		= iCurColumnPos(5)
			C_itemNm 		= iCurColumnPos(6)
			C_SpplSpec      = iCurColumnPos(7)
			C_OrderQty		= iCurColumnPos(8)
			C_OrderUnit		= iCurColumnPos(9)
			C_Popup3		= iCurColumnPos(10)
			C_Cost			= iCurColumnPos(11)
			C_Check			= iCurColumnPos(12)
			C_CostCon		= iCurColumnPos(13)
			C_CostConCd		= iCurColumnPos(14)
			C_OrderAmt		= iCurColumnPos(15)
			C_NetAmt        = iCurColumnPos(16)
			C_OrgNetAmt     = iCurColumnPos(17)
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
			C_TrackingNo	= iCurColumnPos(32)
			C_TrackingNoPop	= iCurColumnPos(33)
			C_Lot_No        = iCurColumnPos(34)
			C_Lot_Seq       = iCurColumnPos(35)
			C_RetCd         = iCurColumnPos(36)
			C_Popup8        = iCurColumnPos(37)
			C_RetNm         = iCurColumnPos(38)
			C_Over			= iCurColumnPos(39)
			C_Under			= iCurColumnPos(40)
			C_Bal_Qty		= iCurColumnPos(41)
			C_Bal_Doc_Amt	= iCurColumnPos(42)
			C_Bal_Loc_Amt	= iCurColumnPos(43)
			C_ExRate		= iCurColumnPos(44)
			C_SeqNo 		= iCurColumnPos(45)
			C_PrNo			= iCurColumnPos(46)
			C_MvmtNo		= iCurColumnPos(47)
			C_PoNo			= iCurColumnPos(48)
			C_PoSeqNo		= iCurColumnPos(49)
			C_MaintSeq		= iCurColumnPos(50)
			C_SoNo			= iCurColumnPos(51)
			C_SoSeqNo		= iCurColumnPos(52)
			C_OrgNetAmt1    = iCurColumnPos(53)
			C_reference     = iCurColumnPos(54)
			C_Stateflg		= iCurColumnPos(55)
			C_Remrk			= iCurColumnPos(56)
	End Select

End Sub	
'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
	Dim IntRetCD
			
	If lblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
		
	arrParam(0) = "N"  'Return Flag
	arrParam(1) = "N"  'Release Flag
	arrParam(2) = ""  'STO Flag

	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function

'------------------------------------------  OpenReqRef()  -------------------------------------------------
'	Name : OpenReqRef()
'	Description :구매요청참조 
'---------------------------------------------------------------------------------------------------------
Function OpenReqRef()

	Dim strRet
	Dim arrParam(4)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_CMODE then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End if 
	
	if frm1.txtRelease.Value = "Y" then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		Exit Function
	End if
	
	if UCase(frm1.hdnRetflg.Value) <> "N" then
		Call DisplayMsgBox("17A012", "X","발주형태" & frm1.txtPotypeCd.Value & "(" & frm1.txtPoTypeNm.value & ")","구매요청참조" )
		Exit Function
	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(4) = Trim(frm1.hdnSubcontraflg.value)
	
	iCalledAspName = AskPRAspName("M2111RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M2111RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetReqRef(strRet)
	End If
		
End Function
'------------------------------------------  SetReqRef()  -------------------------------------------------

Function SetReqRef(strRet)
	Dim Index1,index2,Index3,Count1,Count2
	Dim IntIflg
	Dim strMessage
	Dim intstartRow,intEndRow, TempRow
	Dim iInsRow,intInsertRowsCount
	
	Const C_ReqNo_Ref		= 0
	Const C_PlantCd_Ref		= 1
	Const C_PlantNm_Ref		= 2
	Const C_ItemCd_Ref		= 3
	Const C_ItemNm_Ref		= 4
	Const C_SpplSpec_Ref    = 5                         '품목 규격 추가 
	Const C_Qty_Ref			= 6
	Const C_Unit_Ref		= 7
	Const C_DlvyDt_Ref		= 8
	Const C_Pur_Plan_Dt_Ref	= 9
	Const C_Pr_Type_Ref		= 10
	Const C_Pr_Type_Nm_Ref	= 11
	Const C_SoNo_Ref		= 12
	Const C_SoSeqNo_Ref		= 13
	Const C_TrackingNo_Ref	= 14
	Const C_SLCd_Ref		= 15
	Const C_SLNm_Ref		= 16 
	Const C_HSCd_Ref		= 17
	Const C_HSNm_Ref		= 18


	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	strMessage = ""
	IntIflg=true
	
	with frm1
		
	.vspdData.Redraw = False
	.vspdData.focus
	ggoSpread.Source = .vspdData
	intStartRow = .vspdData.MaxRows + 1
	
	TempRow = .vspdData.MaxRows					'리스트 max값 
	
	intInsertRowsCount = 0 '중복 안될때만 MAXROW에 1을 추가하기 위한변수 
	
	'중복된 요청건참조시 MAXROW값계산 로직 수정 200308
	for index1 = 0 to Count1
	
		.vspdData.Row=Index1+1
		
		If TempRow <> 0 Then
			For Index3 = 1 to TempRow
				if GetSpreadText(.vspdData,C_PrNo,index3,"X","X") = strRet(index1,C_ReqNo_Ref) then
					strMessage = strMessage & strRet(Index1,C_ReqNo_Ref) & ";"
					intIflg=False					
					intInsertRowsCount = 0		'중복될땐 MAXROW를 증가시키지 않음.					
					Exit for
				Else 
					intInsertRowsCount =  1
				End if
			Next
		Else 		
			intInsertRowsCount =  1				
		End If
		
		if IntIflg <> False then
			
			.vspdData.MaxRows = CLng(TempRow) + CLng(intInsertRowsCount) 
			iInsRow = CLng(TempRow) + CLng(intInsertRowsCount) 
			
			TempRow = CLng(TempRow) + CLng(intInsertRowsCount) '다음 MAXROW계산시 베이스가 될 TempRow 를 증가시킴.
			lgBlnFlgChgValue = True
			
			Call .vspdData.SetText(0		,	iInsRow, ggoSpread.InsertFlag)
			Call .vspdData.SetText(C_VatType,	iInsRow, .hdnVATType.value)
			
			If Trim(.hdnVATINCFLG.value) = "2" Then	'포함 
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,0,"X","X")
				Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,0,"X","X")
			Else
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,1,"X","X")
				Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,1,"X","X")
			End If

			If .hdnVATType.value <> "" Then
				call SetVatType(iInsRow)
			End If

			Call .vspdData.SetText(C_VatRate	,	iInsRow, .hdnVATRate.value)
			Call SetSpreadValue(.vspdData,C_CostCon	,iInsRow,1,"X","X")
			Call SetSpreadValue(.vspdData,C_CostConCd	,iInsRow,1,"X","X")

			Call SetState("C",iInsRow)
			
			Call .vspdData.SetText(C_PlantCd	,	iInsRow, strRet(index1,C_PlantCd_Ref))
			Call .vspdData.SetText(C_PlantNm	,	iInsRow, strRet(index1,C_PlantNm_Ref))
			Call .vspdData.SetText(C_itemCd		,	iInsRow, strRet(index1,C_ItemCd_Ref))
			Call .vspdData.SetText(C_itemNm		,	iInsRow, strRet(index1,C_ItemNm_Ref))
			Call .vspdData.SetText(C_SpplSpec	,	iInsRow, strRet(index1,C_SpplSpec_Ref))
			Call .vspdData.SetText(C_OrderQty	,	iInsRow, strRet(index1,C_Qty_Ref))
			Call .vspdData.SetText(C_OrderUnit	,	iInsRow, strRet(index1,C_Unit_Ref))
			Call .vspdData.SetText(C_SoNo		,	iInsRow, strRet(index1,C_SoNo_Ref))
			Call .vspdData.SetText(C_SoSeqNo	,	iInsRow, strRet(index1,C_SoSeqNo_Ref))
			Call .vspdData.SetText(C_DlvyDT		,	iInsRow, strRet(index1,C_DlvyDt_Ref))
			Call .vspdData.SetText(C_SLCd		,	iInsRow, strRet(index1,C_SLCd_Ref))
			Call .vspdData.SetText(C_SLNm		,	iInsRow, strRet(index1,C_SLNm_Ref))
			Call .vspdData.SetText(C_HSCd		,	iInsRow, strRet(index1,C_HSCd_Ref))
			Call .vspdData.SetText(C_HSNm		,	iInsRow, strRet(index1,C_HSNm_Ref))
			Call .vspdData.SetText(C_PrNo		,	iInsRow, strRet(index1,C_ReqNo_Ref))
			Call .vspdData.SetText(C_TrackingNo	,	iInsRow, strRet(index1,C_TrackingNo_Ref))
		Else
			IntIFlg=True
		End if 
	next
	
	intEndRow = iInsRow
	Call SetSpreadColorRef(intStartRow,intEndRow)
	
	if strMessage<>"" then
		Call DisplayMsgBox("17a005", "X",strmessage,"구매요청번호")
		.vspdData.ReDraw = True
		Exit Function
	End if
	
	.vspdData.ReDraw = True
	
	End with
	
	Call ChangeItemPlant(intStartRow,intEndRow)
	
			
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
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
		arrParam(3) = "20!O"
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
		frm1.vspdData.Col = C_ItemCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_ItemNm
		frm1.vspdData.Text = arrRet(1)
		Call ChangeReturnCost()
	End If	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
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
		
		frm1.vspdData.Col = C_PlantCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_PlantNm
		frm1.vspdData.Text = arrRet(1)
		
		Call ChangeReturnCost()
	End If	
		
End Function

'------------------------------------------  OpenHS()  -------------------------------------------------
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
		frm1.vspdData.Col = C_HSCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_HSNm
		frm1.vspdData.Text = arrRet(1)
	End If	
	
End Function
'------------------------------------------  OpenUnit()  -------------------------------------------------
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
		frm1.vspdData.Col=C_OrderUnit
		frm1.vspdData.text= arrRet(0)	
		Call ChangeReturnCost()
	End If	
End Function

'------------------------------------------  OpenSL()  -------------------------------------------------
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
	
	arrParam(4) = "PLANT_CD= " & FilterVar(frm1.vspdData.Text, "''", "S") & ""

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
		frm1.vspdData.Col = C_SLCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_SLNm
		frm1.vspdData.Text = arrRet(1)
	End If	
End Function
'------------------------------------------  OpenVat()  -------------------------------------------------
Function OpenVat()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

'	If IsOpenPop = True Or UCase(frm1.txtVattype.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
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
		Call SetVat(arrRet)
	End If	
End Function
'------------------------------------------  OpenRet()  -------------------------------------------------
Function OpenRet()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

'	If IsOpenPop = True Or UCase(frm1.txtVattype.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
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
		frm1.vspdData.Col = C_RetCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_RetNm
		frm1.vspdData.Text = arrRet(1)
		Call vspdData_Change(C_RetCd, frm1.vspdData.ActiveRow) 
	End If	
End Function
'------------------------------------------  OpenTrackingNo()  -------------------------------------------
Function OpenTrackingNo()

	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	If Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		IsOpenPop = False
		Exit Function
	End if
    
    arrParam(0) = ""
    arrParam(1) = ""
    arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	
	frm1.vspdData.Col=C_SoNo
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	
	arrParam(4) = Trim(frm1.vspdData.Text)
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

'========================================== SetVat()  =============================================
Function SetVat(byval arrRet)	
	
    Dim price, chk_vat_flg
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
		.vspdData.Col		= C_IOFlgCd
		chk_vat_flg	= .vspdData.text
		
		.vspdData.Col = C_VatAmt 
		if chk_vat_flg = "2"		Then
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(price * UNICDbl(arrRet(2))/(100 + UNICDbl(arrRet(2))),frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
		Else
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(price * UNICDbl(arrRet(2))/(100 + UNICDbl(arrRet(2))),frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
		End If
		
	End With
    Call vspdData_Change(C_VatType, frm1.vspdData.ActiveRow)   
	
End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 
'==========================================================================================
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
        Case 6                                                              '과부족허용율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999.9999"
    End Select
         
End Sub
'============================================  2.5.1 TotalSum()  ======================================
'=	Name : TotalSum()																					=
'=	Description : Master L/C Header 화면으로부터 넘겨받은 parameter setting(Cookie 사용)				=
'========================================================================================================
	Sub TotalSum(ByVal row)
		
	    Dim SumTotal, lRow, tmpGrossAmt, tmpVatAmt,tmpamt
		SumTotal = 0
		ggoSpread.source = frm1.vspdData
		SumTotal = UNICDbl(frm1.txtGrossAmt.value)
		frm1.vspdData.Row = row
		frm1.vspdData.Col = C_NetAmt				
		tmpGrossAmt = UNICDbl(frm1.vspdData.Text)
		frm1.vspdData.Col = 0				
		frm1.vspdData.Col = C_OrgNetAmt							
		SumTotal = SumTotal + (tmpGrossAmt - UNICDbl(frm1.vspdData.Text))
        
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
				'Call ChangeItemPlantForUnit(IntRow,IntRow)
				Call ChangeItemPlantForUnit2(IntRow)
			else
				'Call ChangeItemPlant(IntRow,IntRow)
				Call ChangeItemPlant2(IntRow)
			end if
			
		End if
		
	End Sub
	
'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtGrossAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub

'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()
	
	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_Cost,-1, .txtCurr.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_OrderAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'VAT금액 
		ggoSpread.SSSetFloatByCellOfCur C_VatAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloatByCellOfCur C_OrgNetAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
        ggoSpread.SSSetFloatByCellOfCur C_OrgNetAmt1,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
	End With

End Sub	

'================================== =====================================================
' Function Name : InitCollectType
' Function Desc : 소비세유형코드/명/율 저장하기 
' 여기부터 키보드에서 소비세유형코드를 변경시 소비세유형명,소비세율,매입금액,NetAmount를 변경시키는 함수 
'========================================================================================

Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, VbInformation, parent.gLogoName
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
'========================================================================================
Sub SetVatType(byVal iRow)
	Dim VatType, VatTypeNm, VatRate 
	Dim txtVatRate ,txtVatAmt, chk_vat_flg
	     
	With frm1.vspdData
      
       .Row = iRow
	   .Col = C_VatType
	  
		VatType = .text
	
		Call InitCollectType
		Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)
        
       .Col = C_VatNm  
       .text = VatTypeNm
       .Col = C_VatRate
	   .text = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		txtVatRate =  UNICDbl(.text)


	   ' vat 금액계산  
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
		   
 End With
	   
End Sub
'========================================================================================
' Function Name : SetRetCd
' Function Desc : 반납유형 직접 입력시 처리 
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

'------------------------------------  Setreference()  ----------------------------------------------
'	Name : Setreference()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub Setreference()
    
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim ireference

    Err.Clear

	Call CommonQueryRs(" reference ", " b_configuration ", " major_cd = " & FilterVar("M9016", "''", "S") & " and minor_cd = " & FilterVar("CH", "''", "S") & " and seq_no = " & FilterVar("1", "''", "S") & "  ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    ireference = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, VbInformation, parent.gLogoName
		Err.Clear 
		Exit Sub
	End If

    if Trim(lgF0) <> "" then
        frm1.hdnreference.value = UCase(Trim(ireference(0)))
    end if

End Sub


'========================================================================================
' Function Name : setCVatFlg
' Function Desc : 부가세 포함에 따른 의제매입계산 처리 
' Append		: 2002-03-09  L.I.P
'========================================================================================
Sub setCVatFlg(byVal iRow)
	Call setVatType(iRow)
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    ADD CHEN, JAE HYUN - 2005-07-06
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function JumpOrderRun()

    Dim IntRetCd, strVal
    Dim lRow
    Dim ArrIssueCnt, ArrClsFlg, ArrRcptQty
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If frm1.vspdData.MaxRows < 1 Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	'발주확정이후에만 전개할 수 있음.
	If Trim(frm1.hdnReleaseflg.value) <> "Y"  Then
		Call DisplayMsgBox("170011", "x", "x", "x")
		Exit Function
	End If
	
	'외주가공 위탁 발주건에 대해서만 전개 가능함 
	If Trim(frm1.hdnSubContraFlg.value) <> "Y"  Then
		Call DisplayMsgBox("170012", "x", "x", "x")
		Exit Function
	End If
	
	lRow = frm1.vspdData.ActiveRow

	'사급품 출고 여부 조회 
	Call CommonQueryRs(" COUNT(C.ISSUE_QTY) ", " M_PUR_ORD_DTL A, M_CHILD_RESERV_HISTORY B, M_CHILD_RESERV C ", _
		" A.PO_NO = B.PAR_PO_NO AND A.PO_SEQ_NO = B.PAR_PO_SEQ_NO AND B.PR_NO = C.PR_NO AND B.RESVD_SEQ_NO = C.RESVD_SEQ_NO AND C.ISSUE_QTY > 0 " & _
		" AND A.PO_NO =" & FilterVar(frm1.txthdnPoNo.Value, "''", "S") & _
		" AND A.PO_SEQ_NO = " & FilterVar(GetSpreadText(frm1.vspdData,C_SeqNo,lRow,"X","X"), "''", "S")  _
		, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
	If Err.number <> 0 Then
		MsgBox Err.description, VbInformation, parent.gLogoName
		Err.Clear 
		Exit Function
	End If
	
	'사급품을 출고한 것은 전개할 수 없음.
	If InStr(Trim(lgF0), Chr(11)) > 0 Then
		ArrIssueCnt = Split(lgF0, Chr(11))
		If Trim(ArrIssueCnt(0)) > "0" Then
			'170013
			Call DisplayMsgBox("170013",  "X", "X", "X")
			Exit Function
		End If	
	End If	
	
	
	'마감 및 입고 건 조회 
	Call CommonQueryRs(" CLS_FLG, RCPT_QTY ", " M_PUR_ORD_DTL ", _
		" PO_NO =" & FilterVar(frm1.txthdnPoNo.Value, "''", "S") & _
		" AND PO_SEQ_NO = " & FilterVar(GetSpreadText(frm1.vspdData,C_SeqNo,lRow,"X","X"), "''", "S")  _
		, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
	If Err.number <> 0 Then
		MsgBox Err.description, VbInformation, parent.gLogoName
		Err.Clear 
		Exit Function
	End If
	
	'마감된 건은 전개할 수 없음	
	If InStr(Trim(lgF0), Chr(11)) > 0 Then
		ArrClsFlg = Split(lgF0, Chr(11))
		If Trim(ArrIssueCnt(0)) = "Y" Then
			'179033
			Call DisplayMsgBox("179033",  "X", "X", "X")
			Exit Function
		End If
	Else
		Call DisplayMsgBox("173210",  "X", "X", "X")
		Exit Function	
	End If	
	
	'입고된 건은 전개할 수 없음 
	If InStr(Trim(lgF1), Chr(11)) > 0 Then
		ArrRcptQty = Split(lgF1, Chr(11))
		If Trim(ArrRcptQty(0)) > "0" Then
			'170014
			Call DisplayMsgBox("170014",  "X", "X", "X")
			Exit Function
		End If
	End If	
	
	With frm1	
		WriteCookie "txtPlantCd", UCase(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X"))
		WriteCookie "txtPlantNm", GetSpreadText(.vspdData,C_PlantNm,lRow,"X","X")
		WriteCookie "txtItemCd", UCase(GetSpreadText(.vspdData,C_itemCd,lRow,"X","X"))
		WriteCookie "txtItemNm", GetSpreadText(.vspdData,C_itemNm,lRow,"X","X")
		WriteCookie "txtSpecification", GetSpreadText(.vspdData,C_SpplSpec,lRow,"X","X")
		
		WriteCookie "txtPoNo", UCase(Trim(frm1.txthdnPoNo.value))
		WriteCookie "txtPoSeqNo",GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")
		WriteCookie "txtOrderQty", GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X")
		WriteCookie "txtOrderUnit", GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X")
		WriteCookie "txtDlvyDt", GetSpreadText(.vspdData,C_DlvyDT,lRow,"X","X")
		WriteCookie "txtPGMID", "m3112ma1"
	End With	
		
	navigate BIZ_PGM_JUMPORDERRUN_ID	
	
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )   
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		if Trim(frm1.hdnReleaseflg.Value) = "N" then
			Call SetPopupMenuItemInf("1101111111")
		else
			Call SetPopupMenuItemInf("0000111111")
		end if
	
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
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    		
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If
	
End Sub

'========================================================================================
' Function Name : FncSplitColumn
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
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call CurFormatNumSprSheet() 
    Call ggoSpread.ReOrderingSpreadData()
    Call SetSpreadLockAfterQuery()
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim strsstemp1,strsstemp2,strsstemp3
	Dim Qty, Price, DocAmt, VatAmt, VatRate, chk_vat_flg,orgNetAmt
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iNameArr, strPlantCd

	
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

		' OnChange Event 정리 (2005-11-16 by K.D.H)
		Select Case Col
			'공장 
			Case C_PlantCd			
				.Col	= C_ItemCd
				.text 	= ""
				
				.Col 	= C_ItemNM
				.text 	= ""
			'품목 
			Case C_ItemCd			
				.Col 		= C_ItemCd
				strssTemp1 	= Trim(.Text)
				.Col 		= C_PlantCd
				strssTemp2 	= Trim(.Text)
				
				If strssTemp1 = "" Or strssTemp2 = "" then
					ggoSpread.spreadlock C_TrackingNo, Row, C_TrackingNoPop, Row
					.Col 	= C_TrackingNo
					.Text   = ""
					Exit Sub
				End if

				'Call ChangeItemPlant(Row,Row)
				Call ChangeItemPlant2(Row)
			'발주수량, 단가 
			Case C_OrderQty, C_Cost',C_VatRate
				.Col = C_OrderQty
				If Trim(.Text) = "" Or IsNull(.Text) then
					Qty = 0
				Else
					Qty = UNICDbl(.Text)
				End If
				
				.Col = C_Cost
				If Trim(.Text) = "" Or IsNull(.Text) then
					Price = 0
				Else
					Price = UNICDbl(.Text)
				End If
				
				DocAmt 	= Qty * Price
				.Col 	= C_OrderAmt		
				.Text 	= UNIConvNumPCToCompanyByCurrency(DocAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo,"X","X")
				
				Call InitData(Row)
		
		        Call vspdData_Change(C_OrderAmt, Row)
		    ' 단위    
		    Case C_OrderUnit					
'				.Col 		= C_ItemCd
'				strssTemp1 	= Trim(.Text)
'				.Col 		= C_PlantCd
'				strssTemp2 	= Trim(.Text)
'				.Col 		= C_OrderUnit
'				strssTemp3 	= Trim(.Text)
'				
'				If strssTemp1 = "" Or strssTemp2 = "" Or strssTemp3 = "" Then
'					Exit Sub
'				End if

				'Call ChangeItemPlantForUnit(Row, Row)
				Call  ChangeItemPlantForUnit2(Row)
				
			' 단가구분 
			Case C_CostCon
				Call vspdData_ComboSelChange(C_CostCon, Row)	' Line 복사시 SelChange를 강제로 일어나게 한다.
			' 금액 
			Case C_OrderAmt
				.Col = C_OrderAmt
			    DocAmt = UNICDbl(.Text)
			     
			    'VAT 금액 추가  -->
				.Col = C_VatRate ' VAT 율 
				If Trim(.Text) = "" OR IsNull(.Text) then
					VatRate = 0
				Else
					VatRate = UNICDbl(.Text)
				End If
		
				' 부가세 포함/불포함 부가세 계산 변경 2002.3.9 L.I.P
				.Col = C_IOFlgCD
				chk_vat_flg	= .text

				if chk_vat_flg = "2"	Then	'포함 
					VatAmt    = DocAmt * (VatRate/(100+VatRate))
				Else                            '별도 
					VatAmt    = DocAmt * (VatRate/100)
				End If
		
				.Col = C_VatAmt
				.Text = UNIConvNumPCToCompanyByCurrency(VatAmt,frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")		
				VatAmt = UNICDbl(.Text)
				'VAT금액부터 계산후 발주순금액을 계한한다.(금액-VAT금액(함수 적용한 금액))
				if chk_vat_flg = "2"	Then	'포함 
					.Col = C_NetAmt		
				    .Text = UNIConvNumPCToCompanyByCurrency(DocAmt - VatAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo,"X","X")
				    orgNetAmt = .Text
				Else                            '별도 
					.Col = C_NetAmt		
				    .Text = UNIConvNumPCToCompanyByCurrency(DocAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo,"X","X")
				    orgNetAmt = .Text
				End If
						
				'<-- VAT 금액 추가 2002.2.18 L.I.P
				Call TotalSum(Row)					'총품목금액합계 
						
				.Col = C_OrgNetAmt		
				.Text = orgNetAmt
			' VAT포함여부 
			Case C_IOFlg
				.Col = C_IOFlg
				Call vspdData_ComboSelChange(C_IOFlg, Row)	' Line 복사시 SelChange를 강제로 일어나게 한다.
				Call vspdData_Change(C_OrderAmt, Row)
				Call setCVatFlg(Row)	
			' VAT
			Case C_VatType 'or Col = C_VatAmt then
				Call SetVatType(Row)     ' C_VatNm,C_VatRate 세팅 
				call vspdData_Change(C_OrderAmt, Row)
			' 납기일 
			Case C_DlvyDt
				.Col = C_DlvyDt
				strsstemp1 = .Text
				if strsstemp1 = "" then Exit Sub
				strsstemp2 = frm1.txtPoDt.text
				if UniConvDateToYYYYMMDD(strsstemp2,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(strsstemp1,Parent.gDateFormat,"") then
					Call DisplayMsgBox("970023", "X", "납기일", frm1.txtPoDt.Alt)
				end if
			' HS부호 
			Case C_HSCd
    			Err.Clear
				
				.Col = C_HSCd
				Call CommonQueryRs(" HS_NM ", " B_HS_CODE ", " HS_CD = " & FilterVar(.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

				If Err.number <> 0 Then
					MsgBox Err.description, VbInformation, parent.gLogoName
					Err.Clear 
					Exit Sub
				End If

				.Col = C_HSNm
				If Len(lgF0) > 0 Then
					iNameArr = Split(lgF0, Chr(11))
					.Text = iNameArr(0)
				Else
					.Text = ""
					.Col = C_HSCd
					Call DisplayMsgBox("203227", "X", .Text, "X")
					.Text = ""
				End If
			' 창고 
			Case C_SLCd
    			Err.Clear
				.Col = C_PlantCd
				strPlantCd = .Text
				.Col = C_SLCd
				Call CommonQueryRs(" SL_NM ", " B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(.Text, "''", "S") & " AND PLANT_CD = " & FilterVar(strPlantCd, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

				If Err.number <> 0 Then
					MsgBox Err.description, VbInformation, parent.gLogoName
					Err.Clear 
					Exit Sub
				End If

				.Col = C_SLNm
				If Len(lgF0) > 0 Then
					iNameArr = Split(lgF0, Chr(11))
					.Text = iNameArr(0)
				Else
					.Text = ""
					.Col = C_SLCd
					.Text = ""
					Call DisplayMsgBox("169922", "X", "X", "X")
				End If
			' 반품유형 
			Case C_RetCd
				Call SetRetCd()
		End Select
    End With
      
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 
	
End Sub
'=============== vspdData_ComboSelChange() ==================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex 
	With frm1.vspdData
	
		.Row = Row
		.Col = Col

		If Col = C_CostCon Then 
			intIndex = .Value
			.Col = C_CostCon+1
			.Value = intIndex
		Else  
		    intIndex = .Value
			.Col = C_IOFlg+1
			.Value = intIndex
        End If
  End With
 
End Sub
'================ vspdData_ButtonClicked() ================================================
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
		Case C_Check
			Call lookupPrice(Row)
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
		End Select
        
    End If
    
    End With
End Sub
'================ vspdData_LeaveCell() ==========================================================
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
'================ vspdData_TopLeftChange() ==========================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgStrPrevKey <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub


'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData(lRow)
	Dim intIndex 

		frm1.vspdData.Row = lRow

		frm1.vspdData.Col = C_CostConCd
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_CostCon
		frm1.vspdData.value = intindex
End Sub



'================== FncQuery() ===========================================================
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
    'Call ggoOper.ClearField(Document, "2")	'This Function is called if error occurred at Serverside Script(MB)
    
    For intIndex = 1 to frm1.vspdData.MaxCols 
		frm1.vspdData.SetColItemData intindex,0
	Next
	
    ggoSpread.ClearSpreadData
       
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkFieldByCell(frm1.txtPoNo, "A",1)	then
       Exit Function
    End If
    
    '-----------------------
    'Check length of field - Unnecessary
    '-----------------------
    'If Not chkFieldLengthByCell(frm1.txtPoNo, "A",1) Then		
    '   Exit Function
    'End If
   
    '-----------------------
    'Query function call area
    '-----------------------
    frm1.txtQuerytype.value = "Query"
    
    If DbQuery = False Then Exit Function
       
    FncQuery = True	
    Set gActiveElement = document.activeElement
    
End Function
'================== FncNew() ===========================================================
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
    
    Call ggoOper.ClearField(Document, "1") 
    Call ggoOper.ClearField(Document, "2") 
    Call ggoOper.LockField(Document, "N")  
    Call SetDefaultVal()
    Call InitVariables                     
    FncNew = True    
    Set gActiveElement = document.activeElement                      

End Function

'================== FncDelete() ===========================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                      
    
    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then Exit Function
    														
    Err.Clear                              
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then             
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    Call ggoOper.ClearField(Document, "1")         
    Call ggoOper.ClearField(Document, "2")         
    
    FncDelete = True 
    Set gActiveElement = document.activeElement                              
    
End Function

'================== FncSave() ===========================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                
       
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                             
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If

	If Not chkField(Document, "2") OR Not ggoSpread.SSDefaultCheck Then		
		Exit Function
	End If
   
    If DbSave = False Then Exit Function
    
    FncSave = True  
    Set gActiveElement = document.activeElement                                                        
    
End Function
'================== FncCopy() ===========================================================
Function FncCopy() 
    Dim SumTotal,tmpGrossAmt
    if frm1.vspdData.Maxrows < 1	then exit function
    ggoSpread.Source = frm1.vspdData	
    
    ggoSpread.CopyRow
    
    frm1.vspdData.ReDraw = False
    
    Call SetSpreadColor(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)
    
    frm1.vspdData.ReDraw = True
    
    Call SetState("C",frm1.vspdData.ActiveRow)
    
    '복사한 것은 긴급발주로 간주.
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
    frm1.vspdData.Col = C_TrackingNo

 	if Trim(frm1.vspdData.Text) = "*" then
		ggoSpread.spreadlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
	else
	    ggoSpread.spreadUnlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
		ggoSpread.sssetrequired C_TrackingNo, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	end if
	    
    frm1.vspdData.ReDraw = True
   
   '총품목금액합계 
    SumTotal = UNICDbl(frm1.txtGrossAmt.value)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_NetAmt				
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)
    SumTotal = SumTotal + tmpGrossAmt
    frm1.txtGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")     
    
    Set gActiveElement = document.activeElement
End Function

'================== FncCancel() ===========================================================
Function FncCancel() 
	Dim maxrow,maxrow1,SumTotal,tmpGrossAmt,index,index1,orgtmpGrossAmt
	Dim starindex ,endindex,delflag
	if frm1.vspdData.Maxrows < 1	then exit function
	maxrow = frm1.vspdData.Maxrows
	index1 = 0
	
	starindex = frm1.vspdData.SelBlockRow
	endindex  = frm1.vspdData.SelBlockRow2
    
    Redim orgtmpGrossAmt(endindex - starindex)
    Redim chgtmpGrossAmt(endindex - starindex)
    Redim tmpGrossAmt(endindex - starindex)
    Redim delflag(endindex - starindex)
    SumTotal = UNICDbl(frm1.txtGrossAmt.value)
	
	for index = starindex to endindex
	    frm1.vspdData.Row = index
	    frm1.vspdData.Col = C_NetAmt
	    tmpGrossAmt(index1) = UNICDbl(frm1.vspdData.Text)
	    
	    frm1.vspdData.Col = C_OrgNetAmt1
	    orgtmpGrossAmt(index1) = UNICDbl(frm1.vspdData.Text)
	    
	    frm1.vspdData.Col = 0
	    delflag(index1) = frm1.vspdData.Text
	    index1 = index1 + 1
	next
		
	    ggoSpread.Source = frm1.vspdData
		index = frm1.vspdData.ActiveRow - starindex
		
    '//for index = 0 to index1 - 1
        if delflag(index) = ggoSpread.UpdateFlag then
            SumTotal = SumTotal + (orgtmpGrossAmt(index) - tmpGrossAmt(index) )
        elseif  delflag(index) = ggoSpread.DeleteFlag then
            SumTotal = SumTotal + orgtmpGrossAmt(index)
        elseif delflag(index) = ggoSpread.InsertFlag  then
            SumTotal = SumTotal - tmpGrossAmt(index)
        end if
    '//Next   

        ggoSpread.EditUndo                                     
        maxrow1 = frm1.vspdData.Maxrows

    frm1.txtGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")     

	' 취소시 행이 없으면 비활성화 
	if frm1.vspdData.Maxrows < 1 then
	    frm1.btnCallPrice.disabled = true
	    ' --- 2005.07.18 단가 일괄 불러오기 관련 수정 -----------------------------------
	end if

    Set gActiveElement = document.activeElement
    
End Function

'================== FncInsertRow() ===========================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
    Dim IntRetCD
    Dim imRow
    Dim inti
    inti=1
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then Exit Function
	End If
    
	With frm1
        
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        
        For inti= .vspdData.ActiveRow  to .vspdData.ActiveRow +imRow-1
			.Row=inti
			ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
			ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
			ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
			ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd
			Call SetState("C",inti)
    
			'공장 기본값 추가 
			Call .vspdData.SetText(C_PlantCd,	inti, Parent.gPlant)
			Call .vspdData.SetText(C_OrderAmt,	inti, "0")
			Call .vspdData.SetText(C_Cost,		inti, "0")
			Call .vspdData.SetText(C_DlvyDT,	inti, .hdnDlvydt.value)
			Call .vspdData.SetText(C_VatType,	inti, .hdnVATType.value)
			
			If Trim(.hdnVATINCFLG.value) = "2" Then	'포함 
				Call .vspdData.SetText(C_IOFlg,		inti, 0)
				Call .vspdData.SetText(C_IOFlgCd,	inti, 0)
			Else
				Call .vspdData.SetText(C_IOFlg,		inti, 1)
				Call .vspdData.SetText(C_IOFlgCd,	inti, 1)
			End If
						

			If .hdnVATType.value <> "" Then
				Call SetVatType(inti)
			End If

			Call .vspdData.SetText(C_VatRate,		inti, .hdnVATRate.value)
			Call .vspdData.SetText(C_TrackingNo,	inti, "*")
			
			'---------------------------------------------------------
			'ggoSpread.spreadUnlock	C_PlantCd,.vspdData.Row,C_PlantCd,.vspdData.Row
			'ggoSpread.sssetrequired	C_PlantCd,.vspdData.Row,.vspdData.Row
			'ggoSpread.spreadUnlock	C_Popup1,.vspdData.Row,C_Popup1,.vspdData.Row
			'ggoSpread.spreadUnlock	C_ItemCd,.vspdData.Row,C_ItemCd,.vspdData.Row
			'ggoSpread.sssetrequired	C_ItemCd,.vspdData.Row,.vspdData.Row
			'ggoSpread.spreadUnlock	C_Popup2,.vspdData.Row,C_Popup2,.vspdData.Row
			'ggoSpread.spreadUnlock	C_IOFlg,.vspdData.Row,C_IOFlg,.vspdData.Row
			'ggoSpread.sssetrequired	C_IOFlg,.vspdData.Row,.vspdData.Row
	
			'If .hdnImportflg.value = "Y" Then
			'	ggoSpread.spreadUnlock	C_HsCd	,.vspdData.Row,C_Popup5	,.vspdData.Row
			'	ggoSpread.spreadUnlock	C_Popup5,.vspdData.Row,C_Popup5	,.vspdData.Row
			'	ggoSpread.sssetrequired	C_HsCd	,.vspdData.Row,.vspdData.Row
			'End If
			
			Call .vspdData.SetText(C_CostCon,	inti, 1)
			Call .vspdData.SetText(C_CostConCd,	inti, 1)
		Next
		
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		'---------------------------------------------------------
		ggoSpread.spreadUnlock	C_PlantCd,.vspdData.ActiveRow,C_PlantCd,.vspdData.ActiveRow + imRow - 1
		ggoSpread.sssetrequired	C_PlantCd,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
		ggoSpread.spreadUnlock	C_Popup1,.vspdData.ActiveRow,C_Popup1,.vspdData.ActiveRow + imRow - 1
		ggoSpread.spreadUnlock	C_ItemCd,.vspdData.ActiveRow,C_ItemCd,.vspdData.ActiveRow + imRow - 1
		ggoSpread.sssetrequired	C_ItemCd,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
		ggoSpread.spreadUnlock	C_Popup2,.vspdData.ActiveRow,C_Popup2,.vspdData.ActiveRow + imRow - 1
		ggoSpread.spreadUnlock	C_IOFlg,.vspdData.ActiveRow,C_IOFlg,.vspdData.ActiveRow + imRow - 1
		ggoSpread.sssetrequired	C_IOFlg,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
	
		If .hdnImportflg.value = "Y" Then
			ggoSpread.spreadUnlock	C_HsCd	,.vspdData.ActiveRow,C_Popup5	,.vspdData.ActiveRow + imRow - 1
			ggoSpread.spreadUnlock	C_Popup5,.vspdData.ActiveRow,C_Popup5	,.vspdData.ActiveRow + imRow - 1
			ggoSpread.sssetrequired	C_HsCd	,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
		End If
			
        .vspdData.ReDraw = True
    End With
    
    ' 행추가시 단가불러오기 활성화 
    frm1.btnCallPrice.disabled = False

	If Err.number = 0 Then FncInsertRow = True                                                          '☜: Processing is OK
    
    Set gActiveElement = document.ActiveElement   
        
End Function

'================== FncDeleteRow() ===========================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    Dim index,SumTotal,idel
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow

		'.Col = C_Stateflg:	.Row = .ActiveRow
		'if Trim(.Text) = "" then
		'	.Text = "D"
		'End if
		SumTotal = UNICDbl(frm1.txtGrossAmt.value)
		for index = .SelBlockRow to .SelBlockRow2
			.Row = index
			.Col = C_Stateflg
			idel = .text
			.Col = 0

			if Trim(.text) <> ggoSpread.InsertFlag and Trim(idel) <> "D" then
			    .Col = C_NetAmt							
		         SumTotal = SumTotal - UNICDbl(.Text)
                 .Col = C_Stateflg
			     frm1.vspdData.text = "D"
                 frm1.txtGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")     
		    end if
		Next
    End With
    
    Set gActiveElement = document.activeElement
End Function

'================== FncPrint() ===========================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'================== FncPrev() ===========================================================
Function FncPrev() 
    On Error Resume Next 
    Set gActiveElement = document.activeElement                                  
End Function
'================== FncNext() ===========================================================
Function FncNext() 
    On Error Resume Next    
    Set gActiveElement = document.activeElement                               
End Function
'================== FncExcel() ===========================================================
Function FncExcel() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_MULTI)	
    Set gActiveElement = document.activeElement						
End Function
'================== FncFind() ===========================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False)   
    Set gActiveElement = document.activeElement                 
End Function
'================== FncExit() ===========================================================
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
    Set gActiveElement = document.activeElement
    
End Function

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
    
    'If CheckRunningBizProcess = True Then
    '   Exit Function
	'End If        	                                      
	'Call LayerShowHide(1)
    
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

    lgIntFlgMode = Parent.OPMD_UMODE						
    'Call ggoOper.LockField(Document, "Q")	
	Call SetSpreadLockAfterQuery

	Call ToolBarCtrl()

	if Trim(UCase(frm1.hdnReleaseflg.Value)) = "Y" then
		if frm1.hdnClsflg.value = "Y" or frm1.vspdData.Maxrows < 1 then
		    frm1.btnCfmSel.disabled = true		    
		else
		    frm1.btnCfmSel.disabled = False
		end if
		frm1.btnCfm.value = "확정취소"
		' --- 2005.07.18 단가 일괄 불러오기 관련 수정 -----------------------------------
		frm1.btnCallPrice.disabled = True
		' --- 2005.07.18 단가 일괄 불러오기 관련 수정 -----------------------------------
	else
		frm1.btnCfm.value = "확정"
		if frm1.hdnClsflg.value = "Y" or frm1.vspdData.Maxrows < 1 then
		    frm1.btnCfmSel.disabled = true
		    ' --- 2005.07.18 단가 일괄 불러오기 관련 수정 -----------------------------------
		    frm1.btnCallPrice.disabled = true
		    ' --- 2005.07.18 단가 일괄 불러오기 관련 수정 -----------------------------------
		else
		    frm1.btnCfmSel.disabled = False
		    ' --- 2005.07.18 단가 일괄 불러오기 관련 수정 -----------------------------------
		    frm1.btnCallPrice.disabled = False
		    ' --- 2005.07.18 단가 일괄 불러오기 관련 수정 -----------------------------------
		end if
	end if
	
    '=================================================================
    '스프레드에 반품시 반품유형 입력 가능하게 보여줌 반품이 아닐 경우는 안보여주도록 함 2002-02-22
    '=================================================================
     
    if frm1.hdnRetflg.Value = "Y" then
		frm1.vspdData.Col = C_RetCd:		frm1.vspdData.ColHidden = false
		frm1.vspdData.Col = C_Popup8:		frm1.vspdData.ColHidden = false
		frm1.vspdData.Col = C_RetNm:		frm1.vspdData.ColHidden = false
	else
		frm1.vspdData.Col = C_RetCd:		frm1.vspdData.ColHidden = true
		frm1.vspdData.Col = C_Popup8:		frm1.vspdData.ColHidden = true
		frm1.vspdData.Col = C_RetNm:		frm1.vspdData.ColHidden = true	
	end if
    '=================================================================
    frm1.vspdData.focus
	Set gActiveElement = document.activeElement
	
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'================== DbSave() ===========================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal,strDel
	Dim ColSep, RowSep
	Dim iOrderQty
	Dim iCost
	Dim iOrderAmt

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
	
	ColSep = Parent.gColSep															
	RowSep = Parent.gRowSep															

	With frm1
		.txtMode.value = Parent.UID_M0002
		
		lGrpCnt = 0
    
		strVal = ""
		strDel = ""
		iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

		ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
		ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

		iTmpCUBufferCount = -1
		iTmpDBufferCount = -1

		strCUTotalvalLen = 0
		strDTotalvalLen  = 0

    For lRow = 1 To .vspdData.MaxRows step 1
        Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		     Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag	
		     
				if Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))=ggoSpread.InsertFlag then
					strVal = "C" & ColSep
				Else
					strVal = "U" & ColSep
				End if      
				
				If Trim(UNICDbl(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))) = "" Or Trim(UNICDbl(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))) = "0" then
					Call DisplayMsgBox("970021", "X","발주수량", "X")
					.vspdData.Row = lRow
					.vspdData.Action = 0
					Call LayerShowHide(0)
					Exit Function
				End if
					
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")) & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_itemCd,lRow,"X","X")) & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X")),0)  & ColSep
				End If
                   
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X"))  & ColSep
                   
                If Trim(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Cost,lRow,"X","X")),0)  & ColSep
				End If
                    
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_CostConCd,lRow,"X","X"))  & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_OrderAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderAmt,lRow,"X","X")),0)  & ColSep
				End If

                strVal = strVal & Trim(GetSpreadText(.vspdData,C_IOFlgCd,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_VatType,lRow,"X","X"))  & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_VatRate,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_VatRate,lRow,"X","X")),0)  & ColSep
				End If
                   
                If Trim(GetSpreadText(.vspdData,C_VatAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_VatAmt,lRow,"X","X")),0)   & ColSep
				End If
                   
                strVal = strVal & UNIConvDate(Trim(GetSpreadText(.vspdData,C_DlvyDT,lRow,"X","X")))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_HSCd,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_SLCd,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_TrackingNo,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Lot_No,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Lot_Seq,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_RetCd,lRow,"X","X"))  & ColSep
                    
                If Trim(GetSpreadText(.vspdData,C_Over,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Over,lRow,"X","X")),0)   & ColSep
				End If
					      
                If Trim(GetSpreadText(.vspdData,C_Under,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Under,lRow,"X","X")),0)  & ColSep
				End If

                strVal = strVal & Trim(GetSpreadText(.vspdData,C_PrNo,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_MvmtNo,lRow,"X","X"))  & ColSep
                '비고 추가 
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Remrk,lRow,"X","X"))  & ColSep

                strVal = strVal & Trim(GetSpreadText(.vspdData,C_MaintSeq,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_SoNo,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Stateflg,lRow,"X","X"))  & ColSep
	                
	                '반품등록 추가 C_IVNO,C_IVSEQ  27,28
                strVal = strVal & "" & ColSep 'IV No.
                strVal = strVal & "" & ColSep 'IV Seq.

                If Trim(GetSpreadText(.vspdData,C_NetAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_NetAmt,lRow,"X","X")),0)   & ColSep
				End If
					
                iOrderQty=UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X")),0)
                iCost=UNIConvNum(Trim(GetSpreadText(.vspdData,C_Cost,lRow,"X","X")),0)
                iOrderAmt=UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderAmt,lRow,"X","X")),0)

				If UNIConvNum(UNIConvNumPCToCompanyByCurrency(iOrderQty*iCost,frm1.txtCurr.value, Parent.ggAmtOfMoneyNo,"X","X"),0) = iOrderAmt Then
					strVal = strVal & "N" & ColSep
				Else
					strVal = strVal & "Y" & ColSep
				End If

                strVal = strVal & lRow & RowSep

            Case ggoSpread.DeleteFlag
            
		        strDel = "D" & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")) & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_PrNo,lRow,"X","X")) & ColSep				
				strDel = strDel & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep
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
    End With

	frm1.txtMaxRows.value = lGrpCnt
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

	If LayerShowHide(1) = False Then Exit Function
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			
    DbSave = True                                           
    
End Function
'================== DbSaveOk() ===========================================================
Function DbSaveOk()		
									
	Call InitVariables	
	Call MainQuery()	

End Function
'================== DbDelete() ===========================================================
Function DbDelete() 
End Function

'========================================================================================
' Function Name : initFormatField()
' Function Desc : Manual Formatting fields as amount or date 
'========================================================================================
Function  initFormatField()
	
	call FormatDateField(frm1.txtPoDt)		
	call FormatDoubleSingleField(frm1.txtGrossAmt)
	
	'call LockHtmlField(frm1.txtPoNo,"R")	
	'call LockHtmlField(frm1.txtPoTypeCd,"P")
	'call LockHtmlField(frm1.txtPoTypeNm,"P")	
	call LockobjectField(frm1.txtPoDt,"P")
	call LockobjectField(frm1.txtGrossAmt,"P")	
	'call LockHtmlField(frm1.txtSupplierCd,"P")
	'call LockHtmlField(frm1.txtSupplierNm,"P")
	'call LockHtmlField(frm1.txtGroupCd,"P")
	'call LockHtmlField(frm1.txtGroupNm,"P")
	'call LockHtmlField(frm1.txtCurr,"P")
	'call LockHtmlField(frm1.txtCurrNm,"P")
		
End Function                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         


' === 2005.07.15 단가 일괄 불러오기 관련 수정 ===========================================
Sub btnCallPrice_OnClick()
	Dim index
	
	If frm1.vspdData.Maxrows <= 0 then
		Exit Sub
	End if
	
'	If Trim(frm1.txtSupplierCd.value) = "" then
'		Call DisplayMsgBox("SCM003","X","X","X")
'		Call LayerShowHide(0)
'		frm1.txtSupplierCd.focus
'		Exit Sub
'	End If
'	
'	If Trim(frm1.txtCurr.value) = "" then
'		Call DisplayMsgBox("SCM003","X","X","X")
'		Call LayerShowHide(0)
'		frm1.txtCurr.focus
'		Exit Sub
'	End If	
'	
'	Call SetPriceType2
	Call lookupPriceForSelection()
	
	For index = 1 to  frm1.vspdData.Maxrows
'	    frm1.vspdData.row = index
'	    frm1.vspdData.Col = C_SelCheck
'	    
'	    If frm1.vspdData.Text = "1" then
'			frm1.vspdData.Col = 0
			ggoSpread.UpdateRow index
'	    Else
'			'frm1.vspdData.Col = 0
'			'ggoSpread.EditUndo
'	    End If	    
	Next 
	
End Sub

Sub btnCallPrice_Ok()
Dim lRow	
	With frm1
	For lRow = 1 To .vspdData.MaxRows				
'		.vspdData.Row = lRow
'		.vspdData.Col = C_Check
	
'		If .vspdData.Text <> "0" Then
			Call vspdData_Change(C_Cost, lRow)
'		End If	
	Next
	End With
End Sub

Sub SetPriceType()
	Dim IntRetCd, lsPriceType
	
	IntRetCD = CommonQueryRs("MINOR_CD", "B_CONFIGURATION", "(MAJOR_CD = 'M0001' AND REFERENCE = 'Y' )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	lsPriceType = TRIM(REPLACE(lgF0,CHR(11),""))
	
	frm1.hdnPriceType.value = lsPriceType			'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)
	
End Sub

Sub SetPriceType2()
	If frm1.rdoPrcTypeflg1.checked = true then
		lsPriceType = "T"
		frm1.hdnPriceType.value = "T"				'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)
	Else
		lsPriceType = "N"
		frm1.hdnPriceType.value = "N"				'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)
	End if
	
End Sub


Function lookupPriceForSelection()
    Err.Clear
    Dim strVal
    Dim lColSep,lRowSep
    Dim lRow        
    Dim lGrpCnt     
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If
	
	If Not chkField(Document, "2") Then
		Exit Function
	End If

'	If Not chkField(Document, "2") OR Not ggoSpread.SSDefaultCheck Then
'		Exit Function
'	End If
	
	lgBlnFlgChgValue = true
    
    If LayerShowHide(1) = False Then Exit Function

	With frm1		
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    strVal = ""
    
    '-----------------------
    'Data manipulate area
    '-----------------------
	.txtMode.value = "lookupPriceForSelection"	

	For lRow = 1 To .vspdData.MaxRows
				
		.vspdData.Row = lRow
'		.vspdData.Col = C_Check
	
'		If .vspdData.Text <> "0" Then
					
			frm1.vspdData.Row = lRow
			
			strVal = strVal & Trim(frm1.txtSupplierCd.Value) & parent.gColSep			
			frm1.vspdData.Col = C_ItemCd
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			frm1.vspdData.Col = C_PlantCd
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			frm1.vspdData.Col = C_OrderUnit
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			strVal = strVal & Trim(frm1.txtCurr.Value) & parent.gColSep & parent.gColSep		
'			frm1.vspdData.Col = C_PoPrice1
'			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			strVal = strVal & lRow & Parent.gRowSep
					
			lGrpCnt = lGrpCnt + 1

			frm1.vspdData.Col = C_Cost
			frm1.vspdData.Text = 0
'		End If
	Next
	
	If strVal <> "" Then
		If LayerShowHide(1) = False Then Exit Function
		
'		.hdnMaxRows.value = .vspdData.MaxRows
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End If	
	End With
End Function 
		
' === 2005.07.15 단가 일괄 불러오기 관련 수정 ===========================================


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
    strVal = strVal & "&txtStampDt=" & Trim(frm1.txtPoDt.text)
    strVal = strVal & "&txtBpCd=" & Trim(frm1.txtSupplierCd.Value)
	frm1.vspdData.Col = C_itemCd
    strVal = strVal & "&txtItemCd=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_PlantCd
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_OrderUnit
    strVal = strVal & "&txtUnit=" & Trim(frm1.vspdData.text)
    strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurr.value)
    strVal = strVal & "&txtRow=" & Row

    If LayerShowHide(1) = False Then Exit Function

	Call RunMyBizASP(MyBizASP, strVal)

End Function
