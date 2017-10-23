<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--
'**********************************************************************************************
'*  1. Module Name          : Long-term Inv Analysis
'*  2. Function Name        : 
'*  3. Program ID           : I3111QB1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : PI3G130
'*  7. Modified date(First) : 2006/05/25
'*  8. Modified date(Last)  : 2006/09/01
'*  9. Modifier (First)     : KiHong Han
'* 10. Modifier (Last)      : LEE SEUNG WOOK
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "I", "NOCOOKIE","QB")
								
On Error Resume Next
Call HideStatusWnd

Const C_SHEETMAXROWS_D = 100

Const E1_item_cd = 0
Const E1_item_nm = 1
Const E1_spec = 2
Const E1_abc_flag = 3
Const E1_unit = 4
Const E1_inv_price = 5
Const E1_storage_period = 6
Const E1_last_issue_dt = 7
Const E1_PerniciousStockPeriod = 8
Const E1_pernicious_stock_qty = 9
Const E1_pernicious_stock_amt = 10
Const E1_LongtermStockPeriod = 11
Const E1_longterm_stock_qty = 12
Const E1_longterm_stock_amt = 13

Const E2_item_group_cd = 0
Const E2_item_group_nm = 1
Const E2_pernicious_stock_qty = 2
Const E2_pernicious_stock_amt = 3
Const E2_longterm_stock_qty = 4
Const E2_longterm_stock_amt = 5

Const E3_abc_flag = 0
Const E3_pernicious_stock_qty = 1
Const E3_pernicious_stock_amt = 2
Const E3_longterm_stock_qty = 3
Const E3_longterm_stock_amt = 4

Dim PI3G130		

Dim iLngCnt
Dim iLngRow
Dim iLngMaxRows
Dim iStrNextKey

Dim iStrData
Dim TmpBuffer
Dim iTotalStr

Dim iStrPlantCd
Dim iStryyyymm
Dim iStrQueryTargetClass
Dim iStrQueryTargetCd

Dim iVarPlantNm
Dim iVarQueryTargetNm
Dim iVarLongTermStockCalPeriod
Dim iVarPerniciousStockCalPeriod
Dim iArrExport

iStrPlantCd = Request("txtPlantCd")
iStryyyymm = Request("txtYYYYMM")
iStrQueryTargetClass = Request("txtQueryTargetClass")
iStrQueryTargetCd = Request("txtQueryTargetCd")
iLngMaxRows = Request("txtMaxRows")
iStrNextKey = Request("lgStrPrevKey")

Set PI3G130 = Server.CreateObject("PI3G130.cILiLongtermInvAnal")
If CheckSystemError(Err,True) Then					
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Call PI3G130.I_LI_LONGTERM_INV_ANAL_SVR(gstrGlobalCollection, _
											C_SHEETMAXROWS_D, _
											iStrPlantCd, _
											iStryyyymm, _
											iStrQueryTargetClass, _
											iStrQueryTargetCd, _
											iStrNextKey, _
											iVarPlantNm, _
											iVarQueryTargetNm, _
											iVarLongTermStockCalPeriod, _
											iVarPerniciousStockCalPeriod, _
											iArrExport)			

If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Set PI3G130 = Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PI3G130 = Nothing

iLngCnt = UBound(iArrExport, 1)

ReDim TmpBuffer(iLngCnt)

Select Case iStrQueryTargetClass
	Case "1"	'품목별 장기재고현황 
		For iLngRow = 0 To iLngCnt
			'If iLngRow < C_SHEETMAXROWS_D Then
			iStrData = Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E1_item_cd))) _
					& Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E1_item_nm))) _
					& Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E1_spec))) _
					& Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E1_abc_flag))) _
					& Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E1_unit))) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E1_inv_price), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0) _
					& Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E1_storage_period))) _
					& Chr(11) & UNIDateClientFormat(iArrExport(iLngRow, E1_last_issue_dt)) _
					& Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E1_PerniciousStockPeriod))) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E1_pernicious_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E1_pernicious_stock_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) _
					& Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E1_LongtermStockPeriod))) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E1_longterm_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E1_longterm_stock_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) _
					& Chr(11) & iLngMaxRows + iLngRow + 1 _
					& Chr(11) & Chr(12)
				
			TmpBuffer(iLngRow) = iStrData
			'ELSE
			'	iStrNextKey = iArrExport(iLngRow, 0)
			'End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write "<Script language=vbs> " & vbCr         
		Response.Write " With Parent " & vbCr
		'Response.Write "    .lgStrPrevKey = """ & iStrNextKey & """" & vbCr  
		Response.Write "	.ggoSpread.Source = .frm1.vspdData " & vbCr
		Response.Write "	.ggoSpread.SSShowDataByClip  """ & iTotalStr  & """" & vbCr
		'Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then " & vbCr
		'Response.Write "		.DbQuery " & vbCr
		'Response.Write "	Else " & vbCr
		Response.Write "		.frm1.txtPlantNm.Value = """ & ConvSPChars(iVarPlantNm) & """" & vbCr
		Response.Write "		.frm1.txtQueryTargetNm.Value = """ & ConvSPChars(iVarQueryTargetNm) & """" & vbCr
		Response.Write "		.frm1.hPlantCd.Value = """ & iStrPlantCd & """" & vbCr
		Response.Write "		.frm1.hYr.Value = """ & iStrYr & """" & vbCr
		Response.Write "		.frm1.hMnth.Value = """ & iStrMnth & """" & vbCr
		Response.Write "		.frm1.hYyyyMm.value = """ & iStryyyymm & """" & vbCr
		Response.Write "		.frm1.hQueryTargetClass.Value = """ & iStrQueryTargetClass & """" & vbCr
		Response.Write "		.frm1.hQueryTargetCd.Value = """ & iStrQueryTargetCd & """" & vbCr
		Response.Write "		.DbQueryOK " & vbCr
		'Response.Write "	End If " & vbCr
		Response.Write "	.frm1.vspdData.focus " & vbCr
		Response.Write " End With " & vbCr		
		Response.Write "</Script> " & vbCr 

	Case "2"	'품목그룹별 장기재고현황 
		For iLngRow = 0 To iLngCnt
		'If iLngRow < C_SHEETMAXROWS_D Then
			iStrData = Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E2_item_group_cd))) _
					& Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E2_item_group_nm))) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E2_pernicious_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E2_pernicious_stock_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E2_longterm_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E2_longterm_stock_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) _
					& Chr(11) & iLngMaxRows + iLngRow + 1 _
					& Chr(11) & Chr(12)
				
			TmpBuffer(iLngRow) = iStrData
		'ELSE
		'	iStrNextKey = iArrExport(iLngRow, 0)
		'End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write "<Script language=vbs> " & vbCr         
		Response.Write " With Parent " & vbCr
		'Response.Write "    .lgStrPrevKey2 = """ & iStrNextKey & """" & vbCr  
		Response.Write "	.ggoSpread.Source = .frm1.vspdData2 " & vbCr
		Response.Write "	.ggoSpread.SSShowDataByClip  """ & iTotalStr  & """" & vbCr
		'Response.Write "	If .frm1.vspdData2.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData2, 0) And .lgStrPrevKey2 <> """" Then " & vbCr
		'Response.Write "		.DbQuery " & vbCr
		'Response.Write "	Else " & vbCr
		Response.Write "		.frm1.txtPlantNm.Value = """ & ConvSPChars(iVarPlantNm) & """" & vbCr
		Response.Write "		.frm1.txtQueryTargetNm.Value = """ & ConvSPChars(iVarQueryTargetNm) & """" & vbCr
		Response.Write "		.frm1.hPlantCd.Value = """ & iStrPlantCd & """" & vbCr
		Response.Write "		.frm1.hYr.Value = """ & iStrYr & """" & vbCr
		Response.Write "		.frm1.hMnth.Value = """ & iStrMnth & """" & vbCr
		Response.Write "		.frm1.hYyyyMm.value = """ & iStryyyymm & """" & vbCr
		Response.Write "		.frm1.hQueryTargetClass.Value = """ & iStrQueryTargetClass & """" & vbCr
		Response.Write "		.frm1.hQueryTargetCd.Value = """ & iStrQueryTargetCd & """" & vbCr
		Response.Write "		.DbQueryOK " & vbCr
		'Response.Write "	End If " & vbCr
		Response.Write "	.frm1.vspdData2.focus " & vbCr
		Response.Write " End With " & vbCr		
		Response.Write "</Script> " & vbCr 
		
	Case "3"	'ABC구분별 장기재고현황 
		For iLngRow = 0 To iLngCnt
			'If iLngRow < C_SHEETMAXROWS_D Then
			iStrData = Chr(11) & Trim(ConvSPChars(iArrExport(iLngRow, E3_abc_flag))) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E3_pernicious_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E3_pernicious_stock_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E3_longterm_stock_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) _
					& Chr(11) & UniConvNumberDBToCompany(iArrExport(iLngRow, E3_longterm_stock_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) _
					& Chr(11) & iLngMaxRows + iLngRow + 1 _
					& Chr(11) & Chr(12)
				
			TmpBuffer(iLngRow) = iStrData
			'ELSE
			'	iStrNextKey = iArrExport(iLngRow, 0)
			'End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write "<Script language=vbs> " & vbCr         
		Response.Write " With Parent " & vbCr
		'Response.Write "    .lgStrPrevKey3 = """ & iStrNextKey & """" & vbCr  
		Response.Write "	.ggoSpread.Source = .frm1.vspdData3 " & vbCr
		Response.Write "	.ggoSpread.SSShowDataByClip  """ & iTotalStr  & """" & vbCr
		'Response.Write "	If .frm1.vspdData3.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData3, 0) And .lgStrPrevKey3 <> """" Then " & vbCr
		'Response.Write "		.DbQuery " & vbCr
		'Response.Write "	Else " & vbCr
		Response.Write "		.frm1.txtPlantNm.Value = """ & ConvSPChars(iVarPlantNm) & """" & vbCr
		Response.Write "		.frm1.txtQueryTargetNm.Value = """ & ConvSPChars(iVarQueryTargetNm) & """" & vbCr
		Response.Write "		.frm1.hPlantCd.Value = """ & iStrPlantCd & """" & vbCr
		Response.Write "		.frm1.hYr.Value = """ & iStrYr & """" & vbCr
		Response.Write "		.frm1.hMnth.Value = """ & iStrMnth & """" & vbCr
		Response.Write "		.frm1.hYyyyMm.value = """ & iStryyyymm & """" & vbCr
		Response.Write "		.frm1.hQueryTargetClass.Value = """ & iStrQueryTargetClass & """" & vbCr
		Response.Write "		.frm1.hQueryTargetCd.Value = """ & iStrQueryTargetCd & """" & vbCr
		Response.Write "		.DbQueryOK " & vbCr
		'Response.Write "	End If " & vbCr
		Response.Write "	.frm1.vspdData3.focus " & vbCr
		Response.Write " End With " & vbCr		
		Response.Write "</Script> " & vbCr 
		
End Select

'Call ServerMesgBox("LAST: " & Err.number & Err.Description, vbInformation, I_MKSCRIPT)
	
Response.End

%>
