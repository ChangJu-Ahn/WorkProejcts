<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1112MB1
'*  4. Program Name         : 공급처별단가등록 
'*  5. Program Desc         : 공급처별단가등록 
'*  6. Component List       : PM1G121.cMMntSpplItemPriceS
'*  7. Modified date(First) : 2000/05/11
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
  Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")

	Dim lgOpModeCRUD
	
	Dim iPM1G121																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	Dim iPM1G128																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Dim istrData

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
		
	Call HideStatusWnd																 '☜: Hide Processing message

	lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
	
	Select Case lgOpModeCRUD
	    Case CStr(UID_M0001)                                                         '☜: Query
	         Call  SubBizQueryMulti()
	    Case CStr(UID_M0002), CStr(UID_M0005), Cstr(UID_M0005)                       '☜: Save,Update
	         Call SubBizSaveMulti()
	End Select
 
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim E2_ief_Supplied																		'문제발생시 문제를 일으킨 레코드 숫자를 반환한다.
    Dim iErrorPosition
    
    Dim Plant_CD
    Dim Item_CD
    Dim Supplier_CD
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim iDCount
    Dim ii
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   

	Set iPM1G121 = Server.CreateObject("PM1G121.cMMntSpplItemPriceS")

	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End If
	
	Plant_CD = Trim(Request("txtPlantCd2"))
	Item_CD = Trim(Request("txtitemCd2"))
	Supplier_CD = Trim(Request("txtSupplierCd2"))
	
	'Call ServerMesgBox(itxtSpread , vbInformation, I_MKSCRIPT)
	
	Call iPM1G121.M_MAINT_SPPL_ITEM_PRICE_SVR(gStrGlobalCollection, Plant_Cd, item_Cd, Supplier_Cd ,itxtSpread, iErrorPosition)
	
	If CheckSYSTEMError2(Err,True, iErrorPosition(0) & "행:" ,"","","","") = true then 		
		Set iPM1G121 = Nothing															'☜: ComProxy Unload'
		Exit Sub
	End If
	
    Set iPM1G121 = Nothing    
	
	Response.Write "<Script Language=vbscript>"												& vbCr
	Response.Write "With Parent "															& vbCr
	Response.Write " .frm1.txtPlantCd1.Value   = """ & ConvSPChars(Request("txtPlantCd2"))		& """" & vbCr
	Response.Write " .frm1.txtItemCd1.Value   = """ & ConvSPChars(Request("txtItemCd2"))		& """" & vbCr
	Response.Write " .frm1.txtSupplierCd1.Value = """ & ConvSPChars(Request("txtSupplierCd2"))	& """" & vbCr
	Response.Write " .DBSaveOK "           & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr 
End Sub    

	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iMax
	Dim PvArr
	Dim iArrPrevKey
	Const C_SHEETMAXROWS_D  = 100
	Dim I1_m_supplier_item_price
	Const M332_I1_plant_cd = 0
	Const M332_I1_item_cd = 1
	Const M332_I1_bp_cd = 2
	Const M332_I1_valid_fr_dt = 3
	Const M332_I1_valid_to_dt = 4

	Dim I2_m_supplier_item_price_next
	Const M332_I2_pur_unit = 0
	Const M332_I2_pur_cur = 1
	Const M332_I2_valid_fr_dt = 2

	Dim EG1_export_group
	Const M332_EG_pur_unit = 0
	Const M332_EG_pur_cur = 1
	Const M332_EG_valid_fr_dt = 2
	Const M332_EG_pur_prc = 3
	'단가구분 
	Const M332_EG_prc_flg = 4
	Const M332_EG_ext1_cd = 5
	Const M332_EG_ext1_qty = 6
	Const M332_EG_ext1_amt = 7
	Const M332_EG_ext2_cd = 8
	Const M332_EG_ext2_qty = 9
	Const M332_EG_ext2_amt = 10
	Const M332_EG_remark = 11

	Dim E1_b_plant
	Const M332_E1_plant_cd = 0
	Const M332_E1_plant_nm = 1

	Dim E2_b_item
	Const M332_E2_item_cd = 0
	Const M332_E2_item_nm = 1

	Dim E3_b_biz_partner
	Const M332_E3_bp_cd = 0
	Const M332_E3_bp_nm = 1

	Dim E4_m_supplier_item_price_next
	E4_m_supplier_item_price_next=""
	
	ReDim I1_m_supplier_item_price(4)
	ReDim I2_m_supplier_item_price_next(2)
	
	On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear      
	
    Set iPM1G128 = Server.CreateObject("PM1G128.cMLstSpplItemPriceS")
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
	End If
	
	If Request("lgStrPrevKey") <> "" Then
		iArrPrevKey = Split(Request("lgStrPrevKey"),gColSep)
		I2_m_supplier_item_price_next(M332_I2_pur_unit)		= iArrPrevKey(0)
		I2_m_supplier_item_price_next(M332_I2_pur_cur)		= iArrPrevKey(1)
		I2_m_supplier_item_price_next(M332_I2_valid_fr_dt)	= UNIConvDate(iArrPrevKey(2)) 
	Else
		I2_m_supplier_item_price_next(M332_I2_valid_fr_dt)	= "1900-01-01"
	End If

	I1_m_supplier_item_price(M332_I1_plant_cd)		= Trim(Request("txtPlantCd1"))
	I1_m_supplier_item_price(M332_I1_item_cd)		= Trim(Request("txtitemCd1"))
	I1_m_supplier_item_price(M332_I1_bp_cd)			= Trim(Request("txtSupplierCd1"))
    If Len(Trim(Request("txtAppFrDt"))) Then
		I1_m_supplier_item_price(M332_I1_valid_fr_dt)	= UNIConvDate(Request("txtAppFrDt"))
	Else
		I1_m_supplier_item_price(M332_I1_valid_fr_dt)	= "1900-01-01"
	End If 
	
	If Len(Trim(Request("txtAppToDt"))) Then
		I1_m_supplier_item_price(M332_I1_valid_to_dt)	= UNIConvDate(Request("txtAppToDt")) 
	Else
		I1_m_supplier_item_price(M332_I1_valid_to_dt)	= "2999-12-31"
	End If 
	
	Call iPM1G128.M_LIST_SPPL_ITEM_PRICE_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_supplier_item_price, _
	I2_m_supplier_item_price_next,EG1_export_group,E1_b_plant,E2_b_item,E3_b_biz_partner, E4_m_supplier_item_price_next)

	If CheckSYSTEMError(Err,True) = true then 	
		Response.Write "<Script Language=vbscript>"												& vbCr
		Response.Write "With Parent "															& vbCr
		Response.Write " .frm1.hdnPlant.Value   = """ & ConvSPChars(Request("txtPlantCd1"))		& """" & vbCr
		Response.Write " .frm1.hdnItem.Value   = """ & ConvSPChars(Request("txtItemCd1"))		& """" & vbCr
		Response.Write " .frm1.hdnSupplier.Value = """ & ConvSPChars(Request("txtSupplierCd1"))	& """" & vbCr
		Response.Write " .frm1.hdnFrDt.Value   = """ & ConvSPChars(Request("txtAppFrDt"))		& """" & vbCr
		Response.Write " .frm1.hdnToDt.Value = """ & ConvSPChars(Request("txtAppToDt"))			& """" & vbCr
		Response.Write " .DbQueryOk "           & vbCr	
		
		Response.Write " .frm1.txtPlantNm1.value		= """ & ConvSPChars(E1_b_plant(M332_E1_plant_nm)) & """" & vbCr
		Response.Write " .frm1.txtItemNm1.value			= """ & ConvSPChars(E2_b_item(M332_E2_item_nm)) & """" & vbCr
		Response.Write " .frm1.txtSupplierNm1.value		= """ & ConvSPChars(E3_b_biz_partner(M332_E3_bp_nm)) & """" & vbCr	
		Response.Write " .frm1.txtPlantCd2.value		= """ & ConvSPChars(Request("txtPlantCd1")) & """" & vbCr
		Response.Write " .frm1.txtPlantNm2.value		= """ & ConvSPChars(E1_b_plant(M332_E1_plant_nm)) & """" & vbCr
		Response.Write " .frm1.txtItemCd2.value			= """ & ConvSPChars(Request("txtItemCd1")) & """" & vbCr	
		Response.Write " .frm1.txtItemNm2.value			= """ & ConvSPChars(E2_b_item(M332_E2_item_nm)) & """" & vbCr
		Response.Write " .frm1.txtSupplierCd2.value		= """ & ConvSPChars(Request("txtSupplierCd1")) & """" & vbCr	
		Response.Write " .frm1.txtSupplierNm2.value		= """ & ConvSPChars(E3_b_biz_partner(M332_E3_bp_nm)) & """" & vbCr	
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr    
		Set iPM1G128 = Nothing
		Exit Sub
	End If

	Set iPM1G128 = Nothing
	
	iLngMaxRow = CLng(Request("txtMaxRows"))
	
    Response.Write "<Script Language=vbscript>" & vbCr	    
	Response.Write " With parent "	& vbCr
	Response.Write " iLngMaxRow = .frm1.vspdData.MaxRows	 " & VbCr		
	Response.Write " .frm1.txtPlantCd1.value		= """ & ConvSPChars(E1_b_plant(M332_E1_plant_cd)) & """" & vbCr
	Response.Write " .frm1.txtPlantNm1.value		= """ & ConvSPChars(E1_b_plant(M332_E1_plant_nm)) & """" & vbCr
	Response.Write " .frm1.txtPlantCd2.value		= """ & ConvSPChars(E1_b_plant(M332_E1_plant_cd)) & """" & vbCr
	Response.Write " .frm1.txtPlantNm2.value		= """ & ConvSPChars(E1_b_plant(M332_E1_plant_nm)) & """" & vbCr
	Response.Write " .frm1.txtItemCd1.value			= """ & ConvSPChars(E2_b_item(M332_E2_item_cd)) & """" & vbCr	
	Response.Write " .frm1.txtItemNm1.value			= """ & ConvSPChars(E2_b_item(M332_E2_item_nm)) & """" & vbCr
	Response.Write " .frm1.txtItemCd2.value			= """ & ConvSPChars(E2_b_item(M332_E2_item_cd)) & """" & vbCr	
	Response.Write " .frm1.txtItemNm2.value			= """ & ConvSPChars(E2_b_item(M332_E2_item_nm)) & """" & vbCr
	Response.Write " .frm1.txtSupplierCd1.value		= """ & ConvSPChars(E3_b_biz_partner(M332_E3_bp_cd)) & """" & vbCr	
	Response.Write " .frm1.txtSupplierNm1.value		= """ & ConvSPChars(E3_b_biz_partner(M332_E3_bp_nm)) & """" & vbCr	
	Response.Write " .frm1.txtSupplierCd2.value		= """ & ConvSPChars(E3_b_biz_partner(M332_E3_bp_cd)) & """" & vbCr	
	Response.Write " .frm1.txtSupplierNm2.value		= """ & ConvSPChars(E3_b_biz_partner(M332_E3_bp_nm)) & """" & vbCr	
	Response.Write " End With"
    Response.Write "</Script>" & vbCr
	
	iMax = UBound(EG1_export_group,1)
	ReDim PvArr(iMax)
	For iLngRow = 0 To UBound(EG1_export_group, 1)
		If iLngRow = C_SHEETMAXROWS_D Then 
			Exit For
		End If
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M332_EG_pur_unit))
		istrData = istrData & Chr(11) & " "      'PopUp
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M332_EG_pur_cur))
        istrData = istrData & Chr(11) & " "      'PopUp
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, M332_EG_valid_fr_dt))              
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow, M332_EG_pur_prc), 0)
        '단가구분 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M332_EG_prc_flg))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M332_EG_remark))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                               
        istrData = istrData & Chr(11) & Chr(12)

		PvArr(iLngRow) = istrData
		istrData=""
    Next
    istrData = Join(PvArr, "")

    
	Response.Write "<Script Language=vbscript>"												& vbCr
	Response.Write "With Parent "															& vbCr
	Response.Write " .ggoSpread.Source		= .frm1.vspdData"									& vbCr
	Response.Write " .frm1.vspdData.Redraw	= False   "                      & vbCr   
	Response.Write " .ggoSpread.SSShowData     """ & istrData & """" & ",""F""" & vbCr
    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1,-1,.C_Curr,.C_Cost,""C"" ,""I"",""X"",""X"")" & vbCr
	Response.Write " .lgStrPrevKey          = """ & E4_m_supplier_item_price_next			& """" & vbCr
	Response.Write " .frm1.hdnPlant.Value   = """ & ConvSPChars(Request("txtPlantCd1"))		& """" & vbCr
	Response.Write " .frm1.hdnItem.Value	= """ & ConvSPChars(Request("txtItemCd1"))		& """" & vbCr
	Response.Write " .frm1.hdnSupplier.Value= """ & ConvSPChars(Request("txtSupplierCd1"))	& """" & vbCr
	Response.Write " .frm1.hdnFrDt.Value	= """ & ConvSPChars(Request("txtAppFrDt"))		& """" & vbCr
	Response.Write " .frm1.hdnToDt.Value	= """ & ConvSPChars(Request("txtAppToDt"))			& """" & vbCr
	Response.Write " .DbQueryOk "           & vbCr	
	Response.Write  ".frm1.vspdData.Redraw = True   "                      & vbCr   
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr    
End Sub

%>

