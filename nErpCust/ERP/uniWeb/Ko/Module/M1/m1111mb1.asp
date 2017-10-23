<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1111MB1
'*  4. Program Name         : 품목별단가등록 
'*  5. Program Desc         : 품목별단가등록 
'*  6. Component List       : PM1G111.cMMaintItemPriceS / PM1G118.cMListItemPriceS
'*  7. Modified date(First) : 2000/03/27
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

<%
  Call LoadBasisGlobalInf()
  Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
 
    Dim lgOpModeCRUD
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd     
    
    lgOpModeCRUD  = Request("txtMode") 
										                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002), CStr(UID_M0005)                                        '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Dim iPM1G111
	Dim iErrorPosition
	Dim I1_b_plant_cd
	Dim I2_b_item_cd
	Dim txtSpread
	
	'2003-05-27 ksh
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

	Set iPM1G111 = CreateObject("PM1G111.cMMaintItemPriceS")

    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM1G111 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    I1_b_plant_cd = Trim(Request("txtPlantCd1"))
    I2_b_item_cd = Trim(Request("txtItemCd1"))
    '-----------------------
    'Com Action Area
    '-----------------------

    iErrorPosition = iPM1G111.M_MAINT_ITEM_PRICE_SVR(gStrGlobalCollection, I1_b_plant_cd, I2_b_item_cd, itxtSpread )
        
	If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
		Set iPM1G111 = Nothing
		Exit Sub
	End If

	'-----------------------
	'Result data display area
	'-----------------------
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Call parent.DbDeleteOk() "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "             & vbCr
	        
    Set iPM1G111 = Nothing                                                   '☜: Unload Comproxy
End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next 	
	Dim iPM1G118																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iMax
	Dim PvArr

	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim istrData
    Dim arrValue       
	Const C_SHEETMAXROWS_D  = 100

    Dim I1_b_plant_plant_cd
    Dim I2_b_item_item_cd 
    Dim I3_fr_dt
    Dim I4_to_dt

    Dim I5_m_item_pur_price
    Const M334_I5_pur_unit = 0
    Const M334_I5_pur_cur = 1
    Const M334_I5_valid_fr_dt = 2
    ReDim  I5_m_item_pur_price(M334_I5_valid_fr_dt)

    Dim EG1_export_group
    Const M334_EG1_E1_m_item_pur_price_pur_unit = 0
    Const M334_EG1_E1_m_item_pur_price_pur_cur = 1
    Const M334_EG1_E1_m_item_pur_price_valid_fr_dt = 2
    Const M334_EG1_E1_m_item_pur_price_pur_prc = 3
    '이성룡 추가 및 수정 
    Const M334_EG1_E1_m_item_pur_price_prc_flg = 4
    Const M334_EG1_E1_m_item_pur_price_ext1_cd = 5
    Const M334_EG1_E1_m_item_pur_price_ext1_qty = 6
    Const M334_EG1_E1_m_item_pur_price_ext1_amt = 7
    Const M334_EG1_E1_m_item_pur_price_ext2_cd = 8
    Const M334_EG1_E1_m_item_pur_price_ext2_qty = 9
    Const M334_EG1_E1_m_item_pur_price_ext2_amt = 10
    '20050503 비고관련 추가 
    Const M334_EG1_E1_m_item_pur_price_remark = 14
    
    Dim E1_b_plant
    Const M334_E1_plant_cd = 0
    Const M334_E1_plant_nm = 1
    
    Dim E2_b_item
    Const M334_E2_item_cd = 0
    Const M334_E2_item_nm = 1
    
    Dim E3_m_item_pur_price
    Const M334_E3_pur_unit = 0
    Const M334_E3_pur_cur = 1
    Const M334_E3_valid_fr_dt = 2
    
	'⊙: 각 화면당 Relation이 되어 있지 않는 Field들에 대해서는 조회용 Lookup을 행한다.
	'strCode = PlantLookUp(Request("txtPlantCd1"))		' 현재 Function은 안만들어져 있음, GetCalTypeNm를 참조로 각자 만듦.
	
	'⊙: 조회 조건이 >= 인경우는 그냥 통과해서 조회를 진행한다.
	
	'⊙:  조회 조건이 = 인 경우는 Stop시킨다.
	

	If Len(Trim(Request("txtAppFrDt"))) Then
		If UNIConvDate(Request("txtAppFrDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtAppFrDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If

	If Len(Trim(Request("txtAppToDt"))) Then
		If UNIConvDate(Request("txtAppToDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtAppToDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If

    Set iPM1G118 = Server.CreateObject("PM1G118.cMListItemPriceS")

	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM1G118 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_b_plant_plant_cd				      	= Trim(UCase(Request("txtPlantCd1")))
    I2_b_item_item_cd 						= Trim(UCase(Request("txtitemCd1")))
    
    If Request("txtAppFrDt") = "" Then
		I3_fr_dt 	                    	= "1900-01-01"
	Else
		I3_fr_dt                        	= UNIConvDate(Request("txtAppFrDt"))
	End if 

	If Request("txtAppToDt") = "" Then
		I4_to_dt	                    	= "2999-12-31"
	Else
		I4_to_dt	                        = UNIConvDate(Request("txtAppToDt"))
	End if 
	
	lgStrPrevKey = Request("lgStrPrevKey")
	
    If lgStrPrevKey <> "" then	
        arrValue = Split(lgStrPrevKey, gColSep)		
		I5_m_item_pur_price(M334_I5_pur_unit) = arrValue(0)
		I5_m_item_pur_price(M334_I5_pur_cur) = arrValue(1)
		I5_m_item_pur_price(M334_I5_valid_fr_dt) = arrValue(2)
	else			
		I5_m_item_pur_price(M334_I5_pur_unit) = ""
		I5_m_item_pur_price(M334_I5_pur_cur) = ""
		I5_m_item_pur_price(M334_I5_valid_fr_dt) = ""
	End If	
	
	
	Call iPM1G118.M_LIST_ITEM_PRICE_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_b_plant_plant_cd, I2_b_item_item_cd, _
        I3_fr_dt, I4_to_dt, I5_m_item_pur_price, EG1_export_group, E1_b_plant, E2_b_item, E3_m_item_pur_price)
     
     
     If CheckSYSTEMError2(Err,True,"","","","","") = true then 	
		Response.Write "<Script Language=vbscript>"												& vbCr
		Response.Write "With Parent "															& vbCr
		Response.Write " .frm1.hdnPlant.Value   = """ & ConvSPChars(Request("txtPlantCd1"))		& """" & vbCr
		Response.Write " .frm1.hdnItem.Value   = """ & ConvSPChars(Request("txtItemCd1"))		& """" & vbCr
		Response.Write " .frm1.hdnFrDt.value     = """ & Request("txtAppFrDt")                  & """" & vbCr
    	Response.Write " .frm1.hdnToDt.value     = """ & Request("txtAppToDt")     		        & """" & vbCr
		Response.Write " .DbQueryOk "           & vbCr	
		Response.Write "	.frm1.txtPlantCd2.value = """ & ConvSPChars(E1_b_plant(M334_E1_plant_cd))      & """" & vbCr
		Response.Write "	.frm1.txtPlantNm2.value = """ & ConvSPChars(E1_b_plant(M334_E1_plant_nm))      & """" & vbCr
		Response.Write "	.frm1.txtItemCd2.value = """ & ConvSPChars(E2_b_item(M334_E2_item_cd))      & """" & vbCr
		Response.Write "	.frm1.txtItemNm2.value = """ & ConvSPChars(E2_b_item(M334_E2_item_nm))      & """" & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr    
		Set iPM1G118 = Nothing
		Exit Sub
	End If
	
		
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtPlantNm1.value = """ & ConvSPChars(E1_b_plant(M334_E1_plant_nm))      & """" & vbCr
	Response.Write "	.frm1.txtItemNm1.value = """ & ConvSPChars(E2_b_item(M334_E2_item_nm))      & """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr

	If lgStrPrevKey = StrNextKey And UBound(EG1_export_group,1) < 0 Then
		Set iPM1G118 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
    Response.Write "	.frm1.txtPlantCd2.value = """ & ConvSPChars(E1_b_plant(M334_E1_plant_cd))      & """" & vbCr
	Response.Write "	.frm1.txtPlantNm2.value = """ & ConvSPChars(E1_b_plant(M334_E1_plant_nm))      & """" & vbCr
	Response.Write "	.frm1.txtItemCd2.value = """ & ConvSPChars(E2_b_item(M334_E2_item_cd))      & """" & vbCr
	Response.Write "	.frm1.txtItemNm2.value = """ & ConvSPChars(E2_b_item(M334_E2_item_nm))      & """" & vbCr

	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr	

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
	iMax = UBound(EG1_export_group,1)
	ReDim PvArr(iMax)
    
	For iLngRow = 0 To UBound(EG1_export_group,1)
	
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(E3_m_item_pur_price(M334_E3_pur_unit)) & gColSep & ConvSPChars(E3_m_item_pur_price(M334_E3_pur_cur)) _
		                     & gColSep & ConvSPChars(E3_m_item_pur_price(M334_E3_valid_fr_dt) )

           Exit For
        End If 

        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M334_EG1_E1_m_item_pur_price_pur_unit))
        istrData = istrData & Chr(11) & " "      'PopUp
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M334_EG1_E1_m_item_pur_price_pur_cur))
        istrData = istrData & Chr(11) & " "      'PopUp
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow,M334_EG1_E1_m_item_pur_price_valid_fr_dt))
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow,M334_EG1_E1_m_item_pur_price_pur_prc), 0)
        '이성룡 추가 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M334_EG1_E1_m_item_pur_price_prc_flg))
        istrData = istrData & Chr(11) & ""
        '20050503 비고 관련 추가 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M334_EG1_E1_m_item_pur_price_remark))		'remark
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                                   '11
        istrData = istrData & Chr(11) & Chr(12)
        
		PvArr(iLngRow) = istrData
		istrData=""
    Next  

	istrData = Join(PvArr, "")
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr      
    Response.Write "	.ggoSpread.Source          =  .frm1.vspdData "         & vbCr
    Response.Write  "    .frm1.vspdData.Redraw = False   "                     & vbCr   
    Response.Write "	.ggoSpread.SSShowData        """ & istrData & """" & ",""F""" & vbCr
    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1,-1,.C_Curr,.C_Cost,""C"" ,""I"",""X"",""X"")" & vbCr
    Response.Write "	.lgStrPrevKey              = """ & StrNextKey   & """" & vbCr 
    Response.Write " .frm1.hdnPlant.value    = """ & ConvSPChars(UCase(Request("txtPlantCd1")))  & """" & vbCr
	Response.Write " .frm1.hdnItem.value     = """ & ConvSPChars(UCase(Request("txtitemCd1")))   & """" & vbCr
	Response.Write " .frm1.hdnFrDt.value     = """ & Request("txtAppFrDt")                       & """" & vbCr
	Response.Write " .frm1.hdnToDt.value     = """ & Request("txtAppToDt")                       & """" & vbCr
    Response.Write " .DbQueryOk "		    	   & vbCr 
    Response.Write " .frm1.vspdData.focus "		   & vbCr 
    Response.Write " .frm1.vspdData.Redraw = True " & vbCr   
    Response.Write "End With"                      & vbCr
    Response.Write "</Script>"                     & vbCr
  
      
    Set iPM1G118 = Nothing
  
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing

	Dim iPM1G111																'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	Dim I1_b_plant_cd
	Dim I2_b_item_cd
	Dim iErrorPosition
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

    Set iPM1G111 = Server.CreateObject("PM1G111.cMMaintItemPriceS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM1G111 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if

	I1_b_plant_cd 					= UCase(Trim(Request("txtPlantCd2")))
	I2_b_item_cd 					= UCase(Trim(Request("txtItemCd2")))
	
'	Call ServerMesgBox(itxtSpread , vbInformation, I_MKSCRIPT)

	iErrorPosition = iPM1G111.M_MAINT_ITEM_PRICE_SVR(gStrGlobalCollection, I1_b_plant_cd, I2_b_item_cd, itxtSpread, iErrorPosition ) 

   If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
	  Set iPM1G111 = Nothing
	  If Trim(iErrorPosition) <> "" Then
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.SetRow(""" & iErrorPosition & """) " & vbCr
		Response.Write "</Script>" & vbCr	
	  End If
	  Exit Sub
   End If        
    Set iPM1G111 = Nothing                                                   '☜: Unload Comproxy

	Response.Write "<Script language=vbs> " & vbCr 
	Response.Write "With parent " & vbCr
	Response.Write " .frm1.txtPlantCd1.value    = """ & ConvSPChars(UCase(Request("txtPlantCd2")))  & """" & vbCr
	Response.Write " .frm1.txtItemCd1.value     = """ & ConvSPChars(UCase(Request("txtItemCd2")))   & """" & vbCr        
    Response.Write " .DbSaveOk "      & vbCr						
    Response.Write "End With " & vbCr
    Response.Write "</Script> "              
End Sub    


%>  
