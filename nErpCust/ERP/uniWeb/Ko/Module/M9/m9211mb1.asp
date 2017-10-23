<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%

'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M9211MB1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'							  
'*  7. Modified date(First) : 2001/05/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : KO MYOUNG JIN
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  :
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

	On Error Resume Next
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("*", "M","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("*", "M","NOCOOKIE", "MB")
	Call HideStatusWnd

	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	DIM SCheck
	
	Const C_SHEETMAXROWS_D  = 100

	Dim strExpValue
	Dim strExpValue2
	Dim strLastSeq

	lgOpModeCRUD	=	Request("txtMode")	'☜: Read Operation Mode (CRUD)

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0001)                                                         '☜: Query
	             Call SubBizQueryMulti()
	        Case CStr(UID_M0002)
	             Call SubBizSaveMulti()
	        Case CStr ("changeMvmtType")
				 Call DisplayMvmtNm(request("txtMvmtType"))
			Case "changeSpplCd" 
				 Call DisplaySupplierNm(Request("txtSupplierCd"))	 
			Case "changeGroupCd" 
				 Call DisplayGroupNm(Request("txtGroupCd"))	 	
	End Select


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
 
 Dim iPM9G218
 
 Dim FlgData
 Dim I1_m_pur_goods_mvmt_no
 Dim I2_m_pur_goods_mvmt_rcpt_no
 
 Dim E1_b_pur_grp
 Const M9218_E1_pur_grp = 0
 Const M9218_E1_pur_grp_nm = 1
 
 Dim E2_b_biz_partner   
 Const M9218_E2_bp_cd = 0
 Const M9218_E2_bp_nm = 1

 Dim E3_m_pur_goods_mvmt
 
 Dim E4_m_mvmt_type
 Const M9218_E4_io_type_cd = 0
 Const M9218_E4_io_type_nm = 1
 
 
 Dim EG1_export_group
 Const M9218_EG1_E11_item_cd = 1
 Const M9218_EG1_E11_item_nm = 2
 Const M9218_EG1_E11_item_spec = 3
 Const M9218_EG1_E11_mvmt_rcpt_qty = 4
 Const M9218_EG1_E11_mvmt_rcpt_unit = 5
 Const M9218_EG1_E11_tracking_no = 6
 Const M9218_EG1_E11_mvmt_doc_amt = 7
 Const M9218_EG1_E11_mvmt_cur = 8
 Const M9218_EG1_E11_plant_cd = 9
 Const M9218_EG1_E11_plant_nm = 10
 Const M9218_EG1_E11_mvmt_rcpt_sl_cd = 11
 Const M9218_EG1_E11_mvmt_rcpt_sl_nm = 12
 Const M9218_EG1_E11_lot_no = 13
 Const M9218_EG1_E11_lot_sub_no = 14
 Const M9218_EG1_E11_make_lot_no = 15
 Const M9218_EG1_E11_make_lot_sub_no = 16
 Const M9218_EG1_E11_sgi_no = 17
 Const M9218_EG1_E11_sgi_seq_no = 18
 Const M9218_EG1_E11_sto_no = 19
 Const M9218_EG1_E11_sto_seq_no = 20
 Const M9218_EG1_E11_dn_no = 21
 Const M9218_EG1_E11_dn_seq_no = 22
 Const M9218_EG1_E11_base_unit = 23
 Const M9218_EG1_E11_base_price = 24
 Const M9218_EG1_E11_loc_amt = 25
 Const M9218_EG1_E11_mvmt_no = 26
 Const M9218_EG1_E11_Base_qty = 27
 Const M9218_EG1_E11_PUR_GRP = 28
 Const M9218_EG1_E11_PUR_GRP_NM = 29
 Const M9218_EG1_E11_IO_TYPE_CD = 30
 Const M9218_EG1_E11_IO_TYPE_NM = 31
 Const M9218_EG1_E11_BP_CD = 32
 Const M9218_EG1_E11_BP_NM = 33
 Const M9218_EG1_E11_MVMT_RCPT_DT = 34
 Const M9218_EG1_E11_TAX_CD = 35
 Const M9218_EG1_E11_TAX_RT = 36
 'Const M9218_EG1_E11_TAX_REMARK = 37
 
 Dim strGmNo   


    On Error Resume Next 
    Err.Clear                                                               '☜: Protect system from crashing
    
	lgStrPrevKey = Request("lgStrPrevKey")
  
    Set iPM9G218 = Server.CreateObject("PM9G218.cMListStoRcptS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = True Then
				Set iPM9G218 = Nothing
				Exit Sub
	End If
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I2_m_pur_goods_mvmt_rcpt_no  		= FilterVar(UCase(Trim(Request("txtMvmtNo"))),"","SNM")
    
    if Trim(lgStrPrevKey) <> "" then
		I1_m_pur_goods_mvmt_no  	= lgStrPrevKey
	End if
    
    '-----------------------
    'Com action area
    '-----------------------
	Call iPM9G218.M_LIST_STO_RCPT_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_pur_goods_mvmt_no, I2_m_pur_goods_mvmt_rcpt_no, E1_b_pur_grp, E2_b_biz_partner, E3_m_pur_goods_mvmt, E4_m_mvmt_type, EG1_export_group)
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.frm1.txtMvmtNo.value 	= """" " & vbCr
		Response.Write "	Parent.FncNew() " & vbCr
		Response.Write "</Script>" & vbCr
		Set iPM9G218 = Nothing
		Exit Sub
	End If
		
    iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_export_group,1)
    
    IF GroupCount <> 0 then
		If EG1_export_group(GroupCount,M744_EG1_E11_mvmt_no) = E3_m_pur_goods_mvmt Then
			StrNextKey = ""
		Else
			StrNextKey = E3_m_pur_goods_mvmt
		End If
	End if
	
	   Response.Write "<Script Language=VBScript>" & vbCr
	   Response.Write "With parent"                 & vbCr
       Response.Write ".frm1.txtMvmtType.Value      = """ & ConvSPChars(EG1_export_group(GroupCount,M9218_EG1_E11_IO_TYPE_CD)) & """"                              & vbCr
       Response.Write ".frm1.txtMvmtTypeNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M9218_EG1_E11_IO_TYPE_NM)) & """"                              & vbCr       	   		
       Response.Write ".frm1.txtGmDt.text           = """ & UNIDateClientFormat(EG1_export_group(GroupCount,M9218_EG1_E11_MVMT_RCPT_DT)) & """" & vbCr
       Response.Write ".frm1.txtGroupCd.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M9218_EG1_E11_PUR_GRP)) & """"                                   & vbCr
       Response.Write ".frm1.txtGroupNm.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M9218_EG1_E11_PUR_GRP_NM)) & """"                                & vbCr
       Response.Write ".frm1.txtSupplierCd.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M9218_EG1_E11_BP_CD)) & """"                                 & vbCr
       Response.Write ".frm1.txtSupplierNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M9218_EG1_E11_BP_NM)) & """"                                 & vbCr
       'Response.Write ".frm1.txtTaxCd.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M9218_EG1_E11_TAX_CD)) & """"                                 & vbCr
       ' Response.Write ".frm1.txtTaxRATE.TEXT    = """ & UNINumClientFormat(EG1_export_group(GroupCount,M9218_EG1_E11_TAX_RT),ggExchRate.DecPoint,0) & """"                                 & vbCr
       'Response.Write "If trim(.frm1.txtTaxCd.Value) <> """" Then " & vbCr 
       'Response.Write ".frm1.txtTaxNM.VALUE    = """ & ConvSPChars(EG1_export_group(GroupCount,M9218_EG1_E11_TAX_REMARK)) & """"                                 & vbCr
       
      
       

       
       'Response.write "End If "                         & vbCr		
       Response.Write ".frm1.txtMvmtNo1.Value       = """ & ConvSPChars(I2_m_pur_goods_mvmt_rcpt_no) & """"                                   
	   Response.Write "End With"                    & vbCr
	   Response.Write "</Script>"	                & vbCr
     	
	'-----------------------
	'Result data display area
	'----------------------- 
	For iLngRow = 0 To UBound(EG1_export_group,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_export_group(iLngRow, M744_EG1_E11_mvmt_no)) 
           Exit For
        End If  


        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_item_cd))				'입고Seq
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_item_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_item_spec))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M9218_EG1_E11_mvmt_rcpt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_mvmt_rcpt_unit))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_tracking_no))	
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,M9218_EG1_E11_mvmt_doc_amt), ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_mvmt_cur)), ggUnitCostNo,"X","X")   
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_mvmt_cur))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_plant_cd))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_plant_nm))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_mvmt_rcpt_sl_cd))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_mvmt_rcpt_sl_nm))	
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_lot_no))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_lot_sub_no))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_make_lot_no))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_make_lot_sub_no))
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_sgi_no))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_sgi_seq_no))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_sto_no))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_sto_seq_no))
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_dn_no))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_dn_seq_no))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_base_unit))	
        istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_export_group(iLngRow,M9218_EG1_E11_base_price), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
        
        istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_export_group(iLngRow,M9218_EG1_E11_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)   
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_mvmt_no))	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M9218_EG1_E11_Base_qty),ggQty.DecPoint,0)	
        istrData = istrData & Chr(11) & ""	

        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
        
        If strGmNo = "" Then
            strGmNo = ConvSPChars(EG1_export_group(iLngRow,M9218_EG1_E11_sgi_no))
        End if
    Next
    
    Response.Write "<Script Language=VBScript>"      & vbCr
	Response.Write "With parent"                     & vbCr
	Response.Write ".ggoSpread.Source = .frm1.vspdData"       & vbCr
    Response.Write ".ggoSpread.SSShowData        """ & istrData	    & """" & vbCr	   
		
    Response.Write ".lgStrPrevKey     = """ & StrNextKey   & """" & vbCr 

	Response.Write "If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> """" Then " & vbCr
    Response.write ".DbQuery " & vbCr
    Response.write "Else "                           & vbCr
    Response.Write ".frm1.hdnMvmtNo.value = """ & ConvSPChars(Request("txtMvmtNo")) & """" & vbCr
    Response.Write ".frm1.hdnRcptNo.value = """ & ConvSPChars(StrNextKey) & """" & vbCr
    Response.write ".DbQueryOk "                    & vbCr
    Response.write "End If "                         & vbCr		
	Response.Write "End With"                        & vbCr
	Response.Write "</Script>"	                    & vbCr
	
	Set iPM9G218 = Nothing 
	
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'******* 전표No. 만들기. *********
	'*************************************************************************************************************
	If  strGmNo <> "" Or strGmNo <> Null then  

		Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
				
		lgStrSQL = "SELECT document_year FROM i_goods_movement_header " 
		lgStrSQL = lgStrSQL & " WHERE item_document_no =  " & FilterVar(strGmNo , "''", "S") & ""		
		
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.hdnGlType.value	=	""B"" " & vbCr
			Response.Write "</Script>"					& vbCr
		
			Call SubCloseRs(lgObjRs)  
			Call SubCloseDB(lgObjConn)
			Exit Sub
		End if
		'A_GL.Ref_no
		strGmNo	=	strGmNo & "-" & lgObjRs("document_year")
		
		lgStrSQL = "SELECT gl_no FROM a_gl " 
		lgStrSQL = lgStrSQL & " WHERE ref_no =  " & FilterVar(strGmNo , "''", "S") & ""		
				
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			
			lgStrSQL = "SELECT temp_gl_no FROM a_temp_gl " 
			lgStrSQL = lgStrSQL & " WHERE ref_no =  " & FilterVar(strGmNo , "''", "S") & ""		
						
			IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
				Response.Write "<Script Language=VBScript>" & vbCr
				Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""""      " & vbCr  
				Response.Write "</Script>"					& vbCr
		        Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSet		
	    	Else
				Response.Write "<Script Language=VBScript>" & vbCr
				Response.Write "parent.frm1.hdnGlType.value	=	""T""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("temp_gl_no") & """" & vbCr  
				Response.Write "</Script>"					& vbCr
			End if
		    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSet		
	    Else
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.hdnGlType.value	=	""A""	  " & vbCr
			Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("gl_no") & """" & vbCr  
			Response.Write "</Script>"					& vbCr
		End if	
	    
	    Call SubCloseDB(lgObjConn)	                                                '☜ : Release DB Connection		
	Else
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
		Response.Write "</Script>"					& vbCr
	End if
	
 End Sub
        
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================
 Sub subBizSaveMulti()															'☜: 저장 요청을 받음 
 
    Dim iPM9G211
    Dim iCommandSent
    Dim iErrorPosition
 
    Dim arrVal, arrTemp, strStatus		
 
    Dim I1_select_char
    Dim I3_b_biz_partner_bp_cd
    DIm I4_b_pur_grp
    DIm I5_m_mvmt_type_io_type_cd
    Dim I6_m_pur_goods_mvmt
    DIM I7_m_mvmt_tax_cd
    DIM I8_m_mvmt_tax_rt
    
    Const M9211_I6_mvmt_rcpt_no = 0  
    Const M9211_I6_mvmt_rcpt_dt = 1
    Const M9211_I6_gm_no = 2
    
    Redim I6_m_pur_goods_mvmt(M9211_I6_gm_no)

    Dim E1_m_pur_goods_mvmt
    
    On Error Resume Next 		
    Err.Clear																		'☜: Protect system from crashing

	If Len(Request("txtGmDt")) Then
		If UNIConvDate(Request("txtGmDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Response.End	
		End If
	End If
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
		
	iLngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

    Set iPM9G211 = Server.CreateObject("PM9G211.cMMaintStoRcptSvr")    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

    '-----------------------
    'Data manipulate area
    '-----------------------

	I6_m_pur_goods_mvmt(M9211_I6_mvmt_rcpt_no)				= Request("txtMvmtNo1")
	I6_m_pur_goods_mvmt(M9211_I6_mvmt_rcpt_dt)				= UNIConvDate(Request("txtGmDt"))
	I6_m_pur_goods_mvmt(M9211_I6_gm_no)				        = ""							
    I3_b_biz_partner_bp_cd                  				= Trim(Request("txtSupplierCd"))
    I5_m_mvmt_type_io_type_cd                  				= Trim(Request("txtMvmtType"))
    I4_b_pur_grp                            				= Trim(Request("txtGroupCd"))
    I7_m_mvmt_tax_cd										= Trim(Request("txtTaxCd"))
    I8_m_mvmt_tax_rt										= UNIConvNum(Request("txtTaxRate"),0)
    
    arrTemp = Split(Request("txtSpread"), gRowSep)	
    arrVal = Split(arrTemp(0), gColSep)
	
	  
	
	strStatus = arrVal(0)

	if strStatus = "C" then
	   iCommandSent								= "CREATE"
	else 
	   iCommandSent								= "DELETE"		
	end if

   	Call iPM9G211.M_MAINT_STO_RCPT_SVR(gStrGlobalCollection, iCommandSent, I1_select_char, UCase(I3_b_biz_partner_bp_cd), _
	            UCase(I4_b_pur_grp), UCase(I5_m_mvmt_type_io_type_cd), I6_m_pur_goods_mvmt,I7_m_mvmt_tax_cd, I8_m_mvmt_tax_rt,Trim(Request("txtSpread")), _
	            E1_m_pur_goods_mvmt,iErrorPosition)
	
	'If CheckSYSTEMError2(Err, True, iErrorPosition & "","","","","") = True Then
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
	   Set iPM9G211 = Nothing
	   Exit Sub
	End If		
    
    Set iPM9G211 = Nothing                                                   '☜: Unload Comproxy  
        
	Response.Write "<Script language=vbs> " & vbCr 
	Response.Write "With parent " & vbCr
	Response.Write "If """ & lgIntFlgMode & """ = """ & OPMD_CMODE & """ Then " & vbCr
	Response.Write ".frm1.txtMvmtNo.Value = """ & UCase(ConvSPChars(E1_m_pur_goods_mvmt(0))) & """" & vbCr
	Response.Write "End If"				& vbCr	
    Response.Write ".DbSaveOk "      & vbCr						
    Response.Write "End With " & vbCr
    Response.Write "</Script> "    
    
End Sub	

'-----------------------
'Display CodeName
'2002/08/19 Kim Jin Ha
'-----------------------
Sub DisplaySupplierNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)

	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT BP_NM FROM B_BIZ_PARTNER " 
	lgStrSQL = lgStrSQL & " WHERE BP_CD =  " & FilterVar(inCode , "''", "S") & " AND Bp_Type <> " & FilterVar("C", "''", "S") & "  AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND IN_OUT_FLAG = " & FilterVar("I", "''", "S") & "  "		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtSupplierNm.value	=	""" & lgObjRs("BP_NM") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
	    
	    Call SubCloseRs(lgObjRs)  
	Else
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtSupplierCd.value	=	"""" " & vbCr
		Response.Write "	.txtSupplierNm.value	=	"""" " & vbCr
		Response.Write "	.txtSupplierCd.focus			 " & vbCr

		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
		
	End if
	
	Call SubCloseDB(lgObjConn)
	
End Sub 

'-----------------------
'Display DisplayGroupNm
'2002/08/19 Kim Jin Ha
'-----------------------
Sub DisplayGroupNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT PUR_GRP_NM FROM B_PUR_GRP " 
	lgStrSQL = lgStrSQL & " WHERE PUR_GRP =  " & FilterVar(inCode , "''", "S") & " AND USAGE_FLG=" & FilterVar("Y", "''", "S") & " "		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtGroupNm.value	=	""" & lgObjRs("PUR_GRP_NM") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
			
	Else
		Call DisplayMsgBox("125100", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtGroupCd.value	=	"""" " & vbCr
		Response.Write "	.txtGroupNm.value	=	"""" " & vbCr
		Response.Write "	.txtGroupCd.focus			 " & vbCr

		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
	End if
	
	Call SubCloseRs(lgObjRs)  
	Call SubCloseDB(lgObjConn)
End Sub 

Sub DisplayMvmtNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " select c.io_type_nm from "
	lgStrSQL = lgStrSQL & " (select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b " 
	lgStrSQL = lgStrSQL & " where a.rcpt_type = b.io_type_cd    and a.sto_flg = " & FilterVar("Y", "''", "S") & "  AND a.USAGE_FLG=" & FilterVar("Y", "''", "S") & " ) c "
	lgStrSQL = lgStrSQL & " where c.io_type_cd =  " & FilterVar(incode, "''", "S") & ""

	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtMvmtTypeNm.value	=	""" & lgObjRs("io_type_nm") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
			
	Else
		Call DisplayMsgBox("171900", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtMvmtType.value	=	"""" " & vbCr
		Response.Write "	.txtMvmtTypeNm.value	=	"""" " & vbCr
		Response.Write "	.txtMvmtType.focus			 " & vbCr
		

		Response.Write "End With" & vbCr

		Response.Write "</Script>" & vbCr
	End if
	
	Call SubCloseRs(lgObjRs)  
	Call SubCloseDB(lgObjConn)
End Sub 




%>
