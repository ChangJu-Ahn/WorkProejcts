<%

'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3111mb2.asp	
'*  4. Program Name         : 단가확정 
'*  5. Program Desc         : 단가확정 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status
Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------

lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()        
    End Select


'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    


'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub


'============================================================================================================
Sub SubBizQueryMulti()
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
	
    Dim iObjPS3G142       
    
    '-----------------------------------------------
    ' Declare User Variable
    '-----------------------------------------------
    ' 수주번호 / 수주일 / 주문처 / 단가확정여부 
    Dim i1_s_so_hdr1(1)    
    Dim i3_ief_supplied1    
    Dim i5_b_biz_partner1
    Dim i6_s_wks_date1    
    Dim   i3_price_flag1 
    Dim   i3_base_date1
    ReDim i6_s_wks_date1(2)

    Const S348_I4_from_date = 0    ' I6_s_wks_date1 ==> Constrant
    Const S348_I4_to_date = 1
    Const S348_I4_on_date = 2
    
    ' Next Page Variable
    Dim i2_s_so_dtl1(2)
    Dim i4_s_so_hdr1

    ' Reruen Call Variable
    Dim i1_s_so_hdr
    Dim i2_s_so_dtl
    Dim i3_ief_supplied
    Dim i4_s_so_hdr
    Dim i5_b_biz_partner
    Dim i6_s_wks_date

    ' Export Variables
    Dim exp_s_so_dtl
    Dim exp_b_biz_partner
    Dim exp_grp
    Dim exp_s_so_hdr    

    Dim intGroupCount
    Dim StrNextKey  	
    Dim arrValue
    
    Const C_SHEETMAXROWS_D  = 100
    
    ' exp_grp 저장 
    Const c_exp_s_item_sales_price_item_price = 0    
    Const c_exp_s_so_dtl_so_price = 1    
    Const c_exp_s_so_dtl_so_unit = 2
    Const c_exp_s_so_dtl_price_flag = 3
    Const c_exp_s_so_dtl_so_seq = 4
    Const c_exp_s_so_dtl_net_amt = 5
    Const c_exp_s_so_dtl_net_amt_loc = 6
    Const c_exp_s_so_dtl_vat_amt = 7
    Const c_exp_s_so_dtl_vat_amt_loc = 8
    Const c_exp_s_so_dtl_so_qty = 9
    Const c_exp_b_item_item_cd = 10    
    Const c_exp_b_item_item_nm = 11
    Const c_exp_b_minor1_minor_nm = 12    
    Const c_exp_b_minor2_minor_nm = 13    
    Const c_exp_b_biz_partner1_bp_cd = 14    
    Const c_exp_b_biz_partner1_bp_nm = 15
    Const c_exp_s_so_hdr_so_no = 16    
    Const c_exp_s_so_hdr_cur = 17
    Const c_exp_s_so_hdr_deal_type = 18
    Const c_exp_s_so_hdr_pay_meth = 19
    Const c_exp_s_so_hdr_so_dt = 20
    Const c_exp_b_biz_partner2_bp_cd = 21    
    Const c_exp_b_biz_partner2_bp_nm = 22

'2002-12-23 추가    
    Const S348_EG1_E9_sale_grp = 23
	Const S348_EG1_E9_sale_grp_nm = 24
	Const S348_EG1_E10_plant_cd = 25
	Const S348_EG1_E10_plant_nm = 26
	Const c_exp_b_item_spec = 27
	Const c_exp_s_so_dtl_tracking_no = 28
	
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    ' ----------- Condition Field ------------------
    ' 수주번호 / 수주일 / 주문처 / 단가확정여부	
    '-----------------------------------------------
    ' 수주번호 
    '-----------------------------------------------
    i1_s_so_hdr1(0) = Trim(Request("txtSoNo"))
    i1_s_so_hdr1(1) = Trim(Request("txtSalesGrp"))
    i2_s_so_dtl1(1) = Trim(Request("txtPlant"))
    i2_s_so_dtl1(2) = Trim(Request("txtTrackingNo"))
    '-----------------------------------------------
    ' 수주일 
    ' 0 : 시작일자 1 : 종료일자 2 : 서버 날짜 
    '-----------------------------------------------
    i6_s_wks_date1(S348_I4_from_date) = UNIConvDate(Request("txtFromDate"))
    i6_s_wks_date1(S348_I4_to_date) = UNIConvDate(Request("txtToDate"))
    i6_s_wks_date1(S348_I4_on_date) = ""
    '-----------------------------------------------
    ' 주문처 
    '-----------------------------------------------
    i5_b_biz_partner1 = Trim(Request("txtSoldToParty"))
    '-----------------------------------------------
    ' 단가확정여부	
    '-----------------------------------------------
    i3_ief_supplied1 = Trim(Request("txtPostFlag"))
    '-----------------------------------------------
    ' 단가적용규칙여부	
    '-----------------------------------------------
    i3_price_flag1 = Trim(Request("txtPriceFlag"))
    '-----------------------------------------------
    ' 단가규칙적용기준일 
    '-----------------------------------------------
    i3_base_date1 = UNIConvDate(Request("txtBaseDate"))

    'Test Mode --- Condition variables    

	iStrPrevKey = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	

	If iStrPrevKey <> "" then					
	    
		arrValue = Split(iStrPrevKey, gColSep)

        i4_s_so_hdr1 = arrValue(0)
        i2_s_so_dtl1(0) = arrValue(1)
        
	else			

        i2_s_so_dtl1(0) = ""
        i4_s_so_hdr1 = ""

	End If
	
    i1_s_so_hdr			= i1_s_so_hdr1		
    i2_s_so_dtl         = i2_s_so_dtl1        
    i3_ief_supplied     = i3_ief_supplied1    
    i4_s_so_hdr         = i4_s_so_hdr1        
    i5_b_biz_partner    = i5_b_biz_partner1   
    i6_s_wks_date       = i6_s_wks_date1       
 
 
	set iObjPS3G142 = CREATEOBJECT("PS3G142.cCsLcHdrSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
	  		
	' Call the Dll   		
    Call iObjPS3G142.S_FIX_SALES_PRICE_SVR ( gStrGlobalCollection, C_SHEETMAXROWS_D, _
										 I1_s_so_hdr, I2_s_so_dtl, I3_ief_supplied, _
										 I4_s_so_hdr, I5_b_biz_partner, I6_s_wks_date, _
										 E2_s_so_dtl, E3_b_biz_partner, exp_grp, E4_s_so_hdr,  I3_price_flag1, I3_base_date1 )
    

	If CheckSYSTEMError(Err,TRUE) = True Then
        Response.Write "<Script language=vbs>  " & vbCr   
        Response.Write " Parent.frm1.txtSoldToPartyNm.value   = """ & ConvSPChars(Trim(E3_b_biz_partner(1))) & """" & vbCr
        Response.Write "   Parent.frm1.txtSoNo.focus " & vbCr    
        Response.Write "	parent.ButtonVisible(2) " & vbCr
        Response.Write "</Script>      " & vbCr
       Set iObjPS3G142 = Nothing		         
       Exit Sub
    End If   
    
    Set iObjPS3G142 = Nothing	
        
    iLngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count      
    
        
    For iLngRow = 0 To UBound(exp_grp,1)
    		
        If  iLngRow < C_SHEETMAXROWS_D  Then
        Else

            StrNextKey = ConvSPChars(exp_grp(iLngRow, c_exp_s_so_hdr_so_no))
            StrNextKey = StrNextKey & gColSep & UniconvNum(exp_grp(iLngRow, c_exp_s_so_dtl_so_seq), 0)
            
            Exit For
        End If 

		'선택여부 
        istrdata = istrdata & Chr(11) & "0"

        '단가여부 
        If ConvSPChars(exp_grp(iLngRow, c_exp_s_so_dtl_price_flag)) = "Y" then
            istrdata = istrdata & Chr(11) & "확정"
        Else
            istrdata = istrdata & Chr(11) & "미확정"
        End If
                
		'수주번호 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_s_so_hdr_so_no))
		'수주순번 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_s_so_dtl_so_seq))
		'품목코드 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_item_item_cd))
		'품목명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_item_item_nm))
		'가단가 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_s_so_dtl_so_price), 0)
		'진단가 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_s_item_sales_price_item_price), 0)
		'주문처 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_biz_partner1_bp_cd))
		'주문처명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_biz_partner1_bp_nm))
		'수주일 
		istrdata = istrdata & Chr(11) & UNIDateClientFormat(exp_grp(iLngRow, c_exp_s_so_hdr_so_dt))
		'거래유형 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_s_so_hdr_deal_type))
		'거래유형명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_minor1_minor_nm))
		'결제방법 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_s_so_hdr_pay_meth))
		'결제방법명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_minor2_minor_nm))
		'발행처 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_biz_partner2_bp_cd))
		'발행처명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_biz_partner2_bp_nm))
		'수주금액 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_s_so_dtl_net_amt), 0)
		'화폐 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_s_so_hdr_cur))
		
		'영업그룹 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, S348_EG1_E9_sale_grp))
		'영업그룹명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, S348_EG1_E9_sale_grp_nm))
		'공장 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, S348_EG1_E10_plant_cd))
		'공장명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, S348_EG1_E10_plant_nm))
		'규격. 2003-06-11 추가 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_item_spec ))
		'수주단위 2005 02 22 추가 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_s_so_dtl_so_unit ))
		'Tracking No 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_s_so_dtl_tracking_no ))
				        
        ' 공통 구문 
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)       
            
    Next
    
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write " Parent.frm1.txtSoldToPartyNm.value   = """ & Trim(E3_b_biz_partner(1)) & """" & vbCr        
    Response.Write " Parent.frm1.txtHSoldToParty.value   = """ & Trim(Request("txtSoldToParty")) & """" & vbCr    
    Response.Write " Parent.frm1.txtHFromDate.value = """ & Trim(Request("txtFromDate")) & """" & vbCr    
    Response.Write " Parent.frm1.txtHToDate.value   = """ & Trim(Request("txtToDate")) & """" & vbCr
    Response.Write " Parent.frm1.txtHBaseDate.value   = """ & Trim(Request("txtBaseDate")) & """" & vbCr        
    Response.Write " Parent.frm1.txtHSoNo.value = """ & Trim(ConvSPChars(Request("txtSONo"))) & """" & vbCr            
		
    Response.Write " Parent.frm1.vspdData.ReDraw = False"	& vbCr		        
	Response.Write " Parent.SetSpreadColor -1, -1"				& vbCr    
	
    Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData									      " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip        """ & istrData & """, ""F""" & vbCr
            
    Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,-1,-1,Parent.C_Currency,Parent.C_PriceFlagY,   ""C"" ,""I"",""X"",""X"")" & vbCr
    Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,-1,-1,Parent.C_Currency,Parent.C_PriceFlagN,	""C"" ,""I"",""X"",""X"")" & vbCr
    Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,-1,-1,Parent.C_Currency,Parent.C_NetAmt,	""A"" ,""I"",""X"",""X"")" & vbCr
    
    Response.Write " Parent.frm1.vspdData.ReDraw = True"	& vbCr   
    Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey										 & """" & vbCr  
    Response.Write " Parent.DbQueryOk "																			    	& vbCr   
    Response.Write "</Script> "																							& vbCr      
    
    Exit sub 
                   	
End Sub    


'============================================================================================================
Sub SubBizSaveMulti()   

	Dim iObjPS3G141	
	Dim iErrorPosition
		
	On Error Resume Next                                                                 '☜: Protect system from crashing
	    
	Set iObjPS3G141 = Server.CreateObject("PS3G141.CFixSalesPrice")

	If CheckSYSTEMError(Err,True) = True Then
	   Exit Sub
	End If
	   
	Call iObjPS3G141.FixSalesPriceSvr (gStrGlobalCollection, _
	                                   Trim(Request("txtSpread")), iErrorPosition)    
											      
	Set iObjPS3G141 = Nothing

	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
	   Exit Sub
	End If
	                                                       
	Response.Write "<Script language=vbs> " & vbCr         
	Response.Write " Parent.DBSaveOk "      & vbCr   
	Response.Write "</Script> "             & vbCr   
                  
End Sub

'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

