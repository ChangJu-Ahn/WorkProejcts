<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3112mb2.asp	
'*  4. Program Name         : 수주마감 
'*  5. Program Desc         : 수주마감 
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
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
    Call HideStatusWnd                                                               '☜: Hide Processing message
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query			
             Call SubBizQuery()
        Case CStr(UID_M0002)														 '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	Const C_SHEETMAXROWS_D  = 100
	
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
	    
    Dim StrNextKey  	
    Dim arrValue    
    
    Dim I1_s_so_hdr 'imp_next_so_no
    Dim I2_s_wks_date(2)
    Dim I2_s_wks_date1
    Const S345_I2_from_date = 0   
	Const S345_I2_to_date = 1
	
    Dim I3_ief_supplied 'close_flag
    Dim I4_b_biz_partner  'bp_cd
    Dim I5_s_so_hdr 'so_no
    Dim I6_s_so_dtl 'imp_next_so_seq
    
    Dim I7_b_plant_plant_cd
    Dim I8_b_slaes_grp_sales_grp
    Dim I9_s_so_dtl_so_status
    Dim I10_s_so_dtl_Bal_flag
    
    Dim E2_b_biz_partner
	Const S345_E2_bp_cd = 0   
	Const S345_E2_bp_nm = 1
	
	Dim E3_b_plant
    Const S345_E3_plant_cd = 0
    Const S345_E3_plant_nm = 1
    
    Dim E4_b_slaes_grp
    Const S345_E4_sales_grp = 0
    Const S345_E4_sales_grp_nm = 1
    
    Dim E5_s_so_dtl
    Const S345_E5_SoStatus = 0

    Dim EG1_exp_grp
    Const S345_EG1_E1_so_no = 0    
	Const S345_EG1_E2_bp_cd = 1    
	Const S345_EG1_E2_bp_nm = 2
	Const S345_EG1_E3_item_cd = 3   
	Const S345_EG1_E3_item_nm = 4
	Const S345_EG1_E4_so_seq = 5    
	Const S345_EG1_E4_close_flag = 6
	Const S345_EG1_E4_so_qty = 7
	Const S345_EG1_E4_bonus_qty = 8
	Const S345_EG1_E4_req_qty = 9
	Const S345_EG1_E4_req_bonus_qty = 10
	Const S345_EG1_E4_gi_qty = 11
	Const S345_EG1_E4_gi_bonus_qty = 12
	Const S345_EG1_E4_bill_qty = 13
	Const S345_EG1_E4_lc_qty = 14
	Const S345_EG1_E4_so_unit = 15
	Const S345_EG1_E4_so_status = 16
	Const S345_EG1_E4_dlvy_dt = 17
	
	Const S345_EG1_E4_sales_grp = 18
    Const S345_EG1_E4_sales_grp_nm = 19
    Const S345_EG1_E4_plant_cd = 20
    Const S345_EG1_E4_plant_nm = 21
    Const S345_EG1_E3_ItemSpec = 22

	Dim PS3G132

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                       '☜: Clear Error status
            
            
            
	I5_s_so_hdr = Trim(Request("txtConSoNo"))
	I3_ief_supplied = Trim(Request("txtCfmFlag"))
	I4_b_biz_partner = Trim(Request("txtSoldToParty"))
	
	I7_b_plant_plant_cd = Trim(Request("txtPlant"))
	I8_b_slaes_grp_sales_grp = Trim(Request("txtSalesGrp"))
	
	I2_s_wks_date(S345_I2_from_date) = UNIConvDate(Request("txtConSoFrDt"))
	I2_s_wks_date(S345_I2_to_date)   = UNIConvDate(Request("txtConSoToDt"))
	
	I2_s_wks_date1= I2_s_wks_date
	
	I9_s_so_dtl_so_status = Trim(Request("txtStatusFlag"))
    
    I10_s_so_dtl_Bal_flag = Trim(Request("txtBOFlag"))
        
    iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	    
  	
  	If iStrPrevKey <> "" then					
		arrValue = Split(iStrPrevKey, gColSep)
		I1_s_so_hdr = Trim(arrValue(0))			
		I6_s_so_dtl = Trim(arrValue(1))							
	else			
		I1_s_so_hdr = ""
		I6_s_so_dtl = "0"
	End If 
	
	Set PS3G132 = Server.CreateObject("PS3G132.CSListSoDtlForClose")
	
	if CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	end if
   
	Call PS3G132.S_LIST_SO_DTL_FOR_CLOSING(gStrGlobalCollection, Cint(C_SHEETMAXROWS_D), I1_s_so_hdr, I2_s_wks_date1 , _
        I3_ief_supplied, I4_b_biz_partner,I5_s_so_hdr, I6_s_so_dtl, I7_b_plant_plant_cd ,  I8_b_slaes_grp_sales_grp, I9_s_so_dtl_so_status,  _
        I10_s_so_dtl_Bal_flag, E2_b_biz_partner, E3_b_plant, E4_b_slaes_grp, E5_s_so_dtl, EG1_exp_grp)        		    	
    
    If CheckSYSTEMError(Err,True) = True Then 		
    	Response.Write "<Script language=vbs> " & vbCr       
		Response.Write " Parent.frm1.txtSoldToPartyNm.value  = """ & ConvSPChars(E2_b_biz_partner(S345_E2_bp_nm))     & """" & vbCr    
		Response.Write " Parent.frm1.txtPlantNm.value  = """ & ConvSPChars(E3_b_plant(S345_E3_plant_nm))     & """" & vbCr    
		Response.Write " Parent.frm1.txtSalesGrpNm.value  = """ & ConvSPChars(E4_b_slaes_grp(S345_E4_sales_grp_nm))     & """" & vbCr    
		Response.Write " Parent.frm1.txtStatusFlag.value  = """ & ConvSPChars(E5_s_so_dtl(S345_E5_SoStatus))  & """" & vbCr	
						
		Response.Write " Parent.frm1.txtHSoldToParty.value  = """ & ConvSPChars(Request("txtSoldToParty"))  & """" & vbCr		
		Response.Write " Parent.frm1.txtHPlant.value  = """ & ConvSPChars(Request("txtPlant"))  & """" & vbCr		
		Response.Write " Parent.frm1.txtHSalesGrp.value  = """ & ConvSPChars(Request("txtSalesGrp"))  & """" & vbCr	
		Response.Write " Parent.frm1.txtHStatusFlag.value  = """ & ConvSPChars(Request("txtStatusFlag"))  & """" & vbCr	
		Response.Write " Parent.frm1.txtHBOFlag.value  = """ & ConvSPChars(Request("txtBOFlag"))  & """" & vbCr	
		
        Response.Write "   Parent.frm1.txtConSoNo.focus " & vbCr    
		
		Response.Write "</Script> "																							& vbCr                  
		Set PS3G132 = Nothing
		Exit Sub	
	End If
		
	iLngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count          			
	istrData= ""	
	For iLngRow = 0 To UBound(EG1_exp_grp,1)					
		
		If  iLngRow < C_SHEETMAXROWS_D  Then			
		Else 
		   
		   StrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, S345_EG1_E1_so_no))                    & gColSep '0		   		   
		   StrNextKey = StrNextKey & ConvSPChars(EG1_exp_grp(iLngRow, S345_EG1_E4_so_seq))      & gColSep '1		   				   
           Exit For
        End If  
      
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exp_grp( iLngRow,S345_EG1_E4_close_flag )))
		
        istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exp_grp(iLngRow,S345_EG1_E1_so_no )))
		istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exp_grp(iLngRow,S345_EG1_E4_so_seq)))
		
        istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exp_grp(iLngRow,S345_EG1_E3_item_cd )))        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E3_item_nm ))        
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E4_so_unit))		
		
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S345_EG1_E4_so_qty), ggQty.DecPoint, 0)						
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S345_EG1_E4_bonus_qty), ggQty.DecPoint, 0)    
		istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(iLngRow, S345_EG1_E4_dlvy_dt ))		
		
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E2_bp_cd))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E2_bp_nm))
        
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S345_EG1_E4_lc_qty), ggQty.DecPoint, 0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S345_EG1_E4_req_qty ), ggQty.DecPoint, 0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S345_EG1_E4_req_bonus_qty), ggQty.DecPoint, 0)	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S345_EG1_E4_gi_qty), ggQty.DecPoint, 0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S345_EG1_E4_gi_bonus_qty ), ggQty.DecPoint, 0)	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S345_EG1_E4_bill_qty ), ggQty.DecPoint, 0)        	
    		
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E4_sales_grp))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E4_sales_grp_nm))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E4_plant_cd))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E4_plant_nm))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S345_EG1_E3_ItemSpec ))	
		
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)             
        
    Next    
    
    
    
    
    Response.Write "<Script language=vbs> " & vbCr
    Response.Write " Parent.frm1.txtSoldToParty.value    = """ & ConvSPChars(E2_b_biz_partner(S345_E2_bp_cd))      & """" & vbCr    
	Response.Write " Parent.frm1.txtSoldToPartyNm.value  = """ & ConvSPChars(E2_b_biz_partner(S345_E2_bp_nm))      & """" & vbCr    
	Response.Write " Parent.frm1.txtHSoldToParty.value   = """ & ConvSPChars(Request("txtSoldToParty"))  & """" & vbCr
	Response.Write " Parent.frm1.txtPlantNm.value		 = """ & ConvSPChars(E3_b_plant(S345_E3_plant_nm))		   & """" & vbCr    
	Response.Write " Parent.frm1.txtHPlant.value		 = """ & ConvSPChars(Request("txtPlant"))				   & """" & vbCr		
		
	Response.Write " Parent.frm1.txtSalesGrpNm.value	 = """ & ConvSPChars(E4_b_slaes_grp(S345_E4_sales_grp_nm)) & """" & vbCr    
	Response.Write " Parent.frm1.txtHSalesGrp.value		 = """ & ConvSPChars(Request("txtSalesGrp"))			   & """" & vbCr							
		               			
    Response.Write " Parent.frm1.txtHSoNo.value         = """ & ConvSPChars(Request("txtConSoNo"))	   & """" & vbCr
    Response.Write " Parent.frm1.txtHCfmFlag.value      = """ & ConvSPChars(Request("txtCfmFlag"))	   & """" & vbCr

    Response.Write " Parent.frm1.txtHSoFrDt.value       = """ & Request("txtConSoFrDt")	   & """" & vbCr
    Response.Write " Parent.frm1.txtHSoToDt.value       = """ & Request("txtConSoToDt")    & """" & vbCr
    Response.Write " Parent.frm1.txtHBOFlag.value       = """ & Request("txtBOFlag")    & """" & vbCr
    Response.Write " Parent.frm1.txtHStatusFlag.value       = """ & Request("txtStatusFlag")    & """" & vbCr
    
	
    Response.Write " Parent.frm1.vspdData.ReDraw = False"	& vbCr		        
	Response.Write " Parent.SetSpreadColor -1, -1"				& vbCr    
	Response.Write " Parent.frm1.vspdData.ReDraw = True"	& vbCr
		    
    
    Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData									      " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip        """ & istrData										     & """" & vbCr
    Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey										 & """" & vbCr  
    
    Response.Write " Parent.DbQueryOk "																			    	& vbCr   
    Response.Write "</Script> "																							& vbCr      
	
	Set PS3G132 = Nothing	    
	
End Sub    


Sub SubBizSaveMulti()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Dim PS3G122
	Dim pvCommand
	Dim iErrorPosition
		
	Set PS3G122 = Server.CreateObject("PS3G122.cSCloseSoSvr")	
	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    pvCommand = "SAVE"
    Call PS3G122.S_CLOSE_SO_SVR(gStrGlobalCollection, pvCommand, cstr(Trim(Request("txtSpread"))),iErrorPosition )		
	
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set PS3G122 = Nothing
       Exit Sub
	End If	
	
    '-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	
    Set PS3G122 = Nothing    
    
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "               
End Sub

%>
