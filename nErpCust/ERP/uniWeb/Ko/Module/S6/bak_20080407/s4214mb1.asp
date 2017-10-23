<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4214mb1.asp																*
'*  4. Program Name         : Container 배정 
'*  5. Program Desc         : 																*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2005/01/25																*
'*  8. Modified date(Last)  : 																*
'*  9. Modifier (First)     : HJO																*
'* 10. Modifier (Last)      : 																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'*							  2. 2000/04/17 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%												

On Error Resume Next
Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iLngRow
Dim intGroupCount
Dim iPS6G241		

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 
lgStrPrevKey = Request("lgStrPrevKey")
'Response.Write lgstrprevkey & "::"


    
Select Case strMode
Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

	Err.Clear											

    Dim I1_s_cc_hdr_cc_no	
    Dim EG1_s_cc_hdr
    Dim pvCommandSent
        
'    Const E1_cc_no = 0    '[CONVERSION INFORMATION]  View Name : exp s_cc_hdr
 '   Const E1_iv_no = 1
  
    
    Const E1_cc_no = 0
    Const E1_applicant = 1
    Const E2_applicant_nm = 2
    Const E1_iv_no = 3
    Const E1_iv_dt = 4
    Const E1_ship_dt = 5
    Const E1_carton_cnt = 6
    Const E1_tot_packing_cnt = 7
    Const E1_gross_weight= 8
    Const E1_net_weight= 9
    Const E1_measurement = 10

    pvCommandSent = "QUERY"
                															
	I1_s_cc_hdr_cc_no = Trim(Request("txtCCNo"))
	'Response.Write I1_s_cc_hdr_cc_no & ":::I1_s_cc_hdr_cc_no" 
	'---------------------------------- C/C Header Data Query ----------------------------------

	Set iPS6G241 = Server.CreateObject("PS6G241.cSLkExportCcAssignSvr")

'Response.Write err.Description & ":::err"
	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
      'Exit Sub
   End If   

	  		
	Call iPS6G241.S_LOOKUP_EXPORT_CC_HDR_SVR(gStrGlobalCollection, pvCommandSent, I1_s_cc_hdr_cc_no, _
	                                         EG1_s_cc_hdr)
	                                           	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPS6G241 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End
       'Exit Sub
    End If   



    'Set iPS6G241 = Nothing   
	'Dim lgCurrency
	'lgCurrency = ConvSPChars(EG1_s_cc_hdr(E1_cur))
		

	Response.Write "<Script language=vbs>  " & vbCr   			    

	Response.Write " Parent.CurFormatNumericOCX  " & vbCr   		
	Response.Write " Parent.frm1.txtHCCNo.value			= """ & ConvSPChars(Request("txtCCNo"))                                                     & """" & vbcr		
	
	Response.Write " Parent.frm1.txtApplicant.value	= """ & ConvSPChars(EG1_s_cc_hdr(E1_applicant))			& """" & vbcr
	Response.Write " Parent.frm1.txtApplicantNm.value	= """ & ConvSPChars(EG1_s_cc_hdr(E2_applicant_nm))         & """" & vbcr
	Response.Write " Parent.frm1.txtIvNo.value			= """ & ConvSPChars(EG1_s_cc_hdr(E1_iv_no ))         & """" & vbcr
	
	Response.Write " Parent.frm1.txtIvDt.text		= """ & UNIDateClientFormat(EG1_s_cc_hdr(E1_iv_dt ))         & """" & vbcr	
	Response.Write " Parent.frm1.txtShipDt.text			= """ & UNIDateClientFormat(EG1_s_cc_hdr(E1_ship_dt ))         & """" & vbcr
	
	Response.Write " Parent.frm1.txtCarton.value			= """ & UNINumClientFormat(EG1_s_cc_hdr(E1_carton_cnt ), ggQty.DecPoint, 0)    & """" & vbcr
	Response.Write " Parent.frm1.txtPacking.value			= """ & UNINumClientFormat(EG1_s_cc_hdr(E1_tot_packing_cnt), ggQty.DecPoint, 0)    & """" & vbcr	
	Response.Write " Parent.frm1.txtNetW.value		= """ & UNINumClientFormat(EG1_s_cc_hdr(E1_net_weight), ggQty.DecPoint, 0)    & """" & vbcr
	Response.Write " Parent.frm1.txtGrossW.value		= """ & UNINumClientFormat(EG1_s_cc_hdr(E1_gross_weight), ggQty.DecPoint, 0)    & """" & vbcr		
	Response.Write " Parent.frm1.txtMsmnt.value			= """ & UNINumClientFormat(EG1_s_cc_hdr(E1_measurement), ggQty.DecPoint, 0)    & """" & vbcr	

	

	Response.Write " If Parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then Parent.CurFormatNumSprSheet " & vbCr   
	Response.Write " Call parent.CCHdrQueryOk() " & vbCr   
	Response.Write "</Script>      " & vbCr      
'----------------------------------detail query------------------------------------------
    
    Dim I2_s_cc_hdr
    Dim I1_s_cc_dtl
         
    Const C_SHEETMAXROWS_D  = 100
        
    Dim EG1_exp_grp
    
    Const EG1_E1_cont_no = 0
    Const EG1_E1_start_ct_no = 1
    Const EG1_E1_end_ct_no = 2
    Const EG1_E1_packing_cnt = 3
    Const EG1_E1_net_weight = 4
    Const EG1_E1_gross_weight = 5
    Const EG1_E1_measurement = 6
    Const EG1_E1_qty = 7
    Const EG1_E2_unit = 8
    Const EG1_E2_hs_cd = 9
    Const EG1_E3_item_cd = 10
    Const EG1_E3_item_nm = 11
    Const EG1_E3_spec = 12
    Const EG1_E2_lan_no = 13
    Const EG1_E2_plant_cd = 14
    Const EG1_E4_plant_nm = 15
    Const EG1_E3_dn_no = 16
    Const EG1_E3_dn_seq = 17
    Const EG1_E3_so_no = 18
    Const EG1_E3_so_seq = 19
    Const EG1_E3_so_schd_no = 20
    Const EG1_E3_cc_seq = 21
        
    Dim LngLastRow      
    Dim LngMaxRow       
        
    Dim strTemp
    Dim strData
    Dim iStrNextKey
        
    I2_s_cc_hdr = Trim(Request("txtCCNo"))     
        
    If Request("lgStrPrevKey") <> "" then
	  I1_s_cc_dtl = Request("lgStrPrevKey")
    Else
	  I1_s_cc_dtl = 0
    End If
   
	'Set iPS6G241 = Server.CreateObject("PS6G228.cSLtExportCcDtlSvr")          
 
'	If CheckSYSTEMError(Err,True) = True Then
'	   Response.End
 '      'Exit Sub
  '  End If   


	Call iPS6G241.S_LIST_EXPORT_CC_ASSIGN_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I2_s_cc_hdr, _
	                                       I1_s_cc_dtl, EG1_exp_grp)
	                                           	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPS6G241 = Nothing	
       Response.Write "<Script language=vbs>  " & vbCr   
       Response.Write "   Parent.frm1.txtCCNo.focus " & vbCr    
       Response.Write "</Script>      " & vbCr
       Response.End
       'Exit Sub
    End If   

    Set iPS6G241 = Nothing   
                
    LngMaxRow = CLng(Request("txtMaxRows"))										

	For iLngRow = 0 To UBound(EG1_exp_grp,1)
	    If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_cc_seq)) 
           Exit For
        End If  
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E1_cont_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,EG1_E1_start_ct_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E1_end_ct_no ))
        If UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_packing_cnt ), ggQty.DecPoint, 0) <>  0.00 then
        strData = strData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_packing_cnt ), ggQty.DecPoint, 0)
        Else 
        strData = strData & Chr(11) &""
        End If
        If UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_net_weight), ggQty.DecPoint, 0) <>0.00 then
        strData = strData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_net_weight), ggQty.DecPoint, 0)
        Else
        strData = strData & Chr(11) &""
        End If
        If UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_gross_weight ), ggQty.DecPoint, 0) <>0 then
        strData = strData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_gross_weight ), ggQty.DecPoint, 0)        
        Else
        strData = strData & Chr(11) &""        
        End If
        If UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_measurement ), ggQty.DecPoint, 0) <> 0 Then
        strData = strData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_measurement ), ggQty.DecPoint, 0)
        Else
        strData = strData & Chr(11) &""        
        End If
       If UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_qty ), ggQty.DecPoint, 0) <>0 Then
       strData = strData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E1_qty ), ggQty.DecPoint, 0)
       Else
       strData = strData & Chr(11) &""
       End if
        
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E2_unit ))                
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E2_hs_cd))        
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_item_cd ))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_item_nm))        
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_spec))        
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E2_lan_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E2_plant_cd))     
        
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E4_plant_nm))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_dn_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_dn_seq))
        
		strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_so_no ))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_so_seq ))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_so_schd_no))        
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_cc_seq))
        strData = strData & Chr(11) & LngMaxRow + iLngRow
        strData = strData & Chr(11) & Chr(12)
    
    Next            
    
    Response.Write "<Script language=vbs> " & vbCr           
    Response.Write " Parent.frm1.vspdData.ReDraw = False " & vbCr
    Response.Write " Parent.ggoSpread.Source	= Parent.frm1.vspdData " &	 	  vbCr
    Response.Write " Parent.ggoSpread.SSShowData      """ & strData	 & """" & vbCr
    
    Response.Write " Parent.lgStrPrevKey              = """ & iStrNextKey						& """" & vbCr  
    Response.Write " Parent.frm1.txtHCCNo.value       = """ & ConvSPChars(Request("txtCCNo"))   & """" & vbCr                
    Response.Write " Parent.DbQueryOk "															& vbCr   
    Response.Write "</Script> "																	& vbCr      
    
Case CStr(UID_M0002)														'☜: 현재 Save 요청을 받음	
    Dim iErrorPosition		 
		 
    Set iPS6G241 = Server.CreateObject("PS6G241.cSExportCcAssignSvr")       
   
	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
    End If   
    
	Dim reqtxtSpread
	Dim arrRowVal
	Dim count
	Dim arrSumData
	
'	arrSumData = request("txtPacKing")
'	arrSumData = arrSumData & "::" & request("txtGrossW")
'	arrSumData = arrSumData & "::" & request("txtNetW")  
'	arrSumData = arrSumData & "::" & request("txtMsmnt") 
'	arrSumData = arrSumData & "::" & request("txtCarton") 
	
	reqtxtSpread = Request("txtSpread")
	Call iPS6G241.S_MAINT_EXPORT_CC_ASSIGN_SVR(gStrGlobalCollection, Trim(ucase(Request("txtCCNo"))),  _
	                                        Trim(reqtxtSpread), iErrorPosition)
	                                                     
	                                           	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행의 PACKING수량","통관수량","","","") = True Then
       Set iPS6G241 = Nothing
       Response.End
	End If  

    Set iPS6G241 = Nothing 

	Response.Write "<Script language=vbs> " & vbCr      
	Response.Write " Parent.DBSaveOk "		& vbCr   
	Response.Write "</Script> "				& vbCr      													

Case Else
	Response.End
End Select
%>
