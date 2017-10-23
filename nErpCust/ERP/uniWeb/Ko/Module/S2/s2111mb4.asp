<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  4. Program Name         : s2111ma4
'*  5. Program Desc         : 조직별 고객판매계획등록 
'*  5. Program Desc         : 조직별 고객판매계획등록 
'*  6. Comproxy List        : PS2G121.dll, PS2G122.dll, PS2G124.dll
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Mr  Cho
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
    Dim lgOpModeCRUD
    Const lsConfirm  = "CONFIRM"      
	
	On Error Resume Next                                                             '☜: Protect system from crashing 
   
	Call LoadBasisGlobalInf()
	Call loadInfTB19029B( "I", "*", "NOCOOKIE", "MB")     
    Call HideStatusWnd                                                               '☜: Hide Processing message
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query			
             Call SubBizQueryMulti()
        Case CStr(UID_M0002), CStr(UID_M0003)										 '☜: Save,Update, Delete			
             Call SubBizSaveMulti()
        Case CStr(lsConfirm)														 '☜: 확정처리 
			 Call SubConfirm()

    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Dim I1_b_biz_partner 'imp_next	
	
	Dim I2_b_sales_org    'imp	

    Dim I3_s_cust_sales_plan(6) ' imp
    Dim I3_s_cust_sales_plan1 ' imp
    Const S222_I3_sp_year = 0
    Const S222_I3_plan_flag = 1
    Const S222_I3_plan_seq = 2
    Const S222_I3_export_flag = 3
    Const S222_I3_cur = 4
    Const S222_I3_qty_amt_flag = 5
    Const S222_I3_sales_grp = 6

    Dim E1_s_cust_sales_plan 'exp1   
    Const S222_E1_sales_grp = 0
    Const S222_E1_sales_grp_nm = 1
    Const S222_E1_sales_org = 2
    Const S222_E1_sales_org_nm = 3
    Const S222_E1_sp_year = 4
    Const S222_E1_plan_flag = 5
    Const S222_E1_plan_seq = 6
    Const S222_E1_export_flag = 7
    Const S222_E1_plan_flag_nm = 8
    Const S222_E1_export_flag_nm = 9
	'계획차수관련 
	Const S222_E1_plan_seq_nm = 10    

    Dim EG1_exp_grp 'exp2
    Const S222_EG1_E1_s_cust_sales_plan_plan_unit = 0
    Const S222_EG1_E1_s_cust_sales_plan_qty_amt_flag = 1

    Const S222_EG1_E2_s_wks_msp_plan_qty1 = 2
    Const S222_EG1_E2_s_wks_msp_plan_qty2 = 3
    Const S222_EG1_E2_s_wks_msp_plan_qty3 = 4
    Const S222_EG1_E2_s_wks_msp_plan_qty4 = 5
    Const S222_EG1_E2_s_wks_msp_plan_qty5 = 6
    Const S222_EG1_E2_s_wks_msp_plan_qty6 = 7
    Const S222_EG1_E2_s_wks_msp_plan_qty7 = 8
    Const S222_EG1_E2_s_wks_msp_plan_qty8 = 9
    Const S222_EG1_E2_s_wks_msp_plan_qty9 = 10
    Const S222_EG1_E2_s_wks_msp_plan_qty10 = 11
    Const S222_EG1_E2_s_wks_msp_plan_qty11 = 12
    Const S222_EG1_E2_s_wks_msp_plan_qty12 = 13
    Const S222_EG1_E2_s_wks_msp_plan_amt1 = 14
    Const S222_EG1_E2_s_wks_msp_plan_amt2 = 15
    Const S222_EG1_E2_s_wks_msp_plan_amt3 = 16
    Const S222_EG1_E2_s_wks_msp_plan_amt4 = 17
    Const S222_EG1_E2_s_wks_msp_plan_amt5 = 18
    Const S222_EG1_E2_s_wks_msp_plan_amt6 = 19
    Const S222_EG1_E2_s_wks_msp_plan_amt7 = 20
    Const S222_EG1_E2_s_wks_msp_plan_amt8 = 21
    Const S222_EG1_E2_s_wks_msp_plan_amt9 = 22
    Const S222_EG1_E2_s_wks_msp_plan_amt10 = 23
    Const S222_EG1_E2_s_wks_msp_plan_amt11 = 24
    Const S222_EG1_E2_s_wks_msp_plan_amt12 = 25
    Const S222_EG1_E2_s_wks_msp_split_flag1 = 26
    Const S222_EG1_E2_s_wks_msp_split_flag2 = 27
    Const S222_EG1_E2_s_wks_msp_split_flag3 = 28
    Const S222_EG1_E2_s_wks_msp_split_flag4 = 29
    Const S222_EG1_E2_s_wks_msp_split_flag5 = 30
    Const S222_EG1_E2_s_wks_msp_split_flag6 = 31
    Const S222_EG1_E2_s_wks_msp_split_flag7 = 32
    Const S222_EG1_E2_s_wks_msp_split_flag8 = 33
    Const S222_EG1_E2_s_wks_msp_split_flag9 = 34
    Const S222_EG1_E2_s_wks_msp_split_flag10 = 35
    Const S222_EG1_E2_s_wks_msp_split_flag11 = 36
    Const S222_EG1_E2_s_wks_msp_split_flag12 = 37
    Const S222_EG1_E3_b_biz_partner_bp_cd = 38
    Const S222_EG1_E3_b_biz_partner_bp_nm = 39	

	Const C_SHEETMAXROWS_D  = 100
	
	Dim iLngRow	
	Dim iLngMaxRow	
	
	Dim istrData
	Dim iStrPrevKey
	    
    Dim StrNextKey  	
    Dim arrValue    
	
    Dim pS2G122      


    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                       '☜: Clear Error status    
        
    I2_b_sales_org = "G"	
    
    I3_s_cust_sales_plan(S222_I3_sp_year) = Trim(Request("txtConSpYear"))
    I3_s_cust_sales_plan(S222_I3_plan_flag) = Trim(Request("txtConPlanTypeCd"))
    I3_s_cust_sales_plan(S222_I3_export_flag) = Trim(Request("txtConDealTypeCd"))
    I3_s_cust_sales_plan(S222_I3_plan_seq) = Trim(Request("txtConPlanNum"))    
    I3_s_cust_sales_plan(S222_I3_sales_grp) = Trim(Request("txtConSalesOrg"))
    I3_s_cust_sales_plan(S222_I3_cur) = Trim(Request("txtConCurr"))   
        
    iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	      	
  	
  	If iStrPrevKey <> "" then					
		I1_b_biz_partner= iStrPrevKey						
	else			
		I1_b_biz_partner= ""
		
	End If 	
	
	I3_s_cust_sales_plan1 =I3_s_cust_sales_plan	
	
	Set pS2G122 = Server.CreateObject("PS2G122.CsListCustSP")
	
	if CheckSYSTEMError(Err,True) = True Then 
	    Response.Write "<Script language=vbs>  " & vbCr   
		Response.Write " Parent.SetDefaultVal  " & vbCr   
		Response.Write "</Script>      " & vbCr      
		Exit Sub
	end if
   
	Call pS2G122.S_LIST_CUST_SALES_PLAN(gStrGlobalCollection, Cint(C_SHEETMAXROWS_D), I1_b_biz_partner, CStr(I2_b_sales_org), _
         I3_s_cust_sales_plan1, E1_s_cust_sales_plan, EG1_exp_grp )        

	If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "202258" then    
    
		If CheckSYSTEMError(Err,True) = True Then 		
		    Response.Write "<Script language=vbs>  " & vbCr   		
			Response.Write " Parent.frm1.txtConPlanTypeNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan (S222_E1_plan_flag_nm))     & """" & vbCr    
			Response.Write " Parent.frm1.txtConDealTypeNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_export_flag_nm ))   & """" & vbCr        
			Response.Write " Parent.frm1.txtConSalesOrgNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_sales_org_nm))      & """" & vbCr    		
			Response.Write " Parent.frm1.txtConPlanNumNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_plan_seq_nm))      & """" & vbCr    		
		    Response.Write " Parent.SetDefaultVal2  " & vbCr   
			Response.Write "</Script>      " & vbCr      
			Set pS2G122 = Nothing
			Exit Sub
		end if
	Else
		If CheckSYSTEMError(Err,True) = True Then 		
		    Response.Write "<Script language=vbs>  " & vbCr   		
		    Response.Write " Parent.SetDefaultVal  " & vbCr   
			Response.Write " Parent.frm1.txtConPlanTypeNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan (S222_E1_plan_flag_nm))     & """" & vbCr    
			Response.Write " Parent.frm1.txtConDealTypeNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_export_flag_nm ))   & """" & vbCr        
			Response.Write " Parent.frm1.txtConSalesOrgNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_sales_org_nm))      & """" & vbCr    		
			Response.Write " Parent.frm1.txtConPlanNumNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_plan_seq_nm))      & """" & vbCr    		
			Response.Write "</Script>      " & vbCr      
			Set pS2G122 = Nothing
			Exit Sub
		end if
	End If
			
	iLngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count          
	For iLngRow = 0 To UBound(EG1_exp_grp,1)			
		
		If  iLngRow < C_SHEETMAXROWS_D  Then			
		Else '  bp_cd
		   
		   StrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, S222_EG1_E3_b_biz_partner_bp_cd))		   
           Exit For
        End If  				
		
		' 품목 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S222_EG1_E3_b_biz_partner_bp_cd)) 
		istrData = istrData & Chr(11) & ""															
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S222_EG1_E3_b_biz_partner_bp_nm )) 
		  ' 계획단위 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S222_EG1_E1_s_cust_sales_plan_plan_unit  )) 
		istrData = istrData & Chr(11) & ""
		  ' 년계획 수량 합계 
		istrData = istrData & Chr(11) & ""
		  ' 년계획 금액 합계 
		istrData = istrData & Chr(11) & ""														
			
		'1월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty1), ggQty.DecPoint, 0)
		'istrData = istrData & Chr(11) &  UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt1), ggQty.DecPoint, 0)
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt1), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'2월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty2), ggQty.DecPoint, 0)		
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt2), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'3월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty3), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt3), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'4월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty4), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt4), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'5월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty5), ggQty.DecPoint, 0)		
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt5), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'6월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty6), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt6), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'7월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty7), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt7), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'8월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty8), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt8), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'9월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty9), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt9), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'10월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty10), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt10), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'11월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty11), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt11), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
		'12월 수량, 금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_qty12), ggQty.DecPoint, 0)					
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow,S222_EG1_E2_s_wks_msp_plan_amt12), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) 
			
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag1 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag2 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag3 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag4 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag5 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag6 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag7 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag8 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag9 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag10 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag11 )) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp( iLngRow,S222_EG1_E2_s_wks_msp_split_flag12 )) 			
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)               
    Next        
    
    Response.Write "<Script language=vbs> " & vbCr           
    Response.Write " Parent.frm1.vspdData.ReDraw = False " & vbCr
    Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData									      " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip        """ & istrData										     & """" & vbCr
' 추가		
	Response.Write "   Parent.frm1.vspdData.ReDraw = False      " & vbCr	 
	        
        
    For iLngRow = 1   To UBound(EG1_exp_grp,1) +1
		If  iLngRow > C_SHEETMAXROWS_D  Then			
           Exit For
        End If        
		
		Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_ItemCode," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
		Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_ItemName," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
		Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_YearQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
		Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_YearAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				   
 		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag1)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_01PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_01PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_01PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 
		End If
		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag2)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_02PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_02PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_02PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 				
		End If
		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag3)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_03PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_03PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_03PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 				
		End If
				
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag4)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_04PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_04PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_04PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 
		End If
		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag5)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_05PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_05PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_05PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 				
		End If
		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag6)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_06PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_06PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_06PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 				
		End If

		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag7)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_07PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_07PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_07PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 
		End If
		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag8)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_08PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_08PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_08PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 				
		End If
		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag9)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_09PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_09PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_09PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 				
		End If

		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag10)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_10PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_10PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_10PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 
		End If
		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag11)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_11PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_11PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_11PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 				
		End If
		
		If UCase(EG1_exp_grp( iLngRow - 1 ,S222_EG1_E2_s_wks_msp_split_flag12)) = "Y" Then
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_12PlanQty," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				Response.Write " Parent.ggoSpread.SSSetProtected Parent.C_12PlanAmt," & iLngMaxRow +  iLngRow  & " ," & iLngMaxRow + iLngRow  & vbCr
				
			Else
				Response.Write " Call Parent.SplitFlagMonthColor(Parent.C_12PlanQty," & iLngMaxRow +  iLngRow & ")" & vbCr 				
		End If
    
    Next
        
    Response.Write " Parent.frm1.vspdData.ReDraw = True " & vbCr    
    
    Response.Write " Parent.frm1.txtConPlanTypeNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan (S222_E1_plan_flag_nm))     & """" & vbCr    
    Response.Write " Parent.frm1.txtConDealTypeNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_export_flag_nm ))   & """" & vbCr        
    Response.Write " Parent.frm1.txtConSalesOrgNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_sales_org_nm))      & """" & vbCr    
	Response.Write " Parent.frm1.txtConPlanNumNm.value	  = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_plan_seq_nm))       & """" & vbCr    		
    
    Response.Write " Parent.SetDefaultVal " & vbCr            
    Response.Write " Parent.frm1.txtSalesOrg.value   = """ & ConvSPChars(E1_s_cust_sales_plan (S222_E1_sales_org ))     & """" & vbCr    
    Response.Write " Parent.frm1.txtSalesOrgNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_sales_org_nm ))   & """" & vbCr        
    Response.Write " Parent.frm1.txtSpYear.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_sp_year ))      & """" & vbCr    
    Response.Write " Parent.frm1.txtConSpYear.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_sp_year ))      & """" & vbCr    
    
    Response.Write " Parent.frm1.txtPlanTypeCd.value   = """ & ConvSPChars(E1_s_cust_sales_plan (S222_E1_plan_flag))     & """" & vbCr    
    Response.Write " Parent.frm1.txtPlanTypeNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_plan_flag_nm ))   & """" & vbCr        
    Response.Write " Parent.frm1.txtDealTypeCd.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_export_flag))      & """" & vbCr    
    Response.Write " Parent.frm1.txtDealTypeNm.value   = """ & ConvSPChars(E1_s_cust_sales_plan (S222_E1_export_flag_nm))     & """" & vbCr    
    Response.Write " Parent.frm1.txtCurr.value   = """ & gCurrency   & """" & vbCr            
    Response.Write " Parent.frm1.txtPlanNum.value   = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_plan_seq))      & """" & vbCr    
	Response.Write " Parent.frm1.txtPlanNumNm.value	  = """ & ConvSPChars(E1_s_cust_sales_plan(S222_E1_plan_seq_nm))       & """" & vbCr    		    
        
    Response.Write " Parent.frm1.HConSalesOrg.value    = """ & ConvSPChars(Request("txtConSalesOrg"))   & """" & vbCr    
    Response.Write " Parent.frm1.HConSpYear.value      = """ & ConvSPChars(Request("txtConSpYear"))     & """" & vbCr        
    Response.Write " Parent.frm1.HPlanTypeCd.value     = """ & ConvSPChars(Request("txtConPlanTypeCd")) & """" & vbCr    
    Response.Write " Parent.frm1.HConDealTypeCd.value  = """ & ConvSPChars(Request("txtConDealTypeCd")) & """" & vbCr    
    Response.Write " Parent.frm1.HConCurr.value        = """ & ConvSPChars(Request("txtConCurr"))       & """" & vbCr            
    Response.Write " Parent.frm1.HConPlanNum.value     = """ & Request("txtConPlanNum")    & """" & vbCr       			

		
    Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey					& """" & vbCr  
    Response.Write " Parent.DbQueryOk "														& vbCr   
    Response.Write "</Script> "																& vbCr      
    
	Set pS2G122 = Nothing	    
	
End Sub    


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()                                                     '☜: Clear Error status

	Dim itxtSpread
    
    Dim  I1_s_cust_sales_plan      
    Const S218_I1_sp_year = 0
	Const S218_I1_sp_month = 1
	Const S218_I1_plan_flag = 2
	Const S218_I1_plan_seq = 3
	Const S218_I1_export_flag = 4
	Const S218_I1_cur = 5
	Const S218_I1_qty_amt_flag = 6
	Const S218_I1_sale_grp = 7
	Const S218_I1_sale_org = 8
    
    Dim I3_b_sales_org_Div     
    
	Dim E1_s_cust_sales_plan 
	Const S218_E1_plan_seq = 0
	
	Dim E3_b_sales_org
	Const S218_E3_sales_org = 0
	Const S218_E3_sales_org_nm = 1
	
	Dim E4_b_minor    
	Const S218_E4_minor_cd = 0
	Const S218_E4_minor_nm = 1
	
	Dim E5_b_minor    
	Const S218_E5_minor_cd = 0
	Const S218_E5_minor_nm = 1

	
	Dim iErrorPosition	
	Dim PS2G121
	Dim strHang
	
	ReDim  I1_s_cust_sales_plan(8) 
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear  
    
	I3_b_sales_org_Div = "G"	
	iErrorPosition =  ""
    
	I1_s_cust_sales_plan(S218_I1_sp_year) = Trim(Request("txtSpYear"))	
	I1_s_cust_sales_plan(S218_I1_plan_flag) = Trim(Request("txtPlanTypeCd"))
	I1_s_cust_sales_plan(S218_I1_export_flag) = Trim(Request("txtDealTypeCd"))
	I1_s_cust_sales_plan(S218_I1_plan_seq) = Trim(Request("txtPlanNum"))
	I1_s_cust_sales_plan(S218_I1_sale_org) = Trim(Request("txtSalesOrg"))
	I1_s_cust_sales_plan(S218_I1_cur) = gCurrency
	
	I1_s_cust_sales_plan(S218_I1_sale_grp) = ""
	I1_s_cust_sales_plan(S218_I1_qty_amt_flag) = ""
    					
	Set PS2G121 = Server.CreateObject("PS2G121.CsCustSPSv")		
	If CheckSYSTEMError(Err,True) = True Then
      Exit Sub
    End If

	itxtSpread = Trim(Request("txtSpread"))
    
    Call PS2G121.S_CUST_SALES_PLAN_SVR( gStrGlobalCollection,cstr(itxtSpread), I1_s_cust_sales_plan , _
			I3_b_sales_org_Div , E1_s_cust_sales_plan, E3_b_sales_org, E4_b_minor, E5_b_minor,  iErrorPosition)	
	
	If len(Trim(iErrorPosition)) then
		strHang = "행"
	else
		strHang = ""
	end if
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & strHang ,"","","","") = True Then		
		Response.Write "<Script language=vbs> " & vbCr      
		Response.Write " Parent.frm1.txtSalesOrgNm.value    = """ & ConvSPChars(E3_b_sales_org(S218_E3_sales_org_nm)) & """" & vbCr
		Response.Write " Parent.frm1.txtPlanTypeNm.value    = """ & ConvSPChars(E4_b_minor(S218_E4_minor_nm)) & """" & vbCr
		Response.Write " Parent.frm1.txtDealTypeNm.value    = """ & ConvSPChars(E5_b_minor(S218_E5_minor_nm)) & """" & vbCr
		Response.Write " Parent.frm1.txtSalesOrg.focus "	& vbCr      
		
		Response.Write "</Script> "				& vbCr      
       Set PS2G121 = Nothing
       Exit Sub
	End If	

	Response.Write "<Script language=vbs> " & vbCr           	
	If Trim(ConvSPChars(E1_s_cust_sales_plan(S218_E1_plan_seq))) <> "" then
		Response.Write " Parent.frm1.txtPlanNum.value    = """ & ConvSPChars(E1_s_cust_sales_plan(S218_E1_plan_seq)) & """" & vbCr        
		Response.Write " Parent.frm1.txtConPlanNum.value    = """ & ConvSPChars(E1_s_cust_sales_plan(S218_E1_plan_seq)) & """" & vbCr
		Response.Write " Parent.frm1.txtConPlanNum.value    = """ & ConvSPChars(E1_s_cust_sales_plan(S218_E1_plan_seq)) & """" & vbCr
		
	End If
    Response.Write " Parent.DBSaveOk "																			    	& vbCr   
    Response.Write "</Script> "																							& vbCr      
    
	
    '-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
    Set PS2G121 = Nothing   
     

    
End Sub

'==============================================================================
Sub SubConfirm()
	
	Dim I2_s_cust_sales_plan(7)
	Dim I2_s_cust_sales_plan1
	
	Const S223_I2_sp_year = 0
	Const S223_I2_sp_month = 1
	Const S223_I2_plan_flag = 2
	Const S223_I2_plan_seq = 3
	Const S223_I2_export_flag = 4
	Const S223_I2_cur = 5
	Const S223_I2_sales_grp = 6
	Const S223_I2_sales_org = 7
	
	Dim I3_b_sales_org_div
	
	Dim PS2G124

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    	
        	
'	arrValSplit = Split(Request("txtSpread"), gColSep)
	
    I3_b_sales_org_div = "G"
    I2_s_cust_sales_plan(S223_I2_sp_year)     = UCase(Trim(Request("txtSpYear")))
    I2_s_cust_sales_plan(S223_I2_sp_month)    = UCase(Trim(Request("txtSpread")))
    I2_s_cust_sales_plan(S223_I2_plan_flag)   = UCase(Trim(Request("txtPlanTypeCd")))
    I2_s_cust_sales_plan(S223_I2_plan_seq)    = UCase(Trim(Request("txtPlanNum")))
    I2_s_cust_sales_plan(S223_I2_export_flag) = UCase(Trim(Request("txtDealTypeCd")))
    I2_s_cust_sales_plan(S223_I2_cur)         = gCurrency
    I2_s_cust_sales_plan(S223_I2_sales_org)   = UCase(Trim(Request("txtSalesOrg")))
    I2_s_cust_sales_plan(S223_I2_sales_grp)   = ""
            
    I2_s_cust_sales_plan1= I2_s_cust_sales_plan 
    
    Set PS2G124 = Server.CreateObject("PS2G124.CsConfirmCustSP")	
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If


    call PS2G124.S_CONFIRM_CUST_SALES_PLAN(gStrGlobalCollection , I2_s_cust_sales_plan1, I3_b_sales_org_div)
	If CheckSYSTEMError(Err,True) = True Then
		Set PS2G124 = Nothing    
		Exit Sub
    End If
    
    Set PS2G124 = Nothing    
        
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "               
    
    
End Sub

%>

<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function



</Script>
