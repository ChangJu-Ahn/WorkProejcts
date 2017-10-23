<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2113MB1
'*  4. Program Name         : 그룹별 품목판매계획등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : pS2G111.dll, pS2G112.dll, pS2G114.dll
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Mr  Cho
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/03 : 3rd Coding
'*                            -2001/01/03 : 5th Coding
'**********************************************************************************************
%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
    Dim lgOpModeCRUD
    Const lsCONFIRM = "CONFIRM"              'strMode 값:확정처리 
	Const lsSPLIT = "SPLIT"              'strMode 값:공장별배분작업 
	Const lsQtyAmt  = "QtyAmt"              '수량/금액 자동계산 
	Const lsUNIT = "ItemByUnit"             'strMode 값:itemByUnit

	Call LoadBasisGlobalInf()
	Call loadInfTB19029B( "I", "*", "NOCOOKIE", "MB")     
     
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query 
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti() 
		Case CStr(lsCONFIRM)														 '☜: 확정 
			Call SubBizlsCONFIRM()
        Case CStr(lsSPLIT)															 '☜:'품목별배분작업             
			Call SubBizIsSPIT() 
        Case CStr(lsQtyAmt)															 '☜:'수량/금액 자동합계                                                       '☜: Delete
			Call SubBizlsQtyAmt() 
        Case CStr(lsUNIT)															 '☜:'itemByUNIT
			Call SubBizlsUNIT() 
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
 
	Dim iLngRow 
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
    Dim iStrNextKey   
 
	Dim pS2G112 

	Dim I2_s_item_sales_plan

	Dim iarrValue
	Dim I3_b_item  'next import

	Dim I4_b_sales_org ' G:"S", O:
	 
	Dim E1_b_sales_grp
	Dim E2_b_minor 
    Dim E3_b_minor  
    Dim E4_s_item_sales_plan
    Dim E5_b_sales_org
    Dim exp_grp
    Dim E6_b_minor  			    
    
    Const C_SHEETMAXROWS_D  = 100
    
    '[CONVERSION INFORMATION]  IMPORTS View 상수 
    Const S213_I2_sp_year = 0
    Const S213_I2_plan_flag = 1
    Const S213_I2_plan_seq = 2
    Const S213_I2_export_flag = 3
    Const S213_I2_cur = 4
    Const S213_I2_sales_grp = 5

    '[CONVERSION INFORMATION]  EXPORTS View 상수 
    '[CONVERSION INFORMATION]  View Name : exp b_sales_grp
	Const S213_E1_sales_grp = 0
	Const S213_E1_sales_grp_nm = 1

    '[CONVERSION INFORMATION]  View Name : exp_export_flag b_minor
    'Const S213_E2_minor_nm = 0

    '[CONVERSION INFORMATION]  View Name : exp_plan_flag b_minor
    'Const S213_E3_minor_nm = 0

    '[CONVERSION INFORMATION]  View Name : exp s_item_sales_plan
    Const S213_E4_sp_year = 0
    Const S213_E4_plan_flag = 1
    Const S213_E4_plan_seq = 2
    Const S213_E4_export_flag = 3
    Const S213_E4_cur = 4
	Const S213_E4_qty_amt_flag = 5
 
    '[CONVERSION INFORMATION]  View Name : exp b_sales_org
    Const S213_E5_sales_org = 0
    Const S213_E5_sales_org_nm = 1
    
    '[CONVERSION INFORMATION]  Group Name : exp_grp    
    Const S213_EG1_E1_s_item_sales_plan_plan_unit = 0
    Const S213_EG2_E1_b_item_item_cd = 1
    Const S213_EG2_E1_b_item_item_nm = 2
    Const S213_EG2_E1_b_item_spec = 3
    Const S213_EG3_E1_s_wks_msp_plan_qty1 = 4
    Const S213_EG3_E1_s_wks_msp_plan_qty2 = 5
    Const S213_EG3_E1_s_wks_msp_plan_qty3 = 6
    Const S213_EG3_E1_s_wks_msp_plan_qty4 = 7
    Const S213_EG3_E1_s_wks_msp_plan_qty5 = 8
    Const S213_EG3_E1_s_wks_msp_plan_qty6 = 9
    Const S213_EG3_E1_s_wks_msp_plan_qty7 = 10
    Const S213_EG3_E1_s_wks_msp_plan_qty8 = 11
    Const S213_EG3_E1_s_wks_msp_plan_qty9 = 12
    Const S213_EG3_E1_s_wks_msp_plan_qty10 = 13
    Const S213_EG3_E1_s_wks_msp_plan_qty11 = 14
    Const S213_EG3_E1_s_wks_msp_plan_qty12 = 15
    Const S213_EG3_E1_s_wks_msp_split_flag1 = 16
    Const S213_EG3_E1_s_wks_msp_split_flag2 = 17
    Const S213_EG3_E1_s_wks_msp_split_flag3 = 18
    Const S213_EG3_E1_s_wks_msp_split_flag4 = 19
    Const S213_EG3_E1_s_wks_msp_split_flag5 = 20
    Const S213_EG3_E1_s_wks_msp_split_flag6 = 21
    Const S213_EG3_E1_s_wks_msp_split_flag7 = 22
    Const S213_EG3_E1_s_wks_msp_split_flag8 = 23
    Const S213_EG3_E1_s_wks_msp_split_flag9 = 24
    Const S213_EG3_E1_s_wks_msp_split_flag10 = 25
    Const S213_EG3_E1_s_wks_msp_split_flag11 = 26
    Const S213_EG3_E1_s_wks_msp_split_flag12 = 27
    Const S213_EG3_E1_s_wks_msp_plan_amt1 = 28
    Const S213_EG3_E1_s_wks_msp_plan_amt2 = 29
    Const S213_EG3_E1_s_wks_msp_plan_amt3 = 30
    Const S213_EG3_E1_s_wks_msp_plan_amt4 = 31
    Const S213_EG3_E1_s_wks_msp_plan_amt5 = 32
    Const S213_EG3_E1_s_wks_msp_plan_amt6 = 33
    Const S213_EG3_E1_s_wks_msp_plan_amt7 = 34
    Const S213_EG3_E1_s_wks_msp_plan_amt8 = 35
    Const S213_EG3_E1_s_wks_msp_plan_amt9 = 36
    Const S213_EG3_E1_s_wks_msp_plan_amt10 = 37
    Const S213_EG3_E1_s_wks_msp_plan_amt11 = 38
    Const S213_EG3_E1_s_wks_msp_plan_amt12 = 39

 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
 
    Redim I2_s_item_sales_plan(S213_I2_sales_grp)    
'grp
	I4_b_sales_org = "S"   
    I2_s_item_sales_plan(S213_I2_sales_grp)   = Trim(Request("txtConSalesOrg"))
'org
    'I4_b_sales_org   = TRIM(Request("txtConSalesOrg"))

	I2_s_item_sales_plan(S213_I2_sp_year)     = Trim(Request("txtConSpYear"))
	I2_s_item_sales_plan(S213_I2_plan_flag)   = Trim(Request("txtConPlanTypeCd"))
	I2_s_item_sales_plan(S213_I2_export_flag) = Trim(Request("txtConDealTypeCd"))
	I2_s_item_sales_plan(S213_I2_cur)         = Trim(Request("txtConCurr"))
	I2_s_item_sales_plan(S213_I2_plan_seq)    = Trim(Request("txtConPlanNum")) 

	iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                  '☜: Next Key 

	If iStrPrevKey <> "" then     
	 iarrValue = Split(iStrPrevKey, gColSep)
	 I3_b_item = Trim(iarrValue(0))
	else   
	 I3_b_item = ""
	End If          

	Set pS2G112 = Server.CreateObject("pS2G112.cSListItemSalesPlan") 

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

	Call pS2G112.S_LIST_ITEM_SALES_PLAN(gStrGlobalCollection, Cint(C_SHEETMAXROWS_D), I2_s_item_sales_plan, I3_b_item, _
	        I4_b_sales_org, E1_b_sales_grp, E2_b_minor, E3_b_minor, E4_s_item_sales_plan, E5_b_sales_org, exp_grp, E6_b_minor )        
	
	If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "202100" then
	
		If CheckSYSTEMError(Err,True) = True Then

			Response.Write "<Script language=vbs>  " & vbCr   
			Response.Write " With Parent        " & vbCr
			Response.Write "   .frm1.txtConSalesOrgNm.value = """ & ConvSPChars(E1_b_sales_grp(S213_E1_sales_grp_nm))      & """" & vbCr    
			'계획구분 
			Response.Write "   .frm1.txtConPlanTypeNm.value = """ & ConvSPChars(E3_b_minor)                                & """" & vbCr  
			'거래구분 
			Response.Write "   .frm1.txtConDealTypeNm.value = """ & ConvSPChars(E2_b_minor)                                & """" & vbCr  
			'계획차수 
			Response.Write "   .frm1.txtConPlanNumNm.value = """ & ConvSPChars(E6_b_minor)                                      & """" & vbCr  
			Response.Write " End With       " & vbCr                    
			Response.Write "</Script>      " & vbCr      

			Set pS2G112 = Nothing                                                   '☜: Unload Comproxy DLL
			Response.Write "<Script language=vbs>  " & vbCr   
			Response.Write " Parent.SetDefaultVal2  " & vbCr   
			Response.Write "</Script>      " & vbCr      
			Exit Sub
		End If   
	Else
		If CheckSYSTEMError(Err,True) = True Then

			Response.Write "<Script language=vbs>  " & vbCr   
			Response.Write " With Parent        " & vbCr
			Response.Write "   .frm1.txtConSalesOrgNm.value = """ & ConvSPChars(E1_b_sales_grp(S213_E1_sales_grp_nm))      & """" & vbCr    
			'계획구분 
			Response.Write "   .frm1.txtConPlanTypeNm.value = """ & ConvSPChars(E3_b_minor)                                & """" & vbCr  
			'거래구분 
			Response.Write "   .frm1.txtConDealTypeNm.value = """ & ConvSPChars(E2_b_minor)                                & """" & vbCr  
			'계획차수 
			Response.Write "   .frm1.txtConPlanNumNm.value = """ & ConvSPChars(E6_b_minor)                                      & """" & vbCr  
			Response.Write " End With       " & vbCr                    
			Response.Write "</Script>      " & vbCr      

			Set pS2G112 = Nothing                                                   '☜: Unload Comproxy DLL
			Response.Write "<Script language=vbs>  " & vbCr   
			Response.Write " Parent.SetDefaultVal  " & vbCr   
			Response.Write "</Script>      " & vbCr      
			Exit Sub
		End If   
	End If   
	
    Set pS2G112 = Nothing 


    iLngMaxRow  = CLng(Request("txtMaxRows"))           '☜: Fetechd Count      

	For iLngRow = 0 To UBound(exp_grp,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
			iStrNextKey = ConvSPChars(exp_grp(iLngRow, S213_EG2_E1_b_item_item_cd)) 
			Exit For
		End If 
		' 품목 
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow, S213_EG2_E1_b_item_item_cd)) 
		istrData = istrData & Chr(11) & ""               
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow, S213_EG2_E1_b_item_item_nm))   
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow, S213_EG2_E1_b_item_spec))   		
		' 계획단위 
		istrData = istrData & Chr(11) & ConvSPChars(exp_grp(iLngRow, S213_EG1_E1_s_item_sales_plan_plan_unit))   
		istrData = istrData & Chr(11) & ""
		' 년계획 수량 합계 
		istrData = istrData & Chr(11) & ""
		' 년계획 금액 합계 
		istrData = istrData & Chr(11) & ""               
		' 1월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty1), ggQty.DecPoint, 0) 
		' 1월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt1), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 2월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty2), ggQty.DecPoint, 0)  
		' 2월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt2), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 3월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty3), ggQty.DecPoint, 0)  
		' 3월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt3), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 4월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty4), ggQty.DecPoint, 0)  
		' 4월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt4), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 5월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty5), ggQty.DecPoint, 0)  
		' 5월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt5), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 6월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty6), ggQty.DecPoint, 0)  
		' 6월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt6), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 7월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty7), ggQty.DecPoint, 0)  
		' 7월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt7), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 8월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty8), ggQty.DecPoint, 0)  
		' 8월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt8), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 9월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty9), ggQty.DecPoint, 0)  
		' 9월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt9), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 10월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty10), ggQty.DecPoint, 0)  
		' 10월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt10), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 11월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty11), ggQty.DecPoint, 0)  
		' 11월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt11), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 12월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_qty12), ggQty.DecPoint, 0)  
		' 12월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_plan_amt12), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
  
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag1)      
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag2)     
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag3)    
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag4)     
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag5)      
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag6)      
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag7)      
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag8)      
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag9)  
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag10)    
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag11)     
		istrData = istrData & Chr(11) & exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag12)     
       
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12) 

    Next    

    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write " With Parent        " & vbCr
    Response.Write "   .frm1.txtSalesOrg.value   = """ & ConvSPChars(E1_b_sales_grp(S213_E1_sales_grp))         & """" & vbCr 
    Response.Write "   .frm1.txtSalesOrgNm.value = """ & ConvSPChars(E1_b_sales_grp(S213_E1_sales_grp_nm))      & """" & vbCr    
    '계획년도 
    Response.Write "   .frm1.txtSpYear.value  = """ & ConvSPChars(E4_s_item_sales_plan(S213_E4_sp_year))     & """" & vbCr 
    '계획구분 
    Response.Write "   .frm1.txtPlanTypeCd.value = """ & ConvSPChars(E4_s_item_sales_plan(S213_E4_plan_flag))   & """" & vbCr    
    Response.Write "   .frm1.txtPlanTypeNm.value = """ & ConvSPChars(E3_b_minor)                                & """" & vbCr  
    '거래구분 
    Response.Write "   .frm1.txtDealTypeCd.value = """ & ConvSPChars(E4_s_item_sales_plan(S213_E4_export_flag)) & """" & vbCr  
    Response.Write "   .frm1.txtDealTypeNm.value = """ & ConvSPChars(E2_b_minor)                                & """" & vbCr  
    '화폐 
    Response.Write "   .frm1.txtCurr.value   = """ & ConvSPChars(E4_s_item_sales_plan(S213_E4_cur))         & """" & vbCr  
    '계획차수 
    Response.Write "   .frm1.txtPlanNum.value  = """ & ConvSPChars(E4_s_item_sales_plan(S213_E4_plan_seq))    & """" & vbCr  
    Response.Write "   .frm1.txtPlanNumNm.value = """ & ConvSPChars(E6_b_minor)                                      & """" & vbCr  
    '수량금액선택 
    Select Case ConvSPChars(E4_s_item_sales_plan(S213_E4_qty_amt_flag))
	Case "Q"
		Response.Write "   .frm1.rdoSelectQty.checked = True   " & vbCr 
	Case "A"
		Response.Write "   .frm1.rdoSelectAmt.checked = True      " & vbCr  
	End Select
    Response.Write "   .frm1.txtRdoSelect.value  = """ & ConvSPChars(E4_s_item_sales_plan(S213_E4_qty_amt_flag)) & """" & vbCr    
	Response.Write "   .frm1.vspdData.ReDraw = False      " & vbCr  
    'Response.Write "   .SetSpreadColor -1                                    " & vbCr 
    Response.Write "   .ggoSpread.Source          = .frm1.vspdData          " & vbCr
    Response.Write "   .ggoSpread.SSShowDataByClip        """ & istrData        & """" & ", " & """F""" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & iStrNextKey    & """" & vbCr
' 추가  

    For iLngRow = 0 To UBound(exp_grp,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
			Exit For
	        End If
		Response.Write " .ggoSpread.SSSetProtected .C_ItemCode, " & iLngMaxRow + iLngRow +1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
		Response.Write " .ggoSpread.SSSetProtected .C_ItemPopup, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr
		Response.Write " .ggoSpread.SSSetProtected .C_ItemName, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr
		Response.Write " .ggoSpread.SSSetProtected .C_ItemSpec, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr
		Response.Write " .ggoSpread.SSSetRequired .C_PlanUnit, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr
		Response.Write " .ggoSpread.SSSetProtected .C_YearQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr
		Response.Write " .ggoSpread.SSSetProtected .C_YearAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag1)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_01PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_01PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_01PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag2)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_02PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_02PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_02PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag3)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_03PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_03PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_03PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If

		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag4)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_04PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_04PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_04PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag5)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_05PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_05PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_05PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If    
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag6)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_06PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_06PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_06PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If    
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag7)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_07PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_07PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_07PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If    
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag8)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_08PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_08PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_08PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag9)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_09PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_09PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_09PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If    
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag10)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_10PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1  & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_10PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_10PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If    
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag11)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_11PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_11PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_11PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If
		If UCase(exp_grp(iLngRow, S213_EG3_E1_s_wks_msp_split_flag12)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_12PlanQty, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_12PlanAmt, " & iLngMaxRow + iLngRow+1 & ", " & iLngMaxRow + iLngRow+1 & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_12PlanQty, " & iLngMaxRow + iLngRow+1 & ")" & vbCr
		End If    
    
    Next
    
    Response.Write "   .frm1.vspdData.ReDraw = True " & vbCr
    
    Response.Write "   .frm1.HConSalesOrg.value  = """ & ConvSPChars(Request("txtConSalesOrg"))    & """" & vbCr    
    Response.Write "   .frm1.HConSpYear.value  = """ & Request("txtConSpYear")      & """" & vbCr 
    Response.Write "   .frm1.HPlanTypeCd.value  = """ & ConvSPChars(Request("txtConPlanTypeCd"))  & """" & vbCr 
    Response.Write "   .frm1.HConDealTypeCd.value = """ & ConvSPChars(Request("txtConDealTypeCd")) & """" & vbCr 
    Response.Write "   .frm1.HConCurr.value      = """ & Request("txtConCurr")        & """" & vbCr 
    Response.Write "   .frm1.HConPlanNum.value  = """ & Request("txtConPlanNum")     & """" & vbCr 

    Select Case ConvSPChars(E4_s_item_sales_plan(S213_E4_plan_flag))
	Case "1"
		Response.Write "   .frm1.btnSplit.disabled = True     " & vbCr 
	Case "2"
		Response.Write "   .frm1.btnSplit.disabled = False      " & vbCr  
	End Select

    Response.Write "   .DbQueryOk  " & vbCr   
    Response.Write " End With       " & vbCr                    
    Response.Write "</Script>      " & vbCr      

              
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   

	Dim pS2G111 
	Dim iCommandSent
	Dim itxtSpread
	Dim iErrorPosition
	Dim I2_s_item_sales_plan
	Dim I3_b_sales_org
	Dim E4_s_item_sales_plan   
 
    Const S208_I2_sp_year = 0
    Const S208_I2_sp_month = 1
    Const S208_I2_plan_flag = 2
    Const S208_I2_plan_seq = 3
    Const S208_I2_export_flag = 4
    Const S208_I2_cur = 5
    Const S208_I2_qty_amt_flag = 6
    Const S208_I2_sales_grp = 7        

    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                    '☜: Clear Error status                                                            

	 Redim I2_s_item_sales_plan(S208_I2_sales_grp)
	'grp
	 I3_b_sales_org = "S"
	 I2_s_item_sales_plan(S208_I2_sales_grp)    =  UCase(Trim(Request("txtSalesOrg")))
	'org
	 'I3_b_sales_org = UCASE(TRIM(Request("txtSalesOrg")))

	 I2_s_item_sales_plan(S208_I2_sp_year)      =  UCase(Trim(Request("txtSpYear")))
	 I2_s_item_sales_plan(S208_I2_plan_flag)    =  UCase(Trim(Request("txtPlanTypeCd")))
	 I2_s_item_sales_plan(S208_I2_export_flag)  =  UCase(Trim(Request("txtDealTypeCd")))
	 I2_s_item_sales_plan(S208_I2_cur)          =  UCase(Trim(Request("txtCurr")))
	 I2_s_item_sales_plan(S208_I2_qty_amt_flag) =  UCase(Trim(Request("txtRdoSelect")))

	 If Len(Request("txtPlanNum")) Then 
	    I2_s_item_sales_plan(S208_I2_plan_seq)     = UCase(Trim(Request("txtPlanNum")))
	 End if

	iCommandSent = "SAVE"

	Set pS2G111 = Server.CreateObject("PS2G111.cSMonthlySP")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    itxtSpread = Trim(Request("txtSpread"))
    
    Call pS2G111.S_MAINT_MONTHLY_SALES_PLAN(gStrGlobalCollection, iCommandSent, _
                              , I2_s_item_sales_plan , I3_b_sales_org, itxtSpread , _
                           ,iErrorPosition , E4_s_item_sales_plan)
                 
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set pS2G111 = Nothing
       Exit Sub
	End If
 
    Set pS2G111 = Nothing

    Response.Write "<Script language=vbs> " & vbCr            

    Select Case E4_s_item_sales_plan
    Case ""
	 Response.Write "   parent.frm1.txtConPlanNum.value = parent.frm1.txtPlanNum.value " & vbCr    
	Case Else
	 Response.Write "   parent.frm1.txtConPlanNum.value = """   & ConvSPChars(E4_s_item_sales_plan)    & """" & vbCr 
	End Select          

    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr                                                                                                                                              
              
End Sub   
'============================================================================================================
' Name : SubBizlsCONFIRM
' Desc : '월별확정작업             
'============================================================================================================
Sub SubBizlsCONFIRM()

 Dim pS2G111 
 Dim iCommandSent
 Dim I2_s_item_sales_plan
 Dim I3_b_sales_org
 
    Const S208_I2_sp_year = 0
    Const S208_I2_sp_month = 1
    Const S208_I2_plan_flag = 2
    Const S208_I2_plan_seq = 3
    Const S208_I2_export_flag = 4
    Const S208_I2_cur = 5
    Const S208_I2_qty_amt_flag = 6
    Const S208_I2_sales_grp = 7    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

 Redim I2_s_item_sales_plan(S208_I2_sales_grp)

'grp
 I3_b_sales_org = "S"
 I2_s_item_sales_plan(S208_I2_sales_grp)    =  UCase(Trim(Request("txtSalesOrg")))
'org
 'I3_b_sales_org = UCASE(TRIM(Request("txtSalesOrg")))
 
 I2_s_item_sales_plan(S208_I2_sp_year)      =  Trim(Request("txtSpYear"))
 I2_s_item_sales_plan(S208_I2_plan_flag)    =  UCase(Trim(Request("txtPlanTypeCd")))
 I2_s_item_sales_plan(S208_I2_export_flag)  =  UCase(Trim(Request("txtDealTypeCd"))) 
 I2_s_item_sales_plan(S208_I2_cur)          =  UCase(Trim(Request("txtCurr")))
    I2_s_item_sales_plan(S208_I2_plan_seq)     =  Trim(Request("txtPlanNum"))
 I2_s_item_sales_plan(S208_I2_sp_month)    =  Trim(Request("txtBatchMonth"))

 iCommandSent = "CONFIRM"

	Set pS2G111 = Server.CreateObject("PS2G111.cSMonthlySP")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call pS2G111.S_MAINT_MONTHLY_SALES_PLAN(gStrGlobalCollection, iCommandSent, _
                          , I2_s_item_sales_plan , I3_b_sales_org,  , _
                           ,"" ,"" )
                 
	If CheckSYSTEMError(Err,True) = True Then
       Set pS2G111 = Nothing
       Exit Sub
    End If
 
    Set pS2G111 = Nothing

    Response.Write "<Script language=vbs> " & vbCr            
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr                                                                        


End Sub 
'============================================================================================================
' Name : SubBizIsSPIT
' Desc : '품목별배분작업             
'============================================================================================================
Sub SubBizIsSPIT()

	Dim pS2G114 
	Dim I1_b_sales_org
	Dim I3_s_item_sales_plan
 
    Const S211_I3_sp_year = 0
    Const S211_I3_sp_month = 1
    Const S211_I3_plan_flag = 2
    Const S211_I3_plan_seq = 3
    Const S211_I3_export_flag = 4
    Const S211_I3_cur = 5
    Const S211_I3_sales_grp = 6

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

 Redim I3_s_item_sales_plan(S211_I3_sales_grp)

'grp
 I1_b_sales_org = "S"
 I3_s_item_sales_plan(S211_I3_sales_grp)    =  UCase(Trim(Request("txtSalesOrg")))
'org
 'I1_b_sales_org = UCASE(TRIM(Request("txtSalesOrg")))
 
 I3_s_item_sales_plan(S211_I3_sp_year)      =  Trim(Request("txtSpYear"))
 I3_s_item_sales_plan(S211_I3_plan_flag)    =  UCase(Trim(Request("txtPlanTypeCd")))
 I3_s_item_sales_plan(S211_I3_export_flag)  =  UCase(Trim(Request("txtDealTypeCd")))
 I3_s_item_sales_plan(S211_I3_cur)          =  UCase(Trim(Request("txtCurr")))
    I3_s_item_sales_plan(S211_I3_plan_seq)     =  Trim(Request("txtPlanNum"))
 I3_s_item_sales_plan(S211_I3_sp_month)    =  Trim(Request("txtBatchMonth"))

 'iCommandSent = "SPLIT"

	Set pS2G114 = Server.CreateObject("PS2G114.cSInsCfmSalesByItem")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call pS2G114.S_INSERT_CFM_ITEM_SALES_BY_ITEM(gStrGlobalCollection, I1_b_sales_org, I3_s_item_sales_plan )
                 
	If CheckSYSTEMError(Err,True) = True Then
       Set pS2G114 = Nothing
       Exit Sub
    End If
 
    Set pS2G114 = Nothing

    Response.Write "<Script language=vbs> " & vbCr            
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr                                                                        

End Sub 
'============================================================================================================
' Name : SubBizlsQtyAmt
' Desc : 수량/금액 자동합계                                  
'============================================================================================================
Sub SubBizlsQtyAmt()
	Dim pS2G146 
 
	Dim I1_b_item
    Dim I2_s_item_sales_plan
    Dim E1_b_item
    Dim E2_s_item_sales_plan
 
    '[CONVERSION INFORMATION]  View Name : imp s_item_sales_plan
    Const S242_I2_sp_month = 0
    Const S242_I2_cur = 1
    Const S242_I2_plan_unit = 2
    Const S242_I2_qty_amt_flag = 3
    Const S242_I2_plan_qty = 4
    Const S242_I2_plan_amt = 5

    '[CONVERSION INFORMATION]  View Name : exp b_item
    Const S242_E1_item_cd = 0

    '[CONVERSION INFORMATION]  View Name : exp s_item_sales_plan
    Const S242_E2_sp_month = 0
    Const S242_E2_plan_qty = 1
    Const S242_E2_plan_amt = 2

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

    Redim I2_s_item_sales_plan(S242_I2_plan_amt)    
	I1_b_item = Trim(Request("lsItemCode"))
	I2_s_item_sales_plan(S242_I2_sp_month) = Trim(Request("lsPlanMonth"))
	I2_s_item_sales_plan(S242_I2_plan_unit) = Trim(Request("lsPlanUnit"))
	I2_s_item_sales_plan(S242_I2_cur) = Trim(Request("txtCurr"))
	I2_s_item_sales_plan(S242_I2_qty_amt_flag) = Trim(Request("txtRdoSelect"))

	Select Case Request("txtRdoSelect")
	Case "Q"
	 I2_s_item_sales_plan(S242_I2_plan_qty) = Trim(Request("lsPlanQtyAmt"))  
	Case "A"
	 I2_s_item_sales_plan(S242_I2_plan_amt) = Trim(Request("lsPlanQtyAmt"))
	End Select

	Set pS2G146 = Server.CreateObject("PS2G146.CsUpQtyAmtSvr")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call pS2G146.S_UPDATE_QTY_AMT_SVR(gStrGlobalCollection, I1_b_item, I2_s_item_sales_plan, E1_b_item, E2_s_item_sales_plan )
                 
	If CheckSYSTEMError(Err,True) = True Then
       Set pS2G146 = Nothing
       Exit Sub
    End If
 
    Set pS2G146 = Nothing

    Response.Write "<Script language=vbs> " & vbCr            
    Response.Write " With Parent   " & vbCr
    Response.Write "   .frm1.vspdData.Row = """ & Request("CurrentRow")         & """" & vbCr 

	Select Case ConvSPChars(E2_s_item_sales_plan(S242_E2_sp_month))
	Case "01"
	 Response.Write "   .frm1.vspdData.Col = .C_01PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_01PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "02"
	 Response.Write "   .frm1.vspdData.Col = .C_02PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_02PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "03"
	 Response.Write "   .frm1.vspdData.Col = .C_03PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_03PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "04"
	 Response.Write "   .frm1.vspdData.Col = .C_04PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_04PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "05"
	 Response.Write "   .frm1.vspdData.Col = .C_05PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_05PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "06"
	 Response.Write "   .frm1.vspdData.Col = .C_06PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_06PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "07"
	 Response.Write "   .frm1.vspdData.Col = .C_07PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_07PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "08"
	 Response.Write "   .frm1.vspdData.Col = .C_08PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_08PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "09"
	 Response.Write "   .frm1.vspdData.Col = .C_09PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_09PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "10"
	 Response.Write "   .frm1.vspdData.Col = .C_10PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_10PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "11"
	 Response.Write "   .frm1.vspdData.Col = .C_11PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_11PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	Case "12"
	 Response.Write "   .frm1.vspdData.Col = .C_12PlanQty   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UNINumClientFormat(E2_s_item_sales_plan(S242_E2_plan_qty), ggQty.DecPoint, 0)       & """" & vbCr 

	 Response.Write "   .frm1.vspdData.Col = .C_12PlanAmt   " & vbCr 
	 Response.Write "   .frm1.vspdData.Text = """ & UniConvNumberDBToCompany(E2_s_item_sales_plan(S242_E2_plan_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)       & """" & vbCr 
	End Select

	Response.Write " Parent.UpdateQtyAmtSvrOK "      & vbCr   
	Response.Write " End With   " & vbCr                    
	Response.Write "</Script> "             & vbCr            

End Sub 

'============================================================================================================
' Name : SubBizlsUNIT
' Desc : itemByUNIT                                  
'============================================================================================================
Sub SubBizlsUNIT()
 
	Dim S11118
	Dim imp_b_item
	Dim exp_b_unit_of_measure
	Dim GroupView
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

	imp_b_item = UCase(Trim(Request("ItemCd")))

    Set S11118 = Server.CreateObject("PS1G102.cListItemPriceSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

	Call S11118.S_LIST_ITEM_PRICE_SVR(gStrGlobalCollection, , , , , imp_b_item, , exp_b_unit_of_measure, , , , , , ,GroupView)                  

	If CheckSYSTEMError(Err,True) = True Then
       Set S11118 = Nothing
       Exit Sub
    End If
 
    Set S11118 = Nothing

    Response.Write "<Script language=vbs> " & vbCr            
    Response.Write " With Parent.frm1.vspdData   " & vbCr
    Response.Write " .Row = """ & Request("CRow")       & """" & vbCr   
	Response.Write " .Col = parent.C_PlanUnit      " & vbCr      
	Response.Write " .text = """   & ConvSPChars(exp_b_unit_of_measure(0))  & """" & vbCr 
    Response.Write " End With       " & vbCr                    
    Response.Write "</Script> "             & vbCr                                                                        

End Sub 

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
 '------ Developer Coding part (Start ) ------------------------------------------------------------------
 '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
 '------ Developer Coding part (Start ) ------------------------------------------------------------------
 '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
 '------ Developer Coding part (Start ) ------------------------------------------------------------------
 '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>
