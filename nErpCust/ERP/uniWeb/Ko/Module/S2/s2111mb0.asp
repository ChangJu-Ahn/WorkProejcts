<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : s2111mb0
'*  4. Program Name         : 조직별 품목그룹판매계획 
'*  5. Program Desc         : 조직별 품목그룹판매계획 
'*  6. Comproxy List        : PS2G101.dll,PS2G102.dll,PS2G104.dll
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2001/12/18
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
        Case CStr(lsSPLIT)															 '☜:'품목별배분작업             
			Call SubBizIsSPIT() 
        Case CStr(lsQtyAmt)															 '☜:'수량/금액 자동합계                                                       '☜: Delete
			Call SubBizlsQtyAmt() 
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

 
	Dim PS2G102 
	Dim iarrValue
	Dim imp_next_s_bill
	 
	Dim I1_s_wks_user
 
	Dim I3_s_item_group_sales_plan
 
	Dim E1_b_sales_grp
	Dim E3_b_minor 
    Dim E4_b_minor  
    Dim E5_s_item_group_sales_plan
    Dim E6_b_sales_org
    Dim EG1_exp_grp
    Dim E7_b_minor  
        
    Const C_SHEETMAXROWS_D  = 100
    
 
    Const S206_I3_sales_org = 0
    Const S206_I3_sp_year = 1
    Const S206_I3_plan_flag = 2
    Const S206_I3_plan_seq = 3
    Const S206_I3_export_flag = 4
    Const S206_I3_cur = 5
    Const S206_I3_sales_grp = 6
   
    '[CONVERSION INFORMATION]  View Name : exp b_sales_grp

    Const S206_E1_sales_grp = 0
    Const S206_E1_sales_grp_nm = 1

   
    Const S206_E5_sp_year = 0
    Const S206_E5_plan_flag = 1
    Const S206_E5_plan_seq = 2
    Const S206_E5_export_flag = 3
    Const S206_E5_cur = 4
    Const S206_E5_qty_amt_flag = 5

    '[CONVERSION INFORMATION]  EXPORTS View 상수 

    '[CONVERSION INFORMATION]  View Name : exp b_sales_org
    Const S206_E6_sales_org = 0
    Const S206_E6_sales_org_nm = 1
    
    Const S206_EG1_E1_b_item_group_item_group_cd = 0
    Const S206_EG1_E1_b_item_group_item_group_nm = 1
    Const S206_EG1_E2_s_wks_msp_plan_qty1 = 2
    Const S206_EG1_E2_s_wks_msp_plan_qty2 = 3
    Const S206_EG1_E2_s_wks_msp_plan_qty3 = 4
    Const S206_EG1_E2_s_wks_msp_plan_qty4 = 5
    Const S206_EG1_E2_s_wks_msp_plan_qty5 = 6
    Const S206_EG1_E2_s_wks_msp_plan_qty6 = 7
    Const S206_EG1_E2_s_wks_msp_plan_qty7 = 8
    Const S206_EG1_E2_s_wks_msp_plan_qty8 = 9
    Const S206_EG1_E2_s_wks_msp_plan_qty9 = 10
    Const S206_EG1_E2_s_wks_msp_plan_qty10 = 11
    Const S206_EG1_E2_s_wks_msp_plan_qty11 = 12
    Const S206_EG1_E2_s_wks_msp_plan_qty12 = 13
    Const S206_EG1_E2_s_wks_msp_plan_amt1 = 14
    Const S206_EG1_E2_s_wks_msp_plan_amt2 = 15
    Const S206_EG1_E2_s_wks_msp_plan_amt3 = 16
    Const S206_EG1_E2_s_wks_msp_plan_amt4 = 17
    Const S206_EG1_E2_s_wks_msp_plan_amt5 = 18
    Const S206_EG1_E2_s_wks_msp_plan_amt6 = 19
    Const S206_EG1_E2_s_wks_msp_plan_amt7 = 20
    Const S206_EG1_E2_s_wks_msp_plan_amt8 = 21
    Const S206_EG1_E2_s_wks_msp_plan_amt9 = 22
    Const S206_EG1_E2_s_wks_msp_plan_amt10 = 23
    Const S206_EG1_E2_s_wks_msp_plan_amt11 = 24
    Const S206_EG1_E2_s_wks_msp_plan_amt12 = 25
    Const S206_EG1_E2_s_wks_msp_split_flag1 = 26
    Const S206_EG1_E2_s_wks_msp_split_flag2 = 27
    Const S206_EG1_E2_s_wks_msp_split_flag3 = 28
    Const S206_EG1_E2_s_wks_msp_split_flag4 = 29
    Const S206_EG1_E2_s_wks_msp_split_flag5 = 30
    Const S206_EG1_E2_s_wks_msp_split_flag6 = 31
    Const S206_EG1_E2_s_wks_msp_split_flag7 = 32
    Const S206_EG1_E2_s_wks_msp_split_flag8 = 33
    Const S206_EG1_E2_s_wks_msp_split_flag9 = 34
    Const S206_EG1_E2_s_wks_msp_split_flag10 = 35
    Const S206_EG1_E2_s_wks_msp_split_flag11 = 36
    Const S206_EG1_E2_s_wks_msp_split_flag12 = 37
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
 
	iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key 
 
    Redim I3_s_item_group_sales_plan(S206_I3_sales_grp) 
	I3_s_item_group_sales_plan(S206_I3_sales_org)   =  Trim(Request("txtConSalesOrg"))
	I3_s_item_group_sales_plan(S206_I3_sp_year)     =  Trim(Request("txtConSpYear"))
	I3_s_item_group_sales_plan(S206_I3_plan_flag)   =  Trim(Request("txtConPlanTypeCd"))
	I3_s_item_group_sales_plan(S206_I3_plan_seq)    =  Trim(Request("txtConPlanNum"))
	I3_s_item_group_sales_plan(S206_I3_export_flag) =  Trim(Request("txtConDealTypeCd"))
	I3_s_item_group_sales_plan(S206_I3_cur)         =  Trim(Request("txtConCurr"))
	I3_s_item_group_sales_plan(S206_I3_sales_grp)   =  ""
 

	If iStrPrevKey <> "" then     
		iarrValue = Split(iStrPrevKey, gColSep)
		imp_next_s_bill = Trim(iarrValue(0))
	End If          

	Set PS2G102 = Server.CreateObject("PS2G102.cSListItemGpSP") 

	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script language=vbs>    " & vbCr  
		Response.Write "   Parent.SetDefaultVal " & vbCr 
		Response.Write "</Script>      " & vbCr 
		Exit Sub
	End If

	Call PS2G102.S_LIST_ITEM_GROUP_SALES_PLAN(gStrGlobalCollection , C_SHEETMAXROWS_D ,  I1_s_wks_user , _
                imp_next_s_bill, I3_s_item_group_sales_plan, E1_b_sales_grp,  E3_b_minor , _
				E4_b_minor , E5_s_item_group_sales_plan , E6_b_sales_org, EG1_exp_grp, E7_b_minor ) 
	
	If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "202116" then
		
		If CheckSYSTEMError(Err,True) = True Then

			Response.Write "<Script language=vbs>  " & vbCr   
			Response.Write " With Parent        " & vbCr
			Response.Write "   .frm1.txtConSalesOrgNm.value = """ & ConvSPChars(E6_b_sales_org(S206_E6_sales_org_nm))            & """" & vbCr    
			Response.Write "   .frm1.txtConPlanTypeNm.value = """ & ConvSPChars(E3_b_minor)                                      & """" & vbCr  
			Response.Write "   .frm1.txtConDealTypeNm.value = """ & ConvSPChars(E4_b_minor)                                      & """" & vbCr  
			Response.Write "   .frm1.txtConPlanNumNm.value = """ & ConvSPChars(E7_b_minor)                                      & """" & vbCr  
			Response.Write " End With       " & vbCr                    
			Response.Write "</Script>      " & vbCr 

			Set PS2G102 = Nothing                                                   '☜: Unload Comproxy DLL
			Response.Write "<Script language=vbs>  " & vbCr  
			Response.Write "   Parent.SetDefaultVal2 " & vbCr 
			Response.Write "</Script>      " & vbCr 
			Exit Sub
		End If   
	Else
		If CheckSYSTEMError(Err,True) = True Then

			Response.Write "<Script language=vbs>  " & vbCr   
			Response.Write " With Parent        " & vbCr
			Response.Write "   .frm1.txtConSalesOrgNm.value = """ & ConvSPChars(E6_b_sales_org(S206_E6_sales_org_nm))            & """" & vbCr    
			Response.Write "   .frm1.txtConPlanTypeNm.value = """ & ConvSPChars(E3_b_minor)                                      & """" & vbCr  
			Response.Write "   .frm1.txtConDealTypeNm.value = """ & ConvSPChars(E4_b_minor)                                      & """" & vbCr  
			Response.Write "   .frm1.txtConPlanNumNm.value = """ & ConvSPChars(E7_b_minor)                                      & """" & vbCr  
			Response.Write " End With       " & vbCr                    
			Response.Write "</Script>      " & vbCr 

			Set PS2G102 = Nothing                                                   '☜: Unload Comproxy DLL
			Response.Write "<Script language=vbs>  " & vbCr  
			Response.Write "   Parent.SetDefaultVal " & vbCr 
			Response.Write "</Script>      " & vbCr 
			Exit Sub
		End If   
	End If

    Set PS2G102 = Nothing 

    iLngMaxRow  = CLng(Request("txtMaxRows"))           '☜: Fetechd Count      
    
	For iLngRow = 0 To UBound(EG1_exp_grp,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
			iStrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, S206_EG1_E1_b_item_group_item_group_cd)) 
			Exit For
        End If 
		' 품목 

		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S206_EG1_E1_b_item_group_item_group_cd)) 
		istrData = istrData & Chr(11) & ""               
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S206_EG1_E1_b_item_group_item_group_nm))   
		' 계획단위 
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ""
		' 년계획 수량 합계 
		istrData = istrData & Chr(11) & ""
		' 년계획 금액 합계 
		istrData = istrData & Chr(11) & ""               
		' 1월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty1), ggQty.DecPoint, 0) 
		' 1월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt1), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 2월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty2), ggQty.DecPoint, 0)  
		' 2월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt2), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 3월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty3), ggQty.DecPoint, 0)  
		' 3월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt3), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 4월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty4), ggQty.DecPoint, 0)  
		' 4월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt4), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 5월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty5), ggQty.DecPoint, 0)  
		' 5월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt5), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 6월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty6), ggQty.DecPoint, 0)  
		' 6월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt6), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 7월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty7), ggQty.DecPoint, 0)  
		' 7월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt7), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 8월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty8), ggQty.DecPoint, 0)  
		' 8월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt8), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 9월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty9), ggQty.DecPoint, 0)  
		' 9월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt9), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 10월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty10), ggQty.DecPoint, 0)  
		' 10월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt10), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 11월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty11), ggQty.DecPoint, 0)  
		' 11월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt11), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
		' 12월 계획량 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_qty12), ggQty.DecPoint, 0)  
		' 12월 계획금액 
		istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_plan_amt12), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  
  
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag1)      
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag2)     
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag3)    
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag4)     
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag5)      
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag6)      
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag7)      
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag8)      
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag9)  
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag10)    
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag11)     
		istrData = istrData & Chr(11) & EG1_exp_grp(iLngRow, S206_EG1_E2_s_wks_msp_split_flag12)     
		     
		istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
		istrData = istrData & Chr(11) & Chr(12) 
    
    Next    
    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write " With Parent        " & vbCr

    Response.Write "   .frm1.txtSalesOrg.value   = """ & ConvSPChars(E6_b_sales_org(S206_E6_sales_org))               & """" & vbCr 
    Response.Write "   .frm1.txtSalesOrgNm.value = """ & ConvSPChars(E6_b_sales_org(S206_E6_sales_org_nm))            & """" & vbCr    
    '계획년도 
    Response.Write "   .frm1.txtSpYear.value  = """ & ConvSPChars(E5_s_item_group_sales_plan(S206_E5_sp_year))     & """" & vbCr 
    '계획구분 
    Response.Write "   .frm1.txtPlanTypeCd.value = """ & ConvSPChars(E5_s_item_group_sales_plan(S206_E5_plan_flag))   & """" & vbCr    
    Response.Write "   .frm1.txtPlanTypeNm.value = """ & ConvSPChars(E3_b_minor)                                      & """" & vbCr  
    '거래구분 
    Response.Write "   .frm1.txtDealTypeCd.value = """ & ConvSPChars(E5_s_item_group_sales_plan(S206_E5_export_flag)) & """" & vbCr  
    Response.Write "   .frm1.txtDealTypeNm.value = """ & ConvSPChars(E4_b_minor)                                      & """" & vbCr  
    '화폐 
    Response.Write "   .frm1.txtCurr.value   = """ & ConvSPChars(E5_s_item_group_sales_plan(S206_E5_cur))         & """" & vbCr  
    '계획차수 
    Response.Write "   .frm1.txtPlanNum.value = """ & ConvSPChars(E5_s_item_group_sales_plan(S206_E5_plan_seq)) & """" & vbCr  
    Response.Write "   .frm1.txtPlanNumNm.value = """ & ConvSPChars(E7_b_minor)                                      & """" & vbCr  

    '수량금액선택 
    Select Case ConvSPChars(E5_s_item_group_sales_plan(S206_E5_qty_amt_flag))
	Case "Q"
	 Response.Write "   .frm1.rdoSelectQty.checked = True   " & vbCr 
	Case "A"
	 Response.Write "   .frm1.rdoSelectAmt.checked = True      " & vbCr  
	End Select
    Response.Write "   .frm1.txtRdoSelect.value  = """ & ConvSPChars(E5_s_item_group_sales_plan(S206_E5_qty_amt_flag)) & """" & vbCr    
    Response.Write "   .SetSpreadColor -1, -1                                    " & vbCr 
    Response.Write "   .ggoSpread.Source          =   .frm1.vspdData           " & vbCr
    Response.Write "   .ggoSpread.SSShowDataByClip        """ & istrData        & """" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & iStrNextKey    & """" & vbCr
    Response.Write "   .frm1.vspdData.ReDraw = False      " & vbCr
    
    For iLngRow = 0 To UBound(EG1_exp_grp,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		    Exit For
		End If
		Response.Write " .ggoSpread.SSSetProtected .C_ItemCode, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Response.Write " .ggoSpread.SSSetProtected .C_ItemPopup, LngMaxRow + " & iLngRow + 1  & "  , LngMaxRow + " & iLngRow + 1   & vbCr
		Response.Write " .ggoSpread.SSSetProtected .C_ItemName, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr
		Response.Write " .ggoSpread.SSSetProtected .C_YearQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr
		Response.Write " .ggoSpread.SSSetProtected .C_YearAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag1)) = "Y" Then
			Response.Write "    .ggoSpread.SSSetProtected .C_01PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_01PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_01PlanQty , LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag2)) = "Y" Then
			Response.Write "    .ggoSpread.SSSetProtected .C_02PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_02PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_02PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag3)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_03PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_03PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_03PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If

		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag4)) = "Y" Then
			Response.Write "    .ggoSpread.SSSetProtected .C_04PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_04PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_04PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag5)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_05PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_05PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_05PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If    
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag6)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_06PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_06PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_06PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If    
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag7)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_07PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_07PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_07PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If    
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag8)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_08PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_08PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_08PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag9)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_09PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_09PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_09PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If    
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag10)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_10PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_10PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_10PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If    
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag11)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_11PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_11PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_11PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If
		If UCase(EG1_exp_grp(iLngRow  , S206_EG1_E2_s_wks_msp_split_flag12)) = "Y" Then
		    Response.Write "    .ggoSpread.SSSetProtected .C_12PlanQty, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
			Response.Write "    .ggoSpread.SSSetProtected .C_12PlanAmt, LngMaxRow + " & iLngRow + 1  & " , LngMaxRow + " & iLngRow + 1   & vbCr    
		Else
			Response.Write " Call .SplitFlagMonthColor(.C_12PlanQty,LngMaxRow +  " & iLngRow + 1  & ")" & vbCr
		End If    
    
    Next
    Response.Write "  .frm1.vspdData.ReDraw = True " & vbCr
   
    If iStrNextKey <> "" Then
		Response.Write "   .DbQueryOk  " & vbCr
	Else
		Response.Write "   .frm1.HConSalesOrg.value  = """ & ConvSPChars(Request("txtConSalesOrg"))    & """" & vbCr    
        Response.Write "   .frm1.HConSpYear.value  = """ & Request("txtConSpYear")      & """" & vbCr 
        Response.Write "   .frm1.HPlanTypeCd.value  = """ & ConvSPChars(Request("txtConPlanTypeCd"))  & """" & vbCr 
        Response.Write "   .frm1.HConDealTypeCd.value = """ & ConvSPChars(Request("txtConDealTypeCd")) & """" & vbCr 
        Response.Write "   .frm1.HConCurr.value      = """ & Request("txtConCurr")        & """" & vbCr 
        Response.Write "   .frm1.HConPlanNum.value  = """ & Request("txtConPlanNum")     & """" & vbCr 
		Response.Write "   .DbQueryOk  " & vbCr        
	End If
	Select Case ConvSPChars(E5_s_item_group_sales_plan(S206_E5_qty_amt_flag))
     
	Case "A"
		Response.Write "  Call .SetAmtSpreadColor(1,1,""QUERY"") " & vbCr  
	End Select

    Response.Write " End With       " & vbCr                    
    Response.Write "</Script>      " & vbCr      
               
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   
                                                                      
	Dim iS2G101
	Dim itxtSpread
	Dim iErrorPosition
	Dim I2_s_item_group_sales_plan
	Dim E5_s_item_group_sales_plan
	Const S200_I2_sales_org = 0
    Const S200_I2_sp_year = 1
    Const S200_I2_sp_month = 2
    Const S200_I2_plan_flag = 3
    Const S200_I2_plan_seq = 4
    Const S200_I2_export_flag = 5
    Const S200_I2_cur = 6
    Const S200_I2_qty_amt_flag = 7
 
 
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                    '☜: Clear Error status                                                            

	Redim I2_s_item_group_sales_plan(S200_I2_qty_amt_flag)

	I2_s_item_group_sales_plan(S200_I2_sales_org)    =  UCase(Trim(Request("txtSalesOrg")))
	I2_s_item_group_sales_plan(S200_I2_sp_year)      =  UCase(Trim(Request("txtSpYear")))
	I2_s_item_group_sales_plan(S200_I2_plan_flag)    =  UCase(Trim(Request("txtPlanTypeCd")))
	I2_s_item_group_sales_plan(S200_I2_export_flag)  =  UCase(Trim(Request("txtDealTypeCd")))
	I2_s_item_group_sales_plan(S200_I2_cur)          =  UCase(Trim(Request("txtCurr")))
	I2_s_item_group_sales_plan(S200_I2_qty_amt_flag) =  UCase(Trim(Request("txtRdoSelect")))

	If Len(Request("txtPlanNum")) Then 
	    I2_s_item_group_sales_plan(S200_I2_plan_seq)     = UCase(Trim(Request("txtPlanNum")))
	End if
	 
	Set iS2G101 = Server.CreateObject("PS2G101.cSItemGpSalesPl")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
	itxtSpread = Trim(Request("txtSpread"))
	
    Call iS2G101.S_MAINT_ITEM_GROUP_SALES_PL (gStrGlobalCollection,itxtSpread , _
                             I2_s_item_group_sales_plan, "" ,  E5_s_item_group_sales_plan,iErrorPosition)
                 
	Set iS2G101 = Nothing
	
	If Trim(iErrorPosition) <> "" Then
		If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
			Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
			Response.Write " Call Parent.SubSetErrPos(" & iErrorPosition & ")" & vbCr
			Response.Write "</SCRIPT> "		
	       Exit Sub
		End If	
	Else
		If CheckSYSTEMError(Err,True) = True Then
	       Exit Sub
		End If
	End If

    Response.Write "<Script language=vbs> " & vbCr  
    
    Select Case E5_s_item_group_sales_plan
    Case ""
	Response.Write "   parent.frm1.txtConPlanNum.value = parent.frm1.txtPlanNum.value " & vbCr    
	Case Else
	Response.Write "   parent.frm1.txtConPlanNum.value = """   & ConvSPChars(E5_s_item_group_sales_plan)    & """" & vbCr 
	End Select          
    
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr                                                                        
              
End Sub   

'============================================================================================================
' Name : SubBizIsSPIT
' Desc : '품목별배분작업             
'============================================================================================================
Sub SubBizIsSPIT()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

	Dim iS2G104
	Dim I3_s_item_group_sales_plan
    Dim arrValSplit 
	Const S201_I3_sales_org    = 0
    Const S201_I3_sp_year      = 1
    Const S201_I3_sp_month     = 2
    Const S201_I3_plan_flag    = 3
    Const S201_I3_plan_seq     = 4
    Const S201_I3_export_flag  = 5
    Const S201_I3_cur          = 6
    Const S201_I3_qty_amt_flag = 7
    Const S201_I3_sales_grp    = 8
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                    '☜: Clear Error status                                                            

	Redim I3_s_item_group_sales_plan(S201_I3_sales_grp)

	I3_s_item_group_sales_plan(S201_I3_sales_org)    =  UCase(Trim(Request("txtSalesOrg")))
	I3_s_item_group_sales_plan(S201_I3_sp_year)      =  UCase(Trim(Request("txtSpYear")))
	I3_s_item_group_sales_plan(S201_I3_plan_flag)    =  UCase(Trim(Request("txtPlanTypeCd")))
	I3_s_item_group_sales_plan(S201_I3_export_flag)  =  UCase(Trim(Request("txtDealTypeCd")))
	I3_s_item_group_sales_plan(S201_I3_cur)          =  UCase(Trim(Request("txtCurr")))
	I3_s_item_group_sales_plan(S201_I3_qty_amt_flag) =  UCase(Trim(Request("txtRdoSelect")))
 
	arrValSplit = Split(Request("txtSpread"), gColSep)                       '☆: Spread Sheet 내용을 담고 있는 Element명 
    I3_s_item_group_sales_plan(S201_I3_sp_month)     =  Trim(arrValSplit(0))

	If Len(Request("txtPlanNum")) Then 
		I3_s_item_group_sales_plan(S201_I3_plan_seq)     = UCase(Trim(Request("txtPlanNum")))
	End if
  
	Set iS2G104 = Server.CreateObject("PS2G104.cSSplitItemGpSP")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call iS2G104.S_SPLIT_ITEM_GROUP_SALES_PLAN(gStrGlobalCollection,"",I3_s_item_group_sales_plan)
                 

    If CheckSYSTEMError(Err,True) = True Then
       Set iS2G104 = Nothing
       Exit Sub
    End If
 
    Set iS2G104 = Nothing
                                                        
    Response.Write "<Script language=vbs> " & vbCr      
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr       

End Sub 
'============================================================================================================
' Name : SubBizlsQtyAmt
' Desc : 수량/금액 자동합계                                  
'============================================================================================================
Sub SubBizlsQtyAmt()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

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
