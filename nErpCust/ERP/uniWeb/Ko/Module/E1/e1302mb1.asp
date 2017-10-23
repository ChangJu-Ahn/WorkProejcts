<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd_uniSIMS
                                                               '☜: Hide Processing message
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim strYear
    Dim strEmpNo  
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strEmpNo  = lgKeyStream(0)

    Call SubEmpBase(lgKeyStream(0),lgKeyStream(1),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
%>
<Script Language=vbscript>
    With parent.frm1
        .txtEmp_no.Value = "<%=ConvSPChars(emp_no)%>"
        .txtName.Value = "<%=ConvSPChars(Name)%>"
        .txtDept_nm.value = "<%=ConvSPChars(DEPT_NM)%>"    
        .txtroll_pstn.value = "<%=ConvSPChars(roll_pstn)%>"
    End With          
</Script>       
<%

    if emp_no = "" then
        lgErrorStatus = "YES"
        exit sub
    end if 

    strEmpNo  = emp_no
    strYear   = FilterVar(lgKeyStream(2),"'%'", "S")

    Call SubMakeSQLStatements("R","")                                       '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
          Call SetErrorStatus()
       End If
    Else

%>
<Script Language=vbscript>
            
       With Parent	
'소득사항   
            .Frm1.txtNew_pay_tot_amt.Value         = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_pay_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtNew_bonus_tot_amt.Value       = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_bonus_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtafter_bonus_amt.Value         = "<%=UNINumClientFormat(lgObjRs("after_bonus_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txta_amt.value                   = "<%=UNINumClientFormat(lgObjRs("a_amt"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtbonus_tot_amt.value           = "<%=UNINumClientFormat(lgObjRs("bonus_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"          
            .frm1.txtold_after_bonus_amt.value     = "<%=UNINumClientFormat(lgObjRs("old_after_bonus_amt"), ggAmtOfMoney.DecPoint,0)%>"          
            .frm1.txtpay_tot_amt.value             = "<%=UNINumClientFormat(lgObjRs("pay_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.txtb_amt.value                   = "<%=UNINumClientFormat(lgObjRs("b_amt"), ggAmtOfMoney.DecPoint,0)%>"      
            .frm1.txtincome_tot_amt.value          = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtincome_sub_amt.value          = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.txthfa050t_income_amt.value      = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_amt"), ggAmtOfMoney.DecPoint,0)%>"         
            
'인적공재 
            .frm1.txtper_sub_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa050t_per_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtparia_cnt.value               = "<%=ConvSPChars(lgObjRs("hfa050t_paria_cnt"))%>"    
            .frm1.txtparia_sub_amt.value           = "<%=UNINumClientFormat(lgObjRs("hfa050t_paria_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtspouse.value                  = "<%=ConvSPChars(lgObjRs("hfa050t_spouse"))%>" 
            .frm1.txtspouse_sub_amt.value          = "<%=UNINumClientFormat(lgObjRs("hfa050t_spouse_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_cnt1.value                 = "<%=ConvSPChars(lgObjRs("hfa050t_old_cnt"))%>"
            .frm1.txtold_cnt2.value                 = "<%=ConvSPChars(lgObjRs("hfa050t_old_cnt2"))%>"                    
            .frm1.txtold_sub_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa050t_old_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"   
             
            .frm1.txtsupp_old_cnt.value            = "<%=ConvSPChars(lgObjRs("hdf020t_supp_old_cnt"))%>"     
            .frm1.txtsupp_sub_amt.value            = "<%=UNINumClientFormat(lgObjRs("hfa050t_supp_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtlady.value                    = "<%=ConvSPChars(lgObjRs("hfa050t_lady"))%>"  
            .frm1.txtlady_sub_amt.value            = "<%=UNINumClientFormat(lgObjRs("hfa050t_lady_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtsupp_young_cnt.value          = "<%=ConvSPChars(lgObjRs("hdf020t_supp_young_cnt"))%>"    
            .frm1.txtchl_rear.value                = "<%=UNINumClientFormat(lgObjRs("hfa050t_chl_rear"), 0,0)%>"    
            .frm1.txtchl_rear_sub_amt.value        = "<%=UNINumClientFormat(lgObjRs("hfa050t_chl_rear_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtsmall_sub_amt.value           = "<%=UNINumClientFormat(lgObjRs("hfa050t_small_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtd_amt.value                   = "<%=UNINumClientFormat(lgObjRs("d_amt"), ggAmtOfMoney.DecPoint,0)%>"                           
            
'특별세액공제          
            .frm1.txtinsur_amt.value               = "<%=UNINumClientFormat(lgObjRs("hfa050t_med_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtmed_insur_amt.value           = "<%=UNINumClientFormat(lgObjRs("hfa050t_med_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_house_fund_amt.value  = "<%=UNINumClientFormat(lgObjRs("hfa030t_house_fund_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_house_fund_amt.value  = "<%=UNINumClientFormat(lgObjRs("hfa050t_house_fund_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_emp_insur_amt.value   = "<%=UNINumClientFormat(lgObjRs("hfa030t_emp_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_emp_insur_amt.value   = "<%=UNINumClientFormat(lgObjRs("hfa050t_emp_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtlong_house_loan_amt.value     = "<%=UNINumClientFormat(lgObjRs("hfa030t_long_house_loan_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtlong_house_loan_amt1.value     = "<%=UNINumClientFormat(lgObjRs("hfa030t_long_house_loan_amt1"), ggAmtOfMoney.DecPoint,0)%>"  
                        
            .frm1.txthfa030t_other_insur_amt.value = "<%=UNINumClientFormat(lgObjRs("hfa030t_other_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_other_insur_amt.value = "<%=UNINumClientFormat(lgObjRs("hfa050t_other_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_disabled_insur_amt.value = "<%=UNINumClientFormat(lgObjRs("hfa030t_disabled_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_disabled_insur_amt.value = "<%=UNINumClientFormat(lgObjRs("hfa050t_disabled_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txttot_med_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa030t_tot_med_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtmed_sub_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa050t_med_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtlegal_contr_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa030t_legal_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtcontr_sub_amt.value           = "<%=UNINumClientFormat(lgObjRs("hfa050t_contr_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"   
            
            .frm1.txtPoli_contr_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa030t_poli_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtPoli_contr_amt1.value         = "<%=UNINumClientFormat(lgObjRs("hfa030t_poli_contr_amt1"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtOurstock_contr_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa030t_ourstock_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    
                         
            .frm1.txtspeci_med_amt.value           = "<%=UNINumClientFormat(lgObjRs("hfa030t_speci_med_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtapp_contr_amt.value           = "<%=UNINumClientFormat(lgObjRs("hfa030t_app_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtTaxLaw_contr_amt.value           = "<%=UNINumClientFormat(lgObjRs("hfa030t_TaxLaw_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"                    
            .frm1.txtper_edu_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa030t_per_edu_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtDisabled_edu_amt.value        = "<%=UNINumClientFormat(lgObjRs("hfa030t_disabled_edu_amt"), ggAmtOfMoney.DecPoint,0)%>"            
            .frm1.txtedu_sub_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa050t_edu_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtpriv_contr_amt.value          = "<%=UNINumClientFormat(lgObjRs("hfa030t_priv_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtOur_stock_amt.value          = "<%=UNINumClientFormat(lgObjRs("hfa050t_our_stock_amt"), ggAmtOfMoney.DecPoint,0)%>"                    
            .frm1.txtedu_sum_amt.value             = "<%=UNINumClientFormat(lgObjRs("edu_sum_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtstd_sub_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa050t_std_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_indiv_anu_amt.value   = "<%=UNINumClientFormat(lgObjRs("hfa030t_indiv_anu_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_indiv_anu_amt.value   = "<%=UNINumClientFormat(lgObjRs("hfa050t_indiv_anu_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_National_pension_amt.value   = "<%=UNINumClientFormat(lgObjRs("hfa030t_National_pension_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtFore_edu_amt.value   = "<%=UNINumClientFormat(lgObjRs("hfa050t_fore_edu_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_National_pension_sub_amt.value   = "<%=UNINumClientFormat(lgObjRs("hfa050t_National_pension_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtinvest_sub_sum_amt.value      = "<%=UNINumClientFormat(lgObjRs("hfa050t_invest_sub_sum_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtcard_sub_sum_amt.value        = "<%=UNINumClientFormat(lgObjRs("hfa050t_card_sub_sum_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtsum_amt.value                 = "<%=UNINumClientFormat(lgObjRs("sum_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txttax_std_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa050t_tax_std_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtcalu_tax_amt.value            = "<%=UNINumClientFormat(lgObjRs("hfa050t_calu_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtincome_tax_sub_amt.value      = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_tax_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"
			.frm1.txtFore_pay_amt.value      = "<%=UNINumClientFormat(lgObjRs("hfa050t_Fore_pay"), ggAmtOfMoney.DecPoint,0)%>"                    
            .frm1.txthouse_repay_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa050t_house_repay_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txttax_sub_sum_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa050t_tax_sub_sum_amt"), ggAmtOfMoney.DecPoint,0)%>"        
'2004 
            .frm1.hfa030t_Ceremony_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa030t_Ceremony_amt"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.hfa050t_Ceremony_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa050t_Ceremony_amt"), ggAmtOfMoney.DecPoint,0)%>"        
            
'결정세액/차감징수세액 
            .frm1.txtdec_income_tax_amt.value      = "<%=UNINumClientFormat(lgObjRs("hfa050t_dec_income_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtdec_res_tax_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa050t_dec_res_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtdec_farm_tax_amt.value        = "<%=UNINumClientFormat(lgObjRs("hfa050t_dec_farm_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtdec_amt.value                 = "<%=UNINumClientFormat(lgObjRs("dec_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtnew_income_tax_amt.value      = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_income_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtnew_res_tax_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_res_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtnew_farm_tax_amt.value        = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_farm_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtincome_amt.value              = "<%=UNINumClientFormat(lgObjRs("new_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_income_tax_amt.value      = "<%=UNINumClientFormat(lgObjRs("hfa050t_old_income_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_res_tax_amt.value         = "<%=UNINumClientFormat(lgObjRs("hfa050t_old_res_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_farm_tax_amt.value        = "<%=UNINumClientFormat(lgObjRs("hfa050t_old_farm_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_amt.value                 = "<%=UNINumClientFormat(lgObjRs("old_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtincome_tax_amt.value          = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtres_tax_amt.value             = "<%=UNINumClientFormat(lgObjRs("hfa050t_res_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtfarm_tax_amt.value            = "<%=UNINumClientFormat(lgObjRs("hfa050t_farm_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtf_amt.value                   = "<%=UNINumClientFormat(lgObjRs("f_amt"), ggAmtOfMoney.DecPoint,0)%>"     
                        
       End With          
</Script>       
<%     
    End If
	Call SubCloseRs(lgObjRs)
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount

    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
            Case ""
                lgStrSQL = "SELECT a.name Name2, "
                lgStrSQL = lgStrSQL & " b.SPOUSE hfa050t_spouse, b.LADY hfa050t_lady, b.PARIA_CNT hfa050t_paria_cnt, "  
                lgStrSQL = lgStrSQL & " b.OLD_CNT hfa050t_old_cnt,b.OLD_CNT2 hfa050t_old_cnt2, b.CHL_REAR hfa050t_chl_rear, "  
                lgStrSQL = lgStrSQL & " b.INCOME_TOT_AMT hfa050t_income_tot_amt,"
                lgStrSQL = lgStrSQL & " b.INCOME_SUB hfa050t_income_sub_amt, b.INCOME_AMT hfa050t_income_amt, "  
                lgStrSQL = lgStrSQL & " b.PER_SUB hfa050t_per_sub_amt, b.SPOUSE_SUB hfa050t_spouse_sub_amt, "  
                lgStrSQL = lgStrSQL & " b.SUPP_SUB hfa050t_supp_sub_amt, b.OLD_SUB hfa050t_old_sub_amt,"  
                lgStrSQL = lgStrSQL & " b.PARIA_SUB hfa050t_paria_sub_amt, b.LADY_SUB hfa050t_lady_sub_amt,"   
                lgStrSQL = lgStrSQL & " b.CHL_REAR_SUB hfa050t_chl_rear_sub_amt, b.SMALL_SUB hfa050t_small_sub_amt, "  
                lgStrSQL = lgStrSQL & " b.OTHER_INSUR hfa050t_other_insur_amt, b.Disabled_sub_amt hfa050t_disabled_insur_amt, b.MED_INSUR hfa050t_med_insur_amt, "  
                lgStrSQL = lgStrSQL & " b.EMP_INSUR hfa050t_emp_insur_amt, b.MED_SUB hfa050t_med_sub_amt, "  
                lgStrSQL = lgStrSQL & " b.EDU_SUB hfa050t_edu_sub_amt, b.HOUSE_FUND hfa050t_house_fund_amt, "  
                lgStrSQL = lgStrSQL & " b.CONTR_SUB hfa050t_contr_sub_amt, b.STD_SUB hfa050t_std_sub_amt, " 
                lgStrSQL = lgStrSQL & " (b.INDIV_ANU +b.INDIV_ANU2) hfa050t_indiv_anu_amt, b.National_pension_sub_amt hfa050t_National_pension_sub_amt, b.TECH_SUB hfa050t_tech_sub_amt, "  
                lgStrSQL = lgStrSQL & " b.TAX_STD hfa050t_tax_std_amt, b.CALU_TAX hfa050t_calu_tax_amt, "  
                lgStrSQL = lgStrSQL & " b.INCOME_TAX_SUB hfa050t_income_tax_sub_amt,b.fore_pay hfa050t_fore_pay, b.HOUSE_REPAY hfa050t_house_repay_amt,"
                lgStrSQL = lgStrSQL & " b.STOCK_SAVE hfa050t_stock_save_amt, b.TAX_SUB_SUM hfa050t_tax_sub_sum_amt,"   
                lgStrSQL = lgStrSQL & " b.DEC_INCOME_TAX hfa050t_dec_income_tax_amt, b.DEC_FARM_TAX hfa050t_dec_farm_tax_amt,  " 
                lgStrSQL = lgStrSQL & " b.DEC_RES_TAX hfa050t_dec_res_tax_amt, b.OLD_INCOME_TAX hfa050t_old_income_tax_amt, "  
                lgStrSQL = lgStrSQL & " b.OLD_FARM_TAX hfa050t_old_farm_tax_amt, b.OLD_RES_TAX hfa050t_old_res_tax_amt,  " 
                lgStrSQL = lgStrSQL & " b.NEW_INCOME_TAX hfa050t_new_income_tax_amt, b.NEW_FARM_TAX hfa050t_new_farm_tax_amt, "  
                lgStrSQL = lgStrSQL & " b.NEW_RES_TAX hfa050t_new_res_tax_amt, b.INCOME_TAX hfa050t_income_tax_amt,  " 
                lgStrSQL = lgStrSQL & " b.FARM_TAX hfa050t_farm_tax_amt, b.RES_TAX hfa050t_res_tax_amt,  " 
                lgStrSQL = lgStrSQL & " b.CARD_SUB_SUM hfa050t_card_sub_sum_amt,b.fore_edu_sub_amt hfa050t_fore_edu_sub_amt, b.NEW_PAY_TOT hfa050t_new_pay_tot_amt, "  
                lgStrSQL = lgStrSQL & " b.Ceremony_amt hfa050t_Ceremony_amt," '2004
                lgStrSQL = lgStrSQL & " b.NEW_BONUS_TOT hfa050t_new_bonus_tot_amt, b.INVEST_SUB_SUM hfa050t_invest_sub_sum_amt,"
                lgStrSQL = lgStrSQL & " b.INCOME_SHORT, "   
                lgStrSQL = lgStrSQL & " b.TAX_SHORT, "
                lgStrSQL = lgStrSQL & " b.INCOME_REDU, "   
                lgStrSQL = lgStrSQL & " b.TAXES_REDU, "  
                lgStrSQL = lgStrSQL & " b.REDU_SUM, " 
                lgStrSQL = lgStrSQL & " b.NON_TAX1, "  
                lgStrSQL = lgStrSQL & " b.NON_TAX2, "  
                lgStrSQL = lgStrSQL & " b.NON_TAX3, " 
                lgStrSQL = lgStrSQL & " b.NON_TAX4, "  
                lgStrSQL = lgStrSQL & " b.NON_TAX5, "   
                lgStrSQL = lgStrSQL & " b.SAVE_FUND, "  
                lgStrSQL = lgStrSQL & " b.SUPP_CNT, "
                lgStrSQL = lgStrSQL & " b.INSUR_SUB, "
                lgStrSQL = lgStrSQL & " b.SUB_INCOME_AMT,"    
                lgStrSQL = lgStrSQL & " b.FORE_PAY,"    
                lgStrSQL = lgStrSQL & " c.EMP_NO Emp_no2,"   
                lgStrSQL = lgStrSQL & " c.OTHER_INCOME,"
                lgStrSQL = lgStrSQL & " c.FORE_INCOME, "  
                lgStrSQL = lgStrSQL & " c.EDU_SPPORT, "  
                lgStrSQL = lgStrSQL & " c.MED_SPPORT, "  
                lgStrSQL = lgStrSQL & " c.MED_INSUR, "
                lgStrSQL = lgStrSQL & " c.FAM_EDU, "  
                lgStrSQL = lgStrSQL & " c.UNIV_EDU, "  
                lgStrSQL = lgStrSQL & " c.KIND_EDU, "  
                lgStrSQL = lgStrSQL & " c.KIND_EDU_CNT, "  
                lgStrSQL = lgStrSQL & " c.UNIV_EDU_CNT, "  
                lgStrSQL = lgStrSQL & " (c.HOUSE_FUND + c.LONG_HOUSE_LOAN_AMT)  HOUSE_AMT, "
                lgStrSQL = lgStrSQL & " (c.INDIV_ANU +c.INDIV_ANU2) hfa030t_indiv_anu_amt, "  
                lgStrSQL = lgStrSQL & " c.National_pension_amt hfa030t_National_pension_amt, "  
                lgStrSQL = lgStrSQL & " c.SAVE_TAX_SUB, "  
                lgStrSQL = lgStrSQL & " c.HOUSE_REPAY, "  
                lgStrSQL = lgStrSQL & " c.STOCK_SAVE, "  
                lgStrSQL = lgStrSQL & " c.FORE_PAY, "  
                lgStrSQL = lgStrSQL & " c.INCOME_REDU,"   
                lgStrSQL = lgStrSQL & " c.TAXES_REDU, "  
                lgStrSQL = lgStrSQL & " c.TECH_SUB_AMT, "  
                lgStrSQL = lgStrSQL & " c.INVEST_SUB_AMT, "  
                lgStrSQL = lgStrSQL & " c.VENTURE_SUB_AMT, "
                lgStrSQL = lgStrSQL & " c.CEREMONY_AMT hfa030t_Ceremony_amt, "  '2004
				lgStrSQL = lgStrSQL & " c.POLI_CONTRA_AMT1 hfa030t_poli_contr_amt, c.POLI_CONTRA_AMT2 hfa030t_poli_contr_amt1, "  '2004 기부금 
				lgStrSQL = lgStrSQL & " c.OURSTOCK_CONTRA_AMT hfa030t_ourstock_contr_amt, "  '2004 기부금                   
                lgStrSQL = lgStrSQL & " c.OTHER_INSUR hfa030t_other_insur_amt, c.disabled_sub_amt hfa030t_disabled_insur_amt, c.EMP_INSUR hfa030t_emp_insur_amt, "  
                lgStrSQL = lgStrSQL & " c.TOT_MED hfa030t_tot_med_amt, c.SPECI_MED hfa030t_speci_med_amt, "  
                lgStrSQL = lgStrSQL & " c.PER_EDU hfa030t_per_edu_amt,c.disabled_edu_amt hfa030t_disabled_edu_amt, b.our_stock_amt hfa050t_our_stock_amt, "  
                lgStrSQL = lgStrSQL & " c.APP_CONTR hfa030t_app_contr_amt,c.TaxLaw_contr_amt hfa030t_TaxLaw_contr_amt, c.PRIV_CONTR hfa030t_priv_contr_amt, c.PRIV_CONTR hfa030t_priv_contr_amt,"  
                lgStrSQL = lgStrSQL & " c.HOUSE_FUND hfa030t_house_fund_amt, c.LONG_HOUSE_LOAN_AMT hfa030t_long_house_loan_amt,c.LONG_HOUSE_LOAN_AMT1 hfa030t_long_house_loan_amt1,"   
                lgStrSQL = lgStrSQL & " c.AFTER_BONUS_AMT after_bonus_amt, d.SUPP_OLD_CNT hdf020t_supp_old_cnt, "  
                lgStrSQL = lgStrSQL & " d.SUPP_YOUNG_CNT hdf020t_supp_young_cnt, "
                lgStrSQL = lgStrSQL & " T.PAY_TOT_AMT PAY_TOT_AMT, "
                lgStrSQL = lgStrSQL & " T.BONUS_TOT_AMT BONUS_TOT_AMT, "
                lgStrSQL = lgStrSQL & " T.MED_INSUR_AMT, "
                lgStrSQL = lgStrSQL & " T.AFTER_BONUS_AMT   old_after_bonus_amt, "
                lgStrSQL = lgStrSQL & " C.legal_contr   hfa030t_legal_contr_amt, "
                lgStrSQL = lgStrSQL & " (b.NEW_PAY_TOT + b.NEW_BONUS_TOT + c.AFTER_BONUS_AMT) a_amt ,"
                lgStrSQL = lgStrSQL & " (T.PAY_TOT_AMT + T.BONUS_TOT_AMT + T.AFTER_BONUS_AMT ) b_amt ,"
                lgStrSQL = lgStrSQL & " (b.PER_SUB + b.SPOUSE_SUB + b.SUPP_SUB + b.OLD_SUB + b.PARIA_SUB + b.LADY_SUB + "  
                lgStrSQL = lgStrSQL & " b.CHL_REAR_SUB + b.SMALL_SUB ) d_amt ,"
                lgStrSQL = lgStrSQL & " (b.PER_SUB + b.SPOUSE_SUB + b.SUPP_SUB + b.OLD_SUB + b.PARIA_SUB + b.LADY_SUB + "  
                lgStrSQL = lgStrSQL & " b.CHL_REAR_SUB + b.SMALL_SUB + b.STD_SUB + b.INDIV_ANU  + b.National_pension_sub_amt + b.TECH_SUB + b.INVEST_SUB_SUM + "
                lgStrSQL = lgStrSQL & " b.CARD_SUB_SUM + b.our_stock_amt) sum_amt ,"
                lgStrSQL = lgStrSQL & " (c.FAM_EDU + c.UNIV_EDU + c.KIND_EDU ) edu_sum_amt ,"
                lgStrSQL = lgStrSQL & " (b.DEC_INCOME_TAX + b.DEC_FARM_TAX + b.DEC_RES_TAX ) dec_amt ,"
                lgStrSQL = lgStrSQL & " (b.NEW_INCOME_TAX + b.NEW_FARM_TAX + b.NEW_RES_TAX ) new_amt ,"
                lgStrSQL = lgStrSQL & " (b.OLD_INCOME_TAX + b.OLD_FARM_TAX + b.OLD_RES_TAX ) old_amt ,"
                lgStrSQL = lgStrSQL & " (b.INCOME_TAX + b.FARM_TAX + b.RES_TAX ) f_amt "

                lgStrSQL = lgStrSQL & " FROM HAA010T a, "  
                lgStrSQL = lgStrSQL & "		 HFA050T b, "  
                lgStrSQL = lgStrSQL & "		 HFA030T c, "  
                lgStrSQL = lgStrSQL & "		 HDF020T d, "  
                lgStrSQL = lgStrSQL & " (SELECT EMP_NO, SUM(A_PAY_TOT_AMT) PAY_TOT_AMT, SUM(A_BONUS_TOT_AMT) BONUS_TOT_AMT, SUM(A_MED_INSUR) MED_INSUR_AMT," 
                lgStrSQL = lgStrSQL & " SUM(A_AFTER_BONUS_AMT) AFTER_BONUS_AMT FROM HFA040T "
                lgStrSQL = lgStrSQL & " WHERE YEAR_YY = " & FilterVar(lgKeyStream(2),"'%'", "S")
                lgStrSQL = lgStrSQL & " GROUP BY EMP_NO) AS T "
                lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.EMP_NO "
                lgStrSQL = lgStrSQL & " AND b.YEAR_YY = " & FilterVar(lgKeyStream(2), "''", "S")
                lgStrSQL = lgStrSQL & " AND b.YEAR_YY = c.YY "
                lgStrSQL = lgStrSQL & " AND b.EMP_NO = c.EMP_NO "
                lgStrSQL = lgStrSQL & " AND b.EMP_NO *= T.EMP_NO "
                lgStrSQL = lgStrSQL & " AND a.EMP_NO = d.EMP_NO "
                lgStrSQL = lgStrSQL & " AND a.emp_no = " & FilterVar(lgKeyStream(0),"'%'", "S")                       
'                lgStrSQL = lgStrSQL & " AND b.internal_cd like  '%'"
            Case "P"
                lgStrSQL = "Select TOP 1 * " 
                lgStrSQL = lgStrSQL & " From  HAA010T "
                lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
            Case "N"
                lgStrSQL = "Select TOP 1 * " 
                lgStrSQL = lgStrSQL & " From  HAA010T "
                lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
             End Select
      Case "C"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
      Case "U"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
      Case "D"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
    End Select
End Sub
                     

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)    'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)   'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub
%>

<Script Language="VBScript">

    Select Case "<%=lgOpModeCRUD %>"
       Case "UID_M0001"                                                         '☜ : Query
         If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                 .DBQueryOk        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             'parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	

