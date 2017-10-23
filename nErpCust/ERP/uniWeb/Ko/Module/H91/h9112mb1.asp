<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    call LoadBasisGlobalInf()
	Call loadInfTB19029B("Q", "H","NOCOOKIE","MB")
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim YEAR
    Dim EMPNO    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    YEAR  = FilterVar(lgKeyStream(0),"'%'", "S")
    EMPNO = FilterVar(lgKeyStream(1),"'%'", "S")
   
    Call SubMakeSQLStatements("R","")                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the starting data. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else

%>
<Script Language=vbscript>
       Dim DEC_AMT
	             
       With Parent	
            
'소득사항   
            .Frm1.txtNew_pay_tot_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_pay_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtNew_bonus_tot_amt.text       = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_bonus_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtafter_bonus_amt.text         = "<%=UNINumClientFormat(lgObjRs("after_bonus_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txta_amt.text                   = "<%=UNINumClientFormat(lgObjRs("a_amt"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtbonus_tot_amt.text           = "<%=UNINumClientFormat(lgObjRs("bonus_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"          
            .frm1.txtold_after_bonus_amt.text     = "<%=UNINumClientFormat(lgObjRs("old_after_bonus_amt"), ggAmtOfMoney.DecPoint,0)%>"          
            .frm1.txtpay_tot_amt.text             = "<%=UNINumClientFormat(lgObjRs("pay_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.txtb_amt.text                   = "<%=UNINumClientFormat(lgObjRs("b_amt"), ggAmtOfMoney.DecPoint,0)%>"      
            .frm1.txtincome_tot_amt.text          = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_tot_amt"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtincome_sub_amt.text          = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.txthfa050t_income_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_amt"), ggAmtOfMoney.DecPoint,0)%>"         
            
'인적공제
            .frm1.txtper_sub_amt.text             = "<%=UNINumClientFormat(lgObjRs("hfa050t_per_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtparia_cnt.text               = "<%=ConvSPChars(lgObjRs("hfa050t_paria_cnt"))%>"    
            .frm1.txtparia_sub_amt.text           = "<%=UNINumClientFormat(lgObjRs("hfa050t_paria_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtspouse.value                  = "<%=ConvSPChars(lgObjRs("hfa050t_spouse"))%>" 
            .frm1.txtspouse_sub_amt.text          = "<%=UNINumClientFormat(lgObjRs("hfa050t_spouse_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_cnt1.text                 = "<%=ConvSPChars(lgObjRs("hfa050t_old_cnt"))%>"   '2004 경로우대공제(65세이상)
            .frm1.txtold_cnt2.text                 = "<%=ConvSPChars(lgObjRs("hfa050t_old_cnt2"))%>"   '2004 경로우대공제(70세이상)             
            .frm1.txtold_sub_amt1.text             = "<%=UNINumClientFormat(lgObjRs("hfa050t_old_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"  '2004 경로우대공제(65세이상)
            .frm1.txtsupp_old_cnt.text            = "<%=ConvSPChars(lgObjRs("hdf020t_supp_old_cnt"))%>"     
            .frm1.txtsupp_sub_amt.text            = "<%=UNINumClientFormat(lgObjRs("hfa050t_supp_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtlady.value                   = "<%=ConvSPChars(lgObjRs("hfa050t_lady"))%>"  
            .frm1.txtlady_sub_amt.text            = "<%=UNINumClientFormat(lgObjRs("hfa050t_lady_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtsupp_young_cnt.text          = "<%=ConvSPChars(lgObjRs("hdf020t_supp_young_cnt"))%>"    
            .frm1.txtchl_rear.text                = "<%=UNINumClientFormat(lgObjRs("hfa050t_chl_rear"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtchl_rear_sub_amt.text        = "<%=UNINumClientFormat(lgObjRs("hfa050t_chl_rear_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtsmall_sub_amt.text           = "<%=UNINumClientFormat(lgObjRs("hfa050t_small_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtd_amt.text                   = "<%=UNINumClientFormat(lgObjRs("d_amt"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.hfa050t_small_sub_amt_txtsmall_sub_amt9.text           = "<%=UNINumClientFormat(lgObjRs("hfa050t_young_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"   '2007 다자녀추가공제                        
            
'특별세액공제          
            .frm1.txtinsur_amt.text               = "<%=UNINumClientFormat(lgObjRs("hfa030t_med_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtmed_insur_amt.text           = "<%=UNINumClientFormat(lgObjRs("hfa050t_med_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_house_fund_amt.text  = "<%=UNINumClientFormat(lgObjRs("hfa030t_house_fund_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_house_fund_amt.text  = "<%=UNINumClientFormat(lgObjRs("hfa050t_house_fund_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_emp_insur_amt.text   = "<%=UNINumClientFormat(lgObjRs("hfa030t_emp_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_emp_insur_amt.text   = "<%=UNINumClientFormat(lgObjRs("hfa050t_emp_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtlong_house_loan_amt.text     = "<%=UNINumClientFormat(lgObjRs("hfa030t_long_house_loan_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtlong_house_loan_amt1.text     = "<%=UNINumClientFormat(lgObjRs("hfa030t_long_house_loan_amt1"), ggAmtOfMoney.DecPoint,0)%>"    '2004

            .frm1.txthfa030t_other_insur_amt.text = "<%=UNINumClientFormat(lgObjRs("hfa030t_other_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_other_insur_amt.text = "<%=UNINumClientFormat(lgObjRs("hfa050t_other_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_disabled_insur_amt.text = "<%=UNINumClientFormat(lgObjRs("hfa030t_disabled_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_disabled_insur_amt.text = "<%=UNINumClientFormat(lgObjRs("hfa050t_disabled_insur_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txttot_med_amt.text             = "<%=UNINumClientFormat(lgObjRs("hfa030t_tot_med_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtmed_sub_amt.text             = "<%=UNINumClientFormat(lgObjRs("hfa050t_med_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtlegal_contr_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa030t_legal_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtcontr_sub_amt.text           = "<%=UNINumClientFormat(lgObjRs("hfa050t_contr_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtspeci_med_amt.text           = "<%=UNINumClientFormat(lgObjRs("hfa030t_speci_med_amt"), ggAmtOfMoney.DecPoint,0)%>"    
           
            .frm1.txtapp_contr_amt.text           = "<%=UNINumClientFormat(lgObjRs("hfa030t_app_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtper_edu_amt.text             = "<%=UNINumClientFormat(lgObjRs("hfa030t_per_edu_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtedu_sub_amt.text             = "<%=UNINumClientFormat(lgObjRs("hfa050t_edu_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            '2002년 추가 
            .frm1.txtDisable_edu_amt.text          = "<%=UNINumClientFormat(lgObjRs("hfa030t_disabled_edu_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtedu_sum_amt.text             = "<%=UNINumClientFormat(lgObjRs("edu_sum_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtstd_sub_amt.text             = "<%=UNINumClientFormat(lgObjRs("hfa050t_std_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"
            
            '2003
            .frm1.hfa030t_fore_edu_amt.text       = "<%=UNINumClientFormat(lgObjRs("fore_edu_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.hfa050t_fore_edu_sub_amt.text   = "<%=UNINumClientFormat(lgObjRs("fore_edu_sub_amt"), ggAmtOfMoney.DecPoint,0)%>" 
                               
			'2004 결혼장례비 
            .frm1.hfa030t_Ceremony_amt.text       = "<%=UNINumClientFormat(lgObjRs("HFA030T_CEREMONY_AMT"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.hfa050t_Ceremony_amt.text		  = "<%=UNINumClientFormat(lgObjRs("HFA050T_CEREMONY_AMT"), ggAmtOfMoney.DecPoint,0)%>"                    
			'2004 기부금 추가 
            .frm1.txtlegal_contr_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa030t_legal_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtPoli_contr_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa030t_poli_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtOurstock_contr_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa030t_ourstock_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtpriv_contr_amt.text          = "<%=UNINumClientFormat(lgObjRs("hfa030t_priv_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"    

            .frm1.txthfa030t_indiv_anu_amt.text   = "<%=UNINumClientFormat(lgObjRs("hfa030t_indiv_anu_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_indiv_anu_amt.text   = "<%=UNINumClientFormat(lgObjRs("hfa050t_indiv_anu_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa030t_National_pension_amt.text   = "<%=UNINumClientFormat(lgObjRs("hfa030t_National_pension_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthfa050t_National_pension_sub_amt.text   = "<%=UNINumClientFormat(lgObjRs("hfa050t_National_pension_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            '.frm1.txttech_sub_amt.text            = "<%=UNINumClientFormat(lgObjRs("hfa050t_tech_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtinvest_sub_sum_amt.text      = "<%=UNINumClientFormat(lgObjRs("hfa050t_invest_sub_sum_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtcard_sub_sum_amt.text        = "<%=UNINumClientFormat(lgObjRs("hfa050t_card_sub_sum_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtsum_amt.text                 = "<%=UNINumClientFormat(lgObjRs("sum_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txttax_std_amt.text             = "<%=UNINumClientFormat(lgObjRs("hfa050t_tax_std_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtcalu_tax_amt.text            = "<%=UNINumClientFormat(lgObjRs("hfa050t_calu_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtincome_tax_sub_amt.text      = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_tax_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txthouse_repay_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa050t_house_repay_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtPolicontr_tax_sub_amt.text	  = "<%=UNINumClientFormat(lgObjRs("hfa050t_poli_tax_sub"), ggAmtOfMoney.DecPoint,0)%>"   
            .frm1.txtTax_Union_Ded.text          = "<%=UNINumClientFormat(lgObjRs("hfa050t_TAX_UNION_DED"), ggAmtOfMoney.DecPoint,0)%>"   '2005
            .frm1.txttax_sub_sum_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa050t_tax_sub_sum_amt"), ggAmtOfMoney.DecPoint,0)%>"        

'2002 추가 
			.frm1.txtOur_Stock_sub_amt.text       = "<%=UNINumClientFormat(lgObjRs("hfa050t_our_stock_amt"), ggAmtOfMoney.DecPoint,0)%>"        
			.frm1.txtTaxLaw_contr_amt.text        = "<%=UNINumClientFormat(lgObjRs("hfa030t_TaxLaw_contr_amt"), ggAmtOfMoney.DecPoint,0)%>" 
			.frm1.txtFore_pay.text                = "<%=UNINumClientFormat(lgObjRs("hfa050t_fore_pay"), ggAmtOfMoney.DecPoint,0)%>"  
'2005 
			.frm1.txtTaxLaw_contr_amt2.text        = "<%=UNINumClientFormat(lgObjRs("hfa030t_TaxLaw_contr_amt2"), ggAmtOfMoney.DecPoint,0)%>"     
            .frm1.hfa030t_retire_pension.text   = "<%=UNINumClientFormat(lgObjRs("hfa030t_retire_pension"), ggAmtOfMoney.DecPoint,0)%>" '2005
            .frm1.hfa050t_retire_pension.text   = "<%=UNINumClientFormat(lgObjRs("hfa050t_retire_pension"), ggAmtOfMoney.DecPoint,0)%>" '2005
            			      
'결정세액/차감징수세액 
            .frm1.txtdec_income_tax_amt.text      = "<%=UNINumClientFormat(lgObjRs("hfa050t_dec_income_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtdec_res_tax_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa050t_dec_res_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtdec_farm_tax_amt.text        = "<%=UNINumClientFormat(lgObjRs("hfa050t_dec_farm_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtdec_amt.text                 = "<%=UNINumClientFormat(lgObjRs("dec_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtnew_income_tax_amt.text      = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_income_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtnew_res_tax_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_res_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtnew_farm_tax_amt.text        = "<%=UNINumClientFormat(lgObjRs("hfa050t_new_farm_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtincome_amt.text              = "<%=UNINumClientFormat(lgObjRs("new_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_income_tax_amt.text      = "<%=UNINumClientFormat(lgObjRs("hfa050t_old_income_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_res_tax_amt.text         = "<%=UNINumClientFormat(lgObjRs("hfa050t_old_res_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_farm_tax_amt.text        = "<%=UNINumClientFormat(lgObjRs("hfa050t_old_farm_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtold_amt.text                 = "<%=UNINumClientFormat(lgObjRs("old_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtincome_tax_amt.text          = "<%=UNINumClientFormat(lgObjRs("hfa050t_income_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtres_tax_amt.text             = "<%=UNINumClientFormat(lgObjRs("hfa050t_res_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtfarm_tax_amt.text            = "<%=UNINumClientFormat(lgObjRs("hfa050t_farm_tax_amt"), ggAmtOfMoney.DecPoint,0)%>"    
            .frm1.txtf_amt.text                   = "<%=UNINumClientFormat(lgObjRs("f_amt"), ggAmtOfMoney.DecPoint,0)%>"     
                        
       End With          
</Script>       
<%
    End If
    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
	
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '☜ : Update
              Call SubBizSaveSingleUpdate()
    End Select

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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
Select Case pMode 
  Case "R"
       Select Case  lgPrevNext 
            Case ""
                lgStrSQL = "SELECT HAA010T.name Name2, "
                lgStrSQL = lgStrSQL & " HFA050T.SPOUSE hfa050t_spouse, HFA050T.LADY hfa050t_lady, HFA050T.PARIA_CNT hfa050t_paria_cnt, "  
                lgStrSQL = lgStrSQL & " HFA050T.OLD_CNT hfa050t_old_cnt,HFA050T.OLD_CNT2 hfa050t_old_cnt2, HFA050T.CHL_REAR hfa050t_chl_rear, HFA050T.PAY_TOT_AMT pay_tot_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.BONUS_TOT_AMT bonus_tot_amt, HFA050T.INCOME_TOT_AMT hfa050t_income_tot_amt,"
                lgStrSQL = lgStrSQL & " HFA050T.INCOME_SUB hfa050t_income_sub_amt, HFA050T.INCOME_AMT hfa050t_income_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.PER_SUB hfa050t_per_sub_amt, HFA050T.SPOUSE_SUB hfa050t_spouse_sub_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.SUPP_SUB hfa050t_supp_sub_amt, HFA050T.OLD_SUB hfa050t_old_sub_amt, " 
                lgStrSQL = lgStrSQL & " HFA050T.PARIA_SUB hfa050t_paria_sub_amt, HFA050T.LADY_SUB hfa050t_lady_sub_amt,"   
                lgStrSQL = lgStrSQL & " HFA050T.CHL_REAR_SUB hfa050t_chl_rear_sub_amt, HFA050T.SMALL_SUB hfa050t_small_sub_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.OTHER_INSUR hfa050t_other_insur_amt, HFA050T.Disabled_sub_amt hfa050t_disabled_insur_amt, HFA050T.MED_INSUR hfa050t_med_insur_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.EMP_INSUR hfa050t_emp_insur_amt, HFA050T.MED_SUB hfa050t_med_sub_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.EDU_SUB hfa050t_edu_sub_amt, HFA050T.HOUSE_FUND hfa050t_house_fund_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.CONTR_SUB hfa050t_contr_sub_amt, HFA050T.STD_SUB hfa050t_std_sub_amt, "
                lgStrSQL = lgStrSQL & " HFA050T.FORE_EDU_SUB_AMT, " '2003
                                 
                lgStrSQL = lgStrSQL & " (HFA050T.INDIV_ANU + HFA050T.INDIV_ANU2) hfa050t_indiv_anu_amt, HFA050T.National_pension_sub_amt hfa050t_National_pension_sub_amt, HFA050T.TECH_SUB hfa050t_tech_sub_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.TAX_STD hfa050t_tax_std_amt, HFA050T.CALU_TAX hfa050t_calu_tax_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.INCOME_TAX_SUB hfa050t_income_tax_sub_amt, HFA050T.HOUSE_REPAY hfa050t_house_repay_amt, HFA050T.poli_tax_sub hfa050t_poli_tax_sub,HFA050T.TAX_UNION_DED hfa050t_tax_union_ded,"
                lgStrSQL = lgStrSQL & " HFA050T.STOCK_SAVE hfa050t_stock_save_amt, HFA050T.TAX_SUB_SUM hfa050t_tax_sub_sum_amt,"   
                lgStrSQL = lgStrSQL & " HFA050T.DEC_INCOME_TAX hfa050t_dec_income_tax_amt, HFA050T.DEC_FARM_TAX hfa050t_dec_farm_tax_amt,  " 
                lgStrSQL = lgStrSQL & " HFA050T.DEC_RES_TAX hfa050t_dec_res_tax_amt, HFA050T.OLD_INCOME_TAX hfa050t_old_income_tax_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.OLD_FARM_TAX hfa050t_old_farm_tax_amt, HFA050T.OLD_RES_TAX hfa050t_old_res_tax_amt,  " 
                lgStrSQL = lgStrSQL & " HFA050T.NEW_INCOME_TAX hfa050t_new_income_tax_amt, HFA050T.NEW_FARM_TAX hfa050t_new_farm_tax_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.NEW_RES_TAX hfa050t_new_res_tax_amt, HFA050T.INCOME_TAX hfa050t_income_tax_amt,  " 
                lgStrSQL = lgStrSQL & " HFA050T.FARM_TAX hfa050t_farm_tax_amt, HFA050T.RES_TAX hfa050t_res_tax_amt,  " 
                lgStrSQL = lgStrSQL & " HFA050T.CARD_SUB_SUM hfa050t_card_sub_sum_amt, HFA050T.NEW_PAY_TOT hfa050t_new_pay_tot_amt, "  
                lgStrSQL = lgStrSQL & " HFA050T.NEW_BONUS_TOT hfa050t_new_bonus_tot_amt, HFA050T.INVEST_SUB_SUM hfa050t_invest_sub_sum_amt,"
                lgStrSQL = lgStrSQL & " HFA050T.INCOME_SHORT, "   
                lgStrSQL = lgStrSQL & " HFA050T.TAX_SHORT, "
                lgStrSQL = lgStrSQL & " HFA050T.INCOME_REDU, "   
                lgStrSQL = lgStrSQL & " HFA050T.TAXES_REDU, "  
                lgStrSQL = lgStrSQL & " HFA050T.REDU_SUM, " 
                lgStrSQL = lgStrSQL & " HFA050T.NON_TAX1, "  
                lgStrSQL = lgStrSQL & " HFA050T.NON_TAX2, "  
                lgStrSQL = lgStrSQL & " HFA050T.NON_TAX3, " 
                lgStrSQL = lgStrSQL & " HFA050T.NON_TAX4, "  
                lgStrSQL = lgStrSQL & " HFA050T.NON_TAX5, "   
                lgStrSQL = lgStrSQL & " HFA050T.SAVE_FUND, "  
                lgStrSQL = lgStrSQL & " HFA050T.SUPP_CNT, "
                lgStrSQL = lgStrSQL & " HFA050T.INSUR_SUB, "
                lgStrSQL = lgStrSQL & " HFA050T.SUB_INCOME_AMT,"    
                lgStrSQL = lgStrSQL & " HFA050T.FORE_PAY HFA050T_Fore_pay,"    
                lgStrSQL = lgStrSQL & " HFA050T.our_stock_amt hfa050t_our_stock_amt,"    '2002추가 
                lgStrSQL = lgStrSQL & " HFA030T.EMP_NO Emp_no2,"   
                lgStrSQL = lgStrSQL & " HFA030T.OTHER_INCOME,"
                lgStrSQL = lgStrSQL & " HFA030T.FORE_INCOME, "  
                lgStrSQL = lgStrSQL & " HFA030T.EDU_SPPORT, "  
                lgStrSQL = lgStrSQL & " HFA030T.MED_SPPORT, "  
                lgStrSQL = lgStrSQL & " HFA030T.MED_INSUR hfa030t_med_insur_amt, "
                lgStrSQL = lgStrSQL & " HFA030T.FAM_EDU, "  
                lgStrSQL = lgStrSQL & " HFA030T.UNIV_EDU, "  
                lgStrSQL = lgStrSQL & " HFA030T.KIND_EDU, "  
                lgStrSQL = lgStrSQL & " HFA030T.KIND_EDU_CNT, "  
                lgStrSQL = lgStrSQL & " HFA030T.UNIV_EDU_CNT, "  
                lgStrSQL = lgStrSQL & " (HFA030T.HOUSE_FUND + HFA030T.LONG_HOUSE_LOAN_AMT)  HOUSE_AMT, "
                lgStrSQL = lgStrSQL & " HFA030T.FORE_EDU_AMT, "   '2003
                lgStrSQL = lgStrSQL & " HFA030T.RETIRE_PENSION, HFA050T.RETIRE_PENSION , "   '2005
                lgStrSQL = lgStrSQL & " (HFA030T.INDIV_ANU + HFA030T.INDIV_ANU2) hfa030t_indiv_anu_amt, "  
                lgStrSQL = lgStrSQL & " HFA030T.National_pension_amt hfa030t_National_pension_amt, "  
                lgStrSQL = lgStrSQL & " HFA030T.SAVE_TAX_SUB, "  
                lgStrSQL = lgStrSQL & " HFA030T.HOUSE_REPAY, "  
                lgStrSQL = lgStrSQL & " HFA030T.STOCK_SAVE, "  
                lgStrSQL = lgStrSQL & " HFA030T.FORE_PAY, "  
                lgStrSQL = lgStrSQL & " HFA030T.INCOME_REDU,"   
                lgStrSQL = lgStrSQL & " HFA030T.TAXES_REDU, "  
                lgStrSQL = lgStrSQL & " HFA030T.TECH_SUB_AMT, "  
                lgStrSQL = lgStrSQL & " HFA030T.INVEST_SUB_AMT, "  
                lgStrSQL = lgStrSQL & " HFA030T.VENTURE_SUB_AMT, "
                lgStrSQL = lgStrSQL & " HFA030T.OTHER_INSUR hfa030t_other_insur_amt, HFA030T.disabled_sub_amt hfa030t_disabled_insur_amt, HFA030T.EMP_INSUR hfa030t_emp_insur_amt, "  
                lgStrSQL = lgStrSQL & " HFA030T.TOT_MED hfa030t_tot_med_amt, HFA030T.SPECI_MED hfa030t_speci_med_amt,"  
                lgStrSQL = lgStrSQL & " HFA030T.PER_EDU hfa030t_per_edu_amt, HFA030T.LEGAL_CONTR hfa030t_legal_contr_amt, "  
                lgStrSQL = lgStrSQL & " HFA030T.APP_CONTR hfa030t_app_contr_amt,HFA030T.PRIV_CONTR hfa030t_priv_contr_amt, "  
                lgStrSQL = lgStrSQL & " HFA030T.HOUSE_FUND hfa030t_house_fund_amt, HFA030T.LONG_HOUSE_LOAN_AMT hfa030t_long_house_loan_amt, HFA030T.LONG_HOUSE_LOAN_AMT1 hfa030t_long_house_loan_amt1,"   
                lgStrSQL = lgStrSQL & " HFA030T.AFTER_BONUS_AMT after_bonus_amt, HDF020T.SUPP_OLD_CNT hdf020t_supp_old_cnt, "  
                lgStrSQL = lgStrSQL & " HDF020T.SUPP_YOUNG_CNT hdf020t_supp_young_cnt, "
                lgStrSQL = lgStrSQL & " HFA030T.Disabled_edu_amt hfa030t_disabled_edu_amt, "  '2002년 추가 
                lgStrSQL = lgStrSQL & " HFA030T.TaxLaw_contr_amt hfa030t_TaxLaw_contr_amt, "  '2002년 추가 
                lgStrSQL = lgStrSQL & " HFA030T.TaxLaw_contr_amt2 hfa030t_TaxLaw_contr_amt2, "  '2005    
                lgStrSQL = lgStrSQL & " HFA030T.retire_pension hfa030t_retire_pension , "  '2005
                lgStrSQL = lgStrSQL & " HFA050T.retire_pension hfa050t_retire_pension , "  '2005
                lgStrSQL = lgStrSQL & " T.PAY_TOT_AMT, "
                lgStrSQL = lgStrSQL & " T.BONUS_TOT_AMT, "
                lgStrSQL = lgStrSQL & " T.MED_INSUR_AMT, "
                lgStrSQL = lgStrSQL & " T.AFTER_BONUS_AMT   old_after_bonus_amt, "

                lgStrSQL = lgStrSQL & " (HFA050T.NEW_PAY_TOT + HFA050T.NEW_BONUS_TOT + HFA030T.AFTER_BONUS_AMT) a_amt ,"
                lgStrSQL = lgStrSQL & " (T.PAY_TOT_AMT + T.BONUS_TOT_AMT + T.AFTER_BONUS_AMT ) b_amt ,"
                lgStrSQL = lgStrSQL & " (HFA050T.PER_SUB + HFA050T.SPOUSE_SUB + HFA050T.SUPP_SUB + HFA050T.OLD_SUB + HFA050T.PARIA_SUB + HFA050T.LADY_SUB + "  
                lgStrSQL = lgStrSQL & " HFA050T.CHL_REAR_SUB + HFA050T.SMALL_SUB + HFA050T.YOUNG_SUB) d_amt ,"
                lgStrSQL = lgStrSQL & " (HFA050T.PER_SUB + HFA050T.SPOUSE_SUB + HFA050T.SUPP_SUB + HFA050T.OLD_SUB + HFA050T.PARIA_SUB + HFA050T.LADY_SUB + "  
                lgStrSQL = lgStrSQL & " HFA050T.CHL_REAR_SUB + HFA050T.SMALL_SUB + HFA050T.STD_SUB + HFA050T.INDIV_ANU + HFA050T.INDIV_ANU2 + HFA050T.National_pension_sub_amt + HFA050T.TECH_SUB + HFA050T.INVEST_SUB_SUM + "
                lgStrSQL = lgStrSQL & " HFA050T.CARD_SUB_SUM + HFA050t.OUR_STOCK_AMT) sum_amt ,"
                lgStrSQL = lgStrSQL & " (HFA030T.FAM_EDU + HFA030T.UNIV_EDU + HFA030T.KIND_EDU ) edu_sum_amt ,"
                lgStrSQL = lgStrSQL & " (HFA050T.DEC_INCOME_TAX + HFA050T.DEC_FARM_TAX + HFA050T.DEC_RES_TAX ) dec_amt ,"
                lgStrSQL = lgStrSQL & " (HFA050T.NEW_INCOME_TAX + HFA050T.NEW_FARM_TAX + HFA050T.NEW_RES_TAX ) new_amt ,"
                lgStrSQL = lgStrSQL & " (HFA050T.OLD_INCOME_TAX + HFA050T.OLD_FARM_TAX + HFA050T.OLD_RES_TAX ) old_amt ,"
                lgStrSQL = lgStrSQL & " (HFA050T.INCOME_TAX + HFA050T.FARM_TAX + HFA050T.RES_TAX ) f_amt ,"
                
				lgStrSQL = lgStrSQL & " HFA030T.POLI_CONTRA_AMT1 hfa030t_poli_contr_amt, "  '2004 기부금 
				lgStrSQL = lgStrSQL & " HFA030T.OURSTOCK_CONTRA_AMT hfa030t_ourstock_contr_amt, "  '2004 기부금                
				               
				lgStrSQL = lgStrSQL & " HFA030T.CEREMONY_AMT HFA030T_CEREMONY_AMT, HFA050T.CEREMONY_AMT HFA050T_CEREMONY_AMT, "  '2004 결혼장례비 
				lgStrSQL = lgStrSQL & " HFA050T.YOUNG_SUB hfa050t_young_sub_amt "  '2007 다자녀추가공제 

                lgStrSQL = lgStrSQL & " FROM HAA010T, "  
                lgStrSQL = lgStrSQL & " HFA050T, "  
                lgStrSQL = lgStrSQL & " HFA030T, "  
                lgStrSQL = lgStrSQL & " HDF020T, "  
                lgStrSQL = lgStrSQL & " (SELECT EMP_NO, SUM(A_PAY_TOT_AMT) PAY_TOT_AMT, SUM(A_BONUS_TOT_AMT) BONUS_TOT_AMT, SUM(A_MED_INSUR) MED_INSUR_AMT," 
                lgStrSQL = lgStrSQL & " SUM(A_AFTER_BONUS_AMT) AFTER_BONUS_AMT FROM HFA040T "
                lgStrSQL = lgStrSQL & " WHERE YEAR_YY = " & FilterVar(lgKeyStream(0),"'%'", "S")
                lgStrSQL = lgStrSQL & " GROUP BY EMP_NO) AS T "
                lgStrSQL = lgStrSQL & " WHERE HAA010T.emp_no = HFA050T.EMP_NO "
                lgStrSQL = lgStrSQL & " AND HFA050T.YEAR_YY = " & FilterVar(lgKeyStream(0),"'%'", "S")
                lgStrSQL = lgStrSQL & " AND HFA050T.YEAR_YY = HFA030T.YY "
                lgStrSQL = lgStrSQL & " AND HFA050T.EMP_NO = HFA030T.EMP_NO "
                lgStrSQL = lgStrSQL & " AND HFA050T.EMP_NO *= T.EMP_NO "
                lgStrSQL = lgStrSQL & " AND HAA010T.EMP_NO = HDF020T.EMP_NO "
                lgStrSQL = lgStrSQL & " AND HAA010T.emp_no = " & FilterVar(lgKeyStream(1),"'%'", "S")                       
                lgStrSQL = lgStrSQL & " AND HFA050T.internal_cd LIKE   " & FilterVar(lgKeyStream(2) & "%", "''", "S") & ""
'Response.Write lgStrSQL
'Response.End  
 
       End Select
 
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
      '       Parent.DBQueryOk        
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
      '       Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
