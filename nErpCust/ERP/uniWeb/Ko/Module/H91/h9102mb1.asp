<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%

    Call HideStatusWnd                                                               '☜: Hide Processing message

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
 
	DIM lgGetSvrDateTime
	
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    lgGetSvrDateTime = GetSvrDateTime
    
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
    Dim iKey1

 '   On Error Resume Next                                                             '☜: Protect system from crashing
 '   Err.Clear                                                                        '☜: Clear Error status


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
       With Parent	

            .Frm1.txtNon_tax_bas_amt.text          = "<%=UNINumClientFormat(lgObjRs("NON_TAX_BAS"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtNon_tax_limit_amt.text        = "<%=UNINumClientFormat(lgObjRs("NON_TAX_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtNon_dinn_amt.text             = "<%=UNINumClientFormat(lgObjRs("NON_DINN_AMT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtOversea_labor_limit_amt.text  = "<%=UNINumClientFormat(lgObjRs("OVERSEA_LABOR_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtRnd_nontax_limit.text		   = "<%=UNINumClientFormat(lgObjRs("Rnd_nontax_limit"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtIncome_sub_bas_amt.text       = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_BAS"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtIncome_sub_bas0_amt.text      = .frm1.txtIncome_sub_bas_amt.text '추가 
            .frm1.txtIncome_calcu_bas_amt.text     = "<%=UNINumClientFormat(lgObjRs("INCOME_CALCU_BAS"), ggAmtOfMoney.DecPoint,0)%>"      
            .frm1.txtIncome_calcu_bas1_amt.text    = "<%=UNINumClientFormat(lgObjRs("INCOME_CALCU_BAS1"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.txtIncome_calcu_bas2_amt.text    = "<%=UNINumClientFormat(lgObjRs("INCOME_CALCU_BAS2"), ggAmtOfMoney.DecPoint,0)%>"        '추가 
            .frm1.txtIncome_sub_rate1.text         = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_RATE1"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtIncome_sub_rate2.text         = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_RATE2"), ggAmtOfMoney.DecPoint,0)%>"          
            .Frm1.txtIncome_sub_rate3.text         = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_RATE3"), ggAmtOfMoney.DecPoint,0)%>"   
            .Frm1.txtIncome_sub_rate4.text         = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_RATE4"), ggAmtOfMoney.DecPoint,0)%>"      
            .Frm1.txtIncome_sub_rate5.text         = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_RATE5"), ggAmtOfMoney.DecPoint,0)%>"   
            .frm1.txtIncome_sub_bas1_amt.text      = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_BAS1"), ggAmtOfMoney.DecPoint,0)%>"        
            .Frm1.txtIncome_sub_bas2_amt.text      = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_BAS2"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtIncome_sub_bas3_amt.text      = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_BAS3"), ggAmtOfMoney.DecPoint,0)%>" '추가 
            .frm1.txtIncome_sub_limit_amt.text     = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"        
            .Frm1.txtNational_pension_sub_rate.text = "<%=ConvSPChars(lgObjRs("NATIONAL_PENSION_SUB_RATE"))%>"
            .Frm1.txtPer_sub_amt.text              = "<%=UNINumClientFormat(lgObjRs("PER_SUB"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtSpouse_sub_amt.text           = "<%=UNINumClientFormat(lgObjRs("SPOUSE_SUB"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtFam_sub_amt.text              = "<%=UNINumClientFormat(lgObjRs("FAM_SUB"), ggAmtOfMoney.DecPoint,0)%>"          
            .frm1.txtOld_sub_amt1.text              = "<%=UNINumClientFormat(lgObjRs("OLD_SUB"), ggAmtOfMoney.DecPoint,0)%>"     '2004 경로우대공제(65세이상)   
            .frm1.txtOld_sub_amt2.text              = "<%=UNINumClientFormat(lgObjRs("OLD_SUB2"), ggAmtOfMoney.DecPoint,0)%>"    '2004 경로우대공제(70세이상)         
            .frm1.txtParia_sub_amt.text            = "<%=UNINumClientFormat(lgObjRs("PARIA_SUB"), ggAmtOfMoney.DecPoint,0)%>"      
            .frm1.txtChl_rear_sub_amt.text         = "<%=UNINumClientFormat(lgObjRs("CHL_REAR_SUB"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtLady_sub_amt.text             = "<%=UNINumClientFormat(lgObjRs("LADY_SUB"), ggAmtOfMoney.DecPoint,0)%>"        
            .Frm1.txtSmall_sub1_amt.text           = "<%=UNINumClientFormat(lgObjRs("SMALL_SUB1"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtSmall_sub2_amt.text           = "<%=UNINumClientFormat(lgObjRs("SMALL_SUB2"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtIncome_tax_sub_bas_amt.text   = "<%=UNINumClientFormat(lgObjRs("INCOME_TAX_SUB_BAS"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtIncome_tax_rate1.text         = "<%=ConvSPChars(lgObjRs("INCOME_TAX_RATE1"))%>"
            .frm1.txtIncome_tax_bas_amt.text       = "<%=UNINumClientFormat(lgObjRs("INCOME_TAX_BAS"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtIncome_tax_sub_bas1_amt.text  = "<%=UNINumClientFormat(lgObjRs("INCOME_TAX_SUB_BAS"), ggAmtOfMoney.DecPoint,0)%>"          
            .frm1.txtIncome_tax_rate2.text         = "<%=ConvSPChars(lgObjRs("INCOME_TAX_RATE2"))%>"        
            .frm1.txtIncome_tax_limit_amt.text     = "<%=UNINumClientFormat(lgObjRs("INCOME_TAX_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"      
            .frm1.txtOther_insur_limit_amt.text    = "<%=UNINumClientFormat(lgObjRs("OTHER_INSUR_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtDisabled_insur_limit_amt.text = "<%=UNINumClientFormat(lgObjRs("DISABLED_SUB_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtMed_sub_bas_amt.text          = "<%=UNINumClientFormat(lgObjRs("MED_SUB_BAS"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtMed_sub_limit_amt.text        = "<%=UNINumClientFormat(lgObjRs("MED_SUB_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtParia_med_rate.text           = "<%=ConvSPChars(lgObjRs("PARIA_MED_RATE"))%>"
            .Frm1.txtPer_edu_sub.text              = "<%=ConvSPChars(lgObjRs("PER_EDU_SUB"))%>"
            .Frm1.txtFam_edu_sub_amt.text          = "<%=UNINumClientFormat(lgObjRs("FAM_EDU_SUB"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtKind_edu_limit_amt.text       = "<%=UNINumClientFormat(lgObjRs("KIND_EDU_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtUniv_edu_limit_amt.text       = "<%=UNINumClientFormat(lgObjRs("UNIV_EDU_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"          
            .frm1.txtLegal_contr_rate.text         = "<%=ConvSPChars(lgObjRs("LEGAL_CONTR_RATE"))%>"        
            .frm1.txtApp_contr_rate.text           = "<%=ConvSPChars(lgObjRs("APP_CONTR_RATE"))%>"      

            .frm1.txtOurStock_contr_rate.text      = "<%=ConvSPChars(lgObjRs("OURSTOCK_CONTR_RATE"))%>"      '2004 우리사주기부금공제요율 
            
            .frm1.txtHouse_fund_rate.text          = "<%=ConvSPChars(lgObjRs("HOUSE_FUND_RATE"))%>"         
            .frm1.txtHouse_fund_limit_amt.text     = "<%=UNINumClientFormat(lgObjRs("HOUSE_FUND_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtLong_house_loan_limit.text    = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"  '장기주택저당차입금의이자한도액(2003)
            .frm1.txtLong_house_loan_limit1.text    = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_LIMIT1"), ggAmtOfMoney.DecPoint,0)%>"  '2004 장기주택저당차입금의이자한도액 (상환기간 15년이상)

            .frm1.txtFore_edu_rate.text            = "<%=ConvSPChars(lgObjRs("FORE_EDU_RATE"))%>"  '외국인근로자교육비공제율(2003)            
                                
            .Frm1.txtStd_sub_amt.text              = "<%=UNINumClientFormat(lgObjRs("STD_SUB"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtIncome_card_rate1.text        = "<%=ConvSPChars(lgObjRs("INCOME_CARD_RATE1"))%>"
            .Frm1.txtIncome_card_rate2.text        = "<%=ConvSPChars(lgObjRs("INCOME_CARD_RATE2"))%>"
            .Frm1.txtIncome_card2_rate2.text        = "<%=ConvSPChars(lgObjRs("INCOME_CARD2_RATE2"))%>" '직불카드(2003)
            
            .Frm1.txtCeremony_amt.text        = "<%=UNINumClientFormat(lgObjRs("CEREMONY_AMT"), ggAmtOfMoney.DecPoint,0)%>" '2004  결혼장례비공제액 
            .Frm1.txtForeign_separate_tax_rate.text		 = "<%=UNINumClientFormat(lgObjRs("FOREIGN_SEPARATE_TAX_RATE"), ggAmtOfMoney.DecPoint,0)%>" '2004 외국인근로자분리과세금액 
                                    
            .Frm1.txtIncome_card_limit_amt.text    = "<%=UNINumClientFormat(lgObjRs("INCOME_CARD_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtInvest_rate.text              = "<%=ConvSPChars(lgObjRs("INVEST_RATE"))%>"         
            .frm1.txtIndiv_anu_rate.text           = "<%=ConvSPChars(lgObjRs("INDIV_ANU_RATE"))%>"          
            .frm1.txtIndiv_anu_limit_amt.text      = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.txtIndiv_anu2_rate.text          = "<%=ConvSPChars(lgObjRs("INDIV_ANU2_RATE"))%>"          
            .frm1.txtIndiv_anu2_limit_amt.text     = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU2_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.txtHouse_repay_rate.text         = "<%=ConvSPChars(lgObjRs("HOUSE_REPAY_RATE"))%>"      
            
            
            .frm1.txtLong_Stock_save_rate.text     = "<%=ConvSPChars(lgObjRs("LONG_STOCK_SAVE_RATE"))%>"         
            .frm1.txtLong_Stock_save_limit_amt.text = "<%=UNINumClientFormat(lgObjRs("LONG_STOCK_SAVE_LIMIT"), ggAmtOfMoney.DecPoint,0)%>"        
            .frm1.txtFarm_tax.text                 = "<%=ConvSPChars(lgObjRs("FARM_TAX"))%>"         
            .frm1.txtRes_tax.text                  = "<%=ConvSPChars(lgObjRs("RES_TAX"))%>"        
            
            '2002년 추가항목 

            .frm1.txtLong_Stock_save_rate1.text       = "<%=UNINumClientFormat(lgObjRs("Long_Stock_save_rate1"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtDisabled_edu_limit_amt.text      = "<%=UNINumClientFormat(lgObjRs("Disabled_edu_limit_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtOur_Stock_limit_amt.text         = "<%=UNINumClientFormat(lgObjRs("Our_Stock_limit_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtTaxLaw_contr_rate.text           = "<%=ConvSPChars(lgObjRs("TaxLaw_contr_rate"))%>"
            .frm1.txtInvest_rate2.text                = "<%=UNINumClientFormat(lgObjRs("Invest_rate2"), ggAmtOfMoney.DecPoint,0)%>" 

			'2003 급여계산시 특별공제부분(간이세액표기준)
            .Frm1.txtsub_fam1.Text                 = "<%=ConvSPChars(lgObjRs("sub_fam1"))%>"
            .Frm1.txtsub_fam1_amt.Text             = "<%=UNINumClientFormat(lgObjRs("sub_fam1_amt"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtsub_fam2.Text                 = "<%=ConvSPChars(lgObjRs("sub_fam2"))%>"
            .Frm1.txtsub_fam2_amt.Text             = "<%=UNINumClientFormat(lgObjRs("sub_fam2_amt"), ggAmtOfMoney.DecPoint, 0)%>"
            if "<%=ConvSPChars(lgObjRs("sub_fam_flag"))%>" = "Y" then
                .Frm1.txtsub_fam_flag1.checked = true
            else
                .Frm1.txtsub_fam_flag2.checked = true
            end if
            .Frm1.txtStd_sub_amt2.text              = "<%=UNINumClientFormat(lgObjRs("STD_SUB"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtMedPrint.text              = "<%=UNINumClientFormat(lgObjRs("MED_SUB"), ggAmtOfMoney.DecPoint,0)%>"
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

    dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
              
    Call CommonQueryRs("count(*) ","HFA020T"," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If Replace(lgF0, Chr(11), "") = 0 Then
       Call SubBizSaveSingleCreate()  
    Else
       Call SubBizSaveSingleUpdate()
    End If

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HFA020T("
    lgStrSQL = lgStrSQL & " COMP_CD, "
    lgStrSQL = lgStrSQL & " APP_CONTR_LIMIT, "  
    lgStrSQL = lgStrSQL & " PRIV_CONTR_LIMIT, " 
    lgStrSQL = lgStrSQL & " NON_TAX_BAS, "  
    lgStrSQL = lgStrSQL & " NON_TAX_LIMIT, "    
    lgStrSQL = lgStrSQL & " NON_DINN_AMT, " 
    lgStrSQL = lgStrSQL & " OVERSEA_LABOR_LIMIT, " 
    lgStrSQL = lgStrSQL & " Rnd_nontax_limit, "       ' 20040302 by lsn  
    lgStrSQL = lgStrSQL & " INCOME_SUB_BAS, "  
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE1, "  
    lgStrSQL = lgStrSQL & " INCOME_SUB_BAS1, "    
    lgStrSQL = lgStrSQL & " INCOME_CALCU_BAS, "    
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE2, "   
    lgStrSQL = lgStrSQL & " INCOME_CALCU_BAS1, "   
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE3, "  
    lgStrSQL = lgStrSQL & " INCOME_SUB_LIMIT, "    
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE4, " 
    lgStrSQL = lgStrSQL & " INCOME_SUB_BAS2, "   
    lgStrSQL = lgStrSQL & " INCOME_CALCU_BAS2, "   
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE5, "   
    lgStrSQL = lgStrSQL & " INCOME_SUB_BAS3, "    
    lgStrSQL = lgStrSQL & " NATIONAL_PENSION_SUB_RATE, "  
    lgStrSQL = lgStrSQL & " PER_SUB, "  
    lgStrSQL = lgStrSQL & " SPOUSE_SUB, "    
    lgStrSQL = lgStrSQL & " FAM_SUB, "    
    lgStrSQL = lgStrSQL & " OLD_SUB, " 
    lgStrSQL = lgStrSQL & " OLD_SUB2, "        
    lgStrSQL = lgStrSQL & " PARIA_SUB, "   
    lgStrSQL = lgStrSQL & " CHL_REAR_SUB, "  
    lgStrSQL = lgStrSQL & " LADY_SUB, "    
    lgStrSQL = lgStrSQL & " SMALL_SUB1, " 
    lgStrSQL = lgStrSQL & " SMALL_SUB2, "    
    lgStrSQL = lgStrSQL & " INCOME_TAX_SUB_BAS, "  
    lgStrSQL = lgStrSQL & " INCOME_TAX_RATE1, "  
    lgStrSQL = lgStrSQL & " INCOME_TAX_BAS, "    
    lgStrSQL = lgStrSQL & " INCOME_TAX_RATE2, "    
    lgStrSQL = lgStrSQL & " INCOME_TAX_LIMIT, "   
    lgStrSQL = lgStrSQL & " OTHER_INSUR_LIMIT, "   
    lgStrSQL = lgStrSQL & " DISABLED_SUB_LIMIT, "   
    lgStrSQL = lgStrSQL & " MED_SUB_BAS, "  
    lgStrSQL = lgStrSQL & " MED_SUB_LIMIT, "    
    lgStrSQL = lgStrSQL & " PARIA_MED_RATE, " 
    lgStrSQL = lgStrSQL & " PER_EDU_SUB, "    
    lgStrSQL = lgStrSQL & " FAM_EDU_SUB, "  
    lgStrSQL = lgStrSQL & " KIND_EDU_LIMIT, "  
    lgStrSQL = lgStrSQL & " UNIV_EDU_LIMIT, "    
    lgStrSQL = lgStrSQL & " LEGAL_CONTR_RATE, "    
    lgStrSQL = lgStrSQL & " APP_CONTR_RATE, "   
    
    lgStrSQL = lgStrSQL & " OURSTOCK_CONTR_RATE, "  '2004 우리사주기부금공제요율 
    lgStrSQL = lgStrSQL & " OURSTOCK_CONTR, "       '2004 우리사주기부금   
    lgStrSQL = lgStrSQL & " HOUSE_FUND_RATE, "   
    lgStrSQL = lgStrSQL & " HOUSE_FUND_LIMIT, "
    lgStrSQL = lgStrSQL & " LONG_HOUSE_LOAN_LIMIT, "    '장기주택저당차입금의이자한도액(2003)
    lgStrSQL = lgStrSQL & " LONG_HOUSE_LOAN_LIMIT1, "    '2004 장기주택저당차입금의이자한도액 (상환기간 15년이상)   
    lgStrSQL = lgStrSQL & " FORE_EDU_LIMIT, "    '외국인근로자교육비공제율(2003)    

    lgStrSQL = lgStrSQL & " STD_SUB, "    
    lgStrSQL = lgStrSQL & " INCOME_CARD_RATE1, " 
    lgStrSQL = lgStrSQL & " INCOME_CARD_RATE2, "
    lgStrSQL = lgStrSQL & " INCOME_CARD2_RATE2, "    '직불카드(2003)

    lgStrSQL = lgStrSQL & " CEREMONY_AMT, "
    lgStrSQL = lgStrSQL & " FOREIGN_SEPARATE_TAX_RATE, "
        
    lgStrSQL = lgStrSQL & " INCOME_CARD_LIMIT, "  
    lgStrSQL = lgStrSQL & " INVEST_RATE, "  
    lgStrSQL = lgStrSQL & " INDIV_ANU_RATE, "    
    lgStrSQL = lgStrSQL & " INDIV_ANU_LIMIT, "
    lgStrSQL = lgStrSQL & " INDIV_ANU2_RATE, "    
    lgStrSQL = lgStrSQL & " INDIV_ANU2_LIMIT, "
    lgStrSQL = lgStrSQL & " HOUSE_REPAY_RATE, "   
'    lgStrSQL = lgStrSQL & " STOCK_SAVE_RATE, "   
'    lgStrSQL = lgStrSQL & " STOCK_SAVE_LIMIT, "  
    lgStrSQL = lgStrSQL & " LONG_STOCK_SAVE_RATE, "   
    lgStrSQL = lgStrSQL & " LONG_STOCK_SAVE_LIMIT, "  
    lgStrSQL = lgStrSQL & " FARM_TAX, "    
    lgStrSQL = lgStrSQL & " RES_TAX, " 
    
    '2002년 추가 
    lgStrSQL = lgStrSQL & " Long_Stock_save_rate1, "
    lgStrSQL = lgStrSQL & " Disabled_edu_limit_amt, "
    lgStrSQL = lgStrSQL & " Our_Stock_limit_amt, "
    lgStrSQL = lgStrSQL & " TaxLaw_contr_rate, "
    lgStrSQL = lgStrSQL & " Invest_rate2, "

	'2003 급여계산시 특별공제부분(간이세액표기준)
	lgStrSQL = lgStrSQL & " sub_fam1     , "
	lgStrSQL = lgStrSQL & " sub_fam1_amt , "
	lgStrSQL = lgStrSQL & " sub_fam2     , "
	lgStrSQL = lgStrSQL & " sub_fam2_amt , "
	lgStrSQL = lgStrSQL & " sub_fam_flag , "
	lgStrSQL = lgStrSQL & " MED_SUB  , "	'2004
	
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO, "    
    lgStrSQL = lgStrSQL & " ISRT_DT, "  
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO, "  
    lgStrSQL = lgStrSQL & " UPDT_DT) "    
    lgStrSQL = lgStrSQL & " VALUES (" 
    lgStrSQL = lgStrSQL & " " & FilterVar("1", "''", "S") & "  ," 
    lgStrSQL = lgStrSQL & "   5 ,"
    lgStrSQL = lgStrSQL & " 100 ,"
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtNon_tax_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtNon_tax_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtNon_dinn_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOversea_labor_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtRnd_nontax_limit"),0) & ","    ' 20040302 by lsn    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_rate1"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_bas1_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_calcu_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_rate2"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_calcu_bas1_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_rate3"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_rate4"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_bas2_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_calcu_bas2_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_rate5"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_sub_bas3_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtNational_pension_sub_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtPer_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtSpouse_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFam_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOld_sub_amt1"),0) & ","	'2004 경로우대공제(65세이상)
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOld_sub_amt2"),0) & ","	'2004 경로우대공제(70세이상)	
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtParia_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtChl_rear_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLady_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtSmall_sub1_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtSmall_sub2_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_tax_sub_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_tax_rate1"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_tax_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_tax_rate2"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_tax_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOther_insur_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtDisabled_insur_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtMed_sub_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtMed_sub_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtParia_med_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtPer_edu_sub"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFam_edu_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtKind_edu_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtUniv_edu_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLegal_contr_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtApp_contr_rate"),0) & ","
    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("OURSTOCK_CONTR_RATE"),0) & ","	'2004 우리사주기부금공제요율 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("OURSTOCK_CONTR"),0) & ","			'2004 우리사주기부금 
    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtHouse_fund_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtHouse_fund_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLong_house_loan_limit"),0) & "," '장기주택저당차입금의이자한도액(2003)
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLong_house_loan_limit1"),0) & "," '2004 장기주택저당차입금의이자한도액 (상환기간 15년이상)
        
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFore_edu_rate"),0) & "," '외국인근로자교육비공제율(2003)

    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtStd_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_card_rate1"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_card_rate2"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_card2_rate2"),0) & ","    '직불카드(2003)

    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtCeremony_amt"),0) & ","	'2004  결혼장례비공제액 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtForeign_separate_tax_rate"),0) & ","	'2004 외국인근로자분리과세금액 

    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_card_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtInvest_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIndiv_anu_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIndiv_anu_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIndiv_anu2_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIndiv_anu2_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtHouse_repay_rate"),0) & ","
    
    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLong_Stock_save_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLong_Stock_save_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFarm_tax"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtRes_tax"),0) & ","
    
    '2002년 추가항목 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLong_Stock_save_rate1"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtDisabled_edu_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOur_Stock_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtTaxLaw_contr_rate"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtInvest_rate2"),0) & ","

	'2003 급여계산시 특별공제부분(간이세액표기준)
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtsub_fam1"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtsub_fam1_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtsub_fam2"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtsub_fam2_amt"),0) & ","            
    lgStrSQL = lgStrSQL & FilterVar(Request("txtsub_fam_flag"), "''", "S") & ","                           
    lgStrSQL = lgStrSQL & FilterVar(Request("txtMedPrint"), "''", "S") & "," 
                              
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HFA020T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " NON_TAX_BAS = " & UNIConvNum(Request("txtNon_tax_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & " NON_TAX_LIMIT = " & UNIConvNum(Request("txtNon_tax_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " NON_DINN_AMT = " & UNIConvNum(Request("txtNon_dinn_amt"),0) & ","
    lgStrSQL = lgStrSQL & " OVERSEA_LABOR_LIMIT = " & UNIConvNum(Request("txtOversea_labor_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " Rnd_nontax_limit = " & UNIConvNum(Request("txtRnd_nontax_limit"),0) & ","    ' 20040302 by lsn  
      
    lgStrSQL = lgStrSQL & " INCOME_SUB_BAS = " & UNIConvNum(Request("txtIncome_sub_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE1 = " & UNIConvNum(Request("txtIncome_sub_rate1"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_BAS1 = " & UNIConvNum(Request("txtIncome_sub_bas1_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_CALCU_BAS = " & UNIConvNum(Request("txtIncome_calcu_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE2 = " & UNIConvNum(Request("txtIncome_sub_rate2"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_CALCU_BAS1 = " & UNIConvNum(Request("txtIncome_calcu_bas1_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE3 = " & UNIConvNum(Request("txtIncome_sub_rate3"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_LIMIT = " & UNIConvNum(Request("txtIncome_sub_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE4 = " & UNIConvNum(Request("txtIncome_sub_rate4"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_BAS2 = " & UNIConvNum(Request("txtIncome_sub_bas2_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_CALCU_BAS2 = " & UNIConvNum(Request("txtIncome_calcu_bas2_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_RATE5 = " & UNIConvNum(Request("txtIncome_sub_rate5"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_SUB_BAS3 = " & UNIConvNum(Request("txtIncome_sub_bas3_amt"),0) & ","
    
    lgStrSQL = lgStrSQL & " NATIONAL_PENSION_SUB_RATE = " & UNIConvNum(Request("txtNational_pension_sub_rate"),0) & ","
    lgStrSQL = lgStrSQL & " PER_SUB = " & UNIConvNum(Request("txtPer_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & " SPOUSE_SUB = " & UNIConvNum(Request("txtSpouse_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & " FAM_SUB = " & UNIConvNum(Request("txtFam_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & " OLD_SUB = " & UNIConvNum(Request("txtOld_sub_amt1"),0) & ","	'2004 경로우대공제(65세이상)
    lgStrSQL = lgStrSQL & " OLD_SUB2 = " & UNIConvNum(Request("txtOld_sub_amt2"),0) & ","	'2004 경로우대공제(70세이상)  
    lgStrSQL = lgStrSQL & " PARIA_SUB = " & UNIConvNum(Request("txtParia_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & " CHL_REAR_SUB = " & UNIConvNum(Request("txtChl_rear_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & " LADY_SUB = " & UNIConvNum(Request("txtLady_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & " SMALL_SUB1 = " & UNIConvNum(Request("txtSmall_sub1_amt"),0) & ","
    lgStrSQL = lgStrSQL & " SMALL_SUB2 = " & UNIConvNum(Request("txtSmall_sub2_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_TAX_SUB_BAS = " & UNIConvNum(Request("txtIncome_tax_sub_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_TAX_RATE1 = " & UNIConvNum(Request("txtIncome_tax_rate1"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_TAX_BAS = " & UNIConvNum(Request("txtIncome_tax_bas_amt"),0) & ","
    'lgStrSQL = lgStrSQL & " INCOME_TAX_SUB_BAS = " & UNIConvNum(Request("txtIncome_tax_sub_bas1_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_TAX_RATE2 = " & UNIConvNum(Request("txtIncome_tax_rate2"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_TAX_LIMIT = " & UNIConvNum(Request("txtIncome_tax_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " OTHER_INSUR_LIMIT = " & UNIConvNum(Request("txtOther_insur_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " DISABLED_SUB_LIMIT = " & UNIConvNum(Request("txtDisabled_insur_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " MED_SUB_BAS = " & UNIConvNum(Request("txtMed_sub_bas_amt"),0) & ","
    lgStrSQL = lgStrSQL & " MED_SUB_LIMIT = " & UNIConvNum(Request("txtMed_sub_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " PARIA_MED_RATE = " & UNIConvNum(Request("txtParia_med_rate"),0) & ","
    lgStrSQL = lgStrSQL & " PER_EDU_SUB = " & UNIConvNum(Request("txtPer_edu_sub"),0) & ","
    lgStrSQL = lgStrSQL & " FAM_EDU_SUB = " & UNIConvNum(Request("txtFam_edu_sub_amt"),0) & ","
    lgStrSQL = lgStrSQL & " KIND_EDU_LIMIT = " & UNIConvNum(Request("txtKind_edu_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " UNIV_EDU_LIMIT = " & UNIConvNum(Request("txtUniv_edu_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " LEGAL_CONTR_RATE = " & UNIConvNum(Request("txtLegal_contr_rate"),0) & ","
    lgStrSQL = lgStrSQL & " APP_CONTR_RATE = " & UNIConvNum(Request("txtApp_contr_rate"),0) & ","
    
    lgStrSQL = lgStrSQL & " OURSTOCK_CONTR_RATE = " & UNIConvNum(Request("txtOurStock_contr_rate"),0) & ","  '2004 
    
    lgStrSQL = lgStrSQL & " HOUSE_FUND_RATE = " & UNIConvNum(Request("txtHouse_fund_rate"),0) & ","
    lgStrSQL = lgStrSQL & " HOUSE_FUND_LIMIT = " & UNIConvNum(Request("txtHouse_fund_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " LONG_HOUSE_LOAN_LIMIT = " & UNIConvNum(Request("txtLong_house_loan_limit"),0) & "," '장기주택저당차입금의이자한도액(2003)   
    lgStrSQL = lgStrSQL & " LONG_HOUSE_LOAN_LIMIT1 = " & UNIConvNum(Request("txtLong_house_loan_limit1"),0) & ","  '2004 장기주택저당차입금의이자한도액 (상환기간 15년이상)   
    lgStrSQL = lgStrSQL & " FORE_EDU_RATE = " & UNIConvNum(Request("txtFore_edu_rate"),0) & "," '외국인근로자교육비공제율(2003)
    
    lgStrSQL = lgStrSQL & " STD_SUB = " & UNIConvNum(Request("txtStd_sub_amt2"),0) & ","    
    lgStrSQL = lgStrSQL & " INCOME_CARD_RATE1 = " & UNIConvNum(Request("txtIncome_card_rate1"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_CARD_RATE2 = " & UNIConvNum(Request("txtIncome_card_rate2"),0) & ","
    lgStrSQL = lgStrSQL & " INCOME_CARD2_RATE2 = " & UNIConvNum(Request("txtIncome_card2_rate2"),0) & "," '직불카드(2003)

    lgStrSQL = lgStrSQL & " CEREMONY_AMT = " & UNIConvNum(Request("txtCeremony_amt"),0) & ","	'2004  결혼장례비공제액 
    lgStrSQL = lgStrSQL & " FOREIGN_SEPARATE_TAX_RATE = " & UNIConvNum(Request("txtForeign_separate_tax_rate"),0) & ","	'2004 외국인근로자분리과세금액 
    
    lgStrSQL = lgStrSQL & " INCOME_CARD_LIMIT = " & UNIConvNum(Request("txtIncome_card_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INVEST_RATE = " & UNIConvNum(Request("txtInvest_rate"),0) & ","
    lgStrSQL = lgStrSQL & " INDIV_ANU_RATE = " & UNIConvNum(Request("txtIndiv_anu_rate"),0) & ","
    lgStrSQL = lgStrSQL & " INDIV_ANU_LIMIT = " & UNIConvNum(Request("txtIndiv_anu_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " INDIV_ANU2_RATE = " & UNIConvNum(Request("txtIndiv_anu2_rate"),0) & ","
    lgStrSQL = lgStrSQL & " INDIV_ANU2_LIMIT = " & UNIConvNum(Request("txtIndiv_anu2_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " HOUSE_REPAY_RATE = " & UNIConvNum(Request("txtHouse_repay_rate"),0) & ","
    lgStrSQL = lgStrSQL & " LONG_STOCK_SAVE_RATE = " & UNIConvNum(Request("txtLong_Stock_save_rate"),0) & ","
    lgStrSQL = lgStrSQL & " LONG_STOCK_SAVE_LIMIT = " & UNIConvNum(Request("txtLong_Stock_save_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " FARM_TAX = " & UNIConvNum(Request("txtFarm_tax"),0) & ","
    lgStrSQL = lgStrSQL & " RES_TAX = " & UNIConvNum(Request("txtRes_tax"),0) & ","
    
    '2002년 추가항목 
    lgStrSQL = lgStrSQL & " Long_Stock_save_rate1 = " & UNIConvNum(Request("txtLong_Stock_save_rate1"),0) & ","
    lgStrSQL = lgStrSQL & " Disabled_edu_limit_amt = " & UNIConvNum(Request("txtDisabled_edu_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " Our_Stock_limit_amt = " & UNIConvNum(Request("txtOur_Stock_limit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " TaxLaw_contr_rate = " & UNIConvNum(Request("txtTaxLaw_contr_rate"),0) & ","
    lgStrSQL = lgStrSQL & " Invest_rate2 = " & UNIConvNum(Request("txtInvest_rate2"),0) & ","

	'2003 급여계산시 특별공제부분(간이세액표기준)
    lgStrSQL = lgStrSQL & " sub_fam1     = " & UNIConvNum(Request("txtsub_fam1"),0) & ","
    lgStrSQL = lgStrSQL & " sub_fam1_amt = " & UNIConvNum(Request("txtsub_fam1_amt"),0) & "," 
    lgStrSQL = lgStrSQL & " sub_fam2     = " & UNIConvNum(Request("txtsub_fam2"),0) & ","
    lgStrSQL = lgStrSQL & " sub_fam2_amt = " & UNIConvNum(Request("txtsub_fam2_amt"),0) & ","            
    lgStrSQL = lgStrSQL & " sub_fam_flag = " & FilterVar(Request("txtsub_fam_flag"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " MED_SUB = " &  UNIConvNum(Request("txtMedPrint"),0) & ","
                    
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
    lgStrSQL = lgStrSQL & " ISRT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S") & ","                ' datetime
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S")                    ' datetime

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
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
                           lgStrSQL = "Select  COMP_CD, NON_TAX_BAS, NON_TAX_LIMIT, NON_DINN_AMT, INCOME_SUB_BAS, INCOME_SUB_RATE1, INCOME_SUB_RATE2, INCOME_CALCU_BAS2, INCOME_SUB_RATE5, INCOME_SUB_BAS3, " 
                           lgStrSQL = lgStrSQL & " INCOME_CALCU_BAS, INCOME_SUB_LIMIT, PER_SUB,  SPOUSE_SUB, FAM_SUB,  "
                           lgStrSQL = lgStrSQL & " OLD_SUB,OLD_SUB2, PARIA_SUB, LADY_SUB, CHL_REAR_SUB, SMALL_SUB1, SMALL_SUB2, OTHER_INSUR_LIMIT, DISABLED_SUB_LIMIT, "
                           lgStrSQL = lgStrSQL & " MED_SUB_BAS, MED_SUB_LIMIT,  PARIA_MED_RATE, PER_EDU_SUB, FAM_EDU_SUB,  KIND_EDU_LIMIT, "
                           lgStrSQL = lgStrSQL & " UNIV_EDU_LIMIT, LEGAL_CONTR_RATE, APP_CONTR_RATE, APP_CONTR_LIMIT, PRIV_CONTR_LIMIT,   "
                           lgStrSQL = lgStrSQL & " HOUSE_FUND_RATE, HOUSE_FUND_LIMIT, LONG_HOUSE_LOAN_LIMIT ,FORE_EDU_RATE , INDIV_ANU_RATE, INDIV_ANU_LIMIT,INDIV_ANU2_RATE, INDIV_ANU2_LIMIT, INCOME_TAX_SUB_BAS, INCOME_TAX_RATE1,   "
                           lgStrSQL = lgStrSQL & " INCOME_TAX_RATE2, INCOME_TAX_BAS,  INCOME_TAX_LIMIT,  HOUSE_REPAY_RATE,  STOCK_SAVE_RATE, LONG_STOCK_SAVE_RATE,  "
                           lgStrSQL = lgStrSQL & " FARM_TAX,  RES_TAX, ISRT_EMP_NO,  ISRT_DT,  UPDT_EMP_NO,  UPDT_DT,  STOCK_SAVE_LIMIT, LONG_STOCK_SAVE_LIMIT, STD_SUB,  "
                           lgStrSQL = lgStrSQL & " INVEST_RATE,  INCOME_SUB_RATE3,  INCOME_SUB_BAS1,  INCOME_CALCU_BAS1,  INCOME_CARD_RATE1,     "
                           lgStrSQL = lgStrSQL & " INCOME_CARD_RATE2,INCOME_CARD2_RATE2,  INCOME_CARD_LIMIT,  OVERSEA_LABOR_LIMIT,Rnd_nontax_limit ,INCOME_SUB_BAS2, INCOME_SUB_RATE4,  NATIONAL_PENSION_SUB_RATE,  " ' 20040302 by lsn    
                           
                           '2002년 추가항목 
                           lgStrSQL = lgStrSQL & " isnull(Long_Stock_save_rate1,0) Long_Stock_save_rate1, isnull(Disabled_edu_limit_amt,0) Disabled_edu_limit_amt, "
                           lgStrSQL = lgStrSQL & " isnull(Our_Stock_limit_amt,0) Our_Stock_limit_amt, TaxLaw_contr_rate, Invest_rate2," 
                           lgStrSQL = lgStrSQL & " sub_fam1, sub_fam1_amt , sub_fam2 , sub_fam2_amt , sub_fam_flag,"    
                           '2004 출산수당비과세한도액,결혼장례비공제액,외국인근로자분리과세적용여부 , 외국인근로자분리과세금액,2004 장기주택저당차입금의이자한도액 (상환기간 15년이상)
                           lgStrSQL = lgStrSQL & " CEREMONY_AMT, FOREIGN_SEPARATE_TAX_RATE,OURSTOCK_CONTR_RATE,LONG_HOUSE_LOAN_LIMIT1,MED_SUB"	
                           lgStrSQL = lgStrSQL & " From  HFA020T  "
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
        Case "SC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "SD"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "SU"
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
             Parent.DBQueryOk        
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
