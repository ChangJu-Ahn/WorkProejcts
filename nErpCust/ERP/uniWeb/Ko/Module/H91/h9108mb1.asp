<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
    Dim lgSvrDateTime
    
    call LoadBasisGlobalInf()
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")
	    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgSvrDateTime = FilterVar(GetSvrDateTime, "''", "S")
    
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

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
    Dim iKey1, iKey2, iKey3
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    iKey2 = FilterVar(lgKeyStream(1), "''", "S")
    iKey3 = FilterVar(lgKeyStream(2), "''", "S")

    ' 인적사항 및 소득공제 영역 조회 
    Call SubMakeSQLStatements("R",iKey1,iKey2,lgKeyStream(2))                                 '☜ : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists    
       If lgPrevNext = "" Then       
       ElseIf lgPrevNext = "P" Then
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
       iKey2 = FilterVar(ConvSPChars(lgObjRs("EMP_NO")), "''", "S")

%>
<Script Language=vbscript>
       With Parent	     
            .Frm1.txtEmp_no.Value           = "<%=ConvSPChars(lgObjRs("EMP_NO"))%>"            'Set condition area
			.Frm1.txtName.value             = "<%=ConvSPChars(lgObjRs("NAME"))%>"
			.Frm1.txtDept_nm.value          = "<%=ConvSPChars(lgObjRs("DEPT_NM"))%>"
            .frm1.txtRollPstn.value         = "<%=ConvSPChars(lgObjRs("ROLL_PSTN_NM"))%>"       '직위 
            .Frm1.txtPay_grd.value          = "<%=ConvSPChars(lgObjRs("Pay_Grd1_NM"))%>" & "-" & "<%=ConvSPChars(lgObjRs("PAY_GRD2"))%>" '급호 
            .frm1.txtEntr_dt.text           = "<%=UniDateClientFormat(lgObjRs("ENTR_DT"))%>"
            
            If "<%=ConvSPChars(lgObjRs("SPOUSE"))%>" = "Y" Then
                .Frm1.rdoSpouse_t.Checked   = True
            Else
                .Frm1.rdoSpouse_t.Checked   = False
            End If
            If "<%=ConvSPChars(lgObjRs("LADY"))%>" = "Y" Then
                .Frm1.rdoLady_t.Checked   = True
            Else
                .Frm1.rdoLady_t.Checked   = False
            End If
            .Frm1.txtSupp_old_cnt_t.Value   = "<%=ConvSPChars(lgObjRs("SUPP_OLD_CNT"))%>"
            .Frm1.txtSupp_young_cnt_t.Value = "<%=ConvSPChars(lgObjRs("SUPP_YOUNG_CNT"))%>"
            .Frm1.txtOld_cnt_t1.Value        = "<%=ConvSPChars(lgObjRs("OLD_CNT"))%>"
            .Frm1.txtOld_cnt_t2.Value        = "<%=ConvSPChars(lgObjRs("OLD_CNT2"))%>"            
            .Frm1.txtParia_cnt_t.Value      = "<%=ConvSPChars(lgObjRs("PARIA_CNT"))%>"
            .Frm1.txtChl_rear_inwon_t.Value = "<%=ConvSPChars(lgObjRs("CHL_REAR"))%>"
       End With          
</Script>       
<%
    End If

    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet  

    Call SubMakeSQLStatements("C",iKey1,iKey2,lgKeyStream(2))                   '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then          'If data not exists
          lgPrevNext = ""
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()       
    Else
%>
<Script Language=vbscript>
       With Parent	     
            .Frm1.txtOther_insur_amt.text      = "<%=UNINumClientFormat(lgObjRs("OTHER_INSUR"), ggAmtOfMoney.DecPoint,0)%>"            'Set condition area
            .Frm1.txtDisabled_insur_amt.text   = "<%=UNINumClientFormat(lgObjRs("DISABLED_SUB_AMT"), ggAmtOfMoney.DecPoint,0)%>"            'Set condition area
            .Frm1.txtMed_insur_amt.text        = "<%=UNINumClientFormat(lgObjRs("MED_INSUR"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtEmp_insur_amt.text        = "<%=UNINumClientFormat(lgObjRs("EMP_INSUR"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtNational_pension_amt.text = "<%=UNINumClientFormat(lgObjRs("NATIONAL_PENSION_AMT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtPer_edu_amt.text          = "<%=UNINumClientFormat(lgObjRs("PER_EDU"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtFam_edu_amt.text          = "<%=UNINumClientFormat(lgObjRs("FAM_EDU"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtFam_edu_cnt.text          = "<%=UNINumClientFormat(lgObjRs("FAM_EDU_CNT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtKind_edu_amt.text         = "<%=UNINumClientFormat(lgObjRs("KIND_EDU"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtKind_edu_cnt.text         = "<%=UNINumClientFormat(lgObjRs("KIND_EDU_CNT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtUniv_edu_amt.text         = "<%=UNINumClientFormat(lgObjRs("UNIV_EDU"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtUniv_edu_cnt.text         = "<%=UNINumClientFormat(lgObjRs("UNIV_EDU_CNT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtTot_med_amt.text          = "<%=UNINumClientFormat(lgObjRs("TOT_MED"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtSpeci_med_amt.text        = "<%=UNINumClientFormat(lgObjRs("SPECI_MED"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtLegal_contr_amt.text      = "<%=UNINumClientFormat(lgObjRs("LEGAL_CONTR"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtPoli_contr_amt1.text      = "<%=UNINumClientFormat(lgObjRs("POLI_CONTRA_AMT1"), ggAmtOfMoney.DecPoint,0)%>" '2004

            .Frm1.txtOurstock_contr_amt.text  = "<%=UNINumClientFormat(lgObjRs("OURSTOCK_CONTRA_AMT"), ggAmtOfMoney.DecPoint,0)%>" '2004
            
            .Frm1.txtApp_contr_amt.text        = "<%=UNINumClientFormat(lgObjRs("APP_CONTR"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtCeremony_amt.text        = "<%=UNINumClientFormat(lgObjRs("CEREMONY_AMT"), ggAmtOfMoney.DecPoint,0)%>" '2004 결혼장례비     
            .Frm1.txtCeremony_cnt.text        = "<%=UNINumClientFormat(lgObjRs("CEREMONY_CNT"), ggAmtOfMoney.DecPoint,0)%>" '2004 결혼장례비                    
            .Frm1.txtPriv_contr_amt.text       = "<%=UNINumClientFormat(lgObjRs("PRIV_CONTR"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtHouse_fund_amt.text       = "<%=UNINumClientFormat(lgObjRs("HOUSE_FUND"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtLong_house_loan_amt.text  = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_AMT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtLong_house_loan_amt1.text  = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_AMT1"), ggAmtOfMoney.DecPoint,0)%>"
                        
            .Frm1.txtIndiv_anu_amt.text        = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtIndiv_anu2_amt.text       = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU2"), ggAmtOfMoney.DecPoint,0)%>"
            
            .Frm1.txtCard_use_amt.text         = "<%=UNINumClientFormat(lgObjRs("CARD_USE_AMT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtCard2_use_amt.text        = "<%=UNINumClientFormat(lgObjRs("CARD2_USE_AMT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtInstitution_giro.text     = "<%=UNINumClientFormat(lgObjRs("INSTITUTION_GIRO"), ggAmtOfMoney.DecPoint,0)%>"            
            .Frm1.txtRetire_pension.text       = "<%=UNINumClientFormat(lgObjRs("RETIRE_PENSION"), ggAmtOfMoney.DecPoint,0)%>"
            
            .Frm1.txtFore_edu_amt.text         = "<%=UNINumClientFormat(lgObjRs("FORE_EDU_AMT"), ggAmtOfMoney.DecPoint,0)%>"

            .Frm1.txtOther_income_amt.text     = "<%=UNINumClientFormat(lgObjRs("OTHER_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtFore_income_amt.text      = "<%=UNINumClientFormat(lgObjRs("FORE_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtAfter_bonus_amt.text      = "<%=UNINumClientFormat(lgObjRs("AFTER_BONUS_AMT"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtHouse_repay_amt.text      = "<%=UNINumClientFormat(lgObjRs("HOUSE_REPAY"), ggAmtOfMoney.DecPoint,0)%>"
'            .Frm1.txtStock_save_amt.text       = "<%=UNINumClientFormat(lgObjRs("STOCK_SAVE"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtFore_pay_amt.text         = "<%=UNINumClientFormat(lgObjRs("FORE_PAY"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtSave_tax_sub_amt.text     = "<%=UNINumClientFormat(lgObjRs("SAVE_TAX_SUB"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtIncome_redu_amt.text      = "<%=UNINumClientFormat(lgObjRs("INCOME_REDU"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtTaxes_redu_amt.text       = "<%=UNINumClientFormat(lgObjRs("TAXES_REDU"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtTax_Union_Ded.text       = "<%=UNINumClientFormat(lgObjRs("TAX_UNION_DED"), ggAmtOfMoney.DecPoint,0)%>"	'2005         
            
            '2002년 수정 
            .Frm1.txtDisabled_edu_amt.text = "<%=UNINumClientFormat(lgObjRs("Disabled_edu_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtDisabled_edu_cnt.text = "<%=UNINumClientFormat(lgObjRs("Disabled_edu_cnt"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txOur_Stock_amt.Text     = "<%=UNINumClientFormat(lgObjRs("Our_Stock_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtTaxLaw_contr_amt.Text     = "<%=UNINumClientFormat(lgObjRs("TaxLaw_contr_amt"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtTaxLaw_contr_amt2.Text     = "<%=UNINumClientFormat(lgObjRs("TaxLaw_contr_amt2"), ggAmtOfMoney.DecPoint,0)%>"
            .Frm1.txtinvest2_sub_amt.Text      = "<%=UNINumClientFormat(lgObjRs("invest2_sub_amt"), ggAmtOfMoney.DecPoint,0)%>"
            
                                                            
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
              Call SubBizSaveSingleCreate(lgObjRs)  
        Case  OPMD_UMODE                                                             '☜ : Update
              Call SubBizSaveSingleUpdate(lgObjRs)
    End Select

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL =            " DELETE HFA030T"
    lgStrSQL = lgStrSQL & "  WHERE YY         = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "    AND EMP_NO     = " & FilterVar(lgKeyStream(1), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate(lgObjRs)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HFA030T("
    lgStrSQL = lgStrSQL & " YY,                   "
    lgStrSQL = lgStrSQL & " EMP_NO,               "
    lgStrSQL = lgStrSQL & " OTHER_INSUR,          "
    lgStrSQL = lgStrSQL & " DISABLED_SUB_AMT,     "
    lgStrSQL = lgStrSQL & " MED_INSUR,            "
    lgStrSQL = lgStrSQL & " EMP_INSUR,            "
    lgStrSQL = lgStrSQL & " NATIONAL_PENSION_AMT, "
    lgStrSQL = lgStrSQL & " PER_EDU,              "
    lgStrSQL = lgStrSQL & " FAM_EDU,              "
    lgStrSQL = lgStrSQL & " FAM_EDU_CNT,          "
    lgStrSQL = lgStrSQL & " KIND_EDU,             "
    lgStrSQL = lgStrSQL & " KIND_EDU_CNT,         "
    lgStrSQL = lgStrSQL & " UNIV_EDU,             "
    lgStrSQL = lgStrSQL & " UNIV_EDU_CNT,         "
    lgStrSQL = lgStrSQL & " TOT_MED,              "
    lgStrSQL = lgStrSQL & " SPECI_MED,            "
    lgStrSQL = lgStrSQL & " LEGAL_CONTR,          "
    lgStrSQL = lgStrSQL & " POLI_CONTRA_AMT1,          "    '2004
    lgStrSQL = lgStrSQL & " OURSTOCK_CONTRA_AMT,          "    '2004
    lgStrSQL = lgStrSQL & " APP_CONTR,            "
    lgStrSQL = lgStrSQL & " CEREMONY_AMT,         "    '2004 결혼장례비 
    lgStrSQL = lgStrSQL & " CEREMONY_CNT,         "   
            
    lgStrSQL = lgStrSQL & " PRIV_CONTR,           "
    lgStrSQL = lgStrSQL & " HOUSE_FUND,           "
    lgStrSQL = lgStrSQL & " LONG_HOUSE_LOAN_AMT,  "
    lgStrSQL = lgStrSQL & " LONG_HOUSE_LOAN_AMT1,  "    '2004
    
    lgStrSQL = lgStrSQL & " INDIV_ANU,            "
    lgStrSQL = lgStrSQL & " INDIV_ANU2,           "
    lgStrSQL = lgStrSQL & " CARD_USE_AMT,         " 
    lgStrSQL = lgStrSQL & " CARD2_USE_AMT,        " 
    lgStrSQL = lgStrSQL & " INSTITUTION_GIRO,     "		'2005
    lgStrSQL = lgStrSQL & " RETIRE_PENSION,       "     '2005       
    lgStrSQL = lgStrSQL & " FORE_EDU_AMT,         "    
    
    lgStrSQL = lgStrSQL & " OTHER_INCOME,         "
    lgStrSQL = lgStrSQL & " FORE_INCOME,          "
    lgStrSQL = lgStrSQL & " AFTER_BONUS_AMT,      "
    lgStrSQL = lgStrSQL & " HOUSE_REPAY,          "
 '   lgStrSQL = lgStrSQL & " STOCK_SAVE,           "
    lgStrSQL = lgStrSQL & " FORE_PAY,             "
    lgStrSQL = lgStrSQL & " SAVE_TAX_SUB,         "
    lgStrSQL = lgStrSQL & " INCOME_REDU,          "
    lgStrSQL = lgStrSQL & " TAXES_REDU,           "
    lgStrSQL = lgStrSQL & " TAX_UNION_DED,		  "	'2005
    lgStrSQL = lgStrSQL & " TECH_SUB_AMT,         " '현재는 쓰지않는 현장기술인력공제 필드 
    lgStrSQL = lgStrSQL & " MED_SPPORT,           " '의료비지원 필드.. 왜있을까?
    lgStrSQL = lgStrSQL & " EDU_SPPORT,           " '교육비지원 필드.. 왜있을까?
    
    '2002년 수정 
    lgStrSQL = lgStrSQL & " Disabled_edu_amt, "
    lgStrSQL = lgStrSQL & " Disabled_edu_cnt, "
    lgStrSQL = lgStrSQL & " Our_Stock_amt, "
    lgStrSQL = lgStrSQL & " TaxLaw_contr_amt, "
    lgStrSQL = lgStrSQL & " TaxLaw_contr_amt2, "    
    lgStrSQL = lgStrSQL & " Invest2_sub_amt, "
    
    
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO,          " 
    lgStrSQL = lgStrSQL & " ISRT_DT     ,         " 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO ,         " 
    lgStrSQL = lgStrSQL & " UPDT_DT      )        " 
    
    lgStrSQL = lgStrSQL & " VALUES (" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")              & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOther_insur_amt"),0)      & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtDisabled_insur_amt"),0)   & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtMed_insur_amt"),0)        & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtEmp_insur_amt"),0)        & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtNational_pension_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtPer_edu_amt"),0)          & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFam_edu_amt"),0)          & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFam_edu_cnt"),0)          & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtKind_edu_amt"),0)         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtKind_edu_cnt"),0)         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtUniv_edu_amt"),0)         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtUniv_edu_cnt"),0)         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtTot_med_amt"),0)          & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtSpeci_med_amt"),0)        & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLegal_contr_amt"),0)      & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtPoli_contr_amt1"),0)      & ","   '2004
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOurstock_contr_amt"),0)      & ","   '2004
    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtApp_contr_amt"),0)        & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtCeremony_amt"),0)        & ","		'2004 결혼장례비 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtCeremony_cnt"),0)        & ","		'2004 결혼장례비 
        
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtPriv_contr_amt"),0)       & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtHouse_fund_amt"),0)       & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLong_house_loan_amt"),0)  & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLong_house_loan_amt1"),0)  & ","  '2004 상환기간 15년이상  
        
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIndiv_anu_amt"),0)        & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIndiv_anu2_amt"),0)       & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtCard_use_amt"),0)         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtCard2_use_amt"),0)        & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtInstitution_giro"),0)	   & ","	'2005
	lgStrSQL = lgStrSQL & UNIConvNum(Request("txtRetire_pension"),0)       & ","    '2005
        
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFore_edu_amt"),0)         & ","      

    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOther_income_amt"),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFore_income_amt"),0)      & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtAfter_bonus_amt"),0)      & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtHouse_repay_amt"),0)      & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtStock_save_amt"),0)       & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFore_pay_amt"),0)         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtSave_tax_sub_amt"),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtIncome_redu_amt"),0)      & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtTaxes_redu_amt"),0)       & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtTax_Union_Ded"),0)       & ","		'2005
    lgStrSQL = lgStrSQL & "0 , 0, 0,"
    
    '2002년 수정 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtDisabled_edu_amt"),0)       & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtDisabled_edu_cnt"),0)       & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txOur_Stock_amt"),0)       & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtTaxLaw_contr_amt"),0)       & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtTaxLaw_contr_amt2"),0)       & ","     
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtinvest2_sub_amt"),0)       & ","  
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                       & "," 
    lgStrSQL = lgStrSQL & lgSvrDateTime                                   & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                       & "," 
    lgStrSQL = lgStrSQL & lgSvrDateTime
    lgStrSQL = lgStrSQL & ")"
 
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate(lgObjRs)
    Dim ceremonyCnt , ceremonyAmt
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    call CommonQueryRs(" CEREMONY_AMT "," HFA020T "," 1=1",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ceremonyAmt = Replace(lgF0, Chr(11), "")
	ceremonyCnt = UNIConvNum(Request("txtCeremony_cnt"),0)
	ceremonyAmt = ceremonyAmt  * ceremonyCnt

    lgStrSQL =            "UPDATE HFA030T"
    lgStrSQL = lgStrSQL & "   SET OTHER_INSUR          = " & UNIConvNum(Request("txtOther_insur_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       DISABLED_SUB_AMT     = " & UNIConvNum(Request("txtDisabled_insur_amt"),0)   & ","  
    lgStrSQL = lgStrSQL & "       MED_INSUR            = " & UNIConvNum(Request("txtMed_insur_amt"),0)        & ","  
    lgStrSQL = lgStrSQL & "       EMP_INSUR            = " & UNIConvNum(Request("txtEmp_insur_amt"),0)        & ","
    lgStrSQL = lgStrSQL & "       NATIONAL_PENSION_AMT = " & UNIConvNum(Request("txtNational_pension_amt"),0) & ","
    lgStrSQL = lgStrSQL & "       PER_EDU              = " & UNIConvNum(Request("txtPer_edu_amt"),0)          & ","
    lgStrSQL = lgStrSQL & "       FAM_EDU              = " & UNIConvNum(Request("txtFam_edu_amt"),0)          & ","
    lgStrSQL = lgStrSQL & "       FAM_EDU_CNT          = " & UNIConvNum(Request("txtFam_edu_cnt"),0)          & ","
    lgStrSQL = lgStrSQL & "       KIND_EDU             = " & UNIConvNum(Request("txtKind_edu_amt"),0)         & ","
    lgStrSQL = lgStrSQL & "       KIND_EDU_CNT         = " & UNIConvNum(Request("txtKind_edu_cnt"),0)         & ","
    lgStrSQL = lgStrSQL & "       UNIV_EDU             = " & UNIConvNum(Request("txtUniv_edu_amt"),0)         & ","
    lgStrSQL = lgStrSQL & "       UNIV_EDU_CNT         = " & UNIConvNum(Request("txtUniv_edu_cnt"),0)         & ","
    lgStrSQL = lgStrSQL & "       TOT_MED              = " & UNIConvNum(Request("txtTot_med_amt"),0)          & ","
    lgStrSQL = lgStrSQL & "       SPECI_MED            = " & UNIConvNum(Request("txtSpeci_med_amt"),0)        & ","
    lgStrSQL = lgStrSQL & "       LEGAL_CONTR          = " & UNIConvNum(Request("txtLegal_contr_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       POLI_CONTRA_AMT1     = " & UNIConvNum(Request("txtPoli_contr_amt1"),0)      & ","	'2004    
    lgStrSQL = lgStrSQL & "       OURSTOCK_CONTRA_AMT  = " & UNIConvNum(Request("txtOurstock_contr_amt"),0)   & ","	'2004
    
    lgStrSQL = lgStrSQL & "       APP_CONTR            = " & UNIConvNum(Request("txtApp_contr_amt"),0)        & ","

    lgStrSQL = lgStrSQL & "       CEREMONY_CNT            = " & UNIConvNum(Request("txtCeremony_cnt"),0)        & "," '2004 결혼장례비 
    lgStrSQL = lgStrSQL & "       CEREMONY_AMT         = " & UNIConvNum(ceremonyAmt,0)        & ","       
    
    lgStrSQL = lgStrSQL & "       PRIV_CONTR           = " & UNIConvNum(Request("txtPriv_contr_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       HOUSE_FUND           = " & UNIConvNum(Request("txtHouse_fund_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT  = " & UNIConvNum(Request("txtLong_house_loan_amt"),0)  & ","
    lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT1 = " & UNIConvNum(Request("txtLong_house_loan_amt1"),0)  & ","
        
    lgStrSQL = lgStrSQL & "       INDIV_ANU            = " & UNIConvNum(Request("txtIndiv_anu_amt"),0)        & ","
    lgStrSQL = lgStrSQL & "       INDIV_ANU2           = " & UNIConvNum(Request("txtIndiv_anu2_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       CARD_USE_AMT         = " & UNIConvNum(Request("txtCard_use_amt"),0)         & ","
    lgStrSQL = lgStrSQL & "       CARD2_USE_AMT        = " & UNIConvNum(Request("txtCard2_use_amt"),0)		  & ","    
    lgStrSQL = lgStrSQL & "       INSTITUTION_GIRO     = " & UNIConvNum(Request("txtInstitution_giro"),0)     & ","    
    lgStrSQL = lgStrSQL & "       RETIRE_PENSION       = " & UNIConvNum(Request("txtRetire_pension"),0)       & ","    
    
    lgStrSQL = lgStrSQL & "       FORE_EDU_AMT         = " & UNIConvNum(Request("txtFore_edu_amt"),0)         & ","    
    
    lgStrSQL = lgStrSQL & "       OTHER_INCOME         = " & UNIConvNum(Request("txtOther_income_amt"),0)     & ","
    lgStrSQL = lgStrSQL & "       FORE_INCOME          = " & UNIConvNum(Request("txtFore_income_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       AFTER_BONUS_AMT      = " & UNIConvNum(Request("txtAfter_bonus_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       HOUSE_REPAY          = " & UNIConvNum(Request("txtHouse_repay_amt"),0)      & ","
'    lgStrSQL = lgStrSQL & "       STOCK_SAVE           = " &  UNIConvNum(Request("txtStock_save_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       FORE_PAY             = " & UNIConvNum(Request("txtFore_pay_amt"),0)         & ","
    lgStrSQL = lgStrSQL & "       SAVE_TAX_SUB         = " & UNIConvNum(Request("txtSave_tax_sub_amt"),0)     & ","
    lgStrSQL = lgStrSQL & "       INCOME_REDU          = " & UNIConvNum(Request("txtIncome_redu_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       TAXES_REDU           = " & UNIConvNum(Request("txtTaxes_redu_amt"),0)       & "," 
    lgStrSQL = lgStrSQL & "       TAX_UNION_DED        = " & UNIConvNum(Request("txtTax_Union_Ded"),0)       & ","		'2005
        
    '2002년 수정 
    lgStrSQL = lgStrSQL & "       Disabled_edu_amt     = " & UNIConvNum(Request("txtDisabled_edu_amt"),0)       & "," 
    lgStrSQL = lgStrSQL & "       Disabled_edu_cnt     = " & UNIConvNum(Request("txtDisabled_edu_cnt"),0)       & "," 
    lgStrSQL = lgStrSQL & "       Our_Stock_amt        = " & UNIConvNum(Request("txOur_Stock_amt"),0)       & "," 
    lgStrSQL = lgStrSQL & "       TaxLaw_contr_amt     = " & UNIConvNum(Request("txtTaxLaw_contr_amt"),0)       & "," 
    lgStrSQL = lgStrSQL & "       TaxLaw_contr_amt2     = " & UNIConvNum(Request("txtTaxLaw_contr_amt2"),0)       & ","     
    lgStrSQL = lgStrSQL & "       Invest2_sub_amt      = " & UNIConvNum(Request("txtinvest2_sub_amt"),0)       & "," 
    
    lgStrSQL = lgStrSQL & "       UPDT_EMP_NO          = " & FilterVar(gUsrId, "''", "S")                       & "," 
    lgStrSQL = lgStrSQL & "       UPDT_DT              = " & lgSvrDateTime
    lgStrSQL = lgStrSQL & " WHERE YY         = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND EMP_NO     = " & FilterVar(lgKeyStream(1), "''", "S")
'Response.Write	lgStrSQL
'Response.End

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1,pCode2,pCode3)
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
	                Case ""
                         lgStrSQL =            "Select A.EMP_NO EMP_NO, A.NAME NAME, A.DEPT_NM DEPT_NM, A.ROLL_PSTN ROLL_PSTN, "
                         lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", A.ROLL_PSTN) ROLL_PSTN_NM, " 
                         lgStrSQL = lgStrSQL & "       A.PAY_GRD1 PAY_GRD1, dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", A.pay_Grd1) Pay_Grd1_NM,  A.PAY_GRD2 PAY_GRD2, A.ENTR_DT ENTR_DT, " 
                         
                         lgStrSQL = lgStrSQL & "       IsNull(B.SPOUSE," & FilterVar("N", "''", "S") & " ) SPOUSE, IsNull(B.LADY," & FilterVar("N", "''", "S") & " ) LADY, "
                         lgStrSQL = lgStrSQL & "       B.SUPP_OLD_CNT SUPP_OLD_CNT, B.SUPP_YOUNG_CNT SUPP_YOUNG_CNT, "
                         lgStrSQL = lgStrSQL & "       B.OLD_CNT OLD_CNT,B.OLD_CNT2 OLD_CNT2, B.PARIA_CNT PARIA_CNT, B.CHL_REAR CHL_REAR" '2004 경로 
                         lgStrSQL = lgStrSQL & "  From HAA010T A, HDF020T B "
                         lgStrSQL = lgStrSQL & " Where A.EMP_NO = " & pCode2 	
                         lgStrSQL = lgStrSQL & "   And A.EMP_NO = B.EMP_NO"
                         lgStrSQL = lgStrSQL & "   And A.INTERNAL_CD LIKE  " & FilterVar(pCode3 & "%", "''", "S") & "" 
                    Case "P"
                         lgStrSQL =            "Select TOP 1 A.EMP_NO EMP_NO, A.NAME NAME, A.DEPT_NM DEPT_NM, A.ROLL_PSTN ROLL_PSTN, "
                         lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", A.ROLL_PSTN) ROLL_PSTN_NM, " 
                         lgStrSQL = lgStrSQL & "       A.PAY_GRD1 PAY_GRD1, dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", A.pay_Grd1) Pay_Grd1_NM,  A.PAY_GRD2 PAY_GRD2, A.ENTR_DT ENTR_DT, " 
                         
                         lgStrSQL = lgStrSQL & "       IsNull(B.SPOUSE," & FilterVar("N", "''", "S") & " ) SPOUSE, IsNull(B.LADY," & FilterVar("N", "''", "S") & " ) LADY, "
                         lgStrSQL = lgStrSQL & "       B.SUPP_OLD_CNT SUPP_OLD_CNT, B.SUPP_YOUNG_CNT SUPP_YOUNG_CNT, "
                         lgStrSQL = lgStrSQL & "       B.OLD_CNT OLD_CNT,B.OLD_CNT2 OLD_CNT2, B.PARIA_CNT PARIA_CNT, B.CHL_REAR CHL_REAR"
                         lgStrSQL = lgStrSQL & "  From HAA010T A, HDF020T B "
                         lgStrSQL = lgStrSQL & " Where A.EMP_NO < " & pCode2 	
                         lgStrSQL = lgStrSQL & "   And A.EMP_NO = B.EMP_NO"
                         lgStrSQL = lgStrSQL & "   And A.INTERNAL_CD LIKE  " & FilterVar(pCode3 & "%", "''", "S") & "" 
                         lgStrSQL = lgStrSQL & "   And (A.RETIRE_RESN IS NULL OR "
                         lgStrSQL = lgStrSQL & "        A.RETIRE_RESN = '' OR A.RETIRE_RESN = " & FilterVar("6", "''", "S") & ") "                         
                         lgStrSQL = lgStrSQL & " ORDER BY A.EMP_NO DESC"
                      
                    Case "N"
                         lgStrSQL =            "Select TOP 1 A.EMP_NO EMP_NO, A.NAME NAME, A.DEPT_NM DEPT_NM, A.ROLL_PSTN ROLL_PSTN, "
                         lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", A.ROLL_PSTN) ROLL_PSTN_NM, " 
                         lgStrSQL = lgStrSQL & "       A.PAY_GRD1 PAY_GRD1, dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", A.pay_Grd1) Pay_Grd1_NM,  A.PAY_GRD2 PAY_GRD2, A.ENTR_DT ENTR_DT, " 
                         
                         lgStrSQL = lgStrSQL & "       IsNull(B.SPOUSE," & FilterVar("N", "''", "S") & " ) SPOUSE, IsNull(B.LADY," & FilterVar("N", "''", "S") & " ) LADY, "
                         lgStrSQL = lgStrSQL & "       B.SUPP_OLD_CNT SUPP_OLD_CNT, B.SUPP_YOUNG_CNT SUPP_YOUNG_CNT, "
                         lgStrSQL = lgStrSQL & "       B.OLD_CNT OLD_CNT,B.OLD_CNT2 OLD_CNT2, B.PARIA_CNT PARIA_CNT, B.CHL_REAR CHL_REAR"
                         lgStrSQL = lgStrSQL & "  From HAA010T A, HDF020T B "
                         lgStrSQL = lgStrSQL & " Where A.EMP_NO > " & pCode2 	
                         lgStrSQL = lgStrSQL & "   And A.EMP_NO = B.EMP_NO"
                         lgStrSQL = lgStrSQL & "   And A.INTERNAL_CD LIKE  " & FilterVar(pCode3 & "%", "''", "S") & "" 
                         lgStrSQL = lgStrSQL & "   And (A.RETIRE_RESN IS NULL OR "
                         lgStrSQL = lgStrSQL & "        A.RETIRE_RESN = '' OR A.RETIRE_RESN = " & FilterVar("6", "''", "S") & ") "                         
                         lgStrSQL = lgStrSQL & " ORDER BY A.EMP_NO ASC"


            End Select
      
      Case "C"
             Select Case  lgPrevNext 
	                Case ""
                         lgStrSQL =            "Select OTHER_INCOME, FORE_INCOME, EDU_SPPORT, MED_SPPORT, OTHER_INSUR, DISABLED_SUB_AMT, " 
			             lgStrSQL = lgStrSQL & "       MED_INSUR, EMP_INSUR, TOT_MED, SPECI_MED, PER_EDU, FAM_EDU, "
                         lgStrSQL = lgStrSQL & "       UNIV_EDU, KIND_EDU, KIND_EDU_CNT, UNIV_EDU_CNT, LEGAL_CONTR, APP_CONTR, "    
                         lgStrSQL = lgStrSQL & "       POLI_CONTRA_AMT1,OURSTOCK_CONTRA_AMT , " '2004                                  
                         lgStrSQL = lgStrSQL & "       PRIV_CONTR, HOUSE_FUND, INDIV_ANU, SAVE_TAX_SUB, HOUSE_REPAY, "
                         lgStrSQL = lgStrSQL & "       STOCK_SAVE,  FORE_PAY, INCOME_REDU, TAXES_REDU,TAX_UNION_DED, INDIV_ANU2, "
                         lgStrSQL = lgStrSQL & "       CARD_USE_AMT,CARD2_USE_AMT ,INSTITUTION_GIRO,RETIRE_PENSION, FORE_EDU_AMT, FAM_EDU_CNT,  "
                         lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT,LONG_HOUSE_LOAN_AMT1, AFTER_BONUS_AMT, NATIONAL_PENSION_AMT, "
                         lgStrSQL = lgStrSQL & "       Disabled_edu_amt, Disabled_edu_cnt, Our_Stock_amt, "
                         lgStrSQL = lgStrSQL & "       TaxLaw_contr_amt,TaxLaw_contr_amt2, Invest2_sub_amt, "
                         lgStrSQL = lgStrSQL & "       CEREMONY_AMT,CEREMONY_CNT "	'2004  결혼장례비 
                         lgStrSQL = lgStrSQL & " From  HFA030T "
                         lgStrSQL = lgStrSQL & " Where YY     = " & pCode1 	
                         lgStrSQL = lgStrSQL & "   And EMP_NO = " & pCode2 	
                     Case "P"
                         lgStrSQL =            "Select OTHER_INCOME, FORE_INCOME, EDU_SPPORT, MED_SPPORT, OTHER_INSUR, DISABLED_SUB_AMT, " 
			             lgStrSQL = lgStrSQL & "       MED_INSUR, EMP_INSUR, TOT_MED, SPECI_MED, PER_EDU, FAM_EDU, "
                         lgStrSQL = lgStrSQL & "       UNIV_EDU, KIND_EDU, KIND_EDU_CNT, UNIV_EDU_CNT, LEGAL_CONTR, APP_CONTR, "    
                         lgStrSQL = lgStrSQL & "       POLI_CONTRA_AMT1,OURSTOCK_CONTRA_AMT , " '2004                                  
                         lgStrSQL = lgStrSQL & "       PRIV_CONTR, HOUSE_FUND, INDIV_ANU, SAVE_TAX_SUB, HOUSE_REPAY, "
                         lgStrSQL = lgStrSQL & "       STOCK_SAVE,  FORE_PAY, INCOME_REDU, TAXES_REDU,TAX_UNION_DED, INDIV_ANU2, "
                         lgStrSQL = lgStrSQL & "       CARD_USE_AMT,CARD2_USE_AMT ,INSTITUTION_GIRO,RETIRE_PENSION, FORE_EDU_AMT, FAM_EDU_CNT,  "
                         lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT,LONG_HOUSE_LOAN_AMT1, AFTER_BONUS_AMT, NATIONAL_PENSION_AMT, "
                         lgStrSQL = lgStrSQL & "       Disabled_edu_amt, Disabled_edu_cnt, Our_Stock_amt, "
                         lgStrSQL = lgStrSQL & "       TaxLaw_contr_amt,TaxLaw_contr_amt2, Invest2_sub_amt, "
                         lgStrSQL = lgStrSQL & "       CEREMONY_AMT,CEREMONY_CNT "	'2004  결혼장례비                     
                         lgStrSQL = lgStrSQL & " From  HFA030T "
                         lgStrSQL = lgStrSQL & " Where YY     = " & pCode1 	
                         lgStrSQL = lgStrSQL & "   And EMP_NO = " & pCode2 	
                     Case "N"
                         lgStrSQL =            "Select OTHER_INCOME, FORE_INCOME, EDU_SPPORT, MED_SPPORT, OTHER_INSUR, DISABLED_SUB_AMT, " 
			             lgStrSQL = lgStrSQL & "       MED_INSUR, EMP_INSUR, TOT_MED, SPECI_MED, PER_EDU, FAM_EDU, "
                         lgStrSQL = lgStrSQL & "       UNIV_EDU, KIND_EDU, KIND_EDU_CNT, UNIV_EDU_CNT, LEGAL_CONTR, APP_CONTR, "    
                         lgStrSQL = lgStrSQL & "       POLI_CONTRA_AMT1,OURSTOCK_CONTRA_AMT , " '2004                                  
                         lgStrSQL = lgStrSQL & "       PRIV_CONTR, HOUSE_FUND, INDIV_ANU, SAVE_TAX_SUB, HOUSE_REPAY, "
                         lgStrSQL = lgStrSQL & "       STOCK_SAVE,  FORE_PAY, INCOME_REDU, TAXES_REDU,TAX_UNION_DED, INDIV_ANU2, "
                         lgStrSQL = lgStrSQL & "       CARD_USE_AMT,CARD2_USE_AMT ,INSTITUTION_GIRO,RETIRE_PENSION, FORE_EDU_AMT, FAM_EDU_CNT,  "
                         lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT,LONG_HOUSE_LOAN_AMT1, AFTER_BONUS_AMT, NATIONAL_PENSION_AMT, "
                         lgStrSQL = lgStrSQL & "       Disabled_edu_amt, Disabled_edu_cnt, Our_Stock_amt, "
                         lgStrSQL = lgStrSQL & "       TaxLaw_contr_amt,TaxLaw_contr_amt2, Invest2_sub_amt, "
                         lgStrSQL = lgStrSQL & "       CEREMONY_AMT,CEREMONY_CNT "	'2004  결혼장례비                      
                         lgStrSQL = lgStrSQL & " From  HFA030T "
                         lgStrSQL = lgStrSQL & " Where YY     = " & pCode1 	
                         lgStrSQL = lgStrSQL & "   And EMP_NO = " & pCode2 	
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
