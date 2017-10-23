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

    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")   
    lgIntFlgMode      = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002" 
			 Call SubBizSave()                                                    '☜: Save,Update
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

    strEmpNo  = FilterVar(lgKeyStream(0), "''", "S")
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

    strEmpNo  = FilterVar(emp_no, "''", "S")

    strYear   = lgKeyStream(2)
    Call SubMakeSQLStatements("R",strYear,strEmpNo)                                 '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
         ' Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
         ' Call SetErrorStatus()
       End If
          
    Else
   
%>
<Script Language=vbscript>
            
       With Parent	
            if "<%=ConvSPChars(lgObjRs("SPOUSE"))%>" = "Y" THEN
                .frm1.rdoSpouse_t.checked = true
            else
                .frm1.rdoSpouse_t.checked = false
            end if
            if "<%=ConvSPChars(lgObjRs("LADY"))%>" = "Y" THEN
                .frm1.rdoLady_t.checked = true
            else
                .frm1.rdoLady_t.checked = false
            end if
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

    ' 수정가능한 영역 조회 
    Call SubMakeSQLStatements("C",FilterVar(strYear, "''", "S"),FilterVar(emp_no, "''", "S"))                   '☜ : Make sql statements

    If 	FncOpenRs("C",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then          'If data not exists
        
       If lgPrevNext = "" Then
        
            lgErrorStatus = "YES"

          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
          Call SetErrorStatus()
         
       End If
       
    Else
%>
<Script Language=vbscript>
       With Parent	     

            .Frm1.txtOther_insur_amt.Value      = "<%=UNINumClientFormat(lgObjRs("OTHER_INSUR"),ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtMed_insur_amt.Value        = "<%=UNINumClientFormat(lgObjRs("MED_INSUR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtEmp_insur_amt.Value        = "<%=UNINumClientFormat(lgObjRs("EMP_INSUR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtDisabled_insur_amt.Value   = "<%=UNINumClientFormat(lgObjRs("DISABLED_SUB_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtNational_pension_amt.Value = "<%=UNINumClientFormat(lgObjRs("NATIONAL_PENSION_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtPer_edu_amt.Value          = "<%=UNINumClientFormat(lgObjRs("PER_EDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtDisabled_edu_amt.Value     = "<%=UNINumClientFormat(lgObjRs("disabled_edu_amt"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtFore_edu_amt.Value         = "<%=UNINumClientFormat(lgObjRs("FORE_EDU_AMT"), ggAmtOfMoney.DecPoint, 0)%>"                                    
            .Frm1.txtFam_edu_amt.Value          = "<%=UNINumClientFormat(lgObjRs("FAM_EDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtFam_edu_cnt.Value          = "<%=UNINumClientFormat(lgObjRs("FAM_EDU_CNT"), 0, 0)%>"
            .Frm1.txtKind_edu_amt.Value         = "<%=UNINumClientFormat(lgObjRs("KIND_EDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtKind_edu_cnt.Value         = "<%=UNINumClientFormat(lgObjRs("KIND_EDU_CNT"), 0, 0)%>"
            .Frm1.txtUniv_edu_amt.Value         = "<%=UNINumClientFormat(lgObjRs("UNIV_EDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtUniv_edu_cnt.Value         = "<%=UNINumClientFormat(lgObjRs("UNIV_EDU_CNT"), 0, 0)%>"
            .Frm1.txtTot_med_amt.Value          = "<%=UNINumClientFormat(lgObjRs("TOT_MED"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtSpeci_med_amt.Value        = "<%=UNINumClientFormat(lgObjRs("SPECI_MED"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtLegal_contr_amt.Value      = "<%=UNINumClientFormat(lgObjRs("LEGAL_CONTR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtApp_contr_amt.Value        = "<%=UNINumClientFormat(lgObjRs("APP_CONTR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtTaxLaw_contr_amt.Value     = "<%=UNINumClientFormat(lgObjRs("TaxLaw_contr_amt"), ggAmtOfMoney.DecPoint, 0)%>"
            
            .Frm1.txtPoli_contr_amt1.Value     = "<%=UNINumClientFormat(lgObjRs("POLI_CONTRA_AMT1"), ggAmtOfMoney.DecPoint, 0)%>"	'2004
            .Frm1.txtPoli_contr_amt2.Value     = "<%=UNINumClientFormat(lgObjRs("POLI_CONTRA_AMT2"), ggAmtOfMoney.DecPoint, 0)%>"	'2004
            .Frm1.txtOurstock_contr_amt.Value	= "<%=UNINumClientFormat(lgObjRs("OURSTOCK_CONTRA_AMT"), ggAmtOfMoney.DecPoint, 0)%>"	'2004
			.Frm1.txtCeremony_amt.Value			= "<%=UNINumClientFormat(lgObjRs("CEREMONY_AMT"), ggAmtOfMoney.DecPoint, 0)%>"	'2004 
            .Frm1.txtCeremony_cnt.Value			= "<%=UNINumClientFormat(lgObjRs("CEREMONY_CNT"), ggAmtOfMoney.DecPoint, 0)%>" '2004 
            
            .Frm1.txtFam_edu_cnt.Value          = "<%=UNINumClientFormat(lgObjRs("FAM_EDU_CNT"), 0, 0)%>"	                                  
            .Frm1.txtPriv_contr_amt.Value       = "<%=UNINumClientFormat(lgObjRs("PRIV_CONTR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtOur_stock_amt.Value        = "<%=UNINumClientFormat(lgObjRs("our_stock_amt"), ggAmtOfMoney.DecPoint, 0)%>"            
            .Frm1.txtHouse_fund_amt.Value       = "<%=UNINumClientFormat(lgObjRs("HOUSE_FUND"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtLong_house_loan_amt.Value  = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtLong_house_loan_amt1.Value  = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_AMT1"), ggAmtOfMoney.DecPoint, 0)%>"
            
            .Frm1.txtIndiv_anu_amt.Value        = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtIndiv_anu2_amt.Value       = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU2"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtInvest_sub_amt.Value       = "<%=UNINumClientFormat(lgObjRs("INVEST_SUB_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtVenture_sub_amt.Value      = "<%=UNINumClientFormat(lgObjRs("VENTURE_SUB_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtinvest2_sub_amt.Value      = "<%=UNINumClientFormat(lgObjRs("invest2_sub_amt"), ggAmtOfMoney.DecPoint, 0)%>"            
            .Frm1.txtCard_use_amt.Value         = "<%=UNINumClientFormat(lgObjRs("CARD_USE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtCard2_use_amt.Value         = "<%=UNINumClientFormat(lgObjRs("CARD2_USE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"                        
            .Frm1.txtOther_income_amt.Value     = "<%=UNINumClientFormat(lgObjRs("OTHER_INCOME"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtFore_income_amt.Value      = "<%=UNINumClientFormat(lgObjRs("FORE_INCOME"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtAfter_bonus_amt.Value      = "<%=UNINumClientFormat(lgObjRs("AFTER_BONUS_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtHouse_repay_amt.Value      = "<%=UNINumClientFormat(lgObjRs("HOUSE_REPAY"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtFore_pay_amt.Value         = "<%=UNINumClientFormat(lgObjRs("FORE_PAY"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtSave_tax_sub_amt.Value     = "<%=UNINumClientFormat(lgObjRs("SAVE_TAX_SUB"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtIncome_redu_amt.Value      = "<%=UNINumClientFormat(lgObjRs("INCOME_REDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtTaxes_redu_amt.Value       = "<%=UNINumClientFormat(lgObjRs("TAXES_REDU"), ggAmtOfMoney.DecPoint, 0)%>"
       End With          
</Script>       
<%     
    End If

    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet  

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    Select Case lgIntFlgMode
        Case  OPMD_UMODE                                                            '☜ : Update
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
Sub SubBizSaveSingleUpdate(lgObjRs)
    Dim ceremonyCnt , ceremonyAmt
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL =            "UPDATE HFA030T"
    lgStrSQL = lgStrSQL & "   SET TOT_MED              = " & UNIConvNum(Request("txtTot_med_amt"),0)          & ","
    lgStrSQL = lgStrSQL & "       OTHER_INCOME         = " & UNIConvNum(Request("txtOther_income_amt"),0)      & ","  
    lgStrSQL = lgStrSQL & "       MED_INSUR            = " & UNIConvNum(Request("txtMed_insur_amt"),0)        & ","  
    lgStrSQL = lgStrSQL & "       EMP_INSUR            = " & UNIConvNum(Request("txtEmp_insur_amt"),0)        & ","
    lgStrSQL = lgStrSQL & "       DISABLED_SUB_AMT     = " & UNIConvNum(Request("txtDisabled_insur_amt"),0)   & ","
    lgStrSQL = lgStrSQL & "       NATIONAL_PENSION_AMT = " & UNIConvNum(Request("txtNational_pension_amt"),0) & ","
    lgStrSQL = lgStrSQL & "       PER_EDU              = " & UNIConvNum(Request("txtPer_edu_amt"),0)          & ","
    lgStrSQL = lgStrSQL & "       disabled_edu_amt     = " & UNIConvNum(Request("txtDisabled_edu_amt"),0)          & ","
'-- 2003 외국인근로자교육비 추가 
    lgStrSQL = lgStrSQL & "       FORE_EDU_AMT		   = " & UNIConvNum(Request("txtFore_edu_amt"),0)          & ","            
    lgStrSQL = lgStrSQL & "       FAM_EDU              = " & UNIConvNum(Request("txtFam_edu_amt"),0)          & ","
    lgStrSQL = lgStrSQL & "       FAM_EDU_CNT          = " & UNIConvNum(Request("txtFam_edu_cnt"),0)          & ","
    lgStrSQL = lgStrSQL & "       KIND_EDU             = " & UNIConvNum(Request("txtKind_edu_amt"),0)         & ","
    lgStrSQL = lgStrSQL & "       KIND_EDU_CNT         = " & UNIConvNum(Request("txtKind_edu_cnt"),0)         & ","
    lgStrSQL = lgStrSQL & "       UNIV_EDU             = " & UNIConvNum(Request("txtUniv_edu_amt"),0)         & ","
    lgStrSQL = lgStrSQL & "       UNIV_EDU_CNT         = " & UNIConvNum(Request("txtUniv_edu_cnt"),0)         & ","
    lgStrSQL = lgStrSQL & "       SPECI_MED            = " & UNIConvNum(Request("txtSpeci_med_amt"),0)        & ","
    lgStrSQL = lgStrSQL & "       LEGAL_CONTR          = " & UNIConvNum(Request("txtLegal_contr_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       APP_CONTR            = " & UNIConvNum(Request("txtApp_contr_amt"),0)        & ","
    lgStrSQL = lgStrSQL & "       TaxLaw_contr_amt     = " & UNIConvNum(Request("txtTaxLaw_contr_amt"),0)     & ","   
    
    lgStrSQL = lgStrSQL & "       POLI_CONTRA_AMT1     = " & UNIConvNum(Request("txtPoli_contr_amt1"),0)     & ","   '2004
    lgStrSQL = lgStrSQL & "       POLI_CONTRA_AMT2     = " & UNIConvNum(Request("txtPoli_contr_amt2"),0)     & ","   '2004
    lgStrSQL = lgStrSQL & "       OURSTOCK_CONTRA_AMT     = " & UNIConvNum(Request("txtOurstock_contr_amt"),0)     & ","   '2004
    
    lgStrSQL = lgStrSQL & "       PRIV_CONTR           = " & UNIConvNum(Request("txtPriv_contr_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       our_stock_amt        = " & UNIConvNum(Request("txtOur_stock_amt"),0)       & ","    
    lgStrSQL = lgStrSQL & "       HOUSE_FUND           = " & UNIConvNum(Request("txtHouse_fund_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT  = " & UNIConvNum(Request("txtLong_house_loan_amt"),0)  & ","
    lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT1  = " & UNIConvNum(Request("txtLong_house_loan_amt1"),0)  & ","		'2004 
    lgStrSQL = lgStrSQL & "       INDIV_ANU            = " & UNIConvNum(Request("txtIndiv_anu_amt"),0)        & ","
    lgStrSQL = lgStrSQL & "       INDIV_ANU2           = " & UNIConvNum(Request("txtIndiv_anu2_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       INVEST_SUB_AMT       = " & UNIConvNum(Request("txtInvest_sub_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       VENTURE_SUB_AMT      = " & UNIConvNum(Request("txtVenture_sub_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       invest2_sub_amt      = " & UNIConvNum(Request("txtinvest2_sub_amt"),0)      & ","    
    lgStrSQL = lgStrSQL & "       CARD_USE_AMT         = " & UNIConvNum(Request("txtCard_use_amt"),0)         & ","
'-- 2003 직불카드 추가 
    lgStrSQL = lgStrSQL & "       CARD2_USE_AMT         = " & UNIConvNum(Request("txtCard2_use_amt"),0)         & ","    
    lgStrSQL = lgStrSQL & "       OTHER_INSUR          = " & UNIConvNum(Request("txtOther_insur_amt"),0)     & ","
    lgStrSQL = lgStrSQL & "       FORE_INCOME          = " & UNIConvNum(Request("txtFore_income_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       AFTER_BONUS_AMT      = " & UNIConvNum(Request("txtAfter_bonus_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       HOUSE_REPAY          = " & UNIConvNum(Request("txtHouse_repay_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       FORE_PAY             = " & UNIConvNum(Request("txtFore_pay_amt"),0)         & ","
    lgStrSQL = lgStrSQL & "       SAVE_TAX_SUB         = " & UNIConvNum(Request("txtSave_tax_sub_amt"),0)     & ","
    lgStrSQL = lgStrSQL & "       INCOME_REDU          = " & UNIConvNum(Request("txtIncome_redu_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       TAXES_REDU           = " & UNIConvNum(Request("txtTaxes_redu_amt"),0)     & ","
    lgStrSQL = lgStrSQL & "       CEREMONY_CNT           = " & UNIConvNum(Request("txtCeremony_cnt"),0)     & ","	'2004
 
    call CommonQueryRs(" CEREMONY_AMT "," HFA020T "," 1=1",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ceremonyAmt = Replace(lgF0, Chr(11), "")
	ceremonyCnt = UNIConvNum(Request("txtCeremony_cnt"),0)

	ceremonyAmt = ceremonyAmt  * ceremonyCnt
    
    lgStrSQL = lgStrSQL & "       CEREMONY_AMT         = " & UNIConvNum(ceremonyAmt,0)         
    
    lgStrSQL = lgStrSQL & "       WHERE YY             = " & FilterVar(lgKeyStream(2), "''", "S")
    lgStrSQL = lgStrSQL & "       AND EMP_NO           = " & FilterVar(lgKeyStream(0), "''", "S")


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1,pCode2)

    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
	                Case ""

						lgStrSQL =            "Select A.EMP_NO EMP_NO, A.NAME NAME, A.DEPT_NM DEPT_NM, A.ROLL_PSTN ROLL_PSTN, "
						lgStrSQL = lgStrSQL & "       A.PAY_GRD1 PAY_GRD1, A.PAY_GRD2 PAY_GRD2, A.ENTR_DT ENTR_DT, " 
						lgStrSQL = lgStrSQL & "       SUBSTRING(A.RES_NO,1,6) RES_NO1, SUBSTRING(A.RES_NO,7,13) RES_NO2, " 
						lgStrSQL = lgStrSQL & "       A.ZIP_CD ZIP_CD, A.ADDR ADDR, B.SPOUSE SPOUSE, B.LADY LADY, "
						lgStrSQL = lgStrSQL & "       B.SUPP_OLD_CNT SUPP_OLD_CNT, B.SUPP_YOUNG_CNT SUPP_YOUNG_CNT, "
						lgStrSQL = lgStrSQL & "       B.OLD_CNT OLD_CNT,B.OLD_CNT2 OLD_CNT2, B.PARIA_CNT PARIA_CNT, B.CHL_REAR CHL_REAR"  
						lgStrSQL = lgStrSQL & " From HAA010T A, HDF020T B "
						lgStrSQL = lgStrSQL & " Where A.EMP_NO = " & pCode2 	
						lgStrSQL = lgStrSQL & "   And A.EMP_NO = B.EMP_NO"
						lgStrSQL = lgStrSQL & "   And A.INTERNAL_CD LIKE " & FilterVar("%", "''", "S") & "" 


                    Case "P"
                         lgStrSQL =            "Select TOP 1 A.EMP_NO EMP_NO, A.NAME NAME, A.DEPT_NM DEPT_NM, A.ROLL_PSTN ROLL_PSTN, "
                         lgStrSQL = lgStrSQL & "       A.PAY_GRD1 PAY_GRD1, A.PAY_GRD2 PAY_GRD2, A.ENTR_DT ENTR_DT, " 
                         lgStrSQL = lgStrSQL & "       SUBSTRING(A.RES_NO,1,6) RES_NO1, SUBSTRING(A.RES_NO,7,13) RES_NO2, " 
                         lgStrSQL = lgStrSQL & "       A.ZIP_CD ZIP_CD, A.ADDR ADDR, B.SPOUSE SPOUSE, B.LADY LADY, "
                         lgStrSQL = lgStrSQL & "       B.SUPP_OLD_CNT SUPP_OLD_CNT, B.SUPP_YOUNG_CNT SUPP_YOUNG_CNT, "
                         lgStrSQL = lgStrSQL & "       B.OLD_CNT OLD_CNT, B.PARIA_CNT PARIA_CNT, B.CHL_REAR CHL_REAR"  
                         lgStrSQL = lgStrSQL & "  From HAA010T A, HDF020T B "
                         lgStrSQL = lgStrSQL & " Where A.EMP_NO < " & pCode2 	
                         lgStrSQL = lgStrSQL & "   And A.EMP_NO = B.EMP_NO"
                         lgStrSQL = lgStrSQL & "   And A.INTERNAL_CD LIKE " & FilterVar("%", "''", "S") & "" 
                         lgStrSQL = lgStrSQL & "   And (A.RETIRE_RESN IS NULL OR "
                         lgStrSQL = lgStrSQL & "        A.RETIRE_RESN = '' OR A.RETIRE_RESN = " & FilterVar("6", "''", "S") & ") "                         
                         lgStrSQL = lgStrSQL & " ORDER BY A.EMP_NO DESC"
                    Case "N"
                         lgStrSQL =            "Select TOP 1 A.EMP_NO EMP_NO, A.NAME NAME, A.DEPT_NM DEPT_NM, A.ROLL_PSTN ROLL_PSTN, "
                         lgStrSQL = lgStrSQL & "       A.PAY_GRD1 PAY_GRD1, A.PAY_GRD2 PAY_GRD2, A.ENTR_DT ENTR_DT, " 
                         lgStrSQL = lgStrSQL & "       SUBSTRING(A.RES_NO,1,6) RES_NO1, SUBSTRING(A.RES_NO,7,13) RES_NO2, " 
                         lgStrSQL = lgStrSQL & "       A.ZIP_CD ZIP_CD, A.ADDR ADDR, B.SPOUSE SPOUSE, B.LADY LADY, "
                         lgStrSQL = lgStrSQL & "       B.SUPP_OLD_CNT SUPP_OLD_CNT, B.SUPP_YOUNG_CNT SUPP_YOUNG_CNT, "
                         lgStrSQL = lgStrSQL & "       B.OLD_CNT OLD_CNT, B.PARIA_CNT PARIA_CNT, B.CHL_REAR CHL_REAR"  
                         lgStrSQL = lgStrSQL & "  From HAA010T A, HDF020T B "
                         lgStrSQL = lgStrSQL & " Where A.EMP_NO > " & pCode2 	
                         lgStrSQL = lgStrSQL & "   And A.EMP_NO = B.EMP_NO"
                         lgStrSQL = lgStrSQL & "   And A.INTERNAL_CD LIKE " & FilterVar("%", "''", "S") & "" 
                         lgStrSQL = lgStrSQL & "   And (A.RETIRE_RESN IS NULL OR "
                         lgStrSQL = lgStrSQL & "        A.RETIRE_RESN = '' OR A.RETIRE_RESN = " & FilterVar("6", "''", "S") & ") "                         
                         lgStrSQL = lgStrSQL & " ORDER BY A.EMP_NO ASC"
            End Select

      Case "C"
             Select Case  lgPrevNext 
	                Case ""

                         lgStrSQL =            "Select OTHER_INCOME, FORE_INCOME, EDU_SPPORT, MED_SPPORT, OTHER_INSUR, DISABLED_SUB_AMT," 
			             lgStrSQL = lgStrSQL & "       MED_INSUR, EMP_INSUR, TOT_MED, SPECI_MED, PER_EDU, FAM_EDU, "
                         lgStrSQL = lgStrSQL & "       UNIV_EDU, KIND_EDU, KIND_EDU_CNT, UNIV_EDU_CNT, LEGAL_CONTR, APP_CONTR,TaxLaw_contr_amt, "             
                         lgStrSQL = lgStrSQL & "       PRIV_CONTR, HOUSE_FUND, INDIV_ANU, SAVE_TAX_SUB, HOUSE_REPAY, "
                         lgStrSQL = lgStrSQL & "       STOCK_SAVE,  FORE_PAY, INCOME_REDU, TAXES_REDU, INDIV_ANU2, "
                         lgStrSQL = lgStrSQL & "       INVEST_SUB_AMT, VENTURE_SUB_AMT,invest2_sub_amt, CARD_USE_AMT, FAM_EDU_CNT,  "
                         lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT, AFTER_BONUS_AMT, NATIONAL_PENSION_AMT "
						 lgStrSQL = lgStrSQL & "		,  disabled_edu_amt,our_stock_amt,FORE_EDU_AMT,CARD2_USE_AMT,"                                                  
						 lgStrSQL = lgStrSQL & "		CEREMONY_CNT,CEREMONY_AMT,POLI_CONTRA_AMT1,POLI_CONTRA_AMT2,OURSTOCK_CONTRA_AMT,LONG_HOUSE_LOAN_AMT1 "   '2004년 
                         lgStrSQL = lgStrSQL & " From  HFA030T "
                         lgStrSQL = lgStrSQL & " Where YY     = " & pCode1 	
                         lgStrSQL = lgStrSQL & "   And EMP_NO = " & pCode2 	

                     Case "P"
                         lgStrSQL =            "Select OTHER_INCOME, FORE_INCOME, EDU_SPPORT, MED_SPPORT, OTHER_INSUR, DISABLED_SUB_AMT," 
			             lgStrSQL = lgStrSQL & "       MED_INSUR, EMP_INSUR, TOT_MED, SPECI_MED, PER_EDU, FAM_EDU, "
                         lgStrSQL = lgStrSQL & "       UNIV_EDU, KIND_EDU, KIND_EDU_CNT, UNIV_EDU_CNT, LEGAL_CONTR, APP_CONTR,TaxLaw_contr_amt, "             
                         lgStrSQL = lgStrSQL & "       PRIV_CONTR, HOUSE_FUND, INDIV_ANU, SAVE_TAX_SUB, HOUSE_REPAY, "
                         lgStrSQL = lgStrSQL & "       STOCK_SAVE,  FORE_PAY, INCOME_REDU, TAXES_REDU, INDIV_ANU2, "
                         lgStrSQL = lgStrSQL & "       INVEST_SUB_AMT, VENTURE_SUB_AMT,invest2_sub_amt, CARD_USE_AMT, FAM_EDU_CNT,  "
                         lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT, AFTER_BONUS_AMT, NATIONAL_PENSION_AMT "
						 lgStrSQL = lgStrSQL & "		,  disabled_edu_amt,our_stock_amt,FORE_EDU_AMT,CARD2_USE_AMT,"                                                 
						 lgStrSQL = lgStrSQL & "		CEREMONY_CNT,CEREMONY_AMT,POLI_CONTRA_AMT1,POLI_CONTRA_AMT2,OURSTOCK_CONTRA_AMT,LONG_HOUSE_LOAN_AMT1 "   '2004년 
                         lgStrSQL = lgStrSQL & " From  HFA030T "
                         lgStrSQL = lgStrSQL & " Where YY     = " & pCode1 	
                         lgStrSQL = lgStrSQL & "   And EMP_NO = " & pCode2 	
                     Case "N"
                         lgStrSQL =            "Select OTHER_INCOME, FORE_INCOME, EDU_SPPORT, MED_SPPORT, OTHER_INSUR, DISABLED_SUB_AMT, " 
			             lgStrSQL = lgStrSQL & "       MED_INSUR, EMP_INSUR, TOT_MED, SPECI_MED, PER_EDU, FAM_EDU, "
                         lgStrSQL = lgStrSQL & "       UNIV_EDU, KIND_EDU, KIND_EDU_CNT, UNIV_EDU_CNT, LEGAL_CONTR, APP_CONTR, TaxLaw_contr_amt,"             
                         lgStrSQL = lgStrSQL & "       PRIV_CONTR, HOUSE_FUND, INDIV_ANU, SAVE_TAX_SUB, HOUSE_REPAY, "
                         lgStrSQL = lgStrSQL & "       STOCK_SAVE,  FORE_PAY, INCOME_REDU, TAXES_REDU, INDIV_ANU2, "
                         lgStrSQL = lgStrSQL & "       INVEST_SUB_AMT, VENTURE_SUB_AMT, invest2_sub_amt,CARD_USE_AMT, FAM_EDU_CNT,  "
                         lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT, AFTER_BONUS_AMT, NATIONAL_PENSION_AMT "
						 lgStrSQL = lgStrSQL & "		,  disabled_edu_amt,our_stock_amt,FORE_EDU_AMT,CARD2_USE_AMT,"     
						 lgStrSQL = lgStrSQL & "		CEREMONY_CNT,CEREMONY_AMT,POLI_CONTRA_AMT1,POLI_CONTRA_AMT2,OURSTOCK_CONTRA_AMT,LONG_HOUSE_LOAN_AMT1 "   '2004년 
                         lgStrSQL = lgStrSQL & " From  HFA030T "
                         lgStrSQL = lgStrSQL & " Where YY     = " & pCode1 	
                         lgStrSQL = lgStrSQL & "   And EMP_NO = " & pCode2 	
            End Select

     Case "U"
             lgStrSQL =            "Select OTHER_INCOME, FORE_INCOME, EDU_SPPORT, MED_SPPORT, OTHER_INSUR,DISABLED_SUB_AMT, " 
             lgStrSQL = lgStrSQL & "       MED_INSUR, EMP_INSUR, TOT_MED, SPECI_MED, PER_EDU, FAM_EDU, "
             lgStrSQL = lgStrSQL & "       UNIV_EDU, KIND_EDU, KIND_EDU_CNT, UNIV_EDU_CNT, LEGAL_CONTR, APP_CONTR, TaxLaw_contr_amt,"             
             lgStrSQL = lgStrSQL & "       PRIV_CONTR, HOUSE_FUND, INDIV_ANU, SAVE_TAX_SUB, HOUSE_REPAY, "
             lgStrSQL = lgStrSQL & "       STOCK_SAVE,  FORE_PAY, INCOME_REDU, TAXES_REDU, INDIV_ANU2, "
             lgStrSQL = lgStrSQL & "       INVEST_SUB_AMT, VENTURE_SUB_AMT,invest2_sub_amt, CARD_USE_AMT, FAM_EDU_CNT,  "
             lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT, AFTER_BONUS_AMT, NATIONAL_PENSION_AMT "
			 lgStrSQL = lgStrSQL & "		,  disabled_edu_amt,our_stock_amt,FORE_EDU_AMT,CARD2_USE_AMT,"                                      
			 lgStrSQL = lgStrSQL & "		CEREMONY_CNT,CEREMONY_AMT,POLI_CONTRA_AMT1,POLI_CONTRA_AMT2,OURSTOCK_CONTRA_AMT,LONG_HOUSE_LOAN_AMT1 "   '2004년 
             lgStrSQL = lgStrSQL & " From  HFA030T "
             lgStrSQL = lgStrSQL & " Where YY     = " & pCode1 	
             lgStrSQL = lgStrSQL & "   And EMP_NO = " & pCode2 	
      Case "D"
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
            Parent.DBSaveFail
             'Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
