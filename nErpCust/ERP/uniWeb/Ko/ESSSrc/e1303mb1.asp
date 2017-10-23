<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<%
 
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd_uniSIMS

    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgPrevNext        = Request("txtPrevNext")
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
   
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
            If lgKeyStream(3) = "Q" Then
				Call SubBizQuery()
			
			Elseif lgKeyStream(3) = "P" Then
				Call SubBizQuery2()
			Else 
				Call SubBizQuery()
			End if
        
        Case "UID_M0002"             
             Call SubBizSave()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim strYear
    Dim strEmpNo  

    Err.Clear                                                                        '☜: Clear Error status

    strEmpNo  = lgKeyStream(0)
    strYear   = lgKeyStream(2)

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

    lgStrSQL = "SELECT * FROM HFA031T WHERE EMP_NO = " & FilterVar(Emp_no, "''", "S") & " AND YEAR_YY = " & FilterVar(strYear, "''", "S")
 
    If 	FncOpenRs("C",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then          'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
          Call SetErrorStatus()
       End If
    Else
%>
<Script Language=vbscript>
       With Parent
            .Frm1.PAY_TAX_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("PAY_TAX_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.BONUS_TAX_AMT.Value           = "<%=UNINumClientFormat(lgObjRs("BONUS_TAX_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.AFTER_BONUS_AMT.Value         = "<%=UNINumClientFormat(lgObjRs("AFTER_BONUS_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.other_income.Value				= "<%=UNINumClientFormat(lgObjRs("other_income"), ggAmtOfMoney.DecPoint, 0)%>"            
            .Frm1.BEFORE_INCOME_TAX_AMT.Value   = "<%=UNINumClientFormat(lgObjRs("BEFORE_INCOME_TAX_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.BEFORE_RES_TAX_AMT.Value      = "<%=UNINumClientFormat(lgObjRs("BEFORE_RES_TAX_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.OLD_SUPP_CNT.Value            = "<%=ConvSPChars(lgObjRs("OLD_SUPP_CNT"))%>"
            .Frm1.YOUNG_SUPP_CNT.Value          = "<%=ConvSPChars(lgObjRs("YOUNG_SUPP_CNT"))%>"
            .Frm1.OLD_CNT1.Value                 = "<%=ConvSPChars(lgObjRs("OLD_CNT"))%>"
            .Frm1.OLD_CNT2.Value                 = "<%=ConvSPChars(lgObjRs("OLD_CNT2"))%>"            
            .Frm1.PARIA_CNT.Value               = "<%=ConvSPChars(lgObjRs("PARIA_CNT"))%>"

            If "<%=lgObjRs("SPOUSE")%>" = "Y" Then
				.Frm1.SPOUSE.checked = true
            Else
				.Frm1.SPOUSE.checked = false
			End If 
			
            If "<%=lgObjRs("LADY")%>" = "Y" Then
				.Frm1.LADY.checked = true
            Else
				.Frm1.LADY.checked = false
			End If 
						            
            If "<%=lgObjRs("Foreign_separate_tax_yn")%>" = "Y" Then
				.Frm1.txtForeign_separate_tax_yn.checked = true
            Else
				.Frm1.txtForeign_separate_tax_yn.checked = false
			End If
			
            .Frm1.CHL_REAR_CNT.Value            = "<%=ConvSPChars(lgObjRs("CHL_REAR_CNT"))%>"
            .Frm1.MED_INSUR.Value               = "<%=UNINumClientFormat(lgObjRs("MED_INSUR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.DISABLED_INSUR.Value          = "<%=UNINumClientFormat(lgObjRs("DISABLED_SUB_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.EMP_INSUR.Value               = "<%=UNINumClientFormat(lgObjRs("EMP_INSUR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.OTHER_INSUR.Value             = "<%=UNINumClientFormat(lgObjRs("OTHER_INSUR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.MED_SPPORT.Value              = "<%=UNINumClientFormat(lgObjRs("MED_SPPORT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.SPECI_MED.Value               = "<%=UNINumClientFormat(lgObjRs("SPECI_MED"), ggAmtOfMoney.DecPoint, 0)%>"
           
            .Frm1.PER_EDU.Value                 = "<%=UNINumClientFormat(lgObjRs("PER_EDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FAMILY1_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FAMILY1_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FAMILY2_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FAMILY2_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FAMILY3_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FAMILY3_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FAMILY4_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FAMILY4_AMT"), ggAmtOfMoney.DecPoint, 0)%>"

            .Frm1.FORE_EDU_AMT.Value			= "<%=UNINumClientFormat(lgObjRs("FORE_EDU_AMT"), ggAmtOfMoney.DecPoint, 0)%>"            

            .Frm1.FAMILY1_CNT.Value             = "<%=lgObjRs("FAMILY1_CNT")%>"
            .Frm1.FAMILY2_CNT.Value             = "<%=lgObjRs("FAMILY2_CNT")%>"
            .Frm1.FAMILY3_CNT.Value             = "<%=lgObjRs("FAMILY3_CNT")%>"
            .Frm1.FAMILY4_CNT.Value             = "<%=lgObjRs("FAMILY4_CNT")%>"

            .Frm1.HOUSE_FUND.Value              = "<%=UNINumClientFormat(lgObjRs("HOUSE_FUND"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.LONG_HOUSE_LOAN_AMT.Value     = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.LONG_HOUSE_LOAN_AMT1.Value	= "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_AMT1"), ggAmtOfMoney.DecPoint, 0)%>"

            .Frm1.txtLegal_contr_amt.Value		= "<%=UNINumClientFormat(lgObjRs("LEGAL_CONTR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtApp_contr_amt.Value		= "<%=UNINumClientFormat(lgObjRs("APP_CONTR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtTaxLaw_contr_amt.Value		= "<%=UNINumClientFormat(lgObjRs("TaxLaw_contr_amt"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtTaxLaw_contr_amt2.Value	= "<%=UNINumClientFormat(lgObjRs("TaxLaw_contr_amt2"), ggAmtOfMoney.DecPoint, 0)%>"          
            .Frm1.txtCeremony_amt.Value			= "<%=UNINumClientFormat(lgObjRs("CEREMONY_AMT"), ggAmtOfMoney.DecPoint, 0)%>"            
                   
            .Frm1.INDIV_ANU.Value               = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.INDIV_ANU2.Value              = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU2"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.NATIONAL_PENSION_AMT.Value    = "<%=UNINumClientFormat(lgObjRs("NATIONAL_PENSION_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            
            .Frm1.txtinvest2_sub_amt.Value         = "<%=UNINumClientFormat(lgObjRs("invest2_sub_amt"), ggAmtOfMoney.DecPoint, 0)%>"            
            .Frm1.CARD_USE_AMT.Value            = "<%=UNINumClientFormat(lgObjRs("CARD_USE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.CARD2_USE_AMT.Value			= "<%=UNINumClientFormat(lgObjRs("CARD2_USE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtInstitution_giro.Value			= "<%=UNINumClientFormat(lgObjRs("INSTITUTION_GIRO"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtRetire_pension.Value		= "<%=UNINumClientFormat(lgObjRs("Retire_pension"), ggAmtOfMoney.DecPoint, 0)%>"            
                        
            .Frm1.txtOur_stock_amt.Value		= "<%=UNINumClientFormat(lgObjRs("our_stock_amt"), ggAmtOfMoney.DecPoint, 0)%>"            
            .Frm1.HOUSE_REPAY.Value             = "<%=UNINumClientFormat(lgObjRs("HOUSE_REPAY"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FORE_INCOME.Value             = "<%=UNINumClientFormat(lgObjRs("FORE_INCOME"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FORE_PAY.Value                = "<%=UNINumClientFormat(lgObjRs("FORE_PAY"), ggAmtOfMoney.DecPoint, 0)%>"

            .Frm1.INCOME_REDU.Value             = "<%=UNINumClientFormat(lgObjRs("INCOME_REDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.TAXES_REDU.Value              = "<%=UNINumClientFormat(lgObjRs("TAXES_REDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtTax_Union_Ded.Value        = "<%=UNINumClientFormat(lgObjRs("TAX_UNION_DED"), ggAmtOfMoney.DecPoint, 0)%>"	'2005

            .Frm1.txtPoli_contr_amt1.Value		= "<%=UNINumClientFormat(lgObjRs("POLI_CONTRA_AMT1"), ggAmtOfMoney.DecPoint, 0)%>"	'2004
            .Frm1.txtOurstock_contr_amt.Value	= "<%=UNINumClientFormat(lgObjRs("OURSTOCK_CONTRA_AMT"), ggAmtOfMoney.DecPoint, 0)%>"	'2004
			.Frm1.txtCeremony_amt.Value			= "<%=UNINumClientFormat(lgObjRs("CEREMONY_AMT"), ggAmtOfMoney.DecPoint, 0)%>"	'2004 
            .Frm1.txtCeremony_cnt.Value			= "<%=UNINumClientFormat(lgObjRs("CEREMONY_CNT"), ggAmtOfMoney.DecPoint, 0)%>" '2004 

            .Frm1.txtPriv_contr_amt.Value       = "<%=UNINumClientFormat(lgObjRs("UNION_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
       End With
</Script>       
<%     

    End If
   
End Sub    


'============================================================================================================
' Name : SubBizQuery2
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery2()
    Dim strYear
    Dim strEmpNo  
    dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strEmpNo  = lgKeyStream(0)
    strYear   = lgKeyStream(2)

    Call SubEmpBase(lgKeyStream(0),lgKeyStream(1),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
%>
<Script Language=vbscript>
    With parent.frm1
        .txtEmp_no.Value = "<%=emp_no%>"
        .txtName.Value = "<%=Name%>"
        .txtDept_nm.value = "<%=DEPT_NM%>"    
        .txtroll_pstn.value = "<%=roll_pstn%>"
    End With          
</Script>       
<%
    Call SubCreateCommandObject(lgObjComm)

    With lgObjComm
        .CommandText = "usp_hfa031b1"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adChar,adParamInput,Len(gusrID), gusrID)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@year_yy"    ,adChar,adParamInput,Len(strYear), strYear)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@emp_no"     ,adChar,adParamInput,Len(Emp_no), Emp_no)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adChar,adParamOutput,60)

'CREATE procedure usp_hfa031b1 (@usr_id           VARCHAR(13), -- 로그인 ID 
'                               @year_yy          VARCHAR(4),  -- 정산년도 
'                               @emp_no           VARCHAR(13), -- 사번 
'                               @msg_cd           VARCHAR(6)		OUTPUT, -- Error Message Code 
'                               @msg_text         VARCHAR(60)	OUTPUT  -- Error Message Code 
        lgObjComm.Execute ,, adExecuteNoRecords

    End With

   If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        
        if  IntRetCD < 0 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
			Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
            IntRetCD = -1
            lgErrorStatus = "YES"
            Exit Sub
        else
            lgErrorStatus = "NO"
            IntRetCD = 1
        end if
    Else    
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        lgErrorStatus = "YES"       
        IntRetCD = -1
        Exit Sub
    End if

    Call SubCloseCommandObject(lgObjComm)


    lgStrSQL = "SELECT * FROM HFA031T WHERE EMP_NO = " & FilterVar(Emp_no, "''", "S") & " AND YEAR_YY = " & FilterVar(strYear, "''", "S")

    If 	FncOpenRs("C",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then          'If data not exists
       If lgPrevNext = "" Then
          'Call SetErrorStatus()
      
       End If
    Else

%>
<Script Language=vbscript>
       With Parent
            .Frm1.PAY_TAX_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("PAY_TAX_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.BONUS_TAX_AMT.Value           = "<%=UNINumClientFormat(lgObjRs("BONUS_TAX_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.AFTER_BONUS_AMT.Value         = "<%=UNINumClientFormat(lgObjRs("AFTER_BONUS_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.other_income.Value				= "<%=UNINumClientFormat(lgObjRs("other_income"), ggAmtOfMoney.DecPoint, 0)%>"                        
            .Frm1.BEFORE_INCOME_TAX_AMT.Value   = "<%=UNINumClientFormat(lgObjRs("BEFORE_INCOME_TAX_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.BEFORE_RES_TAX_AMT.Value      = "<%=UNINumClientFormat(lgObjRs("BEFORE_RES_TAX_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.OLD_SUPP_CNT.Value            = "<%=lgObjRs("OLD_SUPP_CNT")%>"
            .Frm1.YOUNG_SUPP_CNT.Value          = "<%=lgObjRs("YOUNG_SUPP_CNT")%>"
            .Frm1.OLD_CNT1.Value                 = "<%=lgObjRs("OLD_CNT")%>"
            .Frm1.OLD_CNT2.Value                 = "<%=lgObjRs("OLD_CNT2")%>"
            
            .Frm1.PARIA_CNT.Value               = "<%=lgObjRs("PARIA_CNT")%>"
         
            If "<%=lgObjRs("SPOUSE")%>" = "Y" Then
				.Frm1.SPOUSE.checked = true
            Else
				.Frm1.SPOUSE.checked = false
			End If 
			
            If "<%=lgObjRs("LADY")%>" = "Y" Then
				.Frm1.LADY.checked = true
            Else
				.Frm1.LADY.checked = false
			End If 
						            
            If "<%=lgObjRs("Foreign_separate_tax_yn")%>" = "Y" Then
				.Frm1.txtForeign_separate_tax_yn.checked = true
            Else
				.Frm1.txtForeign_separate_tax_yn.checked = false
			End If                        
            .Frm1.CHL_REAR_CNT.Value            = "<%=lgObjRs("CHL_REAR_CNT")%>"
            .Frm1.MED_INSUR.Value               = "<%=UNINumClientFormat(lgObjRs("MED_INSUR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.DISABLED_INSUR.Value          = "<%=UNINumClientFormat(lgObjRs("DISABLED_SUB_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.EMP_INSUR.Value               = "<%=UNINumClientFormat(lgObjRs("EMP_INSUR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.OTHER_INSUR.Value             = "<%=UNINumClientFormat(lgObjRs("OTHER_INSUR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.MED_SPPORT.Value              = "<%=UNINumClientFormat(lgObjRs("MED_SPPORT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.SPECI_MED.Value               = "<%=UNINumClientFormat(lgObjRs("SPECI_MED"), ggAmtOfMoney.DecPoint, 0)%>"
           
            .Frm1.PER_EDU.Value                 = "<%=UNINumClientFormat(lgObjRs("PER_EDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FAMILY1_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FAMILY1_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FAMILY2_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FAMILY2_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FAMILY3_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FAMILY3_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FAMILY4_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FAMILY4_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FORE_EDU_AMT.Value             = "<%=UNINumClientFormat(lgObjRs("FORE_EDU_AMT"), ggAmtOfMoney.DecPoint, 0)%>"            
           
            .Frm1.FAMILY1_CNT.Value             = "<%=lgObjRs("FAMILY1_CNT")%>"
            .Frm1.FAMILY2_CNT.Value             = "<%=lgObjRs("FAMILY2_CNT")%>"
            .Frm1.FAMILY3_CNT.Value             = "<%=lgObjRs("FAMILY3_CNT")%>"
            .Frm1.FAMILY4_CNT.Value             = "<%=lgObjRs("FAMILY4_CNT")%>"

            .Frm1.HOUSE_FUND.Value              = "<%=UNINumClientFormat(lgObjRs("HOUSE_FUND"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.LONG_HOUSE_LOAN_AMT.Value     = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.LONG_HOUSE_LOAN_AMT1.Value     = "<%=UNINumClientFormat(lgObjRs("LONG_HOUSE_LOAN_AMT1"), ggAmtOfMoney.DecPoint, 0)%>"

            .Frm1.txtLegal_contr_amt.Value             = "<%=UNINumClientFormat(lgObjRs("LEGAL_CONTR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtApp_contr_amt.Value               = "<%=UNINumClientFormat(lgObjRs("APP_CONTR"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtTaxLaw_contr_amt.Value		= "<%=UNINumClientFormat(lgObjRs("TaxLaw_contr_amt"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtTaxLaw_contr_amt2.Value	= "<%=UNINumClientFormat(lgObjRs("TaxLaw_contr_amt2"), ggAmtOfMoney.DecPoint, 0)%>"    
            .Frm1.txtCeremony_amt.Value			= "<%=UNINumClientFormat(lgObjRs("CEREMONY_AMT"), ggAmtOfMoney.DecPoint, 0)%>"            
                      
            .Frm1.INDIV_ANU2.Value              = "<%=UNINumClientFormat(lgObjRs("INDIV_ANU2"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.NATIONAL_PENSION_AMT.Value    = "<%=UNINumClientFormat(lgObjRs("NATIONAL_PENSION_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            
            .Frm1.txtinvest2_sub_amt.Value      = "<%=UNINumClientFormat(lgObjRs("invest2_sub_amt"), ggAmtOfMoney.DecPoint, 0)%>"            
            .Frm1.CARD_USE_AMT.Value            = "<%=UNINumClientFormat(lgObjRs("CARD_USE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.CARD2_USE_AMT.Value            = "<%=UNINumClientFormat(lgObjRs("CARD2_USE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"            
            .Frm1.txtInstitution_giro.Value      = "<%=UNINumClientFormat(lgObjRs("INSTITUTION_GIRO"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtRetire_pension.Value      = "<%=UNINumClientFormat(lgObjRs("Retire_pension"), ggAmtOfMoney.DecPoint, 0)%>"        
	          
            .Frm1.txtOur_stock_amt.Value       = "<%=UNINumClientFormat(lgObjRs("our_stock_amt"), ggAmtOfMoney.DecPoint, 0)%>"                        
            .Frm1.HOUSE_REPAY.Value             = "<%=UNINumClientFormat(lgObjRs("HOUSE_REPAY"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FORE_INCOME.Value             = "<%=UNINumClientFormat(lgObjRs("FORE_INCOME"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.FORE_PAY.Value                = "<%=UNINumClientFormat(lgObjRs("FORE_PAY"), ggAmtOfMoney.DecPoint, 0)%>"

            .Frm1.INCOME_REDU.Value             = "<%=UNINumClientFormat(lgObjRs("INCOME_REDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.TAXES_REDU.Value              = "<%=UNINumClientFormat(lgObjRs("TAXES_REDU"), ggAmtOfMoney.DecPoint, 0)%>"
            .Frm1.txtTax_Union_Ded.Value        = "<%=UNINumClientFormat(lgObjRs("TAX_UNION_DED"), ggAmtOfMoney.DecPoint, 0)%>" '2005

            .Frm1.txtPoli_contr_amt1.Value     = "<%=UNINumClientFormat(lgObjRs("POLI_CONTRA_AMT1"), ggAmtOfMoney.DecPoint, 0)%>"	'2004
            .Frm1.txtOurstock_contr_amt.Value	= "<%=UNINumClientFormat(lgObjRs("OURSTOCK_CONTRA_AMT"), ggAmtOfMoney.DecPoint, 0)%>"	'2004
			.Frm1.txtCeremony_amt.Value			= "<%=UNINumClientFormat(lgObjRs("CEREMONY_AMT"), ggAmtOfMoney.DecPoint, 0)%>"	'2004 
            .Frm1.txtCeremony_cnt.Value			= "<%=UNINumClientFormat(lgObjRs("CEREMONY_CNT"), ggAmtOfMoney.DecPoint, 0)%>" '2004 
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
    Dim ceremonyCnt , ceremonyAmt

    Dim Emp_no

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


    Emp_no = lgKeyStream(0)
'   DB Save
    lgStrSQL =            "UPDATE HFA031T"
    lgStrSQL = lgStrSQL & "   SET PAY_TAX_AMT              = " & UNIConvNum(Request("PAY_TAX_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       BONUS_TAX_AMT            = " & UNIConvNum(Request("BONUS_TAX_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       AFTER_BONUS_AMT          = " & UNIConvNum(Request("AFTER_BONUS_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       BEFORE_INCOME_TAX_AMT    = " & UNIConvNum(Request("BEFORE_INCOME_TAX_AMT"), 0)      & ","
    lgStrSQL = lgStrSQL & "       BEFORE_RES_TAX_AMT       = " & UNIConvNum(Request("BEFORE_RES_TAX_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       other_income				 = " & UNIConvNum(Request("other_income"),0)      & ","
    
    If IsEmpty(Request("FOREIGN_SEPARATE_TAX_YN")) = true Then
        lgStrSQL = lgStrSQL & " FOREIGN_SEPARATE_TAX_YN = " & FilterVar("N", "''", "S") & ","
    ELSE
        lgStrSQL = lgStrSQL & " FOREIGN_SEPARATE_TAX_YN = " & FilterVar("Y", "''", "S") & ","
    END IF
        
    If IsEmpty(Request("SPOUSE")) = true Then
        lgStrSQL = lgStrSQL & " SPOUSE = " & FilterVar("N", "''", "S") & ","
    ELSE
        lgStrSQL = lgStrSQL & " SPOUSE = " & FilterVar("Y", "''", "S") & ","
    END IF
    'lgStrSQL = lgStrSQL & "       SPOUSE                   = '" & Request("SPOUSE")      & "',"
    lgStrSQL = lgStrSQL & "       OLD_SUPP_CNT             = " & UNIConvNum(Request("OLD_SUPP_CNT"),0)      & ","
    lgStrSQL = lgStrSQL & "       YOUNG_SUPP_CNT           = " & UNIConvNum(Request("YOUNG_SUPP_CNT"),0)      & ","
    lgStrSQL = lgStrSQL & "       OLD_CNT                  = " & UNIConvNum(Request("OLD_CNT1"),0)      & ","
    lgStrSQL = lgStrSQL & "       OLD_CNT2                  = " & UNIConvNum(Request("OLD_CNT2"),0)      & ","    
    lgStrSQL = lgStrSQL & "       PARIA_CNT                = " & UNIConvNum(Request("PARIA_CNT"),0)      & ","
    If IsEmpty(Request("LADY")) = true Then
        lgStrSQL = lgStrSQL & " LADY = " & FilterVar("N", "''", "S") & ","
    ELSE
        lgStrSQL = lgStrSQL & " LADY = " & FilterVar("Y", "''", "S") & ","
    END IF
    'lgStrSQL = lgStrSQL & "       LADY                     = " &  Request("LADY")      & ","
    lgStrSQL = lgStrSQL & "       CHL_REAR_CNT             = " & UNIConvNum(Request("CHL_REAR_CNT"),0)      & ","
    lgStrSQL = lgStrSQL & "       MED_INSUR                = " & UNIConvNum(Request("MED_INSUR"),0)      & ","
    lgStrSQL = lgStrSQL & "       DISABLED_SUB_AMT         = " & UNIConvNum(Request("DISABLED_INSUR"),0)      & ","
    lgStrSQL = lgStrSQL & "       EMP_INSUR                = " & UNIConvNum(Request("EMP_INSUR"),0)      & ","
    lgStrSQL = lgStrSQL & "       OTHER_INSUR              = " & UNIConvNum(Request("OTHER_INSUR"),0)      & ","
    lgStrSQL = lgStrSQL & "       MED_SPPORT               = " & UNIConvNum(Request("MED_SPPORT"),0)      & ","
    lgStrSQL = lgStrSQL & "       SPECI_MED                = " & UNIConvNum(Request("SPECI_MED"),0)      & ","
    lgStrSQL = lgStrSQL & "       PER_EDU                  = " & UNIConvNum(Request("PER_EDU"),0)      & ","

    lgStrSQL = lgStrSQL & "       FAMILY1_CNT             = " & UNIConvNum(Request("FAMILY1_CNT"),0)      & ","
    lgStrSQL = lgStrSQL & "       FAMILY2_CNT             = " & UNIConvNum(Request("FAMILY2_CNT"),0)      & ","
    lgStrSQL = lgStrSQL & "       FAMILY3_CNT             = " & UNIConvNum(Request("FAMILY3_CNT"),0)      & ","
    lgStrSQL = lgStrSQL & "       FAMILY4_CNT             = " & UNIConvNum(Request("FAMILY4_CNT"),0)      & ","

    lgStrSQL = lgStrSQL & "       FAMILY1_AMT              = " & UNIConvNum(Request("FAMILY1_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       FAMILY2_AMT              = " & UNIConvNum(Request("FAMILY2_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       FAMILY3_AMT              = " & UNIConvNum(Request("FAMILY3_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       FAMILY4_AMT              = " & UNIConvNum(Request("FAMILY4_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       FORE_EDU_AMT              = " & UNIConvNum(Request("FORE_EDU_AMT"),0)      & ","    
    
    lgStrSQL = lgStrSQL & "       HOUSE_FUND               = " & UNIConvNum(Request("HOUSE_FUND"),0)      & ","
    lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT      = " & UNIConvNum(Request("LONG_HOUSE_LOAN_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       LONG_HOUSE_LOAN_AMT1      = " & UNIConvNum(Request("LONG_HOUSE_LOAN_AMT1"),0)      & ","
    
    lgStrSQL = lgStrSQL & "       LEGAL_CONTR              = " & UNIConvNum(Request("txtLegal_contr_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       APP_CONTR                = " & UNIConvNum(Request("txtApp_contr_amt"),0)      & ","
    lgStrSQL = lgStrSQL & "       TaxLaw_contr_amt         = " & UNIConvNum(Request("txtTaxLaw_contr_amt"),0)      & ","   
    lgStrSQL = lgStrSQL & "       TaxLaw_contr_amt2         = " & UNIConvNum(Request("txtTaxLaw_contr_amt2"),0)      & ","   
     
    lgStrSQL = lgStrSQL & "       CEREMONY_CNT           = " & UNIConvNum(Request("txtCeremony_cnt"),0)     & ","	'2004
 
    call CommonQueryRs(" CEREMONY_AMT "," HFA020T "," 1=1",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ceremonyAmt = Replace(lgF0, Chr(11), "")
	ceremonyCnt = UNIConvNum(Request("txtCeremony_cnt"),0)

	ceremonyAmt = ceremonyAmt  * ceremonyCnt
    
    lgStrSQL = lgStrSQL & "       CEREMONY_AMT         = " & UNIConvNum(ceremonyAmt,0)  & ","   '2004
    
    lgStrSQL = lgStrSQL & "       POLI_CONTRA_AMT1     = " & UNIConvNum(Request("txtPoli_contr_amt1"),0)     & ","   '2004
    lgStrSQL = lgStrSQL & "       OURSTOCK_CONTRA_AMT     = " & UNIConvNum(Request("txtOurstock_contr_amt"),0)     & ","   '2004

    lgStrSQL = lgStrSQL & "       UNION_AMT					= " & UNIConvNum(Request("txtPriv_contr_amt"),0)       & ","
    lgStrSQL = lgStrSQL & "       INDIV_ANU                = " & UNIConvNum(Request("INDIV_ANU"),0)      & ","
    lgStrSQL = lgStrSQL & "       INDIV_ANU2               = " & UNIConvNum(Request("INDIV_ANU2"),0)      & ","
    lgStrSQL = lgStrSQL & "       NATIONAL_PENSION_AMT     = " & UNIConvNum(Request("NATIONAL_PENSION_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       invest2_sub_amt          = " & UNIConvNum(Request("txtinvest2_sub_amt"),0)      & ","    
    lgStrSQL = lgStrSQL & "       CARD_USE_AMT             = " & UNIConvNum(Request("CARD_USE_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       CARD2_USE_AMT             = " & UNIConvNum(Request("CARD2_USE_AMT"),0)      & ","
    lgStrSQL = lgStrSQL & "       INSTITUTION_GIRO			= " & UNIConvNum(Request("txtInstitution_giro"),0)      & ","    
    lgStrSQL = lgStrSQL & "       Retire_pension			= " & UNIConvNum(Request("txtRetire_pension"),0)      & ","    
        
	lgStrSQL = lgStrSQL & "       our_stock_amt             = " & UNIConvNum(Request("txtOur_stock_amt"),0)      & ","    
    lgStrSQL = lgStrSQL & "       HOUSE_REPAY              = " & UNIConvNum(Request("HOUSE_REPAY"),0)      & ","
    lgStrSQL = lgStrSQL & "       FORE_INCOME              = " & UNIConvNum(Request("FORE_INCOME"),0)      & ","
    lgStrSQL = lgStrSQL & "       FORE_PAY                 = " & UNIConvNum(Request("FORE_PAY"),0)         & ","

    lgStrSQL = lgStrSQL & "       INCOME_REDU              = " & UNIConvNum(Request("INCOME_REDU"),0)     & ","
    lgStrSQL = lgStrSQL & "       TAXES_REDU               = " & UNIConvNum(Request("TAXES_REDU"),0)      & ","
    lgStrSQL = lgStrSQL & "       TAX_UNION_DED               = " & UNIConvNum(Request("txtTax_Union_Ded"),0)   
    lgStrSQL = lgStrSQL & " WHERE EMP_NO         = " & FilterVar(Emp_no, "''", "S") &  " AND YEAR_YY = " & FilterVar(lgKeyStream(2), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)

    return

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
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)    'Can not create(Demo code)
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
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
