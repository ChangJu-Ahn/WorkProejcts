<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd_uniSIMS
                                                               '¢Ð: Hide Processing message
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '¢Ð: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '¢Ð: Save,Update
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim strEmpNo  
    Dim strYear  
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strEmpNo  = lgKeyStream(0)
	strYear   = lgKeyStream(1)
	
    Call SubEmpBase(lgKeyStream(0),lgKeyStream(2),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
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
    
    Call SubCreateCommandObject(lgObjComm)

    With lgObjComm
        .CommandText = "usp_hfa051b1"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adXChar,adParamInput,Len(gusrID), gusrID)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@year_yy"    ,adXChar,adParamInput,Len(strYear), strYear)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@emp_no"     ,adXChar,adParamInput,Len(Emp_no), Emp_no)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)

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

    
    
    if emp_no = "" then
        return
    end if 

    strEmpNo  = emp_no

    Call SubMakeSQLStatements("R","")                                       '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       'If lgPrevNext = "" Then
          'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
          'Call SetErrorStatus()
       'End If
    Else
%>
<Script Language=vbscript>
       Dim DEC_AMT
	            
       With Parent	
            
            .Frm1.txtincome_tot_amt.Value	        = "<%=UNINumClientFormat(lgObjRs("txtincome_tot_amt"), 0,0)%>"
            .Frm1.txtincome_sub_amt.Value		    = "<%=UNINumClientFormat(lgObjRs("INCOME_SUB"), 0,0)%>"
            .Frm1.txtIncome_amt.Value				= "<%=UNINumClientFormat(lgObjRs("INCOME_AMT"), 0,0)%>"
            
            .frm1.txtper_sub_amt.value              = "<%=UNINumClientFormat(lgObjRs("PER_SUB"), 0,0)%>"         
            .frm1.txtspouse_sub_amt.value           = "<%=UNINumClientFormat(lgObjRs("SPOUSE_SUB"), 0,0)%>"          
            .frm1.txtsupp_old_cnt.value             = "<%=ConvSPChars(lgObjRs("SUPP_CNT"))%>"    
            .frm1.txtsupp_sub_amt.value             = "<%=UNINumClientFormat(lgObjRs("SUPP_SUB"), 0,0)%>"        
            
            .frm1.txtold_cnt1.value					= "<%=ConvSPChars(lgObjRs("OLD_CNT"))%>"  
            .frm1.txtold_cnt2.value					= "<%=ConvSPChars(lgObjRs("OLD_CNT2"))%>"    
              
            .frm1.txtold_sub_amt1.value				= "<%=UNINumClientFormat(lgObjRs("OLD_SUB"), 0,0)%>"    
                 
            .frm1.txtparia_cnt.value				= "<%=ConvSPChars(lgObjRs("PARIA_CNT"))%>"    
            .frm1.txtparia_sub_amt.value			= "<%=UNINumClientFormat(lgObjRs("PARIA_SUB"), 0,0)%>"         
            .frm1.txtlady_sub_amt.value             = "<%=UNINumClientFormat(lgObjRs("LADY_SUB"), 0,0)%>"    
            .frm1.txtchl_rear.value					= "<%=ConvSPChars(lgObjRs("CHL_REAR"))%>"    
            .frm1.txtchl_rear_sub_amt.value         = "<%=UNINumClientFormat(lgObjRs("CHL_REAR_SUB"), 0,0)%>"    
            
            .frm1.txtsmall_sub_amt.value			= "<%=UNINumClientFormat(lgObjRs("SMALL_SUB"), 0,0)%>"    
            
            .frm1.txtInsur_sub_amt.value			= "<%=UNINumClientFormat(lgObjRs("INSUR_SUB"), 0,0)%>"    
            .frm1.txtMed_sub_amt.value				= "<%=UNINumClientFormat(lgObjRs("MED_SUB"), 0,0)%>"    
            .frm1.txtEdu_amt.value					= "<%=UNINumClientFormat(lgObjRs("EDU_SUB"), 0,0)%>"    
            .frm1.txth_house_fund_amt.value         = "<%=UNINumClientFormat(lgObjRs("HOUSE_FUND"), 0,0)%>"    
            .frm1.txtcontr_sub_amt.value			= "<%=UNINumClientFormat(lgObjRs("CONTR_SUB"), 0,0)%>"    
            .frm1.txtCeremony_amt.value				= "<%=UNINumClientFormat(lgObjRs("CEREMONY_AMT"), 0,0)%>"    
            
            .frm1.txtstd_sub_amt.value				= "<%=UNINumClientFormat(lgObjRs("STD_SUB"), 0,0)%>"    
            
            .frm1.txtSub_income_amt.value           = "<%=UNINumClientFormat(lgObjRs("SUB_INCOME_AMT"), 0,0)%>"    
            .frm1.txtIndiv_anu_amt.value			= "<%=UNINumClientFormat(lgObjRs("INDIV_ANU"), 0,0)%>"    
            .frm1.txtIndiv_anu2_amt.value			= "<%=UNINumClientFormat(lgObjRs("INDIV_ANU2"), 0,0)%>"    
            .frm1.txtInvest_sub_sum_amt.value		= "<%=UNINumClientFormat(lgObjRs("INVEST_SUB_SUM"), 0,0)%>"
            .frm1.txtOur_stock_amt.value			= "<%=UNINumClientFormat(lgObjRs("our_stock_amt"), 0,0)%>"                    
            .frm1.txtcard_sub_sum_amt.value			= "<%=UNINumClientFormat(lgObjRs("CARD_SUB_SUM"), 0,0)%>"
            .frm1.txtRetire_pension.value			= "<%=UNINumClientFormat(lgObjRs("Retire_pension"), 0,0)%>"
            .frm1.txtFore_edu_amt.value				= "<%=UNINumClientFormat(lgObjRs("FORE_EDU_SUB_AMT"), 0,0)%>"                    
            .frm1.txtTax_std_amt.value				= "<%=UNINumClientFormat(lgObjRs("TAX_STD"), 0,0)%>"    
            .frm1.txtCalu_tax_amt.value				= "<%=UNINumClientFormat(lgObjRs("CALU_TAX"), 0,0)%>"    
            
            
            .frm1.txtincome_tax_sub_amt.value		= "<%=UNINumClientFormat(lgObjRs("INCOME_TAX_SUB"), 0,0)%>"    
            .frm1.txthouse_repay_amt.value			= "<%=UNINumClientFormat(lgObjRs("HOUSE_REPAY"), 0,0)%>"    
            .frm1.txtFore_pay_amt.value             = "<%=UNINumClientFormat(lgObjRs("FORE_PAY"), 0,0)%>"    
            .frm1.txtPolicontr_tax_sub_amt.value	= "<%=UNINumClientFormat(lgObjRs("poli_tax_sub"), 0,0)%>"    
            
            .frm1.txttax_sub_sum_amt.value			= "<%=UNINumClientFormat(lgObjRs("TAX_SUB_SUM"), 0,0)%>"    
            .frm1.txtRedu_sum_amt.value				= "<%=UNINumClientFormat(lgObjRs("REDU_SUM"), 0,0)%>"  
            .frm1.txtTax_Union_Ded.value				= "<%=UNINumClientFormat(lgObjRs("TAX_UNION_DED"), 0,0)%>"  '2005  
            
            .frm1.txtDec_income_tax_amt.value		= "<%=UNINumClientFormat(lgObjRs("DEC_INCOME_TAX"), 0,0)%>"    
            .frm1.txtDec_res_tax_amt.value			= "<%=UNINumClientFormat(lgObjRs("DEC_RES_TAX"), 0,0)%>"    
            .frm1.txtDec_farm_tax_amt.value         = "<%=UNINumClientFormat(lgObjRs("DEC_FARM_TAX"), 0,0)%>"    
            .frm1.txtdec_amt.value					= "<%=UNINumClientFormat(lgObjRs("txtdec_amt"), 0,0)%>"    
            .frm1.txtold_income_tax_amt.value       = "<%=UNINumClientFormat(lgObjRs("BEFORE_INCOME_TAX"), 0,0)%>"    
            .frm1.txtBefore_res_tax_amt.value       = "<%=UNINumClientFormat(lgObjRs("BEFORE_RES_TAX"), 0,0)%>"    
            .frm1.txtold_farm_tax_amt.value			= "<%=UNINumClientFormat(lgObjRs("BEFORE_FARM_TAX"), 0,0)%>"    
            .frm1.txtold_amt.value					= "<%=UNINumClientFormat(lgObjRs("txtold_amt"), 0,0)%>"    
            .frm1.txtincome_tax_amt.value           = "<%=UNINumClientFormat(lgObjRs("INCOME_TAX"), 0,0)%>"    
            .frm1.txtRes_tax_amt.value				= "<%=UNINumClientFormat(lgObjRs("RES_TAX"), 0,0)%>"    
            .frm1.txtfarm_tax_amt.value				= "<%=UNINumClientFormat(lgObjRs("FARM_TAX"), 0,0)%>"    
            .frm1.txtf_amt.value					= "<%=UNINumClientFormat(lgObjRs("txtf_amt"), 0,0)%>"    


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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
                lgStrSQL = "SELECT  income_tot_amt txtincome_tot_amt ,"
                lgStrSQL = lgStrSQL & " INCOME_SUB , INCOME_AMT, PER_SUB, SPOUSE_SUB, SUPP_CNT,SUPP_SUB, "
                lgStrSQL = lgStrSQL & " OLD_CNT  , OLD_CNT2,OLD_SUB, PARIA_CNT, PARIA_SUB, LADY_SUB,CHL_REAR	 , "
                lgStrSQL = lgStrSQL & " CHL_REAR_SUB , SMALL_SUB, INSUR_SUB, MED_SUB , EDU_SUB, HOUSE_FUND, "
                lgStrSQL = lgStrSQL & "  CONTR_SUB,CEREMONY_AMT, STD_SUB, SUB_INCOME_AMT, INDIV_ANU, INDIV_ANU2,  "
                lgStrSQL = lgStrSQL & "  INVEST_SUB_SUM,our_stock_amt, CARD_SUB_SUM,RETIRE_PENSION,FORE_EDU_SUB_AMT, TAX_STD, CALU_TAX, INCOME_TAX_SUB, HOUSE_REPAY, FORE_PAY, "
                lgStrSQL = lgStrSQL & "  STOCK_SAVE, poli_tax_sub, TAX_SUB_SUM, REDU_SUM,TAX_UNION_DED, DEC_INCOME_TAX, DEC_RES_TAX ,DEC_FARM_TAX , "
                lgStrSQL = lgStrSQL & " (DEC_INCOME_TAX + DEC_RES_TAX + DEC_FARM_TAX ) txtdec_amt , "
                lgStrSQL = lgStrSQL & "  BEFORE_INCOME_TAX, BEFORE_RES_TAX, BEFORE_FARM_TAX,"
                lgStrSQL = lgStrSQL & " (BEFORE_INCOME_TAX + BEFORE_RES_TAX + BEFORE_FARM_TAX ) txtold_amt , "
                lgStrSQL = lgStrSQL & "  INCOME_TAX, RES_TAX, FARM_TAX,"
                lgStrSQL = lgStrSQL & " (INCOME_TAX + RES_TAX + FARM_TAX) txtf_amt "
                lgStrSQL = lgStrSQL & " FROM HFA051T "  
                lgStrSQL = lgStrSQL & " WHERE HFA051T.emp_no = " & FilterVar(lgKeyStream(0),"'%'", "S")                       
                lgStrSQL = lgStrSQL & " AND HFA051T.internal_cd LIKE  " & FilterVar("%", "''", "S") & ""
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
       Case "UID_M0001"                                                         '¢Ð : Query
         If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                 .DBQueryOk        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
