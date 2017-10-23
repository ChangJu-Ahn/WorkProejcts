<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%    
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtCo_Cd"),gColSep)


    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0),"''", "S")
    iKey2 = FilterVar(request("txtFISC_YEAR"),"''", "S")
    iKey3 = FilterVar(request("cboREP_TYPE"),"''", "S")

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
          Call SetErrorStatus()
    Else

%>
<Script Language=vbscript>
       With Parent	
                .Frm1.txtCO_NM.Value  = "<%=ConvSPChars(lgObjRs("CO_NM"))%>"
                .Frm1.txtCO_ADDR.Value  = "<%=ConvSPChars(lgObjRs("CO_ADDR"))%>"
                .Frm1.txtOWN_RGST_NO.TEXT  = "<%=ConvSPChars(lgObjRs("OWN_RGST_NO"))%>"
                .Frm1.txtLAW_RGST_NO.TEXT  = "<%=ConvSPChars(lgObjRs("LAW_RGST_NO"))%>"
                .Frm1.txtREPRE_NM.Value  = "<%=ConvSPChars(lgObjRs("REPRE_NM"))%>"
                .Frm1.txtREPRE_RGST_NO.text  = "<%=ConvSPChars(lgObjRs("REPRE_RGST_NO"))%>"
                .Frm1.txtTEL_NO.Value  = "<%=ConvSPChars(lgObjRs("TEL_NO"))%>"
                .Frm1.cboCOMP_TYPE1.Value  = "<%=ConvSPChars(lgObjRs("COMP_TYPE1"))%>"
                .Frm1.cboDEBT_MULTIPLE.Value  = "<%=ConvSPChars(lgObjRs("DEBT_MULTIPLE"))%>"
                .Frm1.cboCOMP_TYPE2.Value  = "<%=ConvSPChars(lgObjRs("COMP_TYPE2"))%>"
                .Frm1.txtTAX_OFFICE.Value  = "<%=ConvSPChars(lgObjRs("TAX_OFFICE"))%>"
                .Frm1.txtTAX_OFFICE_NM.Value  = "<%=ConvSPChars(lgObjRs("TAX_OFFICE_NM"))%>"
'                .Frm1.txtFISC_END_DT.Value  = "<%=UNIDateClientFormat(lgObjRs("FISC_END_DT"))%>"
                .Frm1.cboHOLDING_COMP_FLG.Value  = "<%=ConvSPChars(lgObjRs("HOLDING_COMP_FLG"))%>"
                .Frm1.txtIND_CLASS.Value  = "<%=ConvSPChars(lgObjRs("IND_CLASS"))%>"
                .Frm1.txtIND_TYPE.Value  = "<%=ConvSPChars(lgObjRs("IND_TYPE"))%>"
                .Frm1.txtFOUNDATION_DT.text  = "<%=ConvSPChars(lgObjRs("FOUNDATION_DT"))%>"

				.Frm1.txtHOME_TAX_USR_ID.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_USR_ID"))%>"
                .Frm1.txtHOME_TAX_EMAIL.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_E_MAIL"))%>"
                .Frm1.txtHOME_TAX_MAIN_IND.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_MAIN_IND"))%>"
                .Frm1.txtHOME_TAX_MAIN_IND_NM.Value = "<%=ConvSPChars(lgObjRs("HOME_TAX_MAIN_IND_NM"))%>"
               
                .Frm1.txtBANK_CD.Value  = "<%=ConvSPChars(lgObjRs("BANK_CD"))%>"
                .Frm1.txtBANK_NM.Value  = "<%=ConvSPChars(lgObjRs("BANK_NM"))%>"
                .Frm1.txtBANK_BRANCH.Value  = "<%=ConvSPChars(lgObjRs("BANK_BRANCH"))%>"
                .Frm1.txtBANK_DPST.Value  = "<%=ConvSPChars(lgObjRs("BANK_DPST"))%>"
                .Frm1.txtBANK_ACCT_NO.Value  = "<%=ConvSPChars(lgObjRs("BANK_ACCT_NO"))%>"
'-- Ãß°¡
                .Frm1.txtFISC_YEAR_Body.text  = "<%=ConvSPChars(lgObjRs("FISC_YEAR"))%>"
                .Frm1.cboREP_TYPE_Body.Value  = "<%=ConvSPChars(lgObjRs("REP_TYPE"))%>"
                .Frm1.txtFISC_START_DT.text  = "<%=UNIDateClientFormat(lgObjRs("FISC_START_DT"))%>"
                .Frm1.txtFISC_END_DT.text  = "<%=UNIDateClientFormat(lgObjRs("FISC_END_DT"))%>"
                .Frm1.txtHOME_ANY_START_DT.text  = "<%=UNIDateClientFormat(lgObjRs("HOME_ANY_START_DT"))%>"
                .Frm1.txtHOME_ANY_END_DT.text  = "<%=UNIDateClientFormat(lgObjRs("HOME_ANY_END_DT"))%>"
                .Frm1.txtHOME_TAX_USR_ID.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_USR_ID"))%>"
                .Frm1.txtHOME_TAX_EMAIL.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_E_MAIL"))%>"
                .Frm1.txtHOME_TAX_MAIN_IND.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_MAIN_IND"))%>"
                .Frm1.txtHOME_TAX_MAIN_IND_NM.Value = "<%=ConvSPChars(lgObjRs("HOME_TAX_MAIN_IND_NM"))%>"
                
                .Frm1.txtHOME_FILE_MAKE_DT.TEXT  = "<%=ConvSPChars(lgObjRs("HOME_FILE_MAKE_DT"))%>"
				.Frm1.txtINCOM_DT.TEXT  = "<%=ConvSPChars(lgObjRs("INCOM_DT"))%>"
				
                .Frm1.txtAGENT_NM.Value  = "<%=ConvSPChars(lgObjRs("AGENT_NM"))%>"
                .Frm1.txtRECON_BAN_NO.TEXT  = "<%=ConvSPChars(lgObjRs("RECON_BAN_NO"))%>"
                .Frm1.txtRECON_MGT_NO.TEXT  = "<%=ConvSPChars(lgObjRs("RECON_MGT_NO"))%>"
                .Frm1.txtAGENT_TEL_NO.Value  = "<%=ConvSPChars(lgObjRs("AGENT_TEL_NO"))%>"
				.Frm1.txtAGENT_RGST_NO.TEXT  = "<%=ConvSPChars(lgObjRs("AGENT_RGST_NO"))%>"
				.Frm1.txtREQUEST_DT.Value  = "<%=ConvSPChars(lgObjRs("REQUEST_DT"))%>"
				.Frm1.txtAPPO_NO.TEXT  = "<%=ConvSPChars(lgObjRs("APPO_NO"))%>"
				.Frm1.txtAPPO_DT.TEXT  = "<%=ConvSPChars(lgObjRs("APPO_DT"))%>"
				.Frm1.txtAPPO_DESC.Value  = "<%=ConvSPChars(lgObjRs("APPO_DESC"))%>"
				.Frm1.cboEX_RECON_FLG.Value  = "<%=ConvSPChars(lgObjRs("EX_RECON_FLG"))%>"
				.Frm1.cboEX_54_FLG.Value  = "<%=ConvSPChars(lgObjRs("EX_54_FLG"))%>"
				
                .Frm1.txtBANK_CD.Value  = "<%=ConvSPChars(lgObjRs("BANK_CD"))%>"
                .Frm1.txtBANK_NM.Value  = "<%=ConvSPChars(lgObjRs("BANK_NM"))%>"
                .Frm1.txtBANK_BRANCH.Value  = "<%=ConvSPChars(lgObjRs("BANK_BRANCH"))%>"
                .Frm1.txtBANK_DPST.Value  = "<%=ConvSPChars(lgObjRs("BANK_DPST"))%>"
                .Frm1.txtBANK_ACCT_NO.Value  = "<%=ConvSPChars(lgObjRs("BANK_ACCT_NO"))%>"

				.Frm1.cboSUBMIT_FLG.Value  = "<%=ConvSPChars(lgObjRs("SUBMIT_FLG"))%>"
				.Frm1.cboUSE_FLG.Value  = "<%=ConvSPChars(lgObjRs("USE_FLG"))%>"
				.Frm1.txtREVISION_YM.Value  = "<%=C_REVISION_YM%>"
                .Frm1.txtCO_CD.focus

       End With          
</Script>       
<%     
    End If
    Call SubCloseRs(lgObjRs)
End Sub	


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"

		lgStrSQL =			  " SELECT  " & vbCrLf
        lgStrSQL = lgStrSQL & "a.CO_CD, a.CO_NM, a.CO_ADDR, a.OWN_RGST_NO, a.LAW_RGST_NO, a.REPRE_NM, a.REPRE_RGST_NO " & vbCrLf
        lgStrSQL = lgStrSQL & " , a.TEL_NO, a.COMP_TYPE1, a.DEBT_MULTIPLE, a.COMP_TYPE2, a.TAX_OFFICE,  dbo.ufn_GetCodeName('W1079', a.TAX_OFFICE) as TAX_OFFICE_NM " & vbCrLf
        lgStrSQL = lgStrSQL & " , a.HOLDING_COMP_FLG, a.IND_CLASS, a.IND_TYPE, a.FOUNDATION_DT " & vbcrlf
        lgStrSQL = lgStrSQL & ", A.IND_CLASS, A.IND_TYPE " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.FISC_YEAR, A.REP_TYPE " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.HOME_TAX_USR_ID " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.HOME_TAX_E_MAIL, A.HOME_TAX_MAIN_IND, C.DETAIL_NM HOME_TAX_MAIN_IND_NM, A.EX_RECON_FLG, A.EX_54_FLG, A.SUBMIT_FLG, A.USE_FLG " & vbCrLf
	
        lgStrSQL = lgStrSQL & " , A.FISC_START_DT , A.FISC_END_DT, A.HOME_ANY_START_DT, A.HOME_ANY_END_DT " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.HOME_FILE_MAKE_DT, A.INCOM_DT, A.AGENT_NM, A.RECON_BAN_NO, A.RECON_MGT_NO, A.AGENT_TEL_NO, A.AGENT_RGST_NO, A.REQUEST_DT " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.APPO_NO, A.APPO_DT, A.APPO_DESC, A.REVISION_YM, A.USE_FLG " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.BANK_CD, dbo.ufn_GetCodeName('W1020', A.BANK_CD) as BANK_NM, A.BANK_BRANCH, A.BANK_DPST, A.BANK_ACCT_NO " & vbCrLf
        lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY A (nolock) " & vbCrLf
		lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN  tb_std_income_rate C (NOLOCK) ON A.IND_TYPE = C.STD_INCM_RT_CD" & vbCrLf

        lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
		lgStrSQL = lgStrSQL & " 	 AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
	    lgStrSQL = lgStrSQL & "      AND A.REP_TYPE = " & pCode3 	&  vbCrLf

    End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
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
    lgErrorStatus     = "YES"
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
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
        Case "MD"
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
