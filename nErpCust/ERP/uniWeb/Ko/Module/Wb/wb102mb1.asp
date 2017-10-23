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
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtCo_Cd"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

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
    Dim iDx
    Dim iLoopMax
    Dim iKey1, iKey2, iKey3

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0),"''", "S")
    iKey2 = FilterVar(request("txtFISC_YEAR"),"''", "S")
    iKey3 = FilterVar(request("cboREP_TYPE"),"''", "S")

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
         lgStrPrevKey = ""
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
          Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CO_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FISC_YEAR"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REP_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CO_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CO_ADDR"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OWN_RGST_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LAW_RGST_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REPRE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REPRE_RGST_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEL_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COMP_TYPE1"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEBT_MULTIPLE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COMP_TYPE2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TAX_OFFICE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HOLDING_COMP_FLG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IND_CLASS"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IND_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FOUNDATION_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REP_TYPE_CD"))
              
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
               
        Loop 
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)
End Sub	

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Dim iSelCount


    Select Case pMode 
      Case "R"

            iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

			lgStrSQL =			  " SELECT  " & vbCrLf
			lgStrSQL = lgStrSQL & "a.CO_CD, a.CO_NM, a.CO_ADDR, a.OWN_RGST_NO, a.LAW_RGST_NO, a.REPRE_NM, a.REPRE_RGST_NO " & vbCrLf
			lgStrSQL = lgStrSQL & " , a.TEL_NO  " & vbCrLf
			lgStrSQL = lgStrSQL & " , dbo.ufn_GetCodeName('W1009', a.COMP_TYPE1) as COMP_TYPE1" & vbcrlf
			lgStrSQL = lgStrSQL & " , dbo.ufn_GetCodeName('W1010', a.DEBT_MULTIPLE) as DEBT_MULTIPLE" & vbcrlf
			lgStrSQL = lgStrSQL & " , dbo.ufn_GetCodeName('W1079', a.TAX_OFFICE) as TAX_OFFICE" & vbcrlf
			lgStrSQL = lgStrSQL & " , dbo.ufn_GetCodeName('W1013', a.COMP_TYPE2) as COMP_TYPE2" & vbcrlf
			lgStrSQL = lgStrSQL & " , dbo.ufn_GetCodeName('W1003', a.HOLDING_COMP_FLG) as HOLDING_COMP_FLG" & vbcrlf
			lgStrSQL = lgStrSQL & " , a.IND_CLASS, a.IND_TYPE, a.FOUNDATION_DT " & vbcrlf
			lgStrSQL = lgStrSQL & " , A.FISC_YEAR, dbo.ufn_GetCodeName('W1018', A.REP_TYPE) as REP_TYPE , REP_TYPE REP_TYPE_CD" & vbCrLf
			lgStrSQL = lgStrSQL & " , A.HOME_TAX_USR_ID " & vbCrLf
			lgStrSQL = lgStrSQL & " , A.HOME_TAX_E_MAIL, A.HOME_TAX_MAIN_IND, C.DETAIL_NM HOME_TAX_MAIN_IND_NM, A.EX_RECON_FLG " & vbCrLf
	
			lgStrSQL = lgStrSQL & " , A.FISC_START_DT , A.FISC_END_DT, A.HOME_ANY_START_DT, A.HOME_ANY_END_DT " & vbCrLf
			lgStrSQL = lgStrSQL & " , A.HOME_FILE_MAKE_DT, A.AGENT_NM, A.RECON_BAN_NO, A.RECON_MGT_NO, A.AGENT_TEL_NO, A.AGENT_RGST_NO, A.REQUEST_DT " & vbCrLf
			lgStrSQL = lgStrSQL & " , A.APPO_NO, A.APPO_DT, A.APPO_DESC " & vbCrLf
			lgStrSQL = lgStrSQL & " , A.BANK_CD, dbo.ufn_GetCodeName('W1020', A.BANK_CD) as BANK_NM, A.BANK_BRANCH, A.BANK_DPST, A.BANK_ACCT_NO " & vbCrLf
			lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY A (nolock) " & vbCrLf
			lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN  tb_std_income_rate C (NOLOCK) ON A.IND_TYPE = C.STD_INCM_RT_CD" & vbCrLf

            lgStrSQL = lgStrSQL & "  WHERE a.CO_CD = " & pCode1 	 & vbCrLf

			if pCode2 <> "''" then
			    lgStrSQL = lgStrSQL & " 			 AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
            end if
			if pCode3 <> "''" then
	            lgStrSQL = lgStrSQL & "              AND A.REP_TYPE = " & pCode3 	& vbCrlf
            end if

            
            lgStrSQL = lgStrSQL & "  order by A.FISC_YEAR asc, A.rep_type desc " & vbCrLf

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
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk        
	         End with
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