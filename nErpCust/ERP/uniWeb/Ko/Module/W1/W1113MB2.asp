<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% Option Explicit%>
<% session.CodePage=949 %>
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
    Dim iKey1
    Dim strPreCD

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0),"''", "S")

    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
          'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
          Call SetErrorStatus()
    Else


        lgstrData = ""
        
%>
<Script Language=vbscript>
        Parent.Frm1.txtCO_CD_Body.Value  = "<%=Request("txtCo_Cd")%>"
        Parent.Frm1.txtFISC_YEAR_Body.Value  = "<%=Request("txtFISC_YEAR")%>"
        Parent.Frm1.txtREP_TYPE_Body.Value  = "<%=Request("txtREP_TYPE")%>"
        Parent.Frm1.txtBS_PL_FG.Value  = "<%=ConvSPChars(lgObjRs("BS_PL_FG"))%>"
</Script>       
<%     
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GP_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAR_GP_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GP_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FORM_REP_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AMT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUM_FG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GP_LVL"))
            
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
        Loop 

    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)
End Sub	

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1)

    Select Case pMode 
      Case "R"


			lgStrSQL =			  " SELECT a.GP_CD, a.PAR_GP_CD, SPACE((GP_LVL * 2)) + a.GP_PRN_NM AS GP_NM, FORM_REP_NO, 0 AS AMT, SUM_FG, GP_LVL, a.BS_PL_FG " & vbCrLf
            'lgStrSQL = lgStrSQL & "  FROM TB_ACCT_GP a (NOLOCK) " & vbCrLf
            lgStrSQL = lgStrSQL & "  FROM dbo.ufn_TB_ACCT_GP('" & C_REVISION_YM & "') a" & vbCrLf
            lgStrSQL = lgStrSQL & "  WHERE a.BS_PL_FG = '6' " & vbCrLf

            lgStrSQL = lgStrSQL & "  ORDER BY A.GP_SEQ, A.GP_CD  " & vbCrLf

    End Select
PrintLog "SubMakeSQLStatements = " & lgStrSQL
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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk2        
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