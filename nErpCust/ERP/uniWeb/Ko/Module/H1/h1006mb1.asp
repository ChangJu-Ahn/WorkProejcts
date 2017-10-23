<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MA")                                                               '☜: Hide Processing message
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim txtAllow_cd
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    txtAllow_cd = FilterVar(lgKeyStream(0), "''", "S")
    
    Call SubMakeSQLStatements("SR",txtAllow_cd,"X",C_EQ)                             '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then               'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '☜: No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)            '☜: This is the starting data. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)            '☜: This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
        call CommonQueryRs(" ALLOW_NM "," HDA010T "," PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & " and ALLOW_CD = " & txtAllow_cd ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
%>
<Script Language=vbscript>
    With Parent.Frm1
        .txtOneself.text = "<%=UNINumClientFormat(lgObjRs("Oneself"), ggQty.DecPoint, 0)%>"
        .txtspouse_amt.text = "<%=UNINumClientFormat(lgObjRs("spouse_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtfam_amt.text = "<%=UNINumClientFormat(lgObjRs("fam_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtlimit_amt.text = "<%=UNINumClientFormat(lgObjRs("limit_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtlimit_man.text = "<%=UNINumClientFormat(lgObjRs("limit_man"), ggQty.DecPoint, 0)%>"
        .txtAllow_nm.value = "<%=Trim(Replace(lgF0,Chr(11),""))%>"

    End With          
</Script>       
<%     
       Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
        If Request("strFamily_type") = "1" Then     ' 급여마스타참조        
        Else        
            Call SubBizQueryMulti(txtAllow_cd)
        End If
    End If
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
        Case  OPMD_CMODE                                                             '☜: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '☜: Update
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

    lgStrSQL = "DELETE  HDA060T"
    lgStrSQL = lgStrSQL & " WHERE ALLOW_CD = " & FilterVar(Request("txtAllow_cd"), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

    lgStrSQL = "DELETE  HDA070T"
    lgStrSQL = lgStrSQL & " WHERE ALLOW_CD = " & FilterVar(Request("txtAllow_cd"), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pKey1)

    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR",pKey1,"X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        
    Else
    
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("rel_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("rel_cd_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("age"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("over_below"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("over_below_nm"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("allow"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("limit_type"))

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

    Call SubHandleError("MR",lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    Dim txtGlNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HDA060T( ALLOW_CD, ONESELF,SPOUSE_AMT, FAM_AMT, LIMIT_AMT," 
    lgStrSQL = lgStrSQL & " LIMIT_MAN, ISRT_EMP_NO, ISRT_DT , UPDT_EMP_NO , UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtOneself"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtSpouse_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtFam_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLimit_amt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtLimit_man"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDA060T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " ONESELF = " & UNIConvNum(Request("txtOneself"),0) & ","
    lgStrSQL = lgStrSQL & " SPOUSE_AMT = " & UNIConvNum(Request("txtSpouse_amt"),0) & ","
    lgStrSQL = lgStrSQL & " FAM_AMT = " & UNIConvNum(Request("txtFam_amt"),0) & ","
    lgStrSQL = lgStrSQL & " LIMIT_AMT = " & UNIConvNum(Request("txtLimit_amt"),0) & ","
    lgStrSQL = lgStrSQL & " LIMIT_MAN = " & UNIConvNum(Request("txtLimit_man"),0)
    lgStrSQL = lgStrSQL & " WHERE           "
    lgStrSQL = lgStrSQL & " ALLOW_CD = " & FilterVar(lgKeyStream(0), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HDA070T( ALLOW_CD, REL_CD,AGE, OVER_BELOW, ALLOW, LIMIT_TYPE,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT )" 
    lgStrSQL = lgStrSQL & " VALUES( "

    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","    
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(4),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDA070T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " AGE = " & UNIConvNum(arrColVal(4),0) & ","
    lgStrSQL = lgStrSQL & " OVER_BELOW = " & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " ALLOW = " & UNIConvNum(arrColVal(6),0) & ","
    lgStrSQL = lgStrSQL & " LIMIT_TYPE = " & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE ALLOW_CD = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND REL_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HDA070T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND rel_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL = "SELECT * " 
                                   lgStrSQL = lgStrSQL & " FROM  HDA060T "
                                   lgStrSQL = lgStrSQL & " WHERE ALLOW_CD " & pComp & pCode 	
                        End Select
           End Select             
        Case "M"           '                  0               1                 2              3                4                5
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "SELECT TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & "				rel_cd, "
                       lgStrSQL = lgStrSQL & "				dbo.ufn_GetCodeName(" & FilterVar("H0023", "''", "S") & ",rel_cd) rel_cd_nm, " 
                       lgStrSQL = lgStrSQL & "				age, "
                       lgStrSQL = lgStrSQL & "				over_below, " 
                       lgStrSQL = lgStrSQL & "				dbo.ufn_GetCodeName(" & FilterVar("H0051", "''", "S") & ",over_below) over_below_nm, " 
                       lgStrSQL = lgStrSQL & "				allow, " 
                       lgStrSQL = lgStrSQL & "				limit_type "                        
                       lgStrSQL = lgStrSQL & " FROM  HDA070T "
                       lgStrSQL = lgStrSQL & " WHERE ALLOW_CD " & pComp & pCode
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
