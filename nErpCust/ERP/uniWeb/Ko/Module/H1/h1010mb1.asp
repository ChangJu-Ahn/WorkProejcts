<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
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

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)        '☜:"0"(First),"1"(Second),"2"(Third),"3"(...)

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
    Dim iLcNo

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
%>
<Script Language=vbscript>
    With Parent.Frm1
        .txtAllow.text = "<%=UNINumClientFormat(lgObjRs("Allow"), ggAmtOfMoney.DecPoint, 0)%>"
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

    lgStrSQL = "DELETE  HDA090T"
    lgStrSQL = lgStrSQL & " WHERE ALLOW_CD = " & FilterVar(Request("txtAllow_cd"), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

    lgStrSQL = "DELETE  HDA100T"
    lgStrSQL = lgStrSQL & " WHERE ALLOW_CD = " & FilterVar(Request("txtAllow_cd"), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

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
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx = 1
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_nm"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("dilig_cnt"), 0,0)
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

    lgStrSQL = "INSERT INTO HDA090T( allow_cd, allow, ISRT_EMP_NO, ISRT_DT ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO , UPDT_DT   )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtallow"), 0) & ","

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
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

    lgStrSQL = "UPDATE  HDA090T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " allow = " & UNIConvNum(Request("txtallow"), 0) & ","

    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId,"11","S") & "," 
    lgStrSQL = lgStrSQL & " UPDT_Dt = " & FilterVar(GetSvrDateTime,"NULL","S")
    lgStrSQL = lgStrSQL & " WHERE           "
    lgStrSQL = lgStrSQL & " allow_cd = " & FilterVar(lgKeyStream(0), "''", "S")

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

    lgStrSQL = "INSERT INTO HDA100T( allow_cd, dilig_cd, dilig_cnt, ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT, UPDT_EMP_NO," 
    lgStrSQL = lgStrSQL & " UPDT_DT )" 
    lgStrSQL = lgStrSQL & " VALUES( "

    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(4), 0) & ","

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"NULL","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"NULL","S")
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

    lgStrSQL = "UPDATE  HDA100T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " dilig_cnt = " & UNIConvNum(arrColVal(4), 0) & ","

    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,"NULL","S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "   allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")

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

    lgStrSQL = "DELETE  HDA100T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "  allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")

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
                                   lgStrSQL = "Select * " 
                                   lgStrSQL = lgStrSQL & " From  HDA090T "
                                   lgStrSQL = lgStrSQL & " WHERE allow_cd " & pComp & pCode 	
                        End Select
           End Select             
        Case "M"           '                  0               1                 2              3                4                5
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           Select Case Mid(pDataType,2,1)
               Case "R"                
                       lgStrSQL = "Select TOP " & iSelCount 
                       lgStrSQL = lgStrSQL & " dilig_cd,"
                       lgStrSQL = lgStrSQL & " dilig_cd,dbo.ufn_H_GetCodeName(" & FilterVar("HCA010T", "''", "S") & " ,dilig_cd ,'' ) as dilig_nm,"
                       lgStrSQL = lgStrSQL & " dilig_cnt"
                       lgStrSQL = lgStrSQL & " From  HDA100T "
                       lgStrSQL = lgStrSQL & " WHERE allow_cd " & pComp & pCode
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
