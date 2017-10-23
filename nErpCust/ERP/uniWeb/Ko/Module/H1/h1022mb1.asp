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

    Call HideStatusWnd                                                               '��: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    Dim lgGetSvrDateTime    
    lgGetSvrDateTime = GetSvrDateTime

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")                                       '��: "P"(Prev search) "N"(Next search)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    Call SubBizQueryMulti()
End Sub    
	    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = FilterVar(lgKeyStream(1),"'%'", "S")
    
    if Trim(iKey1) = "" & FilterVar("%", "''", "S") & "" Then
    	Call SubMakeSQLStatements("MR",iKey1,"X",C_LIKE)
    Else
    	Call SubMakeSQLStatements("MR",iKey1,"X",C_EQ)
    End if	
    
    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '�� : No data is found.
        Call SetErrorStatus()
        
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CODE_TYPE_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CODE_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SALE_TAG_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SALE_TAG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_TYPE_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CREATE_MTD_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CREATE_MTD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCNT"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCNT_nm"))
    
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

            iDx =  iDx + 1
 
'            If iDx > parent.parent.VisibleRowCnt(.frm1.vspdData, 0) Then
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
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '��: Protect system from crashing

    Err.Clear                                                                        '��: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data
    
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
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
Sub SubBizSaveMultiCreate(arrColVal)
    Dim iclose_dt
    Dim strPay_dt
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iclose_dt=arrColVal(5)
	
    lgStrSQL = "INSERT INTO HDA200T( PROV_TYPE, CODE_TYPE, ALLOW_CD, SALE_TAG, EMP_TYPE, ACCNT," 
    lgStrSQL = lgStrSQL & " CREATE_MTD, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4),"" & FilterVar("*", "''", "S") & " ", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(5),"" & FilterVar("*", "''", "S") & " ", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(8), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "UPDATE  HDA200T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " ACCNT = " & FilterVar(arrColVal(6),"NULL", "S") & ","
    lgStrSQL = lgStrSQL & " CREATE_MTD = " & FilterVar(arrColVal(7),"NULL", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE PROV_TYPE = " & FilterVar(lgKeyStream(0), "NULL", "S")
    lgStrSQL = lgStrSQL & " AND CODE_TYPE = " & FilterVar(arrColVal(2), "NULL", "S")
    lgStrSQL = lgStrSQL & " AND ALLOW_CD = " & FilterVar(arrColVal(3), "NULL", "S")
    lgStrSQL = lgStrSQL & " AND SALE_TAG = " & FilterVar(arrColVal(4), "NULL", "S")
    lgStrSQL = lgStrSQL & " AND EMP_TYPE = " & FilterVar(arrColVal(5),"NULL", "S") 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  HDA200T"
    lgStrSQL = lgStrSQL & " WHERE PROV_TYPE = " & FilterVar(lgKeyStream(0), "NULL", "S")
    lgStrSQL = lgStrSQL & " AND CODE_TYPE = " & FilterVar(arrColVal(2), "NULL", "S")
    lgStrSQL = lgStrSQL & " AND ALLOW_CD = " & FilterVar(arrColVal(3), "NULL", "S")
    lgStrSQL = lgStrSQL & " AND SALE_TAG = " & FilterVar(arrColVal(4), "" & FilterVar("*", "''", "S") & " ", "S")
    lgStrSQL = lgStrSQL & " AND EMP_TYPE = " & FilterVar(arrColVal(5),"" & FilterVar("*", "''", "S") & " ", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements'("MR",iKey1,"X",C_EQ), ("MR",iKey1,"X",C_LIKE)
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select            CODE_TYPE	,"
                       

                        
                        
                       lgStrSQL = lgStrSQL & "	     dbo.ufn_GetCodeName(" & FilterVar("H0121", "''", "S") & ",CODE_TYPE) CODE_TYPE_nm, " 
                       lgStrSQL = lgStrSQL & "		 ALLOW_CD	,"
                       lgStrSQL = lgStrSQL & "       dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ",ALLOW_CD, (case when  hda200t.prov_type = " & FilterVar("2", "''", "S") & " then " & FilterVar("", "''", "S") & " else   hda200t.code_type end) ) ALLOW_nm ,"
				       lgStrSQL = lgStrSQL & "		 SALE_TAG	,"
				       lgStrSQL = lgStrSQL & "	     dbo.ufn_GetCodeName(" & FilterVar("H0071", "''", "S") & ",SALE_TAG) SALE_TAG_nm, " 
                       lgStrSQL = lgStrSQL & "		 EMP_TYPE	,"
                       lgStrSQL = lgStrSQL & "	     dbo.ufn_GetCodeName(" & FilterVar("H0122", "''", "S") & ",EMP_TYPE) EMP_TYPE_nm, "
                       lgStrSQL = lgStrSQL & "		 ACCNT, JNL_NM ACCNT_nm, "
                       lgStrSQL = lgStrSQL & "		 CREATE_MTD, "                       
                       lgStrSQL = lgStrSQL & "	     dbo.ufn_GetCodeName(" & FilterVar("H0123", "''", "S") & ",CREATE_MTD) CREATE_MTD_nm "
                       lgStrSQL = lgStrSQL & " From  HDA200T, A_JNL_ITEM"
                       lgStrSQL = lgStrSQL & " Where PROV_TYPE = " & FilterVar(lgKeyStream(0), "''", "S")
                       lgStrSQL = lgStrSQL & "   And CODE_TYPE " & pComp & " " & pCode
                       lgStrSQL = lgStrSQL & "   And JNL_TYPE = " & FilterVar("HR", "''", "S") & ""
                       lgStrSQL = lgStrSQL & "   And JNL_CD =* ACCNT"
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
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
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '�� : Display data                
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
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
