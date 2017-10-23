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
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)    
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query        
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
   On Error Resume Next
   Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iDx
    Dim iLoopMax
    Dim iKey1, iKey2
    
    On Error Resume Next    
    Err.Clear
      
    iKey1 =  FilterVar(lgKeyStream(0),"'%'", "S")
    iKey2 =  FilterVar(lgKeyStream(1),"'%'", "S")

    Call SubMakeSQLStatements("MR",iKey1,iKey2,C_EQGT,C_LIKE)                              '¢Ð : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       lgStrPrevKey = ""
    
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))                  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_MM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_MM_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_STRT_DD"))
			lgstrData = lgstrData & Chr(11) & "~"
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_MM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_MM_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CRT_END_DD"))

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
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear 

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
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
	On Error Resume Next   
	Err.Clear

	lgStrSQL = "INSERT INTO HDA250T("
	lgStrSQL = lgStrSQL & " DILIG_CD, PAY_CD, CRT_STRT_MM, CRT_STRT_DD, CRT_END_MM, CRT_END_DD, " 
	lgStrSQL = lgStrSQL & " ISRT_DT, ISRT_EMP_NO, UPDT_DT , UPDT_EMP_NO )" 
	lgStrSQL = lgStrSQL & " VALUES(" 
	    
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","

	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(4),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S") & ","      
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S") & ","

	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")         & "," 
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        
	lgStrSQL = lgStrSQL & ")"
	    
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	On Error Resume Next 
	Err.Clear                                                                        '¢Ð: Clear Error status
  
    lgStrSQL = "UPDATE  HDA250T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " CRT_STRT_MM    = " & UNIConvNum(arrColVal(4),0)   & ","
    lgStrSQL = lgStrSQL & " CRT_STRT_DD    = " & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " CRT_END_MM     = " & UNIConvNum(arrColVal(6),0)   & ","
    lgStrSQL = lgStrSQL & " CRT_END_DD     = " & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT        = " & FilterVar(GetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO    = " & FilterVar(gUsrId, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE DILIG_CD = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND PAY_CD   = " & FilterVar(UCase(arrColVal(3)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  HDA250T"
    lgStrSQL = lgStrSQL & " WHERE DILIG_CD = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND PAY_CD   = " & FilterVar(UCase(arrColVal(3)), "''", "S")
   
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp1,pComp2)

	 Dim iSelCount
	 
    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

           Select Case Mid(pDataType,2,1)
               Case "R"               
                       lgStrSQL = "SELECT TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & "	DILIG_CD, "                         
                       lgStrSQL = lgStrSQL & "	dbo.ufn_H_GetCodeName(" & FilterVar("HCA010T", "''", "S") & ", DILIG_CD, '') DILIG_NM, "                       
                       lgStrSQL = lgStrSQL & "  PAY_CD      ,"
                       lgStrSQL = lgStrSQL & "  dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ",PAY_CD) PAY_NM, "
					   lgStrSQL = lgStrSQL & "  CRT_STRT_MM, dbo.ufn_GetCodeName(" & FilterVar("H0088", "''", "S") & ",CRT_STRT_MM) CRT_STRT_MM_nm, "
					   lgStrSQL = lgStrSQL & "  CRT_STRT_DD,  CRT_END_MM, "
					   lgStrSQL = lgStrSQL & "  dbo.ufn_GetCodeName(" & FilterVar("H0088", "''", "S") & ",CRT_END_MM) CRT_END_MM_nm, "
					   lgStrSQL = lgStrSQL & "  CRT_END_DD "
                       lgStrSQL = lgStrSQL & "  FROM HDA250T "
                       lgStrSQL = lgStrSQL & " WHERE DILIG_CD " & pComp1 & pCode 
                       lgStrSQL = lgStrSQL & "   AND PAY_CD " & pComp2 & " " &  pCode1 
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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk        
	         End with        
	         
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
