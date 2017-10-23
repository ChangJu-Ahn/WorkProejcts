<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Dim lgSvrDateTime
    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "A","NOCOOKIE","MB")    
   ' lgSvrDateTime = GetSvrDateTime
    
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

  
    Call SubMakeSQLStatements("MR","x","X",C_EQ)                                 '¡Ù : Make sql statements
    'response.write lgStrSQL
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
    
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_seq"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("document_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("issue_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("issue_dt"),"")
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("ship_dt"),"")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("doc_cur"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("doc_cur"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("xch_rate"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("present_amt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("present_loc_amt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("report_amt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("report_loc_amt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("vat_desc"))
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
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "INSERT INTO a_vat_c_KO441( biz_area_cd, fr_yymm, to_yymm, item_seq,"
    lgStrSQL = lgStrSQL & " document_cd, issue_cd, ISSUE_DT, SHIP_DT, DOC_CUR, vat_desc, "
    lgStrSQL = lgStrSQL & " XCH_RATE, PRESENT_AMT, PRESENT_LOC_AMT, REPORT_AMT, REPORT_LOC_AMT, "
    lgStrSQL = lgStrSQL & " Isrt_emp_no, Isrt_dt     , Updt_emp_no , Updt_dt )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(6))),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(7))),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(8))),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(9))),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(10))),"''","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(16)),"''","S")			& ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(11)),0) 			    & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(12)),0) 			    & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(13)),0) 			    & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(14)),0) 			    & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(15)),0) 			    & ","

    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,"''","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,"''","S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  a_vat_c_KO441"
    lgStrSQL = lgStrSQL & " SET " 

    lgStrSQL = lgStrSQL & " document_cd		= " &  FilterVar(Trim(arrColVal(6)),"''","S") & ","
    lgStrSQL = lgStrSQL & " issue_cd		= " &  FilterVar(Trim(UCase(arrColVal(7))),"''","S") & ","
    lgStrSQL = lgStrSQL & " issue_dt		= " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(8),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " ship_dt			= " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(9),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " doc_cur			= " &  FilterVar(Trim(UCase(arrColVal(10))),"''","S") & ","
    lgStrSQL = lgStrSQL & " xch_rate		= " &  UNIConvNum(Trim(arrColVal(11)),0) & ","
    lgStrSQL = lgStrSQL & " present_amt		= " &  UNIConvNum(Trim(arrColVal(12)),0) & ","
    lgStrSQL = lgStrSQL & " present_loc_amt = " &  UNIConvNum(Trim(arrColVal(13)),0) & ","
    lgStrSQL = lgStrSQL & " report_amt		= " &  UNIConvNum(Trim(arrColVal(14)),0) & ","
    lgStrSQL = lgStrSQL & " report_loc_amt  = " &  UNIConvNum(Trim(arrColVal(15)),0) & ","
    lgStrSQL = lgStrSQL & " vat_desc		= " &  FilterVar(Trim(UCase(arrColVal(16))),"''","S") & ","
    lgStrSQL = lgStrSQL & " Updt_emp_no		= " &  FilterVar(gUsrId,"''","S")   & ","
    lgStrSQL = lgStrSQL & " Updt_dt			= " &  FilterVar(lgSvrDateTime,"''","S")  
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "     biz_area_cd = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & " AND    fr_yymm  = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")
    lgStrSQL = lgStrSQL & " AND    to_yymm  = " &  FilterVar(Trim(UCase(arrColVal(4))),"''","S")
    lgStrSQL = lgStrSQL & " AND    item_seq = " &  FilterVar(Trim(UCase(arrColVal(5))),"''","S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  a_vat_c_KO441"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "     biz_area_cd = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & " AND    fr_yymm  = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")
    lgStrSQL = lgStrSQL & " AND    to_yymm  = " &  FilterVar(Trim(UCase(arrColVal(4))),"''","S")
    lgStrSQL = lgStrSQL & " AND    item_seq = " &  FilterVar(Trim(UCase(arrColVal(5))),"''","S")


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount  & " *"
                       lgStrSQL = lgStrSQL & " FROM  a_vat_c_KO441 "
                       lgStrSQL = lgStrSQL & " WHERE FR_YYMM =  "  & FilterVar(lgKeyStream(0),"''", "S")
                       lgStrSQL = lgStrSQL & " AND TO_YYMM = "     & FilterVar(lgKeyStream(1),"''", "S")
                       lgStrSQL = lgStrSQL & " AND biz_area_cd = " & FilterVar(lgKeyStream(2),"''", "S")
                       lgStrSQL = lgStrSQL & " ORDER BY  item_seq ASC "
                      ' CALL svrmsgbox(lgstrsql, vbinformation, i_mkscript)
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
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
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
