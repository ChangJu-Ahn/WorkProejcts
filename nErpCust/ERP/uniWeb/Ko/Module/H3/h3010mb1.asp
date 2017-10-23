<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
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

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    
    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQ)                                 '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("edu_start_dt"),"")
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("edu_end_dt"),"")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_office"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_office_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_nat"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_nat_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_cont"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_type"))
            lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("edu_score"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("end_dt"),"")
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("edu_fee"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("fee_type"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("repay_amt"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("report_type"))
		    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("add_point"), ggAmtOfMoney.DecPoint, 0)

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

    lgStrSQL = "INSERT INTO HBA030T("
    lgStrSQL = lgStrSQL & "     emp_no,"
    lgStrSQL = lgStrSQL & "     edu_cd,"
    lgStrSQL = lgStrSQL & "     edu_start_dt,"
    lgStrSQL = lgStrSQL & "     edu_end_dt,"
    lgStrSQL = lgStrSQL & "     edu_office,"
    lgStrSQL = lgStrSQL & "     edu_nat,"
    lgStrSQL = lgStrSQL & "     edu_cont,"
    lgStrSQL = lgStrSQL & "     edu_type,"
    lgStrSQL = lgStrSQL & "     edu_score,"
    lgStrSQL = lgStrSQL & "     end_dt,"
    lgStrSQL = lgStrSQL & "     edu_fee,"
    lgStrSQL = lgStrSQL & "     fee_type,"
    lgStrSQL = lgStrSQL & "     repay_amt,"
    lgStrSQL = lgStrSQL & "     report_type,"
    lgStrSQL = lgStrSQL & "     add_point,"
    lgStrSQL = lgStrSQL & "     ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & "     ISRT_DT     ," 
    lgStrSQL = lgStrSQL & "     UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & "     UPDT_DT      )" 
    lgStrSQL = lgStrSQL & "     VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(4),NULL),"NULL","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(8), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(11),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(12),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(13)), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(14),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(15)), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(16),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  HBA030T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & "     edu_end_dt = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & "     edu_office = " & FilterVar(UCase(arrColVal(6)), "''", "S") & ","
    lgStrSQL = lgStrSQL & "     edu_nat = " & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & "     edu_cont = " & FilterVar(arrColVal(8), "''", "S") & ","
    lgStrSQL = lgStrSQL & "     edu_type = " & FilterVar(UCase(arrColVal(9)), "''", "S") & ","
    lgStrSQL = lgStrSQL & "     edu_score = " & UNIConvNum(arrColVal(10),0) & ","
    lgStrSQL = lgStrSQL & "     end_dt = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(11),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & "     edu_fee = " & UNIConvNum(arrColVal(12),0) & ","
    lgStrSQL = lgStrSQL & "     fee_type = " & FilterVar(UCase(arrColVal(13)), "''", "S") & ","
    lgStrSQL = lgStrSQL & "     repay_amt = " & UNIConvNum(arrColVal(14),0) & ","
    lgStrSQL = lgStrSQL & "     report_type = " & FilterVar(UCase(arrColVal(15)), "''", "S") & ","
    lgStrSQL = lgStrSQL & "     add_point = " & UNIConvNum(arrColVal(16),0) & ","
    lgStrSQL = lgStrSQL & "     updt_emp_no = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & "     updt_dt = " & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "     emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND edu_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND edu_start_dt = " & FilterVar(UNIConvDate(arrColVal(4)), "''", "S")

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

    lgStrSQL = "DELETE  HBA030T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "     emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND edu_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND edu_start_dt = " & FilterVar(UNIConvDate(arrColVal(4)), "''", "S")

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
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount   
                       lgStrSQL = lgStrSQL & "				edu_cd				, "
                       lgStrSQL = lgStrSQL & "				dbo.ufn_GetCodeName(" & FilterVar("H0033", "''", "S") & ",edu_cd) edu_nm, "
                       lgStrSQL = lgStrSQL & "				edu_start_dt		, "	
                       lgStrSQL = lgStrSQL & "				edu_end_dt		    , "
                       lgStrSQL = lgStrSQL & "				edu_office			, "
                       lgStrSQL = lgStrSQL & "				dbo.ufn_GetCodeName(" & FilterVar("H0037", "''", "S") & ",edu_office) edu_office_nm, "
                       lgStrSQL = lgStrSQL & "				edu_nat				, "
                       lgStrSQL = lgStrSQL & "				dbo.ufn_H_GetCodeName(" & FilterVar("B_COUNTRY", "''", "S") & ",edu_nat,'') edu_nat_nm, "
                       lgStrSQL = lgStrSQL & "				edu_cont			, "
                       lgStrSQL = lgStrSQL & "				edu_type			, "
                       lgStrSQL = lgStrSQL & "				edu_score			, "
                       lgStrSQL = lgStrSQL & "				end_dt				, "
                       lgStrSQL = lgStrSQL & "				edu_fee				, "
                       lgStrSQL = lgStrSQL & "				fee_type			, "
                       lgStrSQL = lgStrSQL & "				repay_amt			, "
                       lgStrSQL = lgStrSQL & "				report_type			, "
                       lgStrSQL = lgStrSQL & "				add_point			  "                       							
                       lgStrSQL = lgStrSQL & " From  HBA030T "
                       lgStrSQL = lgStrSQL & " Where emp_no " & pComp & pCode
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
