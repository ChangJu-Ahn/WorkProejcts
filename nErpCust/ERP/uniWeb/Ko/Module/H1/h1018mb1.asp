<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Const TAB1 = 1
    Const TAB2 = 2

    Dim lgSvrDateTime    

    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    lgCurrentSpd      = CInt(Request("lgCurrentSpd"))                                '¢Ð: "1"(Spread #1) "2"(Spread #2)

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
    Dim strWhere
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strWhere = FilterVar(lgKeyStream(0),"'%'", "S")

    If lgCurrentSpd = TAB1 Then
       Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)  
       
    Else
       Call SubMakeSQLStatements("MM",strWhere,"X",C_EQGT)       
    End If

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF

            If lgCurrentSpd = TAB1 Then

               lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD")  )        
               lgstrData = lgstrData & Chr(11) & ""  
               lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NM") )              
               lgstrData = lgstrData & Chr(11) & UNINumclientFormat(lgObjRs("BAS_AMT"), ggAmtOfMoney.DecPoint,0)               
               lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("BELOW")))
               lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("BELOW_nm")))
               lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("PROC_BAS")))
               lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("PROC_BAS_nm")))

            Else
               lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ATTEND_TYPE"))
               lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ATTEND_TYPE_nm") )          
               lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DECI_PLACE") )              
               lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("PROC_BAS")))
               lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("PROC_BAS_nm")))
               lgstrData = lgstrData & Chr(11) & ConvSPChars(Get_Format(Trim(lgObjRs("DECI_PLACE"))))

            End If          
                      
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

    On Error Resume Next
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

    If lgCurrentSpd = TAB1 Then      
        lgStrSQL = "INSERT INTO HDA040T( ALLOW_CD , BAS_AMT  , BELOW    , PROC_BAS ,"
        lgStrSQL = lgStrSQL & " ISRT_DT      , ISRT_EMP_NO  , UPDT_DT      , UPDT_EMP_NO)" 
        lgStrSQL = lgStrSQL & " VALUES(" 
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
        lgStrSQL = lgStrSQL & UniConvNum(Trim(arrColVal(4)),0)     & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    Else
        lgStrSQL = "INSERT INTO HDA041T( ATTEND_TYPE ,DECI_PLACE  , PROC_BAS ,"
        lgStrSQL = lgStrSQL & " ISRT_DT , ISRT_EMP_NO  , UPDT_DT , UPDT_EMP_NO)" 
        lgStrSQL = lgStrSQL & " VALUES(" 
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
        lgStrSQL = lgStrSQL & Trim(arrColVal(4))     & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    End If
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
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

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    If lgCurrentSpd = TAB1 Then      
       lgStrSQL = "UPDATE  HDA040T"
       lgStrSQL = lgStrSQL & " SET " 
       lgstrSQL = lgstrSQL & " BAS_AMT  = " & UNIConvNum(arrColVal(4), 0) & ","
       lgStrSQL = lgStrSQL & " BELOW    = " & FilterVar(UCase(arrColVal(5)), "''", "S")   & ","
       lgStrSQL = lgStrSQL & " PROC_BAS = " & FilterVar(UCase(arrColVal(6)), "''", "S")   
       lgStrSQL = lgStrSQL & " WHERE ALLOW_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    Else
       lgStrSQL = "UPDATE  HDA041T"
       lgStrSQL = lgStrSQL & " SET " 
       lgStrSQL = lgStrSQL & " DECI_PLACE = " & Trim(arrColVal(4))  & ","
       lgStrSQL = lgStrSQL & " PROC_BAS = " & FilterVar(UCase(arrColVal(5)), "''", "S")   
       lgStrSQL = lgStrSQL & " WHERE ATTEND_TYPE = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    End If

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

    If lgCurrentSpd = TAB1 Then      
       lgStrSQL = "DELETE  HDA040T"
       lgStrSQL = lgStrSQL & " WHERE "
       lgStrSQL = lgStrSQL & " ALLOW_CD  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    Else
       lgStrSQL = "DELETE  HDA041T"
       lgStrSQL = lgStrSQL & " WHERE "
       lgStrSQL = lgStrSQL & " ATTEND_TYPE = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    End If

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
                       lgStrSQL = "SELECT TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & " ALLOW_CD    ,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ",ALLOW_CD," & FilterVar("", "''", "S") &") ALLOW_nm, "
                       lgStrSQL = lgStrSQL & " BAS_AMT     ," 
                       lgStrSQL = lgStrSQL & " BELOW       ," 
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0051", "''", "S") & ",BELOW) BELOW_nm, "
                       lgStrSQL = lgStrSQL & " PROC_BAS    ,"                       
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0052", "''", "S") & ",PROC_BAS) PROC_BAS_nm "
                       lgStrSQL = lgStrSQL & " FROM  HDA040T "
                       lgStrSQL = lgStrSQL & " WHERE ALLOW_CD " & pComp & pCode  & " ORDER BY ALLOW_CD ASC"              
               Case "M"
                       lgStrSQL = "SELECT TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & " ATTEND_TYPE ,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0124", "''", "S") & ",ATTEND_TYPE) ATTEND_TYPE_nm, "
                       lgStrSQL = lgStrSQL & " DECI_PLACE     ," 
                       lgStrSQL = lgStrSQL & " PROC_BAS       ,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0052", "''", "S") & ",PROC_BAS) PROC_BAS_nm "
                       lgStrSQL = lgStrSQL & " FROM  HDA041T "
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

function get_Format( no )

	DIM retFormat
	
	retFormat = ""
	select case no
		case 0
			retFormat = "###" 
		case 1
			retFormat = "###" & gComNumDec & "0"
		case 2
			retFormat = "###" & gComNumDec & "00"
		case 3
			retFormat = "###" & gComNumDec & "000"
		case 4
			retFormat = "###" & gComNumDec & "0000"
	end select
	
	get_format = retFormat
	
end function

%>

<Script Language="VBScript">
  
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                If "<%=lgCurrentSpd%>" = "<%=TAB1%>" Then
                   .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                Else
                   .ggoSpread.Source     = .frm1.vspdData1
                End If
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .lgStrPrevKey1    = "<%=lgStrPrevKey%>"
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
