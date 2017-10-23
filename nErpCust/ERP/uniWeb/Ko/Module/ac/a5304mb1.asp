<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
    Dim lgStrPrevKey
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                     '☜: Hide Processing message
	Dim strYear

	Call LoadBasisGlobalInf()

    lgErrorStatus  = ""
    lgKeyStream    = Split(Request("txtKeyStream"),gColSep) 
    lgStrPrevKey   = Trim(Request("lgStrPrevKey"))                   '☜: Next Key
    strYear   = Trim(Request("txtYear"))                   '☜: Next Key
 
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim lgStrSQL
    Dim lgstrData
    Dim iDx
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
    
    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrSQL = " SELECT YYYY, SEQ, ACCT_FG, F_ACCT,B.ACCT_NM F_ACCT_NM, T_ACCT, C.ACCT_NM T_ACCT_NM"
	lgStrSQL = lgStrSQL & " FROM A_ACCT_CLOSE_TRANSFER  A "
	lgStrSQL = lgStrSQL & "      LEFT JOIN A_ACCT B ON A.F_ACCT = B.ACCT_CD"
	lgStrSQL = lgStrSQL & "		 LEFT JOIN A_ACCT C ON A.T_ACCT = C.ACCT_CD "
	lgStrSQL = lgStrSQL & " WHERE A.YYYY =  " & FilterVar(strYear, "''", "S") & " "
	'Response.Write lgStrSQL
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
        lgStrPrevKey  = ""
        lgErrorStatus = "YES"
        Exit Sub 
    Else    
       iDx = 1		
       lgstrData = ""
       lgLngMaxRow       = CLng(Request("txtMaxRows"))

       Do While Not lgObjRs.EOF
          lgstrData = lgstrData & Chr(11) & ""
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("F_ACCT"))
          lgstrData = lgstrData & Chr(11) & ""
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("F_ACCT"))
          lgstrData = lgstrData & Chr(11) & ""
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("F_ACCT_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("T_ACCT"))
          lgstrData = lgstrData & Chr(11) & ""
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("T_ACCT_NM"))
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)

          lgObjRs.MoveNext

          iDx =  iDx + 1
          If iDx > C_SHEETMAXROWS_D Then
             Exit Do
         End If   
      Loop 
    End If
    
    If Not lgObjRs.EOF Then
       lgStrPrevKey = lgObjRs("StudentID")
    Else
       lgStrPrevKey = ""
    End If
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
       Exit Sub
    End If   
       
    If lgErrorStatus = "" Then
       Response.Write  " <Script Language=vbscript>                                  " & vbCr
       Response.Write  "    parent.ggoSpread.Source     = Parent.frm1.vspdData       " & vbCr
       Response.Write  "    Parent.lgStrPrevKey         = """ & ConvSPChars(lgStrPrevKey)    & """" & vbCr
       Response.Write  "    parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim itxtSpread
    Dim arrRowVal
    Dim arrColVal
    Dim lgErrorPos
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgErrorPos        = ""                                                           '☜: Set to space

    itxtSpread = Trim(Request("txtSpread"))
    
    If itxtSpread = "" Then
       Exit Sub
    End If   
    
	arrRowVal = Split(itxtSpread, gRowSep)                                 '☜: Split Row    data
	
    For iDx = 0 To UBound(arrRowVal,1) - 1
        arrColVal = Split(arrRowVal(iDx), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C" :  Call SubBizSaveMultiCreate(arrColVal)                        '☜: Create
            Case "U" :  Call SubBizSaveMultiUpdate(arrColVal)                        '☜: Update
            Case "D" :  Call SubBizSaveMultiDelete(arrColVal)                        '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next
    
    If lgErrorStatus = "YES" Then
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  " Parent.SubSetErrPos(Trim(""" & lgErrorPos & """))" & vbCr
       Response.Write  " </Script>                  " & vbCr
    Else
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  " Parent.DBSaveOk            " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If
    
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    Dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    'Response.Write "#01#:" & arrcolval(02)
    'Response.Write "#03#" & arrcolval(03)
    'Response.Write "#04#" & arrcolval(04)
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = " Select ACCT_FG from A_ACCT_CLOSE_TRANSFER "
    lgStrSQL = lgStrSQL & " WHERE YYYY			=  " & FilterVar(strYear, "''", "S") & " "
    lgStrSQL = lgStrSQL & "  AND  F_ACCT		=  " & FilterVar(arrcolval(03), "''", "S") & " "
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                   'R(Read) X(CursorType) X(LockType) 
		If isNull(lgObjRs("ACCT_FG")) = False Then
	   		Call DisplayMsgBox("110102", vbInformation, "", "", I_MKSCRIPT)
	   		
	   		Exit Sub
	   End If
	End if
	
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	Call SubCloseRs(lgObjRs)
    
    lgStrSQL = "INSERT INTO A_ACCT_CLOSE_TRANSFER("
    lgStrSQL = lgStrSQL & " YYYY			,"
    lgStrSQL = lgStrSQL & " SEQ				, ACCT_FG  ,"
    lgStrSQL = lgStrSQL & " F_ACCT			, T_ACCT   , "
    lgStrSQL = lgStrSQL & " INSRT_USER_ID   , INSRT_DT , "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID    , UPDT_DT  )"    '16
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & " " & FilterVar(strYear, "''", "S") & " ,"
    lgStrSQL = lgStrSQL & "" & FilterVar("1", "''", "S") & " ,"
    lgStrSQL = lgStrSQL & "" & FilterVar("1", "''", "S") & " ,"
    lgStrSQL = lgStrSQL & " " & FilterVar(arrcolval(03), "''", "S") & " ,"
    lgStrSQL = lgStrSQL & " " & FilterVar(arrcolval(04), "''", "S") & " ,"
    lgStrSQL = lgStrSQL & FilterVar(gUsrID, "''", "S")                       & "," 
    lgStrSQL = lgStrSQL & " GetDate()," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrID, "''", "S")                       & "," 
    lgStrSQL = lgStrSQL & " GetDate())" 
    
    'Response.Write "Create : " & lgStrSQL
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    Dim lgStrSQL
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	'수정시 수정하는 년도가 마감되었으면 결산마감되었다 체크하는 로직 
	'%1 이미 결산되었습니다.
    lgStrSQL = "SELECT FISC_YR FROM A_GL_SUM " & vbcr
    lgStrSQL = lgStrSQL & " WHERE FISC_DT = " & FilterVar("00", "''", "S") & " " & vbcr
    lgStrSQL = lgStrSQL & " AND FISC_YR + FISC_MNTH >= (SELECT CONVERT (VARCHAR(06),DATEADD(YEAR,1,CONVERT(DATETIME,(SELECT   " & FilterVar(strYear, "''", "S") & "  + SUBSTRING(CONVERT(VARCHAR(6),FISC_START_DT,112),5,2) + " & FilterVar("01", "''", "S") & " ))),112) FROM B_COMPANY) "
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
		Call DisplayMsgBox("111192", vbInformation, "", "", I_MKSCRIPT)
		lgErrorStatus    = "YES"
		ObjectContext.SetAbort
		Call SetErrorStatus
		response.end
	End If

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE A_ACCT_CLOSE_TRANSFER SET "
    lgStrSQL = lgStrSQL & " F_ACCT			=  " & FilterVar(arrcolval(03), "''", "S") & " ,"
    lgStrSQL = lgStrSQL & " T_ACCT			=  " & FilterVar(arrcolval(04), "''", "S") & " ,"
    lgStrSQL = lgStrSQL & " UPDT_USER_ID    = " & FilterVar(gUsrID, "''", "S")		& ","             
    lgStrSQL = lgStrSQL & " UPDT_DT			= GetDate() " 
    lgStrSQL = lgStrSQL & " WHERE YYYY		=  " & FilterVar(strYear, "''", "S") & " "
    lgStrSQL = lgStrSQL & "  AND  F_ACCT	=  " & FilterVar(arrcolval(02), "''", "S") & " "
    
    'Response.Write "Update:" & lgStrSQL
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    Dim lgStrSQL

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "DELETE  FROM A_ACCT_CLOSE_TRANSFER "
    lgStrSQL = lgStrSQL & " WHERE YYYY	=   " & FilterVar(strYear, "''", "S") & " "
    lgStrSQL = lgStrSQL & "  AND  F_ACCT	=  " & FilterVar(arrcolval(02), "''", "S") & " "
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
    
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
