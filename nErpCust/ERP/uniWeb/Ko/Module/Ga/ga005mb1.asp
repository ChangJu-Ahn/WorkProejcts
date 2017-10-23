<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->


<%
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear 
    
    Call LoadBasisGlobalInf()                                                                       '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorStatus1     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)

    Dim lgStrPrevKey
    Dim txtBizUnitNm

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6    
    Dim lgErrorStatus1
    
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)
    'If Len(Trim(Request("lgStrPrevKey")))  Then                                        '¢Ð : Chnage Nextkey str into int value
    '   If Isnumeric(lgStrPrevKey) Then
    '       lgStrPrevKey = CInt(lgStrPrevKey)          
    '    End If   
    ' Else   
    '    lgStrPrevKey = 0
    'End If

   	Const C_SHEETMAXROWS_D  = 100 
   
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '¢Ð: Max fetched data at a time
    

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
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
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
    Dim strBizUnitCd
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

	If Trim(Request("txtBizUnitCd")) <> "" Then   
	  strBizUnitCd = FilterVar(Trim(Request("txtBizUnitCd")),"","SNM")
	End If

	
   IF strBizUnitCd <> "" Then 
		Call CommonQueryRs("BIZ_UNIT_NM","B_BIZ_UNIT","BIZ_UNIT_CD =  " & FilterVar(strBizUnitCd , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		

		If Trim(Replace(lgF0,Chr(11),"X")) = "X" then
			txtBizUnitNm = ""	
		Else
			txtBizUnitNm = Trim(Replace(lgF0,Chr(11),""))
		End if
		
	END IF
	
 
    Call SubMakeSQLStatements(strBizUnitCd)       
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)				'¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKey)

        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("cost_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("cost_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_unit_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_unit_nm"))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
   		    lgObjRs.MoveNext
'------ Developer Coding part (End   ) ------------------------------------------------------------------

            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
       
        Loop 
    End If
    
    If iDx <= lgMaxCount Then
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
    
	lgStrSQL = "SET XACT_ABORT ON  BEGIN TRAN "
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
'            Case "U"
'                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           
           IF lgErrorStatus1 = "YES" Then
				lgStrSQL = "ROLLBACK TRAN "
				lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
				Call DisplayMsgBox("124400", vbInformation, "", "", I_MKSCRIPT)			
			END IF	
		           
           Exit For
        End If
        
    Next
    
    IF lgErrorStatus1 = "YES" Then
		Response.End
	ENd If
	
	
    IF lgErrorStatus1 <> "YES" and lgErrorStatus <> "YES" Then
		lgStrSQL = "COMMIT TRAN  "
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	ENd IF

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

	Call CommonQueryRs("COST_CD", "B_COST_CENTER","COST_TYPE <> " & FilterVar("C", "''", "S") & "  and COST_CD = " & FilterVar(UCase(arrColVal(2)), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Trim(Replace(lgF0,Chr(11),"X")) = "X" then
		lgErrorStatus1 = "YES"
		Call SetErrorStatus
		Exit Sub
	End if
   
    lgStrSQL = "INSERT INTO G_DSTB_CC("
    lgStrSQL = lgStrSQL & " COST_CD," 
    lgStrSQL = lgStrSQL & " COST_TYPE," 
    
    lgStrSQL = lgStrSQL & " INSRT_USER_ID      ," 
    lgStrSQL = lgStrSQL & " INSRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      ," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
   
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","

	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & "getdate()," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")
    lgStrSQL = lgStrSQL & ",getdate())"
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
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

 End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    lgStrSQL = "DELETE  G_DSTB_CC "
    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & " COST_CD  = " & FilterVar(UCase(arrColVal(2)), "''", "S")


    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pCode)
    Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
        
	     iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKey + 1
           
		 lgStrSQL = " Select  a.cost_cd,  b.cost_nm,  c.biz_unit_cd,  c.biz_unit_nm   " 
		 lgStrSQL = lgStrSQL & " from g_dstb_cc a, b_cost_center b, b_biz_unit c  "
		 lgStrSQL = lgStrSQL & " where a.cost_cd = b.cost_cd and b.biz_unit_cd = c.biz_unit_cd  " 
		 lgStrSQL = lgStrSQL & " and c.biz_unit_cd LIKE   " & FilterVar(pCode & "%", "''", "S") & "  ORDER BY a.cost_cd ASC"
 
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
                    'Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)			'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
						'Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    'Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)			'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       'Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
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
          Parent.frm1.txtBizUnitNm.Value = "<%=ConvSPChars(txtBizUnitNm)%>"            
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
