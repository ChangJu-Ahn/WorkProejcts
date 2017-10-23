<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%
	Dim lgStrPrevKey
	
    Dim adCmdText
    Dim adExcuteNoRecords
    Dim lgLngMaxRow2
	Const C_SHEETMAXROWS_D = 10000
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)    
    
    lgStrPrevKey      = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
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
    Dim strWhere
    Dim strSub_type   
    Dim strPAY_CD  
    Dim strDept_nm  
       
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
        
    strWhere =""
           
    Call SubMakeSQLStatements("MR",strWhere,"X","=")                                 '¡Ù : Make sql statements
'Call SvrMsgBox(lgStrSQL,0,1)
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

       ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
         
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("USR_ID"))) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("USR_NM"))) 

	    If lgObjRs("BA_YN") = "Y" then
            lgstrData = lgstrData & Chr(11) & "1"
            Else
            lgstrData = lgstrData & Chr(11) & "0"
	    End If

	    If lgObjRs("PL_YN") = "Y" then
            lgstrData = lgstrData & Chr(11) & "1"
            Else
            lgstrData = lgstrData & Chr(11) & "0"
	    End If

	    If lgObjRs("SG_YN") = "Y" then
            lgstrData = lgstrData & Chr(11) & "1"
            Else
            lgstrData = lgstrData & Chr(11) & "0"
	    End If

	    If lgObjRs("SO_YN") = "Y" then
            lgstrData = lgstrData & Chr(11) & "1"
            Else
            lgstrData = lgstrData & Chr(11) & "0"
	    End If

	    If lgObjRs("PG_YN") = "Y" then
            lgstrData = lgstrData & Chr(11) & "1"
            Else
            lgstrData = lgstrData & Chr(11) & "0"
	    End If

	    If lgObjRs("PO_YN") = "Y" then
            lgstrData = lgstrData & Chr(11) & "1"
            Else
            lgstrData = lgstrData & Chr(11) & "0"
	    End If

            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("NO"))) 
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            
            iDx =  iDx + 1
           ' If iDx > C_SHEETMAXROWS_D Then
            '   lgStrPrevKey = lgStrPrevKey + 1
            '   Exit Do
           ' End If   
        Loop 
    End If
    
   ' If iDx <= C_SHEETMAXROWS_D Then
   '    lgStrPrevKey = ""
   ' End If   

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
        'Call SvrMsgBox(lgErrorStatus,0,1)
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
    Err.Clear   
    '--------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

        
    lgStrSQL = "INSERT INTO Z_USR_ORG_MAST_KO441("
    lgStrSQL = lgStrSQL & " USR_ID,BA_YN,PL_YN," 
    lgStrSQL = lgStrSQL & " SG_YN, SO_YN ,PG_YN, PO_YN, INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT ) " 
    lgStrSQL = lgStrSQL & " VALUES( " 
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"''","S")           & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"''","S")           & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"''","S")           & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"''","S")           & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(6))),"''","S")           & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(7))),"''","S")           & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(8))),"''","S")           & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                                & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")                      & ","  
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                                & ","
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & " )"  

   ' Call svrmsgbox (lgStrSQL, vbinformation, i_mkscript)

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	
End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear   
    
    '--------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
	'    Response.Write                
    lgStrSQL = "UPDATE  Z_USR_ORG_MAST_KO441"
    lgStrSQL = lgStrSQL & " SET BA_YN	=  " & FilterVar(UCase(arrColVal(3))   , "''", "S") & ", "   
    lgStrSQL = lgStrSQL & "  PL_YN	=  " & FilterVar(UCase(arrColVal(4))   , "''", "S") & ", "  
    lgStrSQL = lgStrSQL & "  SG_YN	=  " & FilterVar(UCase(arrColVal(5))   , "''", "S") & ", "  
    lgStrSQL = lgStrSQL & "  SO_YN	=  " & FilterVar(UCase(arrColVal(6))   , "''", "S") & ","  
    lgStrSQL = lgStrSQL & "  PG_YN	=  " & FilterVar(UCase(arrColVal(7))   , "''", "S") & ", "  
    lgStrSQL = lgStrSQL & "  PO_YN	=  " & FilterVar(UCase(arrColVal(8))   , "''", "S") & " "  
    lgStrSQL = lgStrSQL & " WHERE USR_ID = " & FilterVar(Trim(UCase(arrColVal(2))),"''","S")   & " "


    
'	Call svrmsgbox (lgStrSQL, vbinformation, i_mkscript)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
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
    lgStrSQL = "DELETE  Z_USR_ORG_MAST_KO441 "
    lgStrSQL = lgStrSQL & " WHERE USR_ID =  " & FilterVar(Trim(UCase(arrColVal(2))),"''","S")   & " "
     
'    Call svrmsgbox (lgStrSQL, vbinformation, i_mkscript)  
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
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
        
         '  iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                      ' lgStrSQL = "Select TOP " & iSelCount                                                                     
                       

                       lgStrSQL = lgStrSQL & " SELECT A.USR_ID, "
                       lgStrSQL = lgStrSQL & "        A.USR_NM, "
                       lgStrSQL = lgStrSQL & "        B.BA_YN, "
                       lgStrSQL = lgStrSQL & "        B.PL_YN, "
                       lgStrSQL = lgStrSQL & "        B.SG_YN, "
                       lgStrSQL = lgStrSQL & "        B.SO_YN,B.PG_YN,B.PO_YN, "
                       lgStrSQL = lgStrSQL & "        NO = B.USR_ID "
                       lgStrSQL = lgStrSQL & "   FROM Z_USR_MAST_REC A(NOLOCK)  "
                       lgStrSQL = lgStrSQL & "   LEFT OUTER JOIN  Z_USR_ORG_MAST_KO441 B(NOLOCK) ON A.USR_ID = B.USR_ID "
                        lgStrSQL = lgStrSQL & " WHERE A.USR_ID LIKE " & FilterVar(lgKeyStream(0), "'%'", "S")                       
                       lgStrSQL = lgStrSQL & " ORDER BY A.USR_ID "
                       
              Case "S"
                      ' lgStrSQL = "Select TOP " & iSelCount                                                                     
						
					   lgStrSQL = ""	
                       lgStrSQL = lgStrSQL & " SELECT YYYYMMDD, "
                       lgStrSQL = lgStrSQL & "        Disposition, "
                       lgStrSQL = lgStrSQL & "        REPAIR, "
                       lgStrSQL = lgStrSQL & "        REMARK"
                       lgStrSQL = lgStrSQL & "   FROM XQ_BAS_SERIAL_INF_KO382(NOLOCK) "
                       lgStrSQL = lgStrSQL & "  WHERE SERIAL_NO = " & FilterVar(lgKeyStream(0), "''", "S")                       
                       lgStrSQL = lgStrSQL & "  ORDER BY YYYYMMDD "         
                          ' call svrmsgbox(lgStrSQL,0,1)
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


'============================================================================================================
' Name : SplitTime
'============================================================================================================
Function SplitTime(Byval dtDateTime)

    If IsNull(dtDateTime)  Then
        SplitTime = ""
        Exit Function
    End If

    SplitTime = Right("0" & Hour(dtDateTime), 2) & ":" _
            & Right("0" & Minute(dtDateTime), 2) & ":" _
            & Right("0" & Second(dtDateTime), 2)
            
End Function

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
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
