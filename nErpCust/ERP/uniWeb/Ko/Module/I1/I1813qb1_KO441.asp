<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%
Response.Expires = -1                                                                '☜ : will expire the response immediately
Response.Buffer = True                                                               '☜ : The server does not send output to the client until all of the ASP scripts on the current page have been processed

%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adoQuery.vbs" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%


  On Error Resume Next
														'☜: 모든 작업 완료후 작업진행중 표시창을 Hide'
  Err.Clear
  Call HideStatusWnd 
    '---------------------------------------Common-----------------------------------------------------------                                                              '☜: Hide Processing message
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)


Dim IntRetCD
	
	
	
    Call SubOpenDB(lgObjConn)           
   
    Call SubCreateCommandObject(lgObjComm)
      
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
             
    End Select
    
    Call SubCloseDB(lgObjConn)


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
End Sub


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()    
    Dim iLoopMax
    Dim iKey1
    Dim strWhere    
    Dim iDx
                    
    On Error Resume Next    
    Err.Clear                                                               '☜: Clear Error status
               
	 strWhere = FilterVar(lgKeyStream(0), "''", "S") & ", " & _ 
				FilterVar(lgKeyStream(1), "''", "S") & ", " & _ 
				FilterVar(lgKeyStream(2), "''", "S") & ", " & _ 
				FilterVar(lgKeyStream(3), "''", "S") & ", " & _ 
				FilterVar(lgKeyStream(4), "''", "S")


    ' 프로시져를 통해 차이나는 품목들을 저장하는 테이블에 값을 insert한다.
    call  SubBizBatch()
    Call SubMakeSQLStatements("MR", strWhere, "X", C_LIKE)                              '☜ : Make sql statements   
    
  '  Call GetSum
                          
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       lgStrPrevKeyIndex = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
    Else
    
       Call SubSkipRs(lgObjRs, lgMaxCount * lgStrPrevKeyIndex)       
       
       lgStrData = ""
       iDx = 1

       Do While Not lgObjRs.EOF        
           lgStrData = lgStrData & Chr(11) & lgObjRs(0)
           lgStrData = lgStrData & Chr(11) & lgObjRs(1)
           lgStrData = lgStrData & Chr(11) & lgObjRs(2)
           lgStrData = lgStrData & Chr(11) & lgObjRs(3)
           lgStrData = lgStrData & Chr(11) & lgObjRs(4)
           lgStrData = lgStrData & Chr(11) & lgObjRs(5)
           lgStrData = lgStrData & Chr(11) & lgObjRs(6)  
           
           lgStrData = lgStrData & Chr(11) & lgLngMaxRow + iDx
           lgStrData = lgStrData & Chr(11) & Chr(12)                                            
           
           lgObjRs.MoveNext
          
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------                                                                                     
            iDx =  iDx + 1  
            If iDx > lgMaxCount Then               
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1 
               Exit Do
            End If                          
            
        Loop
        
    End If
            
      'lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1       
      If iDx <= lgMaxCount Then
         lgStrPrevKeyIndex = ""
      End If   
      
      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
         ObjectContext.SetAbort
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
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
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
Sub SubBizSaveMultiCreate(arrColVal)

	Dim iCnt, strFg, strType, strOrgId
    On Error Resume Next   
    Err.Clear                                                                        '☜: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
 
 On Error Resume Next 
 Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
       
  
     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
        
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
   
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
    
End Sub



'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear    

	 
    With lgObjComm
        .CommandText = "usp_I1813QA_KO441"
        .CommandType = adCmdStoredProc
        
      
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@yyyymm",	advarXchar,adParamInput,6, lgKeyStream(0))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@item_cd ",	advarXchar,adParamInput,18, lgKeyStream(1))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, lgKeyStream(2))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sl_cd"   ,  advarXchar,adParamInput,10,lgKeyStream(3))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@item_acct",	advarXchar,adParamInput,10,lgKeyStream(4))
        lgObjComm.Execute ,, adExecuteNoRecords
        

    End With

    If  Err.number = 0 Then

    Else           
        Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
    
    
End Sub	

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)


    
    
 Dim iSelCount
 Dim strZeroFg1
    
   ' call svrmsgbox("ㅁ", vbinformation, i_mkscript)
    
    Select Case Mid(pDataType,1,1)
        Case "M"
                  
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
          ' strZeroFg1 = lgKeyStream(9)
           pCode = pCode

           Select Case Mid(pDataType,2,1)           
               Case "R"
                        lgStrSQL = " SELECT *  "
						lgStrSQL = lgStrSQL & " FROM I1813QA_KO441 "
						lgStrSQL = lgStrSQL & " ORDER BY 1 "
           End Select             
    End Select
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
    lgErrorStatus     = "YES"       
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                 
    Err.Clear                            

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
        Case "MB"
			ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub


Response.Write "<Script language=vbs> " & vbCr         
Response.Write " With Parent "      	& vbCr
Response.Write "	If """ & lgErrorStatus & """ = ""NO"" And """ & IntRetCd & """ <> -1 Then "	& vbCr
Response.Write "    .lgStrPrevKey  = """ & NextKey1 & """" & vbCr  
Response.Write "	.ggoSpread.Source	= .frm1.vspdData "				& vbCr
Response.Write "	.ggoSpread.SSShowDataByClip  """ & lgstrData  & """"        & vbCr
Response.Write "		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
Response.Write "			.DbQuery						"				& vbCr
Response.Write "		Else								"				& vbCr
Response.Write "			.DbQueryOK						"				& vbCr
Response.Write "		End If								"				& vbCr
Response.Write "		.frm1.vspdData.focus				"				& vbCr
Response.Write "    End If								"				& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 
Response.End     

%>    
