<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Login History Management, Biz
'*  2. Function Name        : 
'*  3. Program ID             : ZA003mb1.asp
'*  4. Program Name       : 
'*  5. Program Desc         : Lists login history information in details and manages locking status.
'*  6. Comproxy List        : ADO query program.
'*  7. Modified date(First) : 2002/05/20
'*  8. Modified date(Last) : 
'*  9. Modifier (First)        : Park Sang Hoon
'* 10. Modifier (Last)       : Park Sang Hoon
'* 11. Comment              :
'**********************************************************************************************
     
    Dim lgOpModeCRUD
 
    On Error Resume Next                                                             
    Err.Clear                                                                        

    Call HideStatusWnd                                                               

    Call LoadBasisGlobalInf()            
    '------ Developer Coding part (Start ) ------------------------------------------------------------------

    '------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    
    lgOpModeCRUD      = Request("txtMode")                                           
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
    End Select    

'=========================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             
    Err.Clear                                                                        

End Sub    
'=========================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             
    Err.Clear                                                                        
End Sub
'=========================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             
    Err.Clear                                                                        
End Sub

'=========================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLngMaxRow
    Dim iLngRow
    Dim strData
    Dim lgStrPrevKey
    Dim iZa012
    Dim E1_z_msg_logging
    Dim I1_z_msg_logging
    Dim StrNextKey
    Dim iSheetMaxRow

    Const ZA12_I1_LoginDateFrom = 0
    Const ZA12_I1_LoginDateTo = 1
    Const ZA12_I1_Severity1 = 2
    Const ZA12_I1_Severity2 = 3
    Const ZA12_I1_Severity3 = 4
    Const ZA12_I1_Severity4 = 5
    Const ZA12_I1_MsgType1 = 6
    Const ZA12_I1_MsgType2 = 7
    Const ZA12_I1_MsgType3 = 8
    Const ZA12_I1_UserId = 9
    Const ZA12_I1_MsgCd = 10
    Const ZA12_I1_ProgramId = 11
    Const ZA12_I1_Client = 12
    Const ZA12_I1_MajorCd1 = 13
    Const ZA12_I1_MajorCd2 = 14
    
    Const ZA12_E1_OccurDt = 0
    Const ZA12_E1_MsgCd = 1
    Const ZA12_E1_Msg = 2
    Const ZA12_E1_MsgTypeNm = 3
    Const ZA12_E1_UsrId = 4
    Const ZA12_E1_UsrNm = 5
    Const ZA12_E1_Severity = 6
    Const ZA12_E1_ProgramId = 7
    Const ZA12_E1_ClientId = 8
    Const ZA12_E1_ClientIp = 9
    Const ZA12_E1_hOccurDt = 10
    
    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa012 = Server.CreateObject("PZAG012.cListMsgLogging")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_msg_logging(ZA12_I1_MajorCd2)
    
    I1_z_msg_logging(ZA12_I1_LoginDateFrom)  = FilterVar(Request("txtFromDt"), "''", "S")    
    I1_z_msg_logging(ZA12_I1_LoginDateTo)    = FilterVar(Request("txtToDt"), "''", "S")
    
    I1_z_msg_logging(ZA12_I1_Severity1)             = FilterVar(Request("txtS1"), "''", "S")
    I1_z_msg_logging(ZA12_I1_Severity2)             = FilterVar(Request("txtS2"), "''", "S")
    I1_z_msg_logging(ZA12_I1_Severity3)             = FilterVar(Request("txtS3"), "''", "S")
    I1_z_msg_logging(ZA12_I1_Severity4)             = FilterVar(Request("txtS4"), "''", "S")
    
    I1_z_msg_logging(ZA12_I1_MsgType1)             = FilterVar(Request("txtM1"), "''", "S")
    I1_z_msg_logging(ZA12_I1_MsgType2)             = FilterVar(Request("txtM2"), "''", "S")        
    I1_z_msg_logging(ZA12_I1_MsgType3)             = FilterVar(Request("txtM3"), "''", "S")            


    If Request("txtUser") <> "" Then
        I1_z_msg_logging(ZA12_I1_UserId)       = FilterVar(Request("txtUser"), "''", "S")
    Else
        I1_z_msg_logging(ZA12_I1_UserId)       = FilterVar("%", "''", "S")
    End If

    If Request("txtMsg") <> "" Then
        I1_z_msg_logging(ZA12_I1_MsgCd)             = FilterVar(Request("txtMsg"), "''", "S")        
    Else
        I1_z_msg_logging(ZA12_I1_MsgCd)       = FilterVar("%", "''", "S")
    End If

    If Request("txtPgm") <> "" Then
        I1_z_msg_logging(ZA12_I1_ProgramId)             = FilterVar(Request("txtPgm"), "''", "S")                
    Else
        I1_z_msg_logging(ZA12_I1_ProgramId)       = FilterVar("%", "''", "S")
    End If
    
    If Request("txtClient") <> "" Then
        I1_z_msg_logging(ZA12_I1_Client)     = FilterVar(Request("txtClient"), "''", "S")
    Else
        I1_z_msg_logging(ZA12_I1_Client)     = FilterVar("%", "''", "S")
    End If

    I1_z_msg_logging(ZA12_I1_MajorCd1)     = FilterVar("B0007", "''", "S")
    I1_z_msg_logging(ZA12_I1_MajorCd2)     = FilterVar("B0008", "''", "S")    
    
    E1_z_msg_logging = iZa012.ZA_Read_Msg_Logging(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_msg_logging)
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZa012 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E1_z_msg_logging) Then
       Exit Sub
    End If
    
    Set iZa012 = Nothing    


    iLngMaxRow = CLng(Request("txtMaxRows"))
    For iLngRow = 0 To UBound(E1_z_msg_logging, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E1_z_msg_logging(ZA12_E1_hOccurDt,iLngRow)           
           Exit For
        End If
        
        strData = strData & Chr(11) & UNIDateClientFormat(E1_z_msg_logging(ZA12_E1_OccurDt, iLngRow))
        strData = strData & Chr(11) & SplitTime(E1_z_msg_logging(ZA12_E1_OccurDt,iLngRow))      
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_MsgCd,iLngRow))        
        strData = strData & Chr(11) & Replace(ConvSPChars(E1_z_msg_logging(ZA12_E1_Msg,iLngRow)),vbcrlf,"  ")        
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_MsgTypeNm, iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_UsrId,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_UsrNm, iLngRow))                
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_Severity, iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_ProgramId, iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_ClientId, iLngRow))                                
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_ClientIp, iLngRow))                
        strData = strData & Chr(11) & ConvSPChars(E1_z_msg_logging(ZA12_E1_hOccurDt, iLngRow))                        
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
    Response.Write ".frm1.hOccurDt.value = """            & StrNextKey             & """" & vbCr
    Response.Write ".lgStrPrevKey = """     & StrNextKey         & """" & vbCr                
    Response.Write ".DbQueryOk  "                          & vbCr
    Response.Write "End With  "                            & vbCr
    Response.Write "</Script>"                             & vbCr

End Sub    

'=========================================================================================================
Sub SubBizSaveMulti()
    Dim iZa024
    Dim iErrPosition

    On Error Resume Next                                                                 
    Err.Clear                                                                            
End Sub    

'=========================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             
    Err.Clear                                                                        
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'=========================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             
    Err.Clear                                                                        

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'=========================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             
    Err.Clear                                                                        

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'=========================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    
    On Error Resume Next

    '------ Developer Coding part (Start ) ------------------------------------------------------------------

    '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'=========================================================================================================
Sub CommonOnTransactionCommit()
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'=========================================================================================================
Sub CommonOnTransactionAbort()
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'=========================================================================================================
Sub SetErrorStatus()
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'=========================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             
    Err.Clear                                                                        

End Sub

'=========================================================================================================
Function SplitTime(Byval dtDateTime)
    SplitTime = Right("0" & Hour(dtDateTime), 2) & ":" _
            & Right("0" & Minute(dtDateTime), 2) & ":" _
            & Right("0" & Second(dtDateTime), 2)
End Function

%>

