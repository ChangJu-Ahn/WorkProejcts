<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Login History Management, Biz
'*  2. Function Name        : 
'*  3. Program ID           : ZA003mb1.asp
'*  4. Program Name			: 
'*  5. Program Desc         : Lists login history information in details and manages locking status.
'*  6. Comproxy List        : ADO query program.
'*  7. Modified date(First) : 2002/05/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Sang Hoon
'* 10. Modifier (Last)      : Park Sang Hoon
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
    Dim iZa010
    Dim E1_z_log_in_hstry
    Dim I1_z_log_in_hstry
    Dim StrNextKey
    Dim iSheetMaxRow

    Const ZA10_I1_LoginDateFrom = 0
    Const ZA10_I1_LoginDateTo = 1
    Const ZA10_I1_User = 2
    Const ZA10_I1_Client = 3
    Const ZA10_I1_S1 = 4
    Const ZA10_I1_S2 = 5
    Const ZA10_I1_S3 = 6
    Const ZA10_I1_S4 = 7
    Const ZA10_I1_S5 = 8
    Const ZA10_I1_MajorCd = 9
    
    Const ZA10_E1_LoginDate = 0
    Const ZA10_E1_LogoutDate = 1
    Const ZA10_E1_UserId = 2
    Const ZA10_E1_UserNm = 3
    Const ZA10_E1_Status = 4
    Const ZA10_E1_ClientId = 5
    Const ZA10_E1_ClientIp = 6
    Const ZA10_E1_hLoginDateTo = 7    
    

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa010 = Server.CreateObject("PZAG010.cListLogInHstry")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_log_in_hstry(ZA10_I1_MajorCd)

    
    I1_z_log_in_hstry(ZA10_I1_LoginDateFrom)  = FilterVar(Request("txtFromDt"), "''", "S")    
    I1_z_log_in_hstry(ZA10_I1_LoginDateTo)    = FilterVar(Request("txtToDt"), "''", "S")
    
    If Request("txtUser") <> "" Then
        I1_z_log_in_hstry(ZA10_I1_User)       = FilterVar(Request("txtUser"), "''", "S")
    Else
        I1_z_log_in_hstry(ZA10_I1_User)       = FilterVar("%", "''", "S")
    End If
    
    If Request("txtClient") <> "" Then
        I1_z_log_in_hstry(ZA10_I1_Client)     = FilterVar(Request("txtClient"), "''", "S")
    Else
        I1_z_log_in_hstry(ZA10_I1_Client)     = FilterVar("%", "''", "S")
    End If
    
    I1_z_log_in_hstry(ZA10_I1_S1)             = FilterVar(Request("txtS1"), "''", "S")
    I1_z_log_in_hstry(ZA10_I1_S2)             = FilterVar(Request("txtS2"), "''", "S")
    I1_z_log_in_hstry(ZA10_I1_S3)             = FilterVar(Request("txtS3"), "''", "S")
    I1_z_log_in_hstry(ZA10_I1_S4)             = FilterVar(Request("txtS4"), "''", "S")
    I1_z_log_in_hstry(ZA10_I1_S5)             = FilterVar(Request("txtS5"), "''", "S")
    I1_z_log_in_hstry(ZA10_I1_MajorCd)        = FilterVar("Z0010", "''", "S")

    E1_z_log_in_hstry = iZa010.ZA_Read_Log_In_Hstry(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_log_in_hstry)
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZa010 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E1_z_log_in_hstry) Then
       Exit Sub
    End If
    
    Set iZa010 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E1_z_log_in_hstry, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E1_z_log_in_hstry(ZA10_E1_hLoginDateTo,iLngRow)           
           Exit For
        End If
        
        strData = strData & Chr(11) & UNIDateClientFormat(E1_z_log_in_hstry(ZA10_E1_LoginDate, iLngRow))
        strData = strData & Chr(11) & SplitTime(E1_z_log_in_hstry(ZA10_E1_LoginDate,iLngRow))
        strData = strData & Chr(11) & UNIDateClientFormat(E1_z_log_in_hstry(ZA10_E1_LogoutDate,   iLngRow))
        strData = strData & Chr(11) & SplitTime(E1_z_log_in_hstry(ZA10_E1_LogoutDate,   iLngRow))      
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_log_in_hstry(ZA10_E1_UserId,iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_log_in_hstry(ZA10_E1_UserNm,iLngRow))
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_log_in_hstry(ZA10_E1_Status, iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_log_in_hstry(ZA10_E1_ClientId,iLngRow))
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_log_in_hstry(ZA10_E1_ClientIp, iLngRow)))        
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_log_in_hstry(ZA10_E1_hLoginDateTo, iLngRow)))                
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
    Response.Write ".frm1.htxtToDt.value = """            & StrNextKey             & """" & vbCr
    Response.Write ".lgStrPrevKey = """     & StrNextKey         & """" & vbCr                
    Response.Write ".DbQueryOk  "                          & vbCr
    Response.Write "End With  "                            & vbCr
    Response.Write "</Script>"                             & vbCr

End Sub    

'=========================================================================================================
Sub SubBizSaveMulti()
    
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

Function SplitTime(Byval dtDateTime)
    SplitTime = Right("0" & Hour(dtDateTime), 2) & ":" _
            & Right("0" & Minute(dtDateTime), 2) & ":" _
            & Right("0" & Second(dtDateTime), 2)
End Function

%>

