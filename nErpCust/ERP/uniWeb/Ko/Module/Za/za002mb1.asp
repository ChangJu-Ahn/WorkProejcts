<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%
     
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
        '    Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
        '    Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
         '    Call SubBizDelete()
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
    Dim iZa007
    Dim E1_z_logon_gp
    Dim I1_z_logon_gp
    Dim StrNextKey
    Dim iSheetMaxRow
    
    Const ZA07_E1_LogonGp = 0
    Const ZA07_E1_LogonGpNm = 1
    Const ZA07_E1_ApServerId = 2
    Const ZA07_E1_PortNo = 3
    Const ZA07_E1_DbServerIp = 4
    Const ZA07_E1_DbNm = 5
    Const ZA07_E1_DsnNo = 6
    Const ZA07_E1_UsersNo = 7

    Const ZA07_I1_LogonGp = 0


    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa007 = Server.CreateObject("PZAG007.cListLogonGp")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_logon_gp(ZA07_I1_LogonGp)

    If Request("txtLogonGp") <> "" Then
        I1_z_logon_gp(ZA07_I1_LogonGp)  = Request("txtLogonGp")
    Else
        I1_z_logon_gp(ZA07_I1_LogonGp)  = ""
    End If    


    E1_z_logon_gp = iZa007.Z_Read_Logon_Gp(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_logon_gp)
        
    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"            & vbCr
       Response.Write "With Parent "                        & vbCr
       Response.Write ".frm1.txtLogonGp.focus"            & vbCr
       Response.Write ".frm1.txtLogonGp.select"            & vbCr
       Response.Write ".frm1.txtLogonGpNm.value = """""        & vbCr
       Response.Write "End With  "                            & vbCr
       Response.Write "</Script>" & vbCr        
    
       Set iZa007 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E1_z_logon_gp) Then
       Exit Sub
    End If

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "               & vbCr
       Response.Write "Parent.frm1.txtLogonGp.value = """ & Trim(ConvSPChars(E1_z_logon_gp(ZA07_E1_LogonGp,0)))  & """" &vbCr
       Response.Write "Parent.frm1.txtLogonGpNm.value = """ & ConvSPChars(E1_z_logon_gp(ZA07_E1_LogonGpNm,0)) & """" &vbCr
       Response.Write "End With  "                 & vbCr 
       Response.Write "</Script>" & vbCr
    End If
    
    Set iZa007 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E1_z_logon_gp, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E1_z_logon_gp(ZA07_E1_LogonGp,iLngRow)
           Exit For
        End If

        strData = strData & Chr(11) & Trim(UCase(ConvSPChars(E1_z_logon_gp(ZA07_E1_LogonGp, iLngRow))))
        strData = strData & Chr(11) & ConvSPChars(E1_z_logon_gp(ZA07_E1_LogonGpNm,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_logon_gp(ZA07_E1_ApServerId,   iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_logon_gp(ZA07_E1_PortNo,   iLngRow))      
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_logon_gp(ZA07_E1_DbServerIp,iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_logon_gp(ZA07_E1_DbNm,iLngRow))
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_logon_gp(ZA07_E1_DsnNo, iLngRow)))
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_logon_gp(ZA07_E1_UsersNo, iLngRow)))                                
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
    Response.Write ".lgStrPrevKey = """     & StrNextKey         & """" & vbCr                
    Response.Write ".frm1.hLogonGp.value = """           & StrNextKey  & """" & vbCr
    Response.Write ".DbQueryOk  "                          & vbCr
    Response.Write "End With  "                            & vbCr
    Response.Write "</Script>"                             & vbCr

End Sub    

'=========================================================================================================
Sub SubBizSaveMulti()
    Dim iZa006
    Dim iErrPosition

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    Set iZa006 = Server.CreateObject("PZAG006.cCtrlLogonGp")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call iZa006.ZA_Control_Logon_Gp(gStrGlobalCollection,Request("txtSpread"),iErrPosition)

    If CheckSYSTEMError2(Err, True, iErrPosition & "Çà","","","","") = True Then
       Set iZa006 = Nothing                                                         
       Exit Sub
    End If
    
    Set iZa006 = Nothing                                                   
    
    Response.Write "<Script Language=vbscript>"   & vbCr
    Response.Write "Parent.DbSaveOk  "            & vbCr
    Response.Write "</Script>"                    & vbCr
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

%>

