<%@ LANGUAGE=VBSCript %>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Read User Info, Biz
'*  2. Function Name        : 
'*  3. Program ID			: ZA001mb1.asp
'*  4. Program Name			: 
'*  5. Program Desc         : Lists User information in details
'*  6. Comproxy List        : ADO query program.
'*  7. Modified date(First) : 2002/07/02
'*  8. Modified date(Last) 	: 
'*  9. Modifier (First)     : Park Sang Hoon
'* 10. Modifier (Last)		: Park Sang Hoon
'* 11. Comment              :
'**********************************************************************************************
     
    Dim lgOpModeCRUD
 
    On Error Resume Next                                                             
    Err.Clear                                                                        

    Call HideStatusWnd                                                               

    Call LoadBasisGlobalInf()    
    '------ Developer Coding part (Start ) ------------------------------------------------------------------

    '------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Call SubBizQueryMulti()                                                 '¢Ð: Query

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
    Dim iZa004
    Dim E1_z_usr_mast_rec
    Dim I1_z_usr_mast_rec
    Dim StrNextKey
    Dim iSheetMaxRow
    Dim Enc

    Const ZA04_E1_usr_id = 0
    Const ZA04_E1_usr_nm = 1
    Const ZA04_E1_usr_eng_nm = 2
    Const ZA04_E1_password = 3
    Const ZA04_E1_co_cd = 4
    Const ZA04_E1_co_cd_nm = 5
    Const ZA04_E1_log_on_gp = 6
    Const ZA04_E1_log_on_gp_nm = 7
    Const ZA04_E1_usr_valid_dt = 8
    Const ZA04_E1_interface_id = 9
    Const ZA04_E1_pwd_valid_dt = 10
    
    Const ZA04_I1_usr_id = 0

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa004 = Server.CreateObject("PZAG004.cListUsrMastRec")
    Set Enc = Server.CreateObject("EDCodeCom.EDCodeObj.1")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_usr_mast_rec(ZA04_I1_usr_id)
    
    If Request("txtUsrId") <> "" Then
        I1_z_usr_mast_rec(ZA04_I1_usr_id)  = Trim(Request("txtUsrId"))
    Else
        I1_z_usr_mast_rec(ZA04_I1_usr_id)  = "%"
    End If

    E1_z_usr_mast_rec = iZa004.ZA_Read_Usr_Mast_Rec(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_usr_mast_rec)

    If CheckSYSTEMError(Err,True) = True Then
        Response.Write "<Script Language=vbscript>"            & vbCr
        Response.Write "With Parent "                        & vbCr
        Response.Write ".frm1.txtUsrId.focus"            & vbCr
        Response.Write ".frm1.txtUsrId.select"            & vbCr
        Response.Write "End With  "                            & vbCr
        Response.Write "</Script>" & vbCr
        Set iZa004 = Nothing                                                         
        Exit Sub
    End If

    If IsEmpty(E1_z_usr_mast_rec) Then
       Exit Sub
    End If
    
    Set iZa004 = Nothing    
    Set Enc = Nothing

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "               & vbCr
       Response.Write "Parent.frm1.txtUsrId.value = """    & ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_usr_id,0))    & """" &vbCr
       Response.Write "Parent.frm1.txtUsrNm.value = """    & ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_usr_nm,0))    & """" &vbCr
       Response.Write "End With  "                 & vbCr 
       Response.Write "</Script>" & vbCr
    End If

    iLngMaxRow = CLng(Request("txtMaxRows"))    
    
    For iLngRow = 0 To UBound(E1_z_usr_mast_rec, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E1_z_usr_mast_rec(ZA04_E1_usr_id,iLngRow)           
           Exit For
        End If
    
        strData = strData & Chr(11) & ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_usr_id, iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_usr_nm, iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_usr_eng_nm,iLngRow))        
        strData = strData & Chr(11) & "******"
        strData = strData & Chr(11) & ConvSPChars(Trim(E1_z_usr_mast_rec(ZA04_E1_co_cd, iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_co_cd_nm, iLngRow))
        strData = strData & Chr(11) & ConvSPChars(Trim(E1_z_usr_mast_rec(ZA04_E1_log_on_gp,iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_log_on_gp_nm, iLngRow))        
        strData = strData & Chr(11) & UNIDateClientFormat(ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_usr_valid_dt,iLngRow)))        
        strData = strData & Chr(11) & ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_interface_id, iLngRow))        
        strData = strData & Chr(11) & UNIDateClientFormat(ConvSPChars(E1_z_usr_mast_rec(ZA04_E1_pwd_valid_dt,iLngRow)))        
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next

    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
    Response.Write ".frm1.hUsrId.value = """            & StrNextKey             & """" & vbCr
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

Function SplitTime(Byval dtDateTime)
    SplitTime = Right("0" & Hour(dtDateTime), 2) & ":" _
            & Right("0" & Minute(dtDateTime), 2) & ":" _
            & Right("0" & Second(dtDateTime), 2)
End Function
%>

