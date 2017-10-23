<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Login History Management, Biz
'*  2. Function Name        : 
'*  3. Program ID             : ZA014rb2.asp
'*  4. Program Name       : 
'*  5. Program Desc         : Lists login history information in details and manages locking status.
'*  6. Comproxy List        : ADO query program.
'*  7. Modified date(First) : 2002/07/15
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
    Dim iZa005

    Dim E1_z_usr_mast_rec
    Dim E2_z_usr_mast_rec
    Dim E3_z_usr_mast_rec    
    Dim E6_z_usr_mast_rec

    Dim I1_z_usr_mast_rec
    Dim I2_z_usr_mast_rec
    Dim I3_z_usr_mast_rec    
    Dim I6_z_usr_mast_rec
    
    Dim StrNextKey
    Dim iSheetMaxRow

    Const ZA05_I1_usr_id = 0
    
    Const ZA05_I6_usr_id = 0
    Const ZA05_I6_mnu_id = 1
    Const ZA05_I6_mnu_type = 2
    Const ZA05_I6_major_cd1 = 3
    Const ZA05_I6_major_cd2 = 4

    Const ZA05_E1_usr_id = 0
    Const ZA05_E1_usr_nm = 1
    Const ZA05_E1_usr_eng_nm = 2
    Const ZA05_E1_usr_valid_dt = 3
    Const ZA05_E1_password = 4
    Const ZA05_E1_log_on_gp = 5
    Const ZA05_E1_log_on_gp_nm = 6
    Const ZA05_E1_co_cd = 7
    Const ZA05_E1_co_cd_nm = 8
    Const ZA05_E1_use_yn = 9
    Const ZA05_E1_interface_id = 10
    Const ZA05_E1_husr_valid_dt = 11

    Const ZA05_E6_mnu_id = 0
    Const ZA05_E6_mnu_nm = 1
    Const ZA05_E6_mnu_type = 2
    Const ZA05_E6_action_id = 3
    Const ZA05_E6_upper_mnu_id = 4
    
    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa005 = Server.CreateObject("PZAG005.cLookUpUsrMastRec")

    If CheckSYSTEMError(Err,True) = True Then
        Exit Sub
    End If


    Redim I6_z_usr_mast_rec(ZA05_I6_major_cd2)
    
    I6_z_usr_mast_rec(ZA05_I6_usr_id)  = Trim(Request("txtUsr"))
    I6_z_usr_mast_rec(ZA05_I6_mnu_id)    = Trim(Request("txtMnuID"))
    
    I6_z_usr_mast_rec(ZA05_I6_mnu_type)             = Trim(Request("txtMnuType"))
    I6_z_usr_mast_rec(ZA05_I6_major_cd1)             = "Z0006"
    I6_z_usr_mast_rec(ZA05_I6_major_cd2)             = "Z0013"

    Redim I1_z_usr_mast_rec(ZA05_I1_usr_id)    
    I1_z_usr_mast_rec(ZA05_I1_usr_id)       = Trim(Request("txtUsr"))

    E6_z_usr_mast_rec = iZa005.ZA_List_Auth_Gen(gStrGlobalCollection, C_SHEETMAXROWS_D, I6_z_usr_mast_rec)
    Call iZa005.ZA_Lookup_Usr_Mast_Rec(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_usr_mast_rec, I2_z_usr_mast_rec, I3_z_usr_mast_rec, E1_z_usr_mast_rec, E2_z_usr_mast_rec, E3_z_usr_mast_rec)    
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZa005 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E6_z_usr_mast_rec) Then
       Exit Sub
    End If

    If IsEmpty(E1_z_usr_mast_rec) Then
       Exit Sub
    End If
    
    Set iZa005 = Nothing    


    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E6_z_usr_mast_rec, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E6_z_usr_mast_rec(ZA05_E6_mnu_id,iLngRow)           
           Exit For
        End If
        
        strData = strData & Chr(11) & ConvSPChars(E6_z_usr_mast_rec(ZA05_E6_mnu_id, iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E6_z_usr_mast_rec(ZA05_E6_mnu_nm,iLngRow))      
        strData = strData & Chr(11) & ConvSPChars(E6_z_usr_mast_rec(ZA05_E6_mnu_type,iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E6_z_usr_mast_rec(ZA05_E6_action_id,iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E6_z_usr_mast_rec(ZA05_E6_upper_mnu_id, iLngRow))        
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr    
    Response.Write ".frm1.htxtUsr.value = """            & Trim(Request("txtUsr"))      & """" & vbCr
    Response.Write ".frm1.txtUsrNm.value = """            & UCase(ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_nm,0)))      & """" & vbCr        
    Response.Write ".frm1.htxtMnuID.value = """            & StrNextKey             & """" & vbCr
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

 