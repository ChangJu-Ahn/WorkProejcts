<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Read User Info, Biz
'*  2. Function Name        : 
'*  3. Program ID             : ZA001mb1.asp
'*  4. Program Name       : 
'*  5. Program Desc         : Lists User information in details
'*  6. Comproxy List        : ADO query program.
'*  7. Modified date(First) : 2002/07/02
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
    Call SubBizQueryMulti()

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
    Dim iZa027
    Dim E4_z_usr_org_mast
    Dim I4_z_usr_org_mast
    Dim StrNextKey
    Dim iSheetMaxRow
    Dim Enc
    Dim lgOrgType
    Dim lgOrgCd
    
    Const ZA27_E4_usr_id = 0
    Const ZA27_E4_usr_nm = 1
    Const ZA27_E4_usr_eng_nm = 2
    Const ZA27_E4_password = 3
    Const ZA27_E4_co_cd = 4
    Const ZA27_E4_co_cd_nm = 5
    Const ZA27_E4_log_on_gp = 6
    Const ZA27_E4_log_on_gp_nm = 7
    Const ZA27_E4_usr_valid_dt = 8
    Const ZA27_E4_interface_id = 9
    Const ZA27_E4_pwd_valid_dt = 10

    Const ZA27_I4_usr_id = 0
    Const ZA27_I4_org_type = 1
    Const ZA27_I4_org_cd = 2


    Const C_SHEETMAXROWS_D  = 30

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa027 = Server.CreateObject("PZAG027.cListUsrOrgMast")
    Set Enc = Server.CreateObject("EDCodeCom.EDCodeObj.1")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I4_z_usr_org_mast(ZA27_I4_org_cd)
    
    If Request("txtUsrId") <> "" Then
        I4_z_usr_org_mast(ZA27_I4_usr_id)  = Trim(Request("txtUsrId"))
    Else
        I4_z_usr_org_mast(ZA27_I4_usr_id)  = "%"
    End If
    
    I4_z_usr_org_mast(ZA27_I4_org_type)  = Trim(Request("lgOrgType"))
    I4_z_usr_org_mast(ZA27_I4_org_cd)  = Trim(Request("lgOrgCd"))
    E4_z_usr_org_mast = iZa027.ZA_Read_Usr_Org_Mast_Usr_Mast_Rec(gStrGlobalCollection, C_SHEETMAXROWS_D, I4_z_usr_org_mast)
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"            & vbCr
       Response.Write "With Parent "                        & vbCr
       Response.Write ".frm1.txtUsrID.focus"            & vbCr
       Response.Write ".frm1.txtUsrID.select"            & vbCr
       Response.Write ".frm1.txtUsrNm.value = """""        & vbCr
       Response.Write "End With  "                            & vbCr
       Response.Write "</Script>" & vbCr        
    
       Set iZa027 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E4_z_usr_org_mast) Then
       Exit Sub
    End If
    
    Set iZa027 = Nothing    
    Set Enc = Nothing

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "               & vbCr
       Response.Write "Parent.frm1.txtUsrID.value = """    & ConvSPChars(E4_z_usr_org_mast(ZA27_E4_usr_id,0))    & """" &vbCr
       Response.Write "Parent.frm1.txtUsrNm.value = """    & ConvSPChars(E4_z_usr_org_mast(ZA27_E4_usr_nm,0))    & """" &vbCr
       Response.Write "End With  "                 & vbCr 
       Response.Write "</Script>" & vbCr
    End If

   
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E4_z_usr_org_mast, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E4_z_usr_org_mast(ZA27_E4_usr_id,iLngRow)           
           Exit For
        End If
    
        strData = strData & Chr(11) & ConvSPChars(E4_z_usr_org_mast(ZA27_E4_usr_id, iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E4_z_usr_org_mast(ZA27_E4_usr_nm, iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E4_z_usr_org_mast(ZA27_E4_usr_eng_nm,iLngRow))        
        strData = strData & Chr(11) & UNIDateClientFormat(ConvSPChars(E4_z_usr_org_mast(ZA27_E4_usr_valid_dt,iLngRow)))        
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
    Response.Write ".lgStrPrevKey = """     & StrNextKey         & """" & vbCr                
    Response.Write ".frm1.hUsrID.value = """     & StrNextKey         & """" & vbCr                        
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

