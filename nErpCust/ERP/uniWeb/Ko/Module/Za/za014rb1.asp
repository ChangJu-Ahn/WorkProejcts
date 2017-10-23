<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

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
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
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

    Dim lgStrPrevUsrId
    Dim lgStrPrevOccurDt
    Dim lgStrPrevOrgType
    Dim lgStrPrevOrgCd
        
    Dim iZa005

    Dim E1_z_usr_mast_rec
    Dim E2_z_usr_mast_rec
    Dim E3_z_usr_mast_rec
    
    Dim E5_z_usr_mast_rec

    Dim I1_z_usr_mast_rec
    Dim I2_z_usr_mast_rec   
    Dim I3_z_usr_mast_rec              
     
    Dim I5_z_usr_mast_rec
    
    Dim iSheetMaxRow

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
    
    Const ZA05_E5_occur_dt = 0
    Const ZA05_E5_use_yn = 1
    Const ZA05_E5_org_type = 2    
    Const ZA05_E5_org_type_nm = 3
    Const ZA05_E5_org_cd = 4
    Const ZA05_E5_org_nm = 5
    Const ZA05_E5_hoccur_dt = 6

    Const ZA05_I1_usr_id = 0
    
    Const ZA05_I5_usr_Id = 0
    Const ZA05_I5_org_type = 1
    Const ZA05_I5_occur_dt = 2
    Const ZA05_I5_org_cd = 3

    Const C_SHEETMAXROWS_D  = 20

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    Set iZa005 = Server.CreateObject("PZAG005.cLookupUsrMastRec")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    iLngMaxRow = CLng(Request("txtMaxRows"))
        


    Redim I5_z_usr_mast_rec(ZA05_I5_org_cd)    
            
    I5_z_usr_mast_rec(ZA05_I5_usr_id)        = Trim(Request("txtUsrId"))
    I5_z_usr_mast_rec(ZA05_I5_org_type)     = Trim(Request("lgStrPrevOrgType"))                
    I5_z_usr_mast_rec(ZA05_I5_occur_dt)     = Trim(Request("lgStrPrevOccurDt"))                
    I5_z_usr_mast_rec(ZA05_I5_org_cd)       = Trim(Request("lgStrPrevOrgCd"))        

    Redim I1_z_usr_mast_rec(ZA05_I1_usr_id)
    
    I1_z_usr_mast_rec(ZA05_I1_usr_id)       = Trim(Request("txtUsrId"))
    
    E5_z_usr_mast_rec = iZa005.ZA_List_Usr_Org_Mast_Histrory(gStrGlobalCollection, C_SHEETMAXROWS_D, I5_z_usr_mast_rec)
    
    Call iZa005.ZA_Lookup_Usr_Mast_Rec(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_usr_mast_rec, I2_z_usr_mast_rec, I3_z_usr_mast_rec, E1_z_usr_mast_rec, E2_z_usr_mast_rec, E3_z_usr_mast_rec)    
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZa005 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E5_z_usr_mast_rec) Then
       Exit Sub
    End If

    If IsEmpty(E1_z_usr_mast_rec) Then
       Exit Sub
    End If

    If Not IsEmpty(E5_z_usr_mast_rec) Then
        '사용자별 조직관리 조회 
        strData = ""
        For iLngRow = 0 To UBound(E5_z_usr_mast_rec, 2)
            If iLngRow < C_SHEETMAXROWS_D Then
            Else
               lgStrPrevOrgType = E5_z_usr_mast_rec(ZA05_E5_org_type,iLngRow)
               lgStrPrevOrgCd   = E5_z_usr_mast_rec(ZA05_E5_org_cd,iLngRow)            
               lgStrPrevOccurDt = E5_z_usr_mast_rec(ZA05_E5_hoccur_dt,iLngRow)
               Exit For
            End If

            strData = strData & Chr(11) & UNIDateClientFormat(E5_z_usr_mast_rec(ZA05_E5_occur_dt, iLngRow))            
            If UCase(ConvSPChars(E5_z_usr_mast_rec(ZA05_E5_use_yn, iLngRow))) = "Y" Then
               strData = strData & Chr(11) & "1"
            Else
               strData = strData & Chr(11) & "0"
            ENd if            
            strData = strData & Chr(11) & UCase(ConvSPChars(E5_z_usr_mast_rec(ZA05_E5_org_type_nm, iLngRow)))        
            strData = strData & Chr(11) & UCase(ConvSPChars(E5_z_usr_mast_rec(ZA05_E5_org_cd, iLngRow)))
            strData = strData & Chr(11) & UCase(ConvSPChars(E5_z_usr_mast_rec(ZA05_E5_org_nm, iLngRow)))        
            strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
            strData = strData & Chr(11) & Chr(12)
            
        Next
    
        Response.Write "<Script Language=vbscript>"            & vbCr
        Response.Write "With Parent "                          & vbCr
        Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
        Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
        Response.Write ".frm1.htxtUsrId.value = """       & UCase(ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_id,0)))     & """" & vbCr               
        Response.Write ".frm1.txtUsrNm.value = """       & UCase(ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_nm,0)))      & """" & vbCr        
        Response.Write ".lgStrPrevOrgType = """     & lgStrPrevOrgType         & """" & vbCr                    
        Response.Write ".lgStrPrevOrgCd = """       & lgStrPrevOrgCd         & """" & vbCr                    
        Response.Write ".lgStrPrevOccurDt = """     & lgStrPrevOccurDt         & """" & vbCr                                                
        Response.Write ".DbQueryOk  "                          & vbCr                      
        Response.Write "End With  "                            & vbCr
        Response.Write "</Script>"                             & vbCr
    End If

    If Not (iZa005 Is Nothing) Then
        Set iZa005 = Nothing
    End If


End Sub    

'=========================================================================================================
Sub SubBizSaveMulti()    
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

