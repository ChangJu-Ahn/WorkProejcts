<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za006mb1
*  4. Program Name         : Program-Menu Mapping
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.09
*  9. Modifier (First)     : LeeJaeJoon
* 10. Modifier (Last)      : LeeJaeWan
* 11. Comment              :
=======================================================================================================-->
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
    Dim iZa018
    Dim E1_z_lang_co_mast_mnu 
    Dim E2_z_prg_tbl_mapping 
    Dim I1_z_prg_tbl_mapping
    Dim StrNextKey
    Dim iSheetMaxRow
    
    Dim E3_z_prg_tbl_mapping 
    
    Const ZA18_I2_mnu_id = 0
    Const ZA18_I2_table_id = 1
    
    Const ZA18_E1_mnu_id = 0
    Const ZA18_E1_mnu_nm = 1

    Const ZA18_E2_mnu_id = 0
    Const ZA18_E2_mnu_nm = 1
    Const ZA18_E2_table_id = 2
    Const ZA18_E2_table_nm = 3
    Const ZA18_E2_table_type = 4
    Const ZA18_E2_table_type_nm = 5
    
    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa018 = Server.CreateObject("PZAG018.cListPrgTblMapping")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_prg_tbl_mapping(ZA18_I2_table_id)
    I1_z_prg_tbl_mapping(ZA18_I2_mnu_id)   = Request("txtProgramID")
    I1_z_prg_tbl_mapping(ZA18_I2_table_id) = lgStrPrevKey
    
    If Request("txtPrvNext") = "P" or Request("txtPrvNext") = "N" Then
'        Call ServerMesgBox(Request("txtPrvNext") , vbCritical, I_MKSCRIPT)
        E3_z_prg_tbl_mapping = iZa018.ZA_Select_Pre_Next_Prg_Tbl_Mapping(gStrGlobalCollection, Request("txtPrvNext"), Request("txtProgramID"))
        
        If CheckSYSTEMError(Err,True) = True Then
           Set iZa005 = Nothing                                                         
           Exit Sub
        End If

        I1_z_prg_tbl_mapping(ZA18_I2_mnu_id) = CStr(E3_z_prg_tbl_mapping)
        
    End If
    
    E1_z_lang_co_mast_mnu = iZa018.ZA_Lookup_Prg_Tbl_Mapping(gStrGlobalCollection, I1_z_prg_tbl_mapping(ZA18_I2_mnu_id))

    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"            & vbCr
       Response.Write "With Parent "                        & vbCr
       Response.Write ".frm1.txtProgramID.focus"            & vbCr
       Response.Write ".frm1.txtProgramID.select"            & vbCr
       Response.Write ".frm1.txtProgramNm.value = """""        & vbCr
       Response.Write ".frm1.txtProgramIDPrev.value = """""    & vbCr
       Response.Write ".frm1.txtProgramNmPrev.value = """""    & vbCr
       Response.Write "Call .SetToolBar(""1100000000001111"")"    & vbCr
       Response.Write "End With  "                            & vbCr
       Response.Write "</Script>" & vbCr        
    
       Set iZa018 = Nothing                                                         
       Exit Sub
    End If
    
    If IsEmpty(E1_z_lang_co_mast_mnu) Then
       Exit Sub
    End If
        
    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "               & vbCr
       Response.Write "    .Clear "                   & vbCr
       Response.Write ".frm1.txtProgramID.value = """     & ConvSPChars(E1_z_lang_co_mast_mnu(ZA18_E1_mnu_id)) & """" &vbCr
       Response.Write ".frm1.txtProgramNm.value = """     & ConvSPChars(E1_z_lang_co_mast_mnu(ZA18_E1_mnu_nm)) & """" &vbCr
       Response.Write ".frm1.txtProgramIDPrev.value = """ & ConvSPChars(E1_z_lang_co_mast_mnu(ZA18_E1_mnu_id)) & """" &vbCr        '??
       Response.Write ".frm1.txtProgramNmPrev.value = """ & ConvSPChars(E1_z_lang_co_mast_mnu(ZA18_E1_mnu_nm)) & """" &vbCr
       Response.Write "End With  "                 & vbCr 
       Response.Write "</Script>" & vbCr
    End If

    E2_z_prg_tbl_mapping = iZa018.ZA_Read_Prg_Tbl_Mapping(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_prg_tbl_mapping)
        
    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "                & vbCr
       Response.Write ".frm1.txtProgramID.focus"                & vbCr
       Response.Write ".frm1.txtProgramID.select"                & vbCr
       Response.Write "Call .SetToolBar(""1100000011011111"")"    & vbCr
       Response.Write "End With  "                  & vbCr
       Response.Write "</Script>" & vbCr    
    
       Set iZa018 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E2_z_prg_tbl_mapping) Then
       Exit Sub
    End If

    Set iZa018 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E2_z_prg_tbl_mapping,2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = ConvSPChars(E2_z_prg_tbl_mapping(ZA18_E2_table_id,iLngRow))
           Exit For
        End If
    
        strData = strData & Chr(11) & ConvSPChars(E2_z_prg_tbl_mapping(ZA18_E2_table_id,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E2_z_prg_tbl_mapping(ZA18_E2_table_nm,iLngRow))      
        strData = strData & Chr(11) & ConvSPChars(E2_z_prg_tbl_mapping(ZA18_E2_table_type,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E2_z_prg_tbl_mapping(ZA18_E2_table_type_nm,iLngRow))
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """    & strData    & """" & vbCr
    Response.Write ".lgStrPrevKey = """          & StrNextKey & """" & vbCr                
    Response.Write ".frm1.hProgramID.value = """ & StrNextKey & """" & vbCr
    Response.Write ".DbQueryOk  "                          & vbCr
    Response.Write "End With  "                            & vbCr
    Response.Write "</Script>"                             & vbCr

End Sub    

'=========================================================================================================
Sub SubBizSaveMulti()
    Dim iZa017
    Dim iErrPosition

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    Set iZa017 = Server.CreateObject("PZAG017.cCtrlPrgTblMapping")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call iZa017.ZA_Control_Prg_Tbl_Mapping(gStrGlobalCollection,Request("txtSpread"),iErrPosition)
    
    If CheckSYSTEMError2(Err, True, iErrPosition & "Çà:","","","","") = True Then
       Set iZa017 = Nothing                                                         
       Exit Sub
    End If
    
    Set iZa017 = Nothing                                                   
    
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

