<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za005rb1
*  4. Program Name         : Table-Program Mapping Popup
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2002.07.15
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : LeeJaeWan
* 10. Modifier (Last)      : 
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
    Dim iZa0058
    Dim I2_z_table_info
    Dim E2_z_tbl_mapping
    
    Dim StrNextKey
    Dim iSheetMaxRow
    
    Const Za14_I2_table_id = 0
    Const Za14_I2_mnu_id = 1
    
    Const ZA14_E2_table_id = 0
    Const ZA14_E2_table_nm = 1
    Const ZA14_E2_mnu_id = 2
    Const ZA14_E2_mnu_nm = 3

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa0058 = Server.CreateObject("PZAG014.cListTableInfo")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I2_z_table_info(Za14_I2_mnu_id)
    I2_z_table_info(Za14_I2_table_id) = Request("txtCode")
    I2_z_table_info(Za14_I2_mnu_id)   = lgStrPrevKey

    E2_z_tbl_mapping = iZa0058.ZA_Read_Tbl_Prg_Info(gStrGlobalCollection, C_SHEETMAXROWS_D, I2_z_table_info)
        
    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "                & vbCr
       Response.Write ".frm1.txtTable.focus"        & vbCr
       Response.Write ".frm1.txtTable.select"        & vbCr
       Response.Write ".frm1.txtTableNm.value = """""    &vbCr
       Response.Write "End With  "                  & vbCr
       Response.Write "</Script>" & vbCr
    
       Set iZa0058 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E2_z_tbl_mapping) Then
       Exit Sub
    End If

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "Parent.frm1.txtTable.value   = """ & ConvSPChars(E2_z_tbl_mapping(ZA14_E2_table_id,0)) & """" &vbCr
       Response.Write "Parent.frm1.txtTableNm.value = """ & ConvSPChars(E2_z_tbl_mapping(ZA14_E2_table_nm,0)) & """" &vbCr 
       Response.Write "</Script>" & vbCr
    End If
    
    Set iZa0058 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E2_z_tbl_mapping,2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = ConvSPChars(E2_z_tbl_mapping(ZA14_E2_mnu_id,iLngRow))
           Exit For
        End If
    
        strData = strData & Chr(11) & ConvSPChars(E2_z_tbl_mapping(ZA14_E2_mnu_id,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E2_z_tbl_mapping(ZA14_E2_mnu_nm,iLngRow))
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """   & strData             & """" & vbCr
    Response.Write ".lgStrPrevKey = """            & StrNextKey         & """" & vbCr                
    Response.Write ".frm1.htxtTable.value = """ & Request("txtCode") & """" & vbCr
    Response.Write ".DbQueryOk  "                          & vbCr
    Response.Write "End With  "                            & vbCr
    Response.Write "</Script>"                             & vbCr

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

%>

