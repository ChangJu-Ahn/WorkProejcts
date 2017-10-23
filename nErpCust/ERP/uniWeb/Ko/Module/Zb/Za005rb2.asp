<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za005rb1
*  4. Program Name         : Table Layout Popup
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
    Dim I3_z_table_info
    Dim E3_z_tbl_layout
    
    Dim StrNextKey
    Dim iSheetMaxRow
    
    Const ZA14_I3_table_id = 0
    Const ZA14_I3_syscolumn_id = 1
    
    Const ZA14_E3_table_id = 0
    Const ZA14_E3_table_nm = 1
    Const ZA14_E3_syscolumn_id = 2
    Const ZA14_E3_column_nm = 3
    Const ZA14_E3_column_type = 4
    Const ZA14_E3_nullable = 5
    Const ZA14_E3_PK_gubun = 6

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa0058 = Server.CreateObject("PZAG014.cListTableInfo")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I3_z_table_info(ZA14_I3_syscolumn_id)
    I3_z_table_info(ZA14_I3_table_id)     = Request("txtCode")
    I3_z_table_info(ZA14_I3_syscolumn_id) = lgStrPrevKey

    E3_z_tbl_layout = iZa0058.ZA_Read_Tbl_Layout(gStrGlobalCollection, C_SHEETMAXROWS_D, I3_z_table_info)
        
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

    If IsEmpty(E3_z_tbl_layout) Then
       Exit Sub
    End If

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "Parent.frm1.txtTable.value   = """ & ConvSPChars(E3_z_tbl_layout(ZA14_E3_table_id,0)) & """" &vbCr
       Response.Write "Parent.frm1.txtTableNm.value = """ & ConvSPChars(E3_z_tbl_layout(ZA14_E3_table_nm,0)) & """" &vbCr 
       Response.Write "</Script>" & vbCr
    End If
    
    Set iZa0058 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E3_z_tbl_layout,2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = ConvSPChars(E3_z_tbl_layout(ZA14_E3_syscolumn_id,iLngRow))
           Exit For
        End If
    
        strData = strData & Chr(11) & ConvSPChars(E3_z_tbl_layout(ZA14_E3_column_nm,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E3_z_tbl_layout(ZA14_E3_column_type,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E3_z_tbl_layout(ZA14_E3_PK_gubun,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E3_z_tbl_layout(ZA14_E3_nullable,iLngRow))
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """   & strData & """"     & vbCr
    Response.Write ".lgStrPrevKey = """            & StrNextKey         & """" & vbCr                
    Response.Write ".frm1.htxtTable.value = """ & Request("txtCode") & """" & vbCr        'LJW : Confirm
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

