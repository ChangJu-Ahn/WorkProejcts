<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za005mb1
*  4. Program Name         : Table Information Management
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.03
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
    Dim iZa0058
    Dim E1_z_table_info
    Dim I1_z_table_info
    Dim StrNextKey
    Dim iSheetMaxRow
    Dim StrMode
    
    Const ZA14_E1_table_id = 0
    Const ZA14_E1_table_nm = 1
    Const ZA14_E1_module_id = 2
    Const ZA14_E1_module_nm = 3
    Const ZA14_E1_table_type = 4
    Const ZA14_E1_table_type_nm = 5
    Const ZA14_E1_use_yn = 6
    
    Const C_TableId    = 0
    Const C_TableTypeS = 1
    Const C_TableTypeM = 2
    Const C_TableTypeX = 3
    Const C_TableTypeT = 4

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")
    StrMode = Request("StrMode")

    Set iZa0058 = Server.CreateObject("PZAG014.cListTableInfo")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_table_info(C_TableTypeT)
    If lgStrPrevKey <> "" Then
        I1_z_table_info(C_TableId) = lgStrPrevKey
    Else
        I1_z_table_info(C_TableId) = Request("txtCode")
    End If
    
    I1_z_table_info(C_TableTypeS)  = Request("txtChk1")
    I1_z_table_info(C_TableTypeM)  = Request("txtChk2")
    I1_z_table_info(C_TableTypeX)  = Request("txtChk3")
    I1_z_table_info(C_TableTypeT)  = Request("txtChk4")

    E1_z_table_info = iZa0058.ZA_Read_Table_Info(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_table_info,StrMode)

    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "                & vbCr
       Response.Write ".frm1.txtTable.focus"        & vbCr
       Response.Write ".frm1.txtTable.select"        & vbCr
       Response.Write ".frm1.txtTableNm.value = """""    &vbCr
       Response.Write "Call .SetToolBar(""11001101001011"")"    & vbCr
       Response.Write "End With  "                  & vbCr
       Response.Write "</Script>" & vbCr
       
       Set iZa0058 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E1_z_table_info) Then
       Exit Sub
    End If

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "Parent.frm1.txtTable.value   = """ & ConvSPChars(E1_z_table_info(ZA14_E1_table_id,0)) & """" &vbCr
       Response.Write "Parent.frm1.txtTableNm.value = """ & ConvSPChars(E1_z_table_info(ZA14_E1_table_nm,0)) & """" &vbCr 
       Response.Write "</Script>" & vbCr
    End If
    
    Set iZa0058 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E1_z_table_info,2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = ConvSPChars(E1_z_table_info(ZA14_E1_table_id,iLngRow))
           Exit For
        End If

        strData = strData & Chr(11) & ConvSPChars(E1_z_table_info(ZA14_E1_table_id,iLngRow))
        strData = strData & Chr(11) & ""
        strData = strData & Chr(11) & ConvSPChars(E1_z_table_info(ZA14_E1_table_nm,iLngRow))
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_table_info(ZA14_E1_module_id,iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_table_info(ZA14_E1_module_nm,iLngRow))
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_table_info(ZA14_E1_table_type,iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_table_info(ZA14_E1_table_type_nm,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_table_info(ZA14_E1_use_yn,iLngRow))
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """    & strData & """"   & vbCr
    Response.Write ".lgStrPrevKey = """     & StrNextKey         & """" & vbCr                
    Response.Write ".frm1.htxtTable.value = .frm1.htxtTable.value "     & vbCr
    Response.Write ".frm1.hchk1.value = """ & Request("txtChk1") & """" & vbCr
    Response.Write ".frm1.hchk2.value = """ & Request("txtChk2") & """" & vbCr
    Response.Write ".frm1.hchk3.value = """ & Request("txtChk3") & """" & vbCr
    Response.Write ".frm1.hchk4.value = """ & Request("txtChk4") & """" & vbCr
    Response.Write ".DbQueryOk  "                          & vbCr
    Response.Write "End With  "                            & vbCr
    Response.Write "</Script>"                             & vbCr

End Sub    

'=========================================================================================================
Sub SubBizSaveMulti()

    Dim iZa0051
    Dim iErrPosition

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    Set iZa0051 = Server.CreateObject("PZAG013.cCtrlTableInfo")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call iZa0051.ZA_Control_Table_Info(gStrGlobalCollection,Request("txtSpread"),iErrPosition)
    
    If CheckSYSTEMError2(Err, True, iErrPosition & "Çà:","","","","") = True Then
       Set iZa0051 = Nothing                                                         
       Exit Sub
    End If
    
    Set iZa0051 = Nothing                                                   
    
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

