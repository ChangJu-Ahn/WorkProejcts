<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za008mb1
*  4. Program Name         : Audit Management
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.13
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
    Dim iZa022
    Dim E1_z_audit_policy
    Dim I1_z_audit_policy
    Dim StrNextKey
    Dim iSheetMaxRow
    
    Const ZA22_E1_table_id = 0
    Const ZA22_E1_table_nm = 1
    Const ZA22_E1_table_type = 2
    Const ZA22_E1_table_type_nm = 3
    Const ZA22_E1_insert_audit = 4
    Const ZA22_E1_update_audit = 5
    Const ZA22_E1_delete_audit = 6

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa022 = Server.CreateObject("PZAG022.cListAuditPolicy")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    I1_z_audit_policy = Request("txtTableID")
    
    E1_z_audit_policy = iZa022.ZA_Read_Audit_Policy(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_audit_policy)
        
    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "                & vbCr
       Response.Write ".frm1.txtTableId.focus"      & vbCr
       Response.Write ".frm1.txtTableId.select"     & vbCr                        
       Response.Write ".frm1.txtTableNm.value = """""     & vbCr                
       Response.Write "Call .SetToolBar(""11001001000011"")"    & vbCr
       Response.Write ".lgIntFlgMode = .Parent.OPMD_UMODE" & vbCr
       Response.Write "End With  "                  & vbCr
       Response.Write "</Script>" & vbCr
       
       Set iZa022 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E1_z_audit_policy) Then
       Exit Sub
    End If

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "               & vbCr
       Response.Write ".frm1.txtTableId.value = """ & ConvSPChars(E1_z_audit_policy(ZA22_E1_table_id,0)) & """" &vbCr
       Response.Write ".frm1.txtTableNm.value = """ & ConvSPChars(E1_z_audit_policy(ZA22_E1_table_nm,0)) & """" &vbCr
       Response.Write "End With  "                 & vbCr 
       Response.Write "</Script>" & vbCr
    End If
    
    Set iZa022 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E1_z_audit_policy, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = ConvSPChars(E1_z_audit_policy(ZA22_E1_table_id,iLngRow))
           Exit For
        End If
    
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_audit_policy(ZA22_E1_table_id,        iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_policy(ZA22_E1_table_nm,        iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_policy(ZA22_E1_table_type,   iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_policy(ZA22_E1_table_type_nm,iLngRow))      
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_policy(ZA22_E1_insert_audit,    iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_policy(ZA22_E1_update_audit,    iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_policy(ZA22_E1_delete_audit, iLngRow))
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
    Response.Write ".lgStrPrevKey = """     & StrNextKey         & """" & vbCr                
    Response.Write ".frm1.hTableId.value = """               & StrNextKey & """" & vbCr        
    
    Response.Write "Dim i,j " & vbCr
    Response.Write ".frm1.vspddata.redraw = false" & vbCr
    Response.Write "For i = " & iLngMaxRow & " + 1 To .frm1.vspdData.MaxRows " & vbCr
    Response.Write "    .frm1.vspdData.Row = i " & vbCr
    Response.Write "    For j = .C_Insert To .C_Delete " & vbCr
    Response.Write "        .frm1.vspdData.Col = j " & vbCr
    Response.Write "        If .frm1.vspdData.Text = ""2"" Then " & vbCr
    Response.Write "            .frm1.vspdData.CellType = 1 " & vbCr
    Response.Write "            .frm1.vspdData.Text = """" " & vbCr 
    Response.Write "            .frm1.vspdData.Protect = True " & vbCr
    Response.Write "            .frm1.vspdData.Lock = True " & vbCr
    Response.Write "        End If " & vbCr
    Response.Write "    Next " & vbCr
    Response.Write "Next " & vbCr
    Response.Write ".frm1.vspddata.redraw = true" & vbCr
    
    Response.Write ".DbQueryOk  "                          & vbCr
    Response.Write "End With  "                            & vbCr
    Response.Write "</Script>"                             & vbCr

End Sub    

'=========================================================================================================
Sub SubBizSaveMulti()
    Dim iZa021
    Dim iErrPosition

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    Set iZa021 = Server.CreateObject("PZAG021.cCtrlAuditPolicy")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call iZa021.ZA_Control_Audit_Policy(gStrGlobalCollection,Request("txtSpread"),iErrPosition)
    
    If CheckSYSTEMError2(Err, True, iErrPosition & "Çà:","","","","") = True Then
       Set iZa021 = Nothing                                                         
       Exit Sub
    End If
    
    Set iZa021 = Nothing                                                   
    
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

