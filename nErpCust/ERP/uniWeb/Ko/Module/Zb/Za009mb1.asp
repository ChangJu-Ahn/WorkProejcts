<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za009mb1
*  4. Program Name         : Audit Info Overview Query
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.21
*  9. Modifier (First)     : LeeJaeJoon
* 10. Modifier (Last)      : LeeJaeWan
* 11. Comment              :
=======================================================================================================-->
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
    Dim iZa023
    Dim E1_z_audit_mast
    Dim I1_z_audit_mast
    Dim StrNextKey
    Dim iSheetMaxRow
    
    Const ZA23_I1_from_dt = 0
    Const ZA23_I1_to_dt = 1
    Const ZA23_I1_usr_id = 2
    Const ZA23_I1_table_id = 3
    Const ZA23_I1_tran_id_i = 4
    Const ZA23_I1_tran_id_u = 5
    Const ZA23_I1_tran_id_d = 6

    Const ZA23_E1_date = 0
    Const ZA23_E1_time = 1
    Const ZA23_E1_tran_id = 2
    Const ZA23_E1_tran_id_nm = 3
    Const ZA23_E1_usr_id = 4
    Const ZA23_E1_usr_nm = 5
    Const ZA23_E1_table_id = 6

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    Set iZa023 = Server.CreateObject("PZAG023.cListAuditMast")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_audit_mast(ZA23_I1_tran_id_d)
    I1_z_audit_mast(ZA23_I1_from_dt)   = Request("txtFromDt")
    I1_z_audit_mast(ZA23_I1_to_dt)     = Request("txtToDt")            
    I1_z_audit_mast(ZA23_I1_usr_id)    = Request("txtUser")
    I1_z_audit_mast(ZA23_I1_table_id)  = Request("txtTable")
    I1_z_audit_mast(ZA23_I1_tran_id_i) = Request("txtT1")
    I1_z_audit_mast(ZA23_I1_tran_id_u) = Request("txtT2")
    I1_z_audit_mast(ZA23_I1_tran_id_d) = Request("txtT3")

    E1_z_audit_mast = iZa023.ZA_Read_Audit_Mast(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_audit_mast)
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZa023 = Nothing                                                         
       Response.Write "<Script Language=vbscript>"    & vbCr
       Response.Write "With Parent"                    & vbCr
       Response.Write ".SetDetailInCaseOfNoData"    & vbCr
       Response.Write "Call .Parent.SetToolBar(""11000000000011"")" & vbCr
       Response.Write ".InitCondition"                & vbCr
       Response.Write "End With"                    & vbCr
       Response.Write "</Script>"                    & vbCr
       Exit Sub
    End If

    If IsEmpty(E1_z_audit_mast) Then
       Exit Sub
    End If

    Set iZa023 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E1_z_audit_mast, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E1_z_audit_mast(ZA23_E1_date,iLngRow) & " " & E1_z_audit_mast(ZA23_E1_time,iLngRow)
           Exit For
        End If

        strData = strData & Chr(11) & UNIDateClientFormat(E1_z_audit_mast(ZA23_E1_date,       iLngRow))        
        strData = strData & Chr(11) &            E1_z_audit_mast(ZA23_E1_time,          iLngRow)        
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_mast(ZA23_E1_tran_id,      iLngRow))
        
        Select Case ConvSPChars(E1_z_audit_mast(ZA23_E1_tran_id,iLngRow))
        Case "I"
			strData = strData & Chr(11) & "Insert"
		Case "U"
			strData = strData & Chr(11) & "Update"
        Case Else
			strData = strData & Chr(11) & "Delete"
		End Select 
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_mast(ZA23_E1_usr_id,     iLngRow))      
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_mast(ZA23_E1_usr_nm,     iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_audit_mast(ZA23_E1_table_id,   iLngRow))
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
    Response.Write ".lgStrPrevKey = """     & StrNextKey         & """" & vbCr                
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

