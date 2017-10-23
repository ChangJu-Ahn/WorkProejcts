<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za010ma1
*  4. Program Name         : Audit Info Detail Query
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.24
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
    Dim iCount
    
    Dim iLngMaxRow
    Dim iLngRow
    Dim strData
    Dim lgStrPrevKey
    Dim iZa029
    Dim E1_z_table_info
    Dim E2_z_audit_mast_dtl
    Dim E3_bk_column
    Dim E4_bk_column_size
    Dim I1_z_audit_mast_dtl
    Dim StrNextKey
    Dim iColSize
    
    Const ZA29_I1_table_id = 0
    Const ZA29_I1_from_dt = 1
    Const ZA29_I1_to_dt = 2
    Const ZA29_I1_tran_i = 3
    Const ZA29_I1_tran_u = 4
    Const ZA29_I1_tran_d = 5
    Const ZA29_I1_usr_id = 6

    Const ZA29_E1_table_id = 0
    Const ZA29_E1_table_nm = 1

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa029 = Server.CreateObject("PZAG029.cListAuditMastDtl")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_audit_mast_dtl(ZA29_I1_usr_id)
    I1_z_audit_mast_dtl(ZA29_I1_table_id)    = Request("txtTable")
    I1_z_audit_mast_dtl(ZA29_I1_from_dt)    = Request("txtFromDt")
    If lgStrPrevKey <> "" then
        I1_z_audit_mast_dtl(ZA29_I1_to_dt)    = lgStrPrevKey
    Else
        I1_z_audit_mast_dtl(ZA29_I1_to_dt)    = Request("txtToDt")
    End If
    I1_z_audit_mast_dtl(ZA29_I1_tran_i)        = Request("txtT1")
    I1_z_audit_mast_dtl(ZA29_I1_tran_u)        = Request("txtT2")
    I1_z_audit_mast_dtl(ZA29_I1_tran_d)        = Request("txtT3")
    I1_z_audit_mast_dtl(ZA29_I1_usr_id)        = Request("txtUser")

    Call iZa029.ZA_Read_Audit_Mast_Dtl(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_audit_mast_dtl, E1_z_table_info, E2_z_audit_mast_dtl, E3_bk_column, E4_bk_column_size)
    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "                & vbCr
       Response.Write ".frm1.txtTable.focus"        & vbCr
       Response.Write ".frm1.txtTable.select"        & vbCr                        
       Response.Write ".frm1.txtTableNm.value = """""     & vbCr                
       Response.Write "Call .Parent.SetToolBar(""11000000000011"")"    & vbCr
       Response.Write "Call .InitSpreadSheet()"        & vbCr
       Response.Write ".InitCondition"                & vbCr
       Response.Write "End With  "                  & vbCr
       Response.Write "</Script>" & vbCr
    
       Set iZa029 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E2_z_audit_mast_dtl) Then
       Exit Sub
    End If

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "               & vbCr
       
       Response.Write ".frm1.vspdData.Col = .frm1.vspdData.MaxCols"        & vbCr
       Response.Write ".frm1.vspdData.ColHidden = False"                & vbCr
       
       Response.Write ".frm1.vspdData.MaxCols = " & CStr(UBound(E3_bk_column,1)) & vbCr
       Response.Write ".frm1.vspdData.Col = .frm1.vspdData.MaxCols"        & vbCr
       Response.Write ".frm1.vspdData.ColHidden = True"                    & vbCr
       Response.Write ".ggoSpread.Source = .frm1.vspdData"                & vbCr
       Response.Write ".frm1.vspdData.ReDraw = False"                    & vbCr
       Response.Write ".ggoSpread.Spreadinit"                            & vbCr

       For iCount = 5 To UBound(E3_bk_column,1) - 1
            If Len(E3_bk_column(iCount)) > (CInt(E4_bk_column_size(0,iCount-2))) Then
                iColSize = Len(E3_bk_column(iCount)) + 2
            Else
                iColSize = CInt(E4_bk_column_size(0,iCount-2)) + 2
            End If
            Response.Write ".ggoSpread.SSSetEdit " & CStr(iCount) & ", """ & ConvSPChars(E3_bk_column(iCount)) & """, " & CStr(iColSize) & vbCr
       Next
       
       Response.Write ".ggoSpread.SpreadLock 1, -1" & vbCr
       Response.Write ".frm1.vspdData.ReDraw = True" & vbCr
                   
       Response.Write ".frm1.txtTable.value = """    & ConvSPChars(E1_z_table_info(ZA29_E1_table_id))     & """" & vbCr
       Response.Write ".frm1.txtTableNm.value = """ & ConvSPChars(E1_z_table_info(ZA29_E1_table_nm))       & """" & vbCr
       Response.Write ".frm1.htxtTable.value = """   & ConvSPChars(I1_z_audit_mast_dtl(ZA29_I1_table_id)) & """" & vbCr
       Response.Write "End With  "                   & vbCr 
       Response.Write "</Script>"                     & vbCr
    End If

    Set iZa029 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E2_z_audit_mast_dtl, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E2_z_audit_mast_dtl(0,iLngRow) & " " & E2_z_audit_mast_dtl(1,iLngRow)
           Exit For
        End If

        strData = strData & Chr(11) & UNIDateClientFormat(E2_z_audit_mast_dtl(0, iLngRow))
        strData = strData & Chr(11) &            E2_z_audit_mast_dtl(1, iLngRow)                
        strData = strData & Chr(11) &           E2_z_audit_mast_dtl(UBound(E2_z_audit_mast_dtl, 1), iLngRow)
        For iDx = 4 To UBound(E2_z_audit_mast_dtl, 1) - 1
            If IsDate(E2_z_audit_mast_dtl(iDx, iLngRow)) Then
                  strData = strData & Chr(11) & UNIDateClientFormat(E2_z_audit_mast_dtl(iDx, iLngRow))
               Else
                  strData = strData & Chr(11) & ConvSPChars(E2_z_audit_mast_dtl(iDx, iLngRow))        
            End If          
        Next
        
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

