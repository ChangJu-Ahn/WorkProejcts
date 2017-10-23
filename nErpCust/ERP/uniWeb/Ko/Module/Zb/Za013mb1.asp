<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name          : Basis Architect
*  2. Function Name        : System Management
*  3. Program ID           : za013mb1.asp
*  4. Program Name         : Object Management
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000/05/09
*  8. Modified date(Last)  : 2002/06/06
*  9. Modifier (First)     : ParkSangHoon
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
    Dim iZa025
    Dim E1_z_object_info
    Dim I1_z_object_info
    Dim StrNextKey
    Dim iSheetMaxRow
    Dim txtModuleID
    
    Const ZA25_I1_obj_type    = 0
    Const ZA25_I1_obj_id    = 1
    Const ZA25_I1_module_id = 2
    
    Const ZA25_E1_obj_type = 0
    Const ZA25_E1_Aminor_nm = 1
    Const ZA25_E1_obj_id = 2
    Const ZA25_E1_obj_nm = 3
    Const ZA25_E1_module_id = 4
    Const ZA25_E1_Bminor_nm = 5
    Const ZA25_E1_reg_type = 6
    Const ZA25_E1_Cminor_nm = 7
    Const ZA25_E1_obj_user = 8
    Const ZA25_E1_Dminor_nm = 9
    Const ZA25_E1_use_yn = 10
    Const ZA25_E1_obj_path = 11

    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")
    txtModuleID  = Request("txtModuleID")
    If txtModuleID = "*" Then
        txtModuleID = ""
    End If

    Set iZa025 = Server.CreateObject("PZAG025.cListObjectInfo")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    Redim I1_z_object_info(ZA25_I1_module_id)
    I1_z_object_info(ZA25_I1_obj_type)  = Request("cboObjType")
    I1_z_object_info(ZA25_I1_obj_id)    = Request("txtObjID")
    I1_z_object_info(ZA25_I1_module_id) = txtModuleID

    E1_z_object_info = iZa025.ZA_Read_Object_Info(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_object_info)
        
    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "                & vbCr
       Response.Write ".frm1.txtObjId.focus"        & vbCr
       Response.Write ".frm1.txtObjId.select"        & vbCr
       Response.Write ".frm1.txtObjNm.value = """"" & vbCr
       Response.Write "Call .SetToolBar(""11001101001011"")"    & vbCr
       Response.Write "End With  "                  & vbCr
       Response.Write "</Script>" & vbCr    
    
       Set iZa025 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E1_z_object_info) Then
       Exit Sub
    End If

    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "               & vbCr
       Response.Write ".frm1.txtObjId.value = """    & ConvSPChars(E1_z_object_info(ZA25_E1_obj_id,0))    & """" &vbCr
       Response.Write ".frm1.txtObjNm.value = """    & ConvSPChars(E1_z_object_info(ZA25_E1_obj_nm,0))    & """" &vbCr
       Response.Write "End With  "                 & vbCr 
       Response.Write "</Script>" & vbCr
    End If
    
    Set iZa025 = Nothing    
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E1_z_object_info, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = ConvSPChars(E1_z_object_info(ZA25_E1_obj_id,iLngRow))
           Exit For
        End If
    
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_object_info(ZA25_E1_obj_type, iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_object_info(ZA25_E1_Aminor_nm,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_object_info(ZA25_E1_obj_id,   iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_object_info(ZA25_E1_obj_nm,   iLngRow))      
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_object_info(ZA25_E1_module_id,iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_object_info(ZA25_E1_Bminor_nm,iLngRow))
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_object_info(ZA25_E1_reg_type, iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_object_info(ZA25_E1_Cminor_nm,iLngRow))
        strData = strData & Chr(11) & UCase(ConvSPChars(E1_z_object_info(ZA25_E1_obj_user, iLngRow)))
        strData = strData & Chr(11) & ConvSPChars(E1_z_object_info(ZA25_E1_Dminor_nm,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_object_info(ZA25_E1_use_yn,   iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E1_z_object_info(ZA25_E1_obj_path, iLngRow))
        strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "With Parent "                          & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
    Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
    Response.Write ".lgStrPrevKey = """     & StrNextKey         & """" & vbCr                
    Response.Write ".frm1.htxtObjType.value = """           & Request("cboObjType")  & """" & vbCr
    Response.Write ".frm1.htxtModuleID.value = """           & Request("txtModuleID") & """" & vbCr
    Response.Write ".frm1.htxtObjID.value = """            & StrNextKey             & """" & vbCr
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

    Set iZa024 = Server.CreateObject("PZAG024.cCtrlObjectInfo")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call iZa024.ZA_Control_Object_Info(gStrGlobalCollection,Request("txtSpread"),iErrPosition)
    
    If CheckSYSTEMError2(Err, True, iErrPosition & "Çà:","","","","") = True Then
       Set iZa024 = Nothing                                                         
       Exit Sub
    End If
    
    Set iZa024 = Nothing                                                   
    
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

