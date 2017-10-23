<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

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
<Script Language=vbscript src="../../inc/incUni2KTV.vbs"></Script>

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
    Dim iZa027
    Dim E1_z_usr_org_mast
    Dim E2_z_usr_org_mast
    Dim E3_z_usr_org_mast        
    Dim I1_z_usr_org_mast
    Dim I2_z_usr_org_mast
    Dim I3_z_usr_org_mast        
    Dim StrNextKey
    Dim iSheetMaxRow
    
    Const ZA27_E1_minor_cd = 0
    Const ZA27_E1_minor_nm = 1

    Const ZA27_E2_org_type = 0
    Const ZA27_E2_org_cd = 1
    Const ZA27_E2_org_nm = 2

    Const ZA27_E3_org_cd = 0
    Const ZA27_E3_org_nm = 1
    Const ZA27_E3_usr_id = 2
    Const ZA27_E3_usr_nm = 3
    Const ZA27_E3_occur_dt = 4
    Const ZA27_E3_use_yn = 5
    Const ZA27_E3_org_type = 6
    Const ZA27_E3_hoccur_dt = 7

    Const ZA27_I1_major_cd = 0

    Const ZA27_I2_major_cd = 0

    Const ZA27_I3_major_cd = 0
    Const ZA27_I3_org_type = 1
    Const ZA27_I3_org_cd = 2
    Const ZA27_I3_usr_id = 3
    
    Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    lgStrPrevKey = Request("lgStrPrevKey")

    Set iZa027 = Server.CreateObject("PZAG027.cListUsrOrgMast")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If


    If UCase(Request("strCmd")) = "INIT" Then
        Redim I1_z_usr_org_mast(ZA27_I1_major_cd)
        I1_z_usr_org_mast(ZA27_I1_major_cd)  = "Z0001"
    
        Redim I2_z_usr_org_mast(ZA27_I2_major_cd)
        I2_z_usr_org_mast(ZA27_I2_major_cd)  = "Z0001"        
    Else    
        Redim I3_z_usr_org_mast(ZA27_I3_usr_id)
        I3_z_usr_org_mast(ZA27_I3_major_cd)  = "Z0001"
        I3_z_usr_org_mast(ZA27_I3_org_type)  = Request("strType")
        I3_z_usr_org_mast(ZA27_I3_org_cd)  = Request("strCd")
        I3_z_usr_org_mast(ZA27_I3_usr_id)  = Request("lgStrPrevKey")        
    End If

    Call iZa027.ZA_Read_Usr_Org_Mast(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_usr_org_mast,I2_z_usr_org_mast,I3_z_usr_org_mast,E1_z_usr_org_mast,E2_z_usr_org_mast,E3_z_usr_org_mast)

    Dim i,j
    If CheckSYSTEMError(Err,True) = True Then
       Set iZa027 = Nothing                                                         
       Exit Sub
    End If

    Set iZa027 = Nothing    

    If UCase(Request("strCmd")) = "INIT" Then  
      
        Response.Write "<Script Language=vbscript>" & vbCr
        Response.Write "Dim Nod " & vbCr       
        Response.Write "</Script>" & vbCr
    
        For iLngRow = 0 To UBound(E1_z_usr_org_mast, 2)
            Response.Write "<Script Language=vbscript>" & vbCr
            Response.Write "Set Nod = Parent.frm1.uniTree1.Nodes.Add (, tvwChild, """ & ConvSPChars(E1_z_usr_org_mast(ZA27_E1_minor_cd,iLngRow)) & """,""" & ConvSPChars(E1_z_usr_org_mast(ZA27_E1_minor_nm,iLngRow)) & """, C_Folder, C_Open )" & vbCr              
            Response.Write "Nod.ExpandedImage = C_Open " & vbCr
            Response.Write "</Script>" & vbCr
        Next

        Response.Write "<Script Language=vbscript>" & vbCr
        Response.Write "Set Nod = Nothing " & vbCr                   
        Response.Write "Parent.frm1.uniTree1.MousePointer = 0 " & vbCr                                    
        Response.Write "</Script>" & vbCr

        For iLngRow = 0 To UBound(E2_z_usr_org_mast, 2)
            Response.Write "<Script Language=vbscript>" & vbCr
            Response.Write "Set Nod = Parent.frm1.uniTree1.Nodes.Add (""" & ConvSPChars(E2_z_usr_org_mast(ZA27_E2_org_type,iLngRow)) & """, tvwChild," & """" & ConvSPChars(E2_z_usr_org_mast(ZA27_E2_org_type,iLngRow)) & "::" & ConvSPChars(E2_z_usr_org_mast(ZA27_E2_org_cd,iLngRow)) & "::" & ConvSPChars(E2_z_usr_org_mast(ZA27_E2_org_nm,iLngRow)) & """,""" & ConvSPChars(E2_z_usr_org_mast(ZA27_E2_org_nm,iLngRow)) & """, C_URL, C_None )" & vbCr              
            Response.Write "Nod.ExpandedImage = C_None " & vbCr
            Response.Write "</Script>" & vbCr
        Next
        
        Response.Write "<Script Language=vbscript>" & vbCr
        Response.Write "Set Nod = Nothing " & vbCr                   
        Response.Write "Parent.frm1.uniTree1.MousePointer = 0 " & vbCr                                    
        Response.Write "</Script>" & vbCr
        
    End If
    
    iLngMaxRow = CLng(Request("txtMaxRows"))

    If IsEmpty(E3_z_usr_org_mast) Then
       Exit Sub
    End If

    For iLngRow = 0 To UBound(E3_z_usr_org_mast, 2)
        If iLngRow < C_SHEETMAXROWS_D Then
        Else
           StrNextKey = E3_z_usr_org_mast(ZA27_E3_usr_id,iLngRow)
           Exit For
        End If

        strData = strData & Chr(11) & UCase(ConvSPChars(E3_z_usr_org_mast(ZA27_E3_org_cd, iLngRow)))    
        strData = strData & Chr(11) & ConvSPChars(E3_z_usr_org_mast(ZA27_E3_org_nm,iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E3_z_usr_org_mast(ZA27_E3_usr_id,   iLngRow))
        strData = strData & Chr(11) & ConvSPChars(E3_z_usr_org_mast(ZA27_E3_usr_nm,   iLngRow))      
        strData = strData & Chr(11) & UNIDateClientFormat(E3_z_usr_org_mast(ZA27_E3_occur_dt,iLngRow))    
        strData = strData & Chr(11) & ConvSPChars(E3_z_usr_org_mast(ZA27_E3_use_yn,iLngRow))        
        strData = strData & Chr(11) & SplitTime(E3_z_usr_org_mast(ZA27_E3_occur_dt, iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E3_z_usr_org_mast(ZA27_E3_org_type,iLngRow))        
        strData = strData & Chr(11) & ConvSPChars(E3_z_usr_org_mast(ZA27_E3_hoccur_dt, iLngRow))        
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
    Dim iZa026
    Dim iErrPosition

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    Set iZa026 = Server.CreateObject("PZAG026.cCtrlUsrOrgMast")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call iZa026.ZA_Control_Usr_Org_Mast(gStrGlobalCollection,Request("txtSpread"),iErrPosition)
    
    If CheckSYSTEMError2(Err, True, iErrPosition & "행:","","","","") = True Then
       Set iZa026 = Nothing                                                         
       Exit Sub
    End If
    
    Set iZa026 = Nothing                                                   
    
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
'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
Function SplitTime(Byval dtDateTime)
    SplitTime = Right("0" & Hour(dtDateTime), 2) & ":" _
            & Right("0" & Minute(dtDateTime), 2) & ":" _
            & Right("0" & Second(dtDateTime), 2)

End Function
%>

