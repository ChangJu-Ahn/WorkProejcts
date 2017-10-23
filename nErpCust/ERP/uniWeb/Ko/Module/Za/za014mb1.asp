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
    Dim iZa001
    Dim iErrPosition
    Dim importArray    
    
    On Error Resume Next                                                                 
    Err.Clear        

    Const C_SelectChar = 0
    Const C_UsrId = 1                                                          

    Redim ImportArray(C_UsrId)    

    ImportArray(C_SelectChar) = "D"
    ImportArray(C_UsrId) = Request("txtUsrId")
            
    Set iZa001 = Server.CreateObject("PZAG001.cCtrlUsrMastRec")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    Call iZa001.ZA_Control_Usr_Mast_Rec(gStrGlobalCollection,importArray, "", "",iErrPosition)
    
    If CheckSYSTEMError2(Err, True, iErrPosition & "행:","","","","") = True Then
       Set iZa001 = Nothing                                                         
       Exit Sub
    End If
    
    Set iZa001 = Nothing                                                   
    
    Response.Write "<Script Language=vbscript>"   & vbCr
    Response.Write "Parent.DbDeleteOk "            & vbCr
    Response.Write "</Script>"                    & vbCr            
End Sub

'=========================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLngMaxRow
    Dim iLngMaxRow1
    Dim iLngRow
    Dim strData
    Dim lgStrPrevKey
    Dim lgStrPrvNext
    Dim lgCurrentSpd
    Dim iZa005
    
    Dim E1_z_usr_mast_rec
    Dim E2_z_usr_mast_rec
    Dim E3_z_usr_mast_rec
    
    Dim E4_z_usr_mast_rec                    
    
    Dim I1_z_usr_mast_rec
    Dim I2_z_usr_mast_rec
    Dim I3_z_usr_mast_rec    
        
    Dim lgStrPrevUsrRoleId
    Dim lgStrPrevOrgType
    Dim lgStrPrevOccurDt        
    Dim lgStrPrevOrgCd
    
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
    Const ZA05_E1_USR_KIND = 12
    Const ZA05_E1_email = 13
    
    Const ZA05_E2_usr_role_id = 0
    Const ZA05_E2_usr_role_nm = 1
    Const ZA05_E2_compst_role_type = 2
    
    Const ZA05_E3_use_yn = 0
    Const ZA05_E3_org_type = 1
    Const ZA05_E3_org_type_nm = 2
    Const ZA05_E3_org_cd = 3
    Const ZA05_E3_org_nm = 4
    Const ZA05_E3_occur_dt = 5
    Const ZA05_E3_hoccur_dt = 6
    
    Const ZA05_E4_usr_id = 0

    Const ZA05_I1_usr_id = 0

    Const ZA05_I2_usr_id = 0
    Const ZA05_I2_usr_role_id = 1

    Const ZA05_I3_usr_Id = 0
    Const ZA05_I3_org_type = 1
    Const ZA05_I3_occur_dt = 2
    Const ZA05_I3_org_cd = 3

    Const C_SHEETMAXROWS_D  = 30

    On Error Resume Next                                                                 
    Err.Clear                                                                            

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    Set iZa005 = Server.CreateObject("PZAG005.cLookupUsrMastRec")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    iLngMaxRow = CLng(Request("txtMaxRows"))
    iLngMaxRow1 = CLng(Request("txtMaxRows1"))
        


    Redim I1_z_usr_mast_rec(ZA05_I1_usr_id)    
    I1_z_usr_mast_rec(ZA05_I1_usr_id)        = Trim(Request("txtUsrId"))        

    lgCurrentSpd = Request("lgCurrentSpd")
    
    If Trim(lgCurrentSpd) = "A" or Trim(lgCurrentSpd) = "*" Then
        Redim I2_z_usr_mast_rec(ZA05_I2_usr_role_id)    
        I2_z_usr_mast_rec(ZA05_I2_usr_id)       = Trim(Request("txtUsrId"))
        I2_z_usr_mast_rec(ZA05_I2_usr_role_id)  = Trim(Request("lgStrPrevUsrRoleId"))        
    End If

    If Trim(lgCurrentSpd) = "B" or Trim(lgCurrentSpd) = "*" Then
        Redim I3_z_usr_mast_rec(ZA05_I3_org_cd)            
        I3_z_usr_mast_rec(ZA05_I3_usr_id)        = Trim(Request("txtUsrId"))
        I3_z_usr_mast_rec(ZA05_I3_org_type)     = Trim(Request("txtOrgType"))                
        I3_z_usr_mast_rec(ZA05_I3_occur_dt)     = Trim(Request("txtOccurDt"))                
        I3_z_usr_mast_rec(ZA05_I3_org_cd)       = Trim(Request("txtOrgCd"))        
    End If

    If Request("lgStrPrvNext") = "P" or Request("lgStrPrvNext") = "N" Then
        E4_z_usr_mast_rec = iZa005.ZA_Select_Pre_Next_Usr_Mast_Rec(gStrGlobalCollection, Request("lgStrPrvNext"), Request("txtUsrId"))

        If CheckSYSTEMError(Err,True) = True Then
           Response.Write "<Script Language=vbscript>"            & vbCr
           Response.Write "With Parent "                        & vbCr
           Response.Write ".frm1.txtUsrId1.focus"            & vbCr
           Response.Write ".frm1.txtUsrId1.select"            & vbCr
           Response.Write ".lgStrPrvNext    = """""    & vbCr
           Response.Write "End With  "                            & vbCr
           Response.Write "</Script>" & vbCr        
    
           Set iZa005 = Nothing                                                         
           Exit Sub
        Else
           Response.Write "<Script Language=vbscript>"            & vbCr
           Response.Write "With Parent "                          & vbCr           
           Response.Write "Call .EraseContents "        & vbCr
           Response.Write "End With  "                            & vbCr           
           Response.Write "</Script>" & vbCr                    
        End If
        
        Redim I2_z_usr_mast_rec(ZA05_I2_usr_role_id)    
        Redim I3_z_usr_mast_rec(ZA05_I3_org_cd)            
        I1_z_usr_mast_rec(ZA05_I1_usr_id)  = CStr(E4_z_usr_mast_rec)
        I2_z_usr_mast_rec(ZA05_I2_usr_id)  = CStr(E4_z_usr_mast_rec)
        I3_z_usr_mast_rec(ZA05_I3_usr_id)  = CStr(E4_z_usr_mast_rec)
        
    End If
    
    Call iZa005.ZA_Lookup_Usr_Mast_Rec(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_usr_mast_rec, I2_z_usr_mast_rec, I3_z_usr_mast_rec, E1_z_usr_mast_rec, E2_z_usr_mast_rec, E3_z_usr_mast_rec)

    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"            & vbCr
       Response.Write "With Parent "                        & vbCr
       Response.Write ".frm1.txtUsrId1.focus"            & vbCr
       Response.Write ".frm1.txtUsrId1.select"            & vbCr
       Response.Write ".frm1.txtUsrNm1.value = """""        & vbCr
       Response.Write ".SetToolBar(""1110110100001111"") "        & vbCr
       Response.Write "End With  "                            & vbCr
       Response.Write "</Script>" & vbCr        
    
       Set iZa005 = Nothing                                                         
       Exit Sub
    End If

    If IsEmpty(E1_z_usr_mast_rec) Then
       Exit Sub
    End If
    
    If lgStrPrevKey = "" Then
       Response.Write "<Script Language=vbscript>" & vbCr
       Response.Write "With Parent "               & vbCr
       Response.Write ".frm1.txtUsrId1.value = """       & ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_id,0))       & """" & vbCr
       Response.Write ".frm1.txtUsrId2.value = """       & ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_id,0))       & """" & vbCr       
       Response.Write ".frm1.htxtUsrId.value = """       & ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_id,0))       & """" & vbCr       
       Response.Write ".frm1.txtUsrNm1.value = """       & ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_nm,0))       & """" & vbCr
       Response.Write ".frm1.txtUsrNm2.value = """       & ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_nm,0))       & """" & vbCr       
       Response.Write ".frm1.txtUsrEngNm.value = """    & ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_usr_eng_nm,0))   & """" & vbCr        
       Response.Write ".frm1.txtUsrValidDt.text = """   & UNIDateClientFormat(E1_z_usr_mast_rec(ZA05_E1_usr_valid_dt,0)) & """" & vbCr 
       Response.Write ".frm1.txtPassword.value = """    & "******"   & """" & vbCr        
       Response.Write ".frm1.txtLogOnGrp.value = """    & Trim(UCase(ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_log_on_gp,0))))     & """" & vbCr         
       Response.Write ".frm1.txtLogOnGrpNm.value = """ & UCase(ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_log_on_gp_nm,0)))  & """" & vbCr
       Response.Write ".frm1.txtCoCd.value = """        & UCase(ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_co_cd,0)))        & """" & vbCr
       Response.Write ".frm1.txtCoCdNm.value = """      & UCase(ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_co_cd_nm,0)))     & """" & vbCr              
       Response.Write ".frm1.hPassword.value = """      & ConvSPChars(LCase(E1_z_usr_mast_rec(ZA05_E1_password,0)))     & """" & vbCr
       Response.Write ".frm1.hUsrIdValidDt.value = """    & ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_husr_valid_dt,0))     & """" & vbCr                                     
       Response.Write ".frm1.cboUserKind.value = """    & E1_z_usr_mast_rec(ZA05_E1_USR_KIND,0)  & """" & vbCr
       Response.Write ".frm1.txtEmail.value = """    & ConvSPChars(E1_z_usr_mast_rec(ZA05_E1_email ,0))  & """" & vbCr
       
       If Trim(E1_z_usr_mast_rec(ZA05_E1_use_yn,0)) = "Y" Then
            Response.Write "Parent.frm1.rdoUseYn1.checked = """ & True & """" &vbCr                       
       ElseIf Trim(E1_z_usr_mast_rec(ZA05_E1_use_yn,0)) = "N" Then
            Response.Write "Parent.frm1.rdoUseYn2.checked = """ & True & """" &vbCr                       
       End If
       
       Response.Write "Parent.frm1.txtInterfaceId.value = """ & E1_z_usr_mast_rec(ZA05_E1_interface_id,0) & """" &vbCr        
       Response.Write "Parent.DbQueryOk  "                          & vbCr              
       Response.Write "End With  "                 & vbCr 
       Response.Write "</Script>" & vbCr
    End If

    Set iZa005 = Nothing  

    strData = ""

    If Not IsEmpty(E2_z_usr_mast_rec) Then
        For iLngRow = 0 To UBound(E2_z_usr_mast_rec, 2)
            If iLngRow < C_SHEETMAXROWS_D Then
            Else
               lgStrPrevUsrRoleId = E2_z_usr_mast_rec(ZA05_E2_usr_role_id,iLngRow)
               Exit For
            End If

            strData = strData & Chr(11) & UCase(ConvSPChars(E2_z_usr_mast_rec(ZA05_E2_usr_role_id, iLngRow)))
            strData = strData & Chr(11) & " "      'PopUp
            strData = strData & Chr(11) & ConvSPChars(E2_z_usr_mast_rec(ZA05_E2_usr_role_nm,iLngRow))
            
            If CStr(Trim(E2_z_usr_mast_rec(ZA05_E2_compst_role_type,iLngRow))) = "1" Then
                strData = strData & Chr(11) & "Composite Role"
            Else
                strData = strData & Chr(11) & "Menu Role"
            End If            
            
            strData = strData & Chr(11) & iLngMaxRow + iLngRow + 1
            strData = strData & Chr(11) & Chr(12)
        Next
    
        Response.Write "<Script Language=vbscript>"            & vbCr
        Response.Write "With Parent "                          & vbCr
        Response.Write ".ggoSpread.Source = .frm1.vspdData "   & vbCr
        Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
        Response.Write ".lgStrPrevUsrRoleId = """     & lgStrPrevUsrRoleId         & """" & vbCr                
        Response.Write ".DbQueryOk  "                          & vbCr
        Response.Write "End With  "                            & vbCr
        Response.Write "</Script>"                             & vbCr
    End If
    
'    Response.Write IsEmpty(E3_z_usr_mast_rec)
'    Response.End
    
    If Not IsEmpty(E3_z_usr_mast_rec) Then
        '사용자별 조직관리 조회 
        strData = ""
        For iLngRow = 0 To UBound(E3_z_usr_mast_rec, 2)
            If iLngRow < C_SHEETMAXROWS_D Then
            Else
               lgStrPrevOrgType = E3_z_usr_mast_rec(ZA05_E3_org_type,iLngRow)
               lgStrPrevOccurDt = E3_z_usr_mast_rec(ZA05_E3_occur_dt,iLngRow)
               lgStrPrevOrgCd   = E3_z_usr_mast_rec(ZA05_E3_org_cd,iLngRow)
               Exit For
            End If

            strData = strData & Chr(11) & UCase(ConvSPChars(E3_z_usr_mast_rec(ZA05_E3_use_yn, iLngRow)))
            strData = strData & Chr(11) & UCase(ConvSPChars(E3_z_usr_mast_rec(ZA05_E3_org_type, iLngRow)))
            strData = strData & Chr(11) & " "      'PopUp
            strData = strData & Chr(11) & UCase(ConvSPChars(E3_z_usr_mast_rec(ZA05_E3_org_type_nm, iLngRow)))        
            strData = strData & Chr(11) & UCase(ConvSPChars(E3_z_usr_mast_rec(ZA05_E3_org_cd, iLngRow)))
            strData = strData & Chr(11) & " "      'PopUp                
            strData = strData & Chr(11) & UCase(ConvSPChars(E3_z_usr_mast_rec(ZA05_E3_org_nm, iLngRow)))        
            strData = strData & Chr(11) & UNIDateClientFormat(E3_z_usr_mast_rec(ZA05_E3_occur_dt, iLngRow))
            strData = strData & Chr(11) & SplitTime(E3_z_usr_mast_rec(ZA05_E3_occur_dt, iLngRow))        
            strData = strData & Chr(11) & UCase(ConvSPChars(E3_z_usr_mast_rec(ZA05_E3_hoccur_dt, iLngRow)))                                                
            strData = strData & Chr(11) & iLngMaxRow1 + iLngRow + 1
            strData = strData & Chr(11) & Chr(12)
        Next
    
        Response.Write "<Script Language=vbscript>"            & vbCr
        Response.Write "With Parent "                          & vbCr
        Response.Write ".ggoSpread.Source = .frm1.vspdData1 "   & vbCr
        Response.Write ".ggoSpread.SSShowData """ & strData & """"   & vbCr
        Response.Write ".lgStrPrevOrgType = """     & lgStrPrevOrgType         & """" & vbCr                    
        Response.Write ".lgStrPrevOccurDt = """     & lgStrPrevOccurDt         & """" & vbCr                
        Response.Write ".frm1.hOrgType.value = """           & lgStrPrevOrgType  & """" & vbCr
        Response.Write ".frm1.hOrgCd.value = """           & lgStrPrevOrgCd  & """" & vbCr        
        Response.Write ".frm1.hOccurDt.value = """           & lgStrPrevOccurDt  & """" & vbCr    
        Response.Write ".DbQueryOk  "                          & vbCr
        Response.Write "End With  "                            & vbCr
        Response.Write "</Script>"                             & vbCr
    End If
    
End Sub    

'=========================================================================================================
Sub SubBizSaveMulti()
    Dim iZa001
    Dim iErrPosition
    Dim importArray
    Dim lgStrPwdUpdate
    Dim lgIntFlgMode
    
    On Error Resume Next                                                                 
    Err.Clear        

    '사용자 정보 관리 
    Const C_SelectChar = 0
    Const C_UsrId = 1                                                          
    Const C_UsrNm = 2                                                            
    Const C_UsrEngNm = 3                                                           
    Const C_UsrValidDt = 4                                                           
    Const C_InterfaceId = 5                                                           
    Const C_Password = 6                                                         
    Const C_CoCd = 7                                                            
    Const C_LogOnGrp = 8                                                            
    Const C_UseYn = 9                                                                
    Const C_PwdUpdateOrNot = 10
    Const C_UserKind = 11
    Const C_Email = 12
    
    
    'Role할당 
    Const C_CurrentRow = 1
    Const C_UsrMastRecId = 2        
    Const C_UsrRoleId = 3

    '사용자별 조직관리 
    Const C_UsrOrgMastUseYn    = 3
    Const C_OrgType  = 4
    Const C_OrgCd    = 5
    Const C_OrgNm    = 6    
    Const C_hOccurDt = 7

    Set iZa001 = Server.CreateObject("PZAG001.cCtrlUsrMastRec")

    If Request("txtBlnFlgChgValue") = "True" Then            

        'Redim ImportArray(C_PwdUpdateOrNot)        
        Redim ImportArray(C_Email)        
        
        lgIntFlgMode = CInt(Request("txtFlgMode"))    
        lgStrPwdUpdate = Request("txtPwdUpdateOrNot")        
            
        If lgIntFlgMode     = OPMD_CMODE Then
            ImportArray(C_SelectChar) = "C"
        ElseIf lgIntFlgMode = OPMD_UMODE Then
            ImportArray(C_SelectChar) = "U"
        End If

        If lgStrPwdUpdate     = "Y" Then    
            ImportArray(C_PwdUpdateOrNot) = "Y"
        ElseIf lgStrPwdUpdate = "N" Then
            ImportArray(C_PwdUpdateOrNot) = "N"
        End If
        
        If CheckSYSTEMError(Err,True) = True Then
           Exit Sub
        End If
            
        ImportArray(C_UsrId)       = Trim(Request("txtUsrId2"))
        ImportArray(C_UsrNm)       = Request("txtUsrNm2")
        ImportArray(C_UsrEngNm)    = Request("txtUsrEngNm")
        ImportArray(C_UsrValidDt)  = UNIConvDate(Request("txtUsrValidDt"))
        ImportArray(C_InterfaceId) = Trim(Request("txtInterfaceId"))
        ImportArray(C_Password)    = Trim(Request("hPassword"))        
        ImportArray(C_CoCd)        = Trim(Request("txtCoCd"))
        ImportArray(C_LogOnGrp)    = Trim(Request("txtLogOnGrp"))
        ImportArray(C_UseYn)       = Trim(Request("txtUseYn"))
     
        If Trim(Request("cboUserKind")) = "" Then
           ImportArray(C_UserKind)    = "U"
        Else   
           ImportArray(C_UserKind)    = Trim(Request("cboUserKind"))
        End If
        
        ImportArray(C_Email)		= Trim(Request("txtEmail"))
        
    End If
    
    Call iZa001.ZA_Control_Usr_Mast_Rec(gStrGlobalCollection,importArray, Request("txtSpread"), Request("txtSpread1"),iErrPosition)
    
       If Not IsEmpty(importArray) Then
           Erase ImportArray
       End If
        
    If CheckSYSTEMError2(Err, True, iErrPosition & "행:","","","","") = True Then
       Set iZa001 = Nothing                                                         
       Exit Sub
    End If
    
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
'=========================================================================================================
Function SplitTime(Byval dtDateTime)

    If IsNull(dtDateTime)  Then
        SplitTime = ""
        Exit Function
    End If

    SplitTime = Right("0" & Hour(dtDateTime), 2) & ":" _
            & Right("0" & Minute(dtDateTime), 2) & ":" _
            & Right("0" & Second(dtDateTime), 2)
            
End Function


%>

