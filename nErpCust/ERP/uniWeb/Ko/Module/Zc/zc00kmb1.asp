<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
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

    lgOpModeCRUD = Request("txtMode")
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)    
            Call SubBizQueryMulti()
        Case CStr(UID_M0002),CStr(UID_M0005),CStr(UID_M0003)
            Call SubBizSaveMulti()        
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


    Dim iZC0014
    Dim istrCode
    Dim istrNm
    Dim lgStrPrevkey
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    Dim iStrData
    Dim iStrFlag
    Dim    lgstrPrvNext
    Dim istrMnuId
    Dim istrMnuNm
    
    Dim E1_Z_Prcs_Mnu_Asso
    Const C_SHEETMAXROWS_D = 100
    
    Const ZC14_E1_PRCS_SEQ     = 0
    Const ZC14_E1_PRCS_SUB_SEQ = 1
    Const ZC14_E1_MNU_ID       = 2
    Const ZC14_E1_MNU_NM       = 3
    Const ZC14_E1_OPTN_FLAG    = 4
    Const ZC14_E1_REMARK       = 5
    Const ZC14_E1_PRCS_CD      = 6
    Const ZC14_E1_PRCS_NM      = 7
    Const ZC14_E1_SYS_FLAG     = 8
    
    On Error Resume Next 
    Err.Clear
    
    istrCode = Request("txtPrcsCd")
    
    lgstrPrvNext = Request("txtPrvNext")
    
    Set iZC0014 = Server.CreateObject("PZCG014.cListPrcsMnuAsso")

    If CheckSYSTEMError(Err,True) = True Then
       Set iZC0014 = Nothing                                                         
       Exit Sub                                                        
    End If


    If lgstrPrvNext = "P" OR lgstrPrvNext = "N" then     
        E1_Z_Prcs_Mnu_Asso = iZC0014.ZC_SELECT_PRE_NEXT_PRCS_MNU_ASSO (gStrGlobalCollection,lgstrPrvNext,istrCode)
        
        If CheckSYSTEMError(Err,True)= True Then 
            Response.Write "<Script language=vbscript>"                            & vbCr    
            Response.Write "    parent.frm1.txtPrcsCd.focus "                    & vbCr                                        
            Response.Write "    parent.frm1.txtPrcsCd.select  "                    & vbCr 
            Response.Write "    parent.SetToolBar ""1111111111111111""  "         & vbCr                                                                                                                                                              
            Response.Write "</script>"                                            & vbCr                            
            Set iZBG011 = Nothing 
            Exit Sub
        End If 
        
        istrCode = Trim(E1_Z_Prcs_Mnu_Asso(0,0))
        istrNm = Trim(E1_Z_Prcs_Mnu_Asso(1,0))
        istrFlag = Trim(E1_Z_Prcs_Mnu_Asso(2,0))
        istrMnuId = Trim(E1_Z_Prcs_Mnu_Asso(3,0))
        istrMnuNm = Trim(E1_Z_Prcs_Mnu_Asso(4,0))
                    
        Response.Write "<Script language=vbscript>"                                        & vbCr    
        Response.Write "With Parent "                                                    & vbCr    
        Response.Write "    .Clear    "                                                    & vbCr                    
        Response.Write "    .frm1.txtPrcsCd.Value = """ & ConvSPChars(istrCode) & """"    & vbCr
        Response.Write "    .frm1.txtPrcsNm.Value = """ & ConvSPChars(istrNm) & """"    & vbCr
        Response.Write "    .frm1.txtPrcsCd2.Value = """ & ConvSPChars(istrCode) & """"    & vbCr
        Response.Write "    .frm1.txtPrcsNm2.Value = """ & ConvSPChars(istrNm) & """"    & vbCr    
        Response.Write "    .frm1.txtUppMnu.Value = """ & ConvSPChars(istrMnuId) & """"    & vbCr'khy2003.03
        Response.Write "    .frm1.txtUppMnuNm.Value = """ & ConvSPChars(istrMnuNm) & """"    & vbCr'khy2003.03                                    
        Response.Write "    .frm1.txtPrcsCd.focus "                                        & vbCr                                        
        Response.Write "    If " & istrFlag & "= 1    Then    "                            & vbCr            
        Response.Write "        .frm1.cbxSysFlag.Checked = True    "                        & vbCr
        Response.Write "    End If        "                                                & vbCr
        Response.Write "End With "                                                        & vbCr
        Response.Write "</script>"                                                        & vbCr
        
        istrCode = istrCode
    Else     
        E1_Z_Prcs_Mnu_Asso = iZC0014.ZC_SELECT_PRE_NEXT_PRCS_MNU_ASSO (gStrGlobalCollection,"",istrCode)

        If CheckSYSTEMError(Err,True)= True Then                             
            Response.Write "<Script language=vbscript>"                            & vbCr    
            Response.Write "    parent.frm1.txtPrcsCd.focus "                    & vbCr                                        
            Response.Write "</script>"                                            & vbCr
            Set iZBG011 = Nothing 
            Exit Sub
        End If 
        
        
        istrCode = Trim(E1_Z_Prcs_Mnu_Asso(0,0))
        istrNm = Trim(E1_Z_Prcs_Mnu_Asso(1,0))
        istrFlag = Trim(E1_Z_Prcs_Mnu_Asso(2,0))
        istrMnuId = Trim(E1_Z_Prcs_Mnu_Asso(3,0))
        istrMnuNm = Trim(E1_Z_Prcs_Mnu_Asso(4,0))
        

        Response.Write "<Script language=vbscript>"                                        & vbCr    
        Response.Write "With Parent "                                                    & vbCr    
        Response.Write "    .Clear    "                                                    & vbCr                    
        Response.Write "    .frm1.txtPrcsCd.Value = """ & ConvSPChars(istrCode) & """"    & vbCr
        Response.Write "    .frm1.txtPrcsNm.Value = """ & ConvSPChars(istrNm) & """"    & vbCr
        Response.Write "    .frm1.txtPrcsCd2.Value = """ & ConvSPChars(istrCode) & """"    & vbCr
        Response.Write "    .frm1.txtPrcsNm2.Value = """ & ConvSPChars(istrNm) & """"    & vbCr    
        Response.Write "    .frm1.txtUppMnu.Value = """ & ConvSPChars(istrMnuId) & """"    & vbCr'khy2003.03
        Response.Write "    .frm1.txtUppMnuNm.Value = """ & ConvSPChars(istrMnuNm) & """"    & vbCr'khy2003.03                                        
        Response.Write "    .frm1.txtPrcsCd.focus "                                        & vbCr                                                
        Response.Write "    If " & istrFlag & "= 1    Then    "                            & vbCr            
        Response.Write "        .frm1.cbxSysFlag.Checked = True    "                        & vbCr
        Response.Write "    End If        "                                                & vbCr            
        Response.Write "End With "                                                        & vbCr
        Response.Write "</script>"                                                        & vbCr
        
        istrCode = istrCode
        
    End if 

    Response.Write "<Script Language=vbscript>"                        & vbCr
    Response.Write "With Parent "                                    & vbCr
    Response.Write "    .SetToolBar ""1111111111011111""  "         & vbCr
    Response.Write "End With "                                        & vbCr
    Response.Write "</Script>"
        
    E1_Z_Prcs_Mnu_Asso=""

    E1_Z_Prcs_Mnu_Asso = iZC0014.ZC_PRCS_MNU_ASSO (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)
    
    
    If CheckSYSTEMError(Err,True) = True Then        
       Set iZC0014 = Nothing            
       Exit Sub
    End If
    
       
    Set iZC0014 = Nothing

    iLngMaxRow  = CLng(Request("txtMaxRows"))
    
    istrFlag = ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_SYS_FLAG,iLngMaxRow))
    

    Response.Write "<Script Language=vbscript>"                        & vbCr
    Response.Write "With Parent "                                    & vbCr        
    Response.Write "    .frm1.txtPrcsNm.value  = """ & ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_PRCS_NM,iLngMaxRow)) & """" & vbCr
    'Response.Write "strFlag  = """ & ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_SYS_FLAG,iLngMaxRow)) & """" & vbCr
    Response.Write "If " & istrFlag & "= ""1"" Then " & vbCr
    Response.Write "    .frm1.cbxSysFlag.checked = True " & vbCr
    Response.Write "Else" & vbCr
    Response.Write "    .frm1.cbxSysFlag.checked = False " & vbCr
    Response.Write "End If" & vbCr
    Response.Write "     .frm1.vspdData.ReDraw = False " & vbCr
    Response.Write "End With "                                        & vbCr
    Response.Write "</Script>"    
    
    For iLngRow =0 To UBound(E1_Z_Prcs_Mnu_Asso,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            iStrNextKey = E1_Z_Prcs_Mnu_Asso(0,iLngRow)
            Exit For
        End if 
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_PRCS_SEQ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_PRCS_SUB_SEQ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_MNU_ID,iLngRow))
        istrData = istrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_MNU_NM,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_OPTN_FLAG,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_REMARK,iLngRow))
        iStrData = iStrData & Chr(11) & iLngMaxRow + ConvSPChars(iLngRow) 
        iStrData = iStrData & Chr(11) & Chr(12)
    Next
    

    Response.Write "<Script Language=vbscript>"                        & vbCr
    Response.Write "With Parent "                                    & vbCr
    Response.Write "    .ggoSpread.Source = .frm1.vspdData "        & vbCr
    Response.Write "    .ggoSpread.SSShowData """ & iStrData        & """" & vbCr
    Response.Write "    .frm1.vspdData.ReDraw = True "               & vbCr    
    Response.Write "    .DbQueryOk  "                               & vbCr
    Response.Write "End With "                                        & vbCr
    Response.Write "</Script>"        
    
    
End Sub

        
'=========================================================================================================        
Sub SubBizSaveMulti()

    Dim iZC013
    Dim iErrorPosition
    Dim iStrSpread    
    Dim iStrSingle    
    
    Dim    txtMode
    Dim LngMaxRow
    
    
    Const ZC13_I1_PRCS_CD     = 0
    Const ZC13_I1_PRCS_NM     = 1
    Const ZC13_I1_Sys_Flag    = 2
    Const ZC13_I1_Upp_Mnu     = 3'khy2003.03
    
    On Error Resume Next 
    Err.Clear
    
    LngMaxRow = Request("txtMaxRows")

    Set iZC013 = Server.CreateObject("PZCG013.cCtrlPrcsMnuAsso")

    If CheckSYSTEMError(Err,True) = True Then
       Set iZC013 = Nothing    
       Exit Sub
    End If

    istrSpread = Request("txtSpread")
    
    
    If lgOpModeCRUD = CStr(UID_M0002) then
            txtMode = "CREATE"    
    ElseIf lgOpModeCRUD = CStr(UID_M0003) Then 
            txtMode = "DELETE"
    else
            txtMode = "UPDATE"
    End if    
    

    'khy2003.03
    Redim iStrSingle(ZC13_I1_Upp_Mnu)
    iStrSingle(ZC13_I1_PRCS_CD) = UCase(Request("txtPrcsCd2"))
    iStrSingle(ZC13_I1_PRCS_NM) = Request("txtPrcsNm2")
    iStrSingle(ZC13_I1_Sys_Flag) = Request("hSysFlag")
    iStrSingle(ZC13_I1_Upp_Mnu) = UCase(Request("txtUppMnu"))'khy2003.03
    

    Call iZC013.ZC_CTRL_PRCS_MNU_ASSO (gStrGlobalCollection,istrSpread,iStrSingle,LngMaxRow,txtMode,iErrorPosition)

    
    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then              
       Set iZC013 = Nothing    
       Exit Sub        
    End If
    
    Set iZC013 = Nothing    
      
    
    If txtMode = "DELETE" Then 
        Response.Write "<Script Language=vbscript>"                    & vbCr
        Response.Write "    Parent.frm1.txtPrcsCd.value = """""        & vbCr
        Response.Write "    Parent.frm1.txtPrcsNm.value = """""        & vbCr
        Response.Write "    Parent.frm1.txtPrcsCd2.value = """""        & vbCr
        Response.Write "    Parent.frm1.txtPrcsNm2.value = """""        & vbCr
        Response.Write "    Parent.frm1.txtUppMnu.value = """""        & vbCr 'khy2003.03
        Response.Write "    Parent.frm1.txtUppMnuNm.value = """""        & vbCr'khy2003.03
        Response.Write "    Parent.frm1.cbxSysFlag.Checked = False"    & vbCr        
        Response.Write "    Parent.Clear "                            & vbCr
        Response.Write "</Script>"                                    & vbCr    
    Else
        Response.Write "<Script Language=vbscript>"                    & vbCr            
        Response.Write "    Parent.frm1.txtPrcsCd.value = Parent.frm1.txtPrcsCd2.value"   & vbCr        
        Response.Write "    Parent.DbSaveOk "                    & vbCr        
        Response.Write "</Script>"                            & vbCr    
    End If 
    

End Sub
   
%>

