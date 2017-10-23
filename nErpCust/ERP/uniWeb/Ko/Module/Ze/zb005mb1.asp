<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%    

    
    Dim strMode
    
    On Error Resume Next
    Err.Clear 
    
    Call HideStatusWnd                                                               

    Call LoadBasisGlobalInf()        
    '------ Developer Coding part (Start ) ------------------------------------------------------------------

    '------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    strMode = Request("txtMode")
    
    Select Case StrMode
        Case CStr(UID_M0001)
            Call SubBizQueryMulti()
        Case CStr(UID_M0002)
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

    Dim iZb011
    
    Dim istrCode 
    Dim istrNm
    Dim lgStrNextKey        
    Dim lgStrPrevKey
    Dim iLngMaxRow        
    Dim iStrData
    Dim iLngRow

    Dim I1_z_usr_role_compst
    Dim E1_z_usr_role_compst
    Dim E2_z_usr_role_compst

    Const ZB11_I1_usr_role_id = 0

    Const ZB11_E1_usr_role_id = 0
    Const ZB11_E1_usr_role_nm = 1
    
    Const C_SHEETMAXROWS_D = 100
    
    
    On Error Resume Next 
    Err.Clear 

    '---------- Developer Coding part (Start) ---------------------------------------------------------------    
    
    lgStrPrevKey = Request("lgStrPrevKey")
        
    Set iZb011 = Server.CreateObject("PZBG011.cListUsrRoleComRole")
    
    If CheckSYSTEMError(Err,True)=True then
        Set iZb011= Nothing 
        Exit Sub
    End If
    
    Redim I1_z_usr_role_compst(ZB11_I1_usr_role_id)
    I1_z_usr_role_compst(ZB11_I1_usr_role_id) = Request("txtCompositeID")

    If Request("txtPrvNext") = "P" or Request("txtPrvNext") = "N" Then
        E2_z_usr_role_compst = iZb011.ZB_Select_Pre_Next_Compst_Role(gStrGlobalCollection, Request("txtPrvNext"), Request("txtCompositeID"))
        
        If CheckSYSTEMError(Err,True) = True Then
           Response.Write "<Script language=vbscript>"                & vbCr            
           Response.Write "Parent.frm1.txtCompositeID.focus  "        & vbCr
           Response.Write "Parent.frm1.txtCompositeID.select  "    & vbCr                       
           Response.Write "Parent.frm1.txtCD.Value =          """ & ConvSPChars(E2_z_usr_role_compst(0)) & """" & vbCr
		   Response.Write "Parent.frm1.txtNM.Value =          """ & ConvSPChars(E2_z_usr_role_compst(1)) & """" & vbCr
           Response.Write "Parent.lgStrQueryFlag  ="""" " & vbCr                                           
           Response.Write "</script>"                                & vbCr
        
           Set iZb011 = Nothing                                                         
           Exit Sub
        Else
           Response.Write "<Script Language=vbscript>"            & vbCr
           Response.Write "With Parent "                          & vbCr    
           Response.Write ".frm1.txtCompositeID.Value = """ & ConvSPChars(E2_z_usr_role_compst(0)) & """" & vbCr
	       Response.Write ".frm1.txtCompositeNm.Value = """ & ConvSPChars(E2_z_usr_role_compst(1)) & """" & vbCr
		   Response.Write ".frm1.txtCD.Value =          """ & ConvSPChars(E2_z_usr_role_compst(0)) & """" & vbCr
		   Response.Write ".frm1.txtNM.Value =          """ & ConvSPChars(E2_z_usr_role_compst(1)) & """" & vbCr       
           Response.Write "Call .EraseContents "        & vbCr
           Response.Write "End With  "                            & vbCr           
           Response.Write "</Script>" & vbCr                    
        End If

        I1_z_usr_role_compst(ZB11_I1_usr_role_id) = CStr(E2_z_usr_role_compst(0))
        
    End If

    E1_z_usr_role_compst = iZb011.ZB_LIST_USR_ROLE_COMPST_ROLE (gStrGlobalCollection ,C_SHEETMAXROWS_D, I1_z_usr_role_compst(ZB11_I1_usr_role_id))
    
    If CheckSYSTEMError(Err,True)= True Then                     
        Response.Write "<Script language=vbscript>"                & vbCr            
        Response.Write "Parent.frm1.txtCompositeID.focus  "        & vbCr
        Response.Write "Parent.frm1.txtCompositeID.select  "    & vbCr            
        Response.Write "Parent.frm1.txtCompositeID.value = """       &  UCase(ConvSPChars(I1_z_usr_role_compst(ZB11_I1_usr_role_id)))       & """" & vbCr     
        'Response.Write "Parent.frm1.txtCompositeNm.value="""" " & vbCr
        Response.Write "Parent.frm1.txtCd.value=Parent.frm1.txtCompositeID.value" & vbCr
        Response.Write "Parent.frm1.txtNm.value=Parent.frm1.txtCompositeNm.value" & vbCr                
        Response.Write "Parent.SetToolBar(""1100101011011111"") "        & vbCr        
        Response.Write "</script>"                                & vbCr
    
        Set iZb011 = Nothing 
        Exit Sub
    End If 

    If IsEmpty(E1_z_usr_role_compst) Then
       Exit Sub
    End If
    
    Set iZb011 = nothing 
            
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To UBound(E1_z_usr_role_compst,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            lgStrPrvNext = ConvSPChars(E1_z_usr_role_compst(ZB11_E1_usr_role_id,iLngRow))
            Exit For
        End If
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_usr_role_compst(ZB11_E1_usr_role_id,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_usr_role_compst(ZB11_E1_usr_role_nm,iLngRow))
        iStrData = iStrData & Chr(11) & iLngMaxRow + iLngRow + 1
        iStrData = iStrData & Chr(11) & Chr(12)
    Next 

    Response.Write "<Script Language=vbscript>"                            & vbCr
    Response.Write "With Parent "                                        & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "                & vbCr
    Response.Write ".ggoSpread.SSShowData """ & iStrData                & """" & vbCr
    Response.Write ".lgStrPrevKey =  lgStrNextKey"                    &  vbCr    
    Response.Write ".frm1.txtCompositeID.value = """       &  UCase(ConvSPChars(I1_z_usr_role_compst(ZB11_I1_usr_role_id)))       & """" & vbCr     '다시 체크               
    Response.Write ".frm1.hCompositeID.value = .frm1.txtCompositeID.value"    & vbCr
    Response.Write ".frm1.hCompositeNm.value = .frm1.txtCompositeNm.value"    & vbCr
    Response.Write ".frm1.txtCD.value = .frm1.txtCompositeID.value"    & vbCr
    Response.Write ".frm1.txtNm.value = .frm1.txtCompositeNm.value"    & vbCr
    Response.Write ".DBQueryOK "                                        & vbCr
    
    Response.Write "End With "                                            & vbCr
    Response.Write "</Script>"                                            & vbCr
    

End Sub
'=========================================================================================================
Sub SubBizSaveMulti()

    Dim iZBG010
    Dim iErrorPosition
    
    On Error Resume Next                                    
    Err.Clear            
    
    Set iZBG010 = Server.CreateObject("PZBG010.cCtrlUsrRoleComRole")
    
    If CheckSYSTEMError(Err,True) = True Then 
        Set iZBG010 = Nothing                                                         
        Exit Sub
    End if 


    Call iZBG010.ZB_CTRL_USR_ROLE_COM_ROLE(gStrGlobalCollection,FilterVar(Request("txtSpread"),"","SNM"),iErrorPosition)
                 
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then          
       Set iZb011 = Nothing                                                         
       Exit Sub                                                        
    End If
    
    Set iZb011 = Nothing
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "Parent.DbSaveOk "                    & vbCr
    Response.Write "</Script>"                            & vbCr

End Sub

%>
