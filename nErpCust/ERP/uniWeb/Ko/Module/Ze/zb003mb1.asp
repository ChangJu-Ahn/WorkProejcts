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

    Dim iZB09
    
    Dim istrCode 
    Dim istrNm
    
    Dim lgstrPrvNext    
    Dim lgStrNextKey    
    
    Dim lgstrNextActId  
    Dim lgStrPrevActId  
    
    Dim lgstrPrvKey
    Dim iLngMaxRow      
    Dim iStrData
    Dim iLngRow
    Dim lgstrMenu
        
    Dim E1_ZB_LIST_USR_ROLE_MNU_AUTH
    
    Const C_SHEETMAXROWS_D = 100
    
    Const ZB09_I1_Role_Id = 0
    Const ZB09_I1_Mnu_Type = 1    
    Const ZB09_I1_Mnu_Id = 2
    Const ZB09_I1_Lang_Cd = 3
    
    
    
    On Error Resume Next
    Err.Clear 
    
    Redim istrCode(ZB09_I1_Lang_Cd)
    
    istrCode(ZB09_I1_Role_Id) = Request("txtRoleID")
    istrCode(ZB09_I1_Mnu_Type) = Request("txtMenuType")
    istrCode(ZB09_I1_Mnu_Id) = Request("lgStrPrevKey")
    istrCode(ZB09_I1_Lang_Cd) = gLang
    
    istrNm = Request("txtroleNm")
    lgstrPrvNext = Request("txtPrvNext")
    
    Set iZB09 = Server.CreateObject("PZBG009.cListUsrRoleMnuAuth")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZB09 = Nothing                                                       
       Exit Sub
    End If
    

    If lgstrPrvNext = "P" OR lgstrPrvNext = "N" then             
        E1_ZB_LIST_USR_ROLE_MNU_AUTH= iZB09.ZB_SELECT_PRE_NEXT_USR_ROLE (gStrGlobalCollection,lgstrPrvNext,istrCode(ZB09_I1_Role_Id))
        
    
        If CheckSYSTEMError(Err,True)= True Then 
            Response.Write "<Script language=vbscript>" & vbCr        
            Response.Write "Parent.frm1.txtRoleID.focus " & vbCr
            Response.Write "Parent.frm1.txtRoleID.select " & vbCr
            Response.Write "</script>" & vbCr    
            Set iZB09 = Nothing 
            Exit Sub
        End If 
        
        Response.Write "<Script language=vbscript>" & vbCr        
        Response.Write "Parent.Clear    " & vbCr                            
        Response.Write "Parent.frm1.txtRoleID.Value = """ & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,0)) & """" & vbCr
        Response.Write "Parent.frm1.txtRoleNm.Value = """ & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(1,0)) & """" & vbCr
        Response.Write "Parent.frm1.txtCD.Value =          """ & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,0)) & """" & vbCr
        Response.Write "Parent.frm1.txtNM.Value =          """ & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(1,0)) & """" & vbCr
        Response.Write "</script>" & vbCr
        
        
        istrCode(ZB09_I1_Role_Id) = E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,0)
                
    Else 
        E1_ZB_LIST_USR_ROLE_MNU_AUTH= iZB09.ZB_SELECT_PRE_NEXT_USR_ROLE (gStrGlobalCollection,"",istrCode(ZB09_I1_Role_Id))
                
        If CheckSYSTEMError(Err,True)= True Then
            Response.Write "<Script language=vbscript>" & vbCr        
            Response.Write "Parent.frm1.txtRoleID.focus " & vbCr
            Response.Write "Parent.frm1.txtRoleID.select " & vbCr
            Response.Write "Parent.frm1.txtRoleNm.value ="""" " & vbCr
            Response.Write "Parent.frm1.txtMenuID.value ="""" " & vbCr
            Response.Write "Parent.frm1.txtCd.value ="""" " & vbCr
            Response.Write "Parent.frm1.txtNm.value ="""" " & vbCr
            Response.Write "</script>" & vbCr        
            Set iZB09 = Nothing 
            Exit Sub
        End If 
        
        Response.Write "<Script language=vbscript>" & vbCr                
        Response.Write "Parent.frm1.txtRoleID.Value = """ & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,0)) & """" & vbCr
        Response.Write "Parent.frm1.txtRoleNm.Value = """ & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(1,0)) & """" & vbCr
        Response.Write "Parent.frm1.txtCD.Value =          """ & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,0)) & """" & vbCr
        Response.Write "Parent.frm1.txtNM.Value =          """ & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(1,0)) & """" & vbCr        
        Response.Write "</script>" & vbCr
        
        
    End if 
    
    Erase E1_ZB_LIST_USR_ROLE_MNU_AUTH 
    

        
    E1_ZB_LIST_USR_ROLE_MNU_AUTH = iZB09.ZB_LIST_USR_ROLE_MNU_AUTH (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)
    
    
    
    If CheckSYSTEMError(Err,True) = True Then        
       Set iZB09 = Nothing           
       Exit Sub
    End If
    
    If IsEmpty(E1_ZB_LIST_USR_ROLE_MNU_AUTH) then    
        Set iZB09 = Nothing 
        Exit Sub
    End If
    
    Set iZB09 = Nothing 
    
    If IsEmpty(E1_ZB_LIST_USR_ROLE_MNU_AUTH) Then
        Set iZB09 = Nothing        
        Response.Write "<Script Language=vbscript>"                            & vbCr        
        Response.Write "Parent.DBQueryOK "                                        & vbCr
        Response.Write "</Script>"
                                                   
       Exit Sub
    End If
    
    iLngMaxRow =CLng(Request("txtMaxRows"))
    
    For iLngRow = 0 To UBound(E1_ZB_LIST_USR_ROLE_MNU_AUTH,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            lgStrNextkey = ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(1,iLngRow))
            Exit For
        End if 
    
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(1,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(2,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(3,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(4,iLngRow))
            
        If ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(4,iLngRow))="A" Then 
            lgstrMenu = "All"
        ElseIf ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(4,iLngRow))="E" Then 
            lgstrmenu = "Excel / Print "
        ElseIf ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(4,iLngRow))="Q" Then 
            lgstrMenu="Query"
        Else
            lgstrMenu = "None"
        End if 
        iStrData = iStrData & Chr(11) & lgstrMenu
        iStrData = iStrData & Chr(11) & iLngMaxRow + ConvSPChars(iLngRow)
        iStrData = iStrData & Chr(11) & Chr(12)
    Next

    lgStrNextKey = E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,iLngRow)
    lgStrNextActId = E1_ZB_LIST_USR_ROLE_MNU_AUTH(3,iLngRow)
    
    
    If istrData =Chr(11) & chr(12) Then 
        istrData =""
    End If 
        
    Response.Write "<Script Language=vbscript>"                            & vbCr
    Response.Write "With Parent "                                        & vbCr
    Response.Write ".ggoSpread.Source = .frm1.vspdData "                & vbCr
    Response.Write ".ggoSpread.SSShowData """ & iStrData                & """" & vbCr
    Response.Write ".lgStrPrevKey =""" & lgStrNextKey                    & """" & vbCr    
    Response.Write ".lgStrPrevActId =""" & lgStrNextActId                & """" & vbCr            
    Response.Write ".frm1.hRoleID.value = .frm1.txtRoleID.value"           & vbCr    
    Response.Write ".DBQueryOK "                                        & vbCr
    
    Response.Write "End With "                                            & vbCr
    Response.Write "</Script>"
    
   
   
End Sub
'=========================================================================================================
Sub SubBizSaveMulti()                                                              
    
    Dim iZBG008
    Dim iErrorPosition    
    Dim strCode                             
    
    On Error Resume Next
    Err.Clear
    
    Set iZBG008 = Server.CreateObject("PZBG008.cCtrlUsrRoleMnuAuth")

    If CheckSYSTEMError(Err,True) = True Then
       Set iZBG008 = Nothing        
       Exit Sub                        
    End If

    Call iZBG008.ZB_CTRL_USR_ROLE_MENU_AUTH(gStrGlobalCollection,Request("txtSpread"),iErrorPosition)

    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then          
       Set iZBG008 = Nothing                                                   
       Response.End 
    End If
    
    Set iZBG008 = Nothing

    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "Parent.DbSaveOk "                    & vbCr
    Response.Write "</Script>"                            & vbCr

End Sub 
%>


