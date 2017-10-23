<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%                                        

    Dim lgOpModeCRUD

    On Error Resume Next                                                       
    err.Clear 

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
        Case CStr(UID_M0003)
            
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
    
    Dim iZB006                                                                    

    Dim istrCode
    Dim lgStrPrevKey    
    Dim iStrNextKey        
    Dim iLngMaxRow        
    Dim iStrData
    Dim iLngRow
    Dim iRoleType
    Dim istrMode
    
    Dim E1_Z_Usr_Role    
    
    Const C_SHEETMAXROWS_D = 100
    
    Const ZB06_I1_USR_ROLE_ID = 0
    Const ZB06_I1_USR_ROLE_NM = 1
    
    Const ZB06_I1_COMPST_ROLE_TYPE = 2
    
    On Error Resume Next 
    Err.Clear 
    
    Redim istrCode(ZB06_I1_USR_ROLE_ID)
    
    istrCode(ZB06_I1_USR_ROLE_ID) = Request("txtRoleID")

    lgStrPrevKey = Request("lgStrPrevKey")
    
    
        
    if lgStrPrevKey <> "" then
         istrCode(ZB06_I1_USR_ROLE_ID) = lgStrPrevKey
    end if 
        
    Set iZB006 = Server.CreateObject("PZBG006.cListUsrRole")

    If CheckSYSTEMError(Err,True) = True Then    
        Set iZB006 = Nothing    
       Exit Sub
    End If
    
    istrMode ="L"
    E1_Z_Usr_Role = iZB006.ZB_LIST_USR_ROLE (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode, istrMode)
    
    If CheckSYSTEMError(Err,True) = True Then        
        Response.Write "<Script language=vbscript>" & vbCr        
        Response.Write "    parent.frm1.txtRoleID.focus " & vbCr
        Response.Write "    parent.frm1.txtRoleID.select " & vbCr                                    
        Response.Write "</script>" & vbCr
        Set iZB006 = Nothing                                    
        Exit sub                                                    
    End If
       
    If lgStrPrevKey = "" then        
        Response.Write "<Script language=vbscript>" & vbCr        
        Response.Write "    parent.frm1.txtRoleID.value = """ & ConvSPChars(E1_Z_Usr_Role(ZB06_I1_USR_ROLE_ID, 0)) & """" & vbCr
        Response.Write "    parent.frm1.txtRoleNm.value = """ & ConvSPChars(E1_Z_Usr_Role(ZB06_I1_USR_ROLE_NM, 0)) & """" & vbCr                                    
        Response.Write "</script>" & vbCr
    End if 
    
    Set iZB006 = nothing 
            
    iLngMaxRow = CLng(Request("txtMaxRows"))

    For iLngRow = 0 To  UBound(E1_Z_Usr_Role,2) 
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            iStrNextKey = ConvSPChars(E1_Z_Usr_Role(0,iLngRow))
            Exit For
        End if                 
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role(ZB06_I1_USR_ROLE_ID, iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role(ZB06_I1_USR_ROLE_NM, iLngRow))             
        If ConvSPChars(E1_Z_Usr_Role(ZB06_I1_COMPST_ROLE_TYPE, iLngRow))= "1" Then 
            iRoleType="Composite Role"
        Else
            iRoleType="Menu Role"
        End If
        iStrData = iStrData & Chr(11) & iRoleType
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & iLngMaxRow + ConvSPChars(iLngRow)              
        iStrData = iStrData & Chr(11) & Chr(12)     
        
     Next
     
     
     Response.Write "<Script Language=vbscript>"                        & vbCr
     Response.Write "With Parent "                                        & vbCr
     Response.Write "    .ggoSpread.Source = .frm1.vspdData "            & vbCr
     Response.Write "    .ggoSpread.SSShowData """ & iStrData            & """" & vbCr
     Response.Write "    .lgStrPrevKey = """ & iStrNextKey                & """" & vbCr
     Response.Write "    .frm1.hRoleID.value = .frm1.txtRoleID.value "    & vbCr
     Response.Write "    .frm1.txtRoleID.Focus        "                    & vbCr
     Response.Write "    .DbQueryOk "                                    & vbCr
     Response.Write "End With "                                            & vbCr
     Response.Write "</Script>"                                            & vbCr
     
End Sub
'=========================================================================================================
Sub SubBizSaveMulti()
                                    

    Dim iZB003
    Dim iErrorPosition 
    Dim istrSpread
    
    On Error Resume Next 
    Err.Clear                                                                        
    
    Set iZB003 = Server.CreateObject("PZBG003.cCtrlUsrRole")   
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZB003 = Nothing
       Exit Sub
    End If    
    
    istrSpread = Request("txtSpread")
    

    Call iZB003.ZB_CTRL_USR_ROLE (gStrGlobalCollection,istrSpread,iErrorPosition)    
    
    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then          
       Set iZB003 = Nothing                                                         
       Exit Sub                                                        
    End If
    
    Set iZB003 = Nothing                                                   
    
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "Parent.DbSaveOk "                    & vbCr
    Response.Write "</Script>"                            & vbCr
    
End Sub
%>
