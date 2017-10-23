<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/IncServer.asp" -->

<%    

    Dim lgOpModeCRUD
    
    On Error Resume Next
    Err.Clear
    
    Call HideStatusWnd    

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
    Dim iZC009
    Dim istrCode 
    Dim lgStrPrevKey
            
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    Dim iStrData 

    Dim E1_Z_Co_Frm
    Const C_SHEETMAXROWS_D = 100
    
    Const ZC09_I1_Lang_Cd = 0
    Const ZC09_I1_Frm_Id = 1
    Const ZC09_I1_Use_Yn = 2
    Const ZC09_I1_Mnu_ID = 3
    
    Const ZC09_E1_Lang_Cd = 0
    Const ZC09_E1_Frm_Id = 1
    Const ZC09_E1_Frm_Nm = 2
    Const ZC09_E1_Use_Yn = 3
    Const ZC09_E1_Mnu_Id = 4
    Const ZC09_E1_Mnu_Nm = 5
        
        
    On Error Resume Next
    Err.Clear
    
    ReDim istrCode(ZC09_I1_Mnu_ID)    

    istrCode(ZC09_I1_Lang_Cd) = Request("txtLangCd")
    istrCode(ZC09_I1_Frm_Id) = Request("txtFrmID")
    istrCode(ZC09_I1_Use_Yn) = Request("cboUseYN")
    istrCode(ZC09_I1_Mnu_ID) = Request("txtFrmID")
    
    
    Set iZC009 = Server.CreateObject("PZCG009.cListCoFrm")

    If CheckSYSTEMError(Err,True) = True Then
       Set iZC009 = Nothing    
       Exit Sub                
    End If
    
    E1_Z_Co_Frm = iZC009.ZC_LIST_CO_FRM (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)
    
    If CheckSYSTEMError(Err,True) = True Then    
       Set iZC009 = Nothing                                                       
       Exit Sub
    End If
    
    If IsEmpty(E1_Z_Co_Frm) Then
        Set iZC009 = Nothing                                                       
       Exit Sub
    End If
    
    Set iZC009 = Nothing 
    
    iLngMaxRow = CLng(Request("txtMaxRows"))
    
    For iLngRow =0 to UBound(E1_Z_Co_Frm,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            iStrNextkey = E1_Z_Co_Frm(1,iLngRow)
            Exit For
        End if 
        
        
        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(E1_Z_Co_Frm(ZC09_E1_Lang_Cd ,iLngRow)))        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Frm(ZC09_E1_Frm_Id ,iLngRow))                            
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Frm(ZC09_E1_Frm_Nm ,iLngRow))        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Frm(ZC09_E1_Use_Yn ,iLngRow))        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Frm(ZC09_E1_Mnu_Id ,iLngRow))        
        iStrData = iStrData & Chr(11) & " "      'PopUp                
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Frm(ZC09_E1_Mnu_Nm ,iLngRow))    
        iStrData = iStrData & Chr(11) & iLngMaxRow +ConvSPChars(iLngRow)                        
        iStrData = iStrData & Chr(11) & Chr(12)
    Next 
    
    Response.Write "<Script Language=vbscript>"                        & vbCr
    Response.Write "With Parent "                                    & vbCr
    Response.Write "    .ggoSpread.Source = .frm1.vspdData "        & vbCr
    Response.Write "    .ggoSpread.SSShowData """ & iStrData        & """" & vbCr
    Response.Write "    .lgStrPrevKey = """ & iStrNextkey & """" & vbCr    
    Response.Write "    .frm1.vspdData.ReDraw = True "               & vbCr   
    Response.Write "    .frm1.hLangCd.value = .frm1.txtLangCd.value" & vbCr 
    Response.Write "    .frm1.hFrmID.value = .frm1.txtFrmID.value " & vbCr            
    Response.Write "    .frm1.hUseYN.value = .frm1.cboUseYN.value " & vbCr    
    
    Response.Write "End With "                                        & vbCr
    Response.Write "</Script>"                                        & vbCr
    
    
End Sub
'=========================================================================================================
Sub SubBizSaveMulti()
                                    

    Dim iZC008
    Dim iErrorPosition
    Dim iStrSpread        
    
    
    On Error Resume Next 
    Err.Clear

    Set iZC008 = Server.CreateObject("PZCG008.cCtrlCoFrm")

    If CheckSYSTEMError(Err,True) = True Then
       Set iZC008 = Nothing    
       Exit Sub
    End If
    
    iStrSpread = Request("txtSpread")
    
    Call iZC008.ZC_CTRL_CO_FRM (gStrGlobalCollection,istrSpread,iErrorPosition)

    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then          
       Set iZC008 = Nothing    
       Exit Sub                
    End If
    
    Set iZC008 = Nothing    

    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "Parent.DbSaveOk "                    & vbCr
    Response.Write "</Script>"    
    
        
End Sub    
%>

