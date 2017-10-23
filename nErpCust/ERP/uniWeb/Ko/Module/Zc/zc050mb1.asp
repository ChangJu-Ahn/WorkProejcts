<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%    

    Dim lgOpModeCRUD
    
    On Error Resume Next                                                             

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

    Dim iZC003
    Dim istrCode 
    Dim lgStrPrevKey
            
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    Dim iStrData
    
    
    Dim E1_Z_Co_Mnu
    
    Const C_SHEETMAXROWS_D     = 100
    
    Const ZC03_I1_Lang_Cd      =  0
    Const ZC03_I1_Mnu_Id       =  1
        
    Const C_E1_MNU_ID          =  0
    Const C_E1_MNU_NM          =  1
    Const C_E1_ALLOW_YN        =  2
    Const C_E1_BIZ_AREA_YN     =  3
    Const C_E1_INTERNAL_YN     =  4
    Const C_E1_SUB_INTERNAL_YN =  5
    Const C_E1_PERSONAL_YN     =  6
    Const C_E1_PLANT_YN        =  7
    Const C_E1_PUR_ORG_YN      =  8
    Const C_E1_PUR_GRP_YN      =  9
    Const C_E1_SALES_ORG_YN    = 10
    Const C_E1_SALES_GRP_YN    = 11
    Const C_E1_SL_YN           = 12
    Const C_E1_WC_YN           = 13
    
    On Error Resume Next                                                             
        
    Redim istrCode(ZC03_I1_Mnu_Id)
    
    istrCode(ZC03_I1_Lang_Cd)  = Request("txtLangCd")
    istrCode(ZC03_I1_Mnu_Id)   = Request("txtMnuID")
    
    lgStrPrevKey = Request("lgStrPrevKey") 
    
    If lgStrPrevKey <>"" Then 
       istrCode(ZC03_I1_Mnu_Id) = lgStrPrevKey
    End if         
        
    Set iZC003 = Server.CreateObject("PZCG060.cListDataAuthFlag")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC003 = Nothing                                                       
       Exit Sub
    End If

    
    E1_Z_Co_Mnu = iZC003.ZC_DATA_AUTH_FLAG_LIST (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)

    If CheckSYSTEMError(Err,True) = True Then
       Set iZC003 = Nothing                                                       
       Exit Sub
    End If
    
    If IsEmpty(E1_Z_Co_Mnu) Then
        Set iZC003 = Nothing                                                       
       Exit Sub
    End If
    
    Set iZC003 = Nothing 
            
    iLngMaxRow =CLng(Request("txtMaxRows"))
    
    For iLngRow = 0 To UBound(E1_Z_Co_Mnu,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            iStrNextkey = E1_Z_Co_Mnu(0,iLngRow)
            Exit For
        End if 

        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_MNU_ID         ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_MNU_NM         ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_ALLOW_YN       ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_BIZ_AREA_YN    ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_INTERNAL_YN    ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SUB_INTERNAL_YN,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PERSONAL_YN    ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PLANT_YN       ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PUR_ORG_YN     ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PUR_GRP_YN     ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SALES_ORG_YN   ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SALES_GRP_YN   ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SL_YN          ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_WC_YN          ,iLngRow))

        iStrData = iStrData & Chr(11) & iLngMaxRow +ConvSPChars(iLngRow)                        
        iStrData = iStrData & Chr(11) & Chr(12)
    Next 
    
    Response.Write "<Script Language=vbscript>"                      & vbCr
    Response.Write "With Parent "                                    & vbCr
    Response.Write "    .ggoSpread.Source = .frm1.vspdData "         & vbCr
    Response.Write "    .ggoSpread.SSShowDataByClip  """ & iStrData  & """,""F""" & vbCrLf
    Response.Write "    .lgStrPrevKey = """ & iStrNextkey & """"     & vbCr    
    Response.Write "    .frm1.vspdData.ReDraw = True "               & vbCr   
    Response.Write "    .frm1.hMnuID.value = .frm1.txtMnuID.value "  & vbCr        
    Response.Write "    .DbQueryOk  "                                & vbCr    
    Response.Write "End With "                                       & vbCr
    Response.Write "</Script>"                                       & vbCr
    
    
End Sub

'=========================================================================================================
Sub SubBizSaveMulti()

    Dim iZC001
    Dim iErrorPosition
    Dim iStrSpread

    On Error Resume Next                                                             
    
    Set iZC001 = Server.CreateObject("PZCG060.cCtrlDataAuthFlag")

    If CheckSYSTEMError(Err,True) = True Then
       Set iZC001 = Nothing                                                       
       Exit Sub
    End If

    iStrSpread = Request("txtSpread")
    
    Call iZC001.ZC_CTRL_DATA_AUTH_FLAG(gStrGlobalCollection,istrSpread,iErrorPosition)
    
    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then          
       Set iZC001 = Nothing    
       Exit Sub                
    End If
    
    Set iZC001 = Nothing    

    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "Parent.DbSaveOk "                    & vbCr
    Response.Write "</Script>"    


End Sub

%>
