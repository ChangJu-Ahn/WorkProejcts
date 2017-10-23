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

    Dim iZC003
    Dim istrCode 
    Dim lgStrPrevKey
            
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    Dim iStrData
    
    
    Dim E1_Z_Co_Mnu
    Const C_SHEETMAXROWS_D = 100
    
    Const ZC03_I1_Lang_Cd = 0
    Const ZC03_I1_Mnu_Id = 1
    Const Zc03_I1_Mnu_Type = 2
    Const Zc03_I1_Mnu_Yn = 3
        
    Const ZC03_E1_Lang_Cd = 0
    Const ZC03_E1_Mnu_Id = 1
    Const ZC03_E1_Mnu_Nm = 2
    Const ZC03_E1_Mnu_Type = 3
    Const ZC03_E1_Sys_Lve = 4    
    Const ZC03_E1_Called_Frm_Id = 5
    Const ZC03_E1_Mnu_Seq = 6
    Const ZC03_E1_Use_Yn = 7
    Const ZC03_E1_Upper_Mnu_Id = 8
        
    On Error Resume Next
    Err.Clear
        
    Redim istrCode(ZC03_E1_Upper_Mnu_Id)
    
    istrCode(ZC03_I1_Lang_Cd)  = Request("txtLangCd")
    istrCode(ZC03_I1_Mnu_Id)   = Request("txtMnuID")
    istrCode(Zc03_I1_Mnu_Type) = Request("cboMnuType")
    istrCode(Zc03_I1_Mnu_Yn)   = Request("cboUseYN")
    
    
    lgStrPrevKey = Request("lgStrPrevKey") 
    
    If lgStrPrevKey <>"" Then 
        istrCode(ZC03_I1_Mnu_Id) = lgStrPrevKey
    End if         
    
        
    Set iZC003 = Server.CreateObject("PZCG003.cListCoMastMnu")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC003 = Nothing                                                       
       Exit Sub
    End If

    
    E1_Z_Co_Mnu = iZC003.ZC_LIST_CO_MNU (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)

    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"            & vbCr       
       Response.Write "    Parent.frm1.txtLangCd.focus() "        & vbCr                       
       Response.Write " Parent.frm1.txtLangCd.Select() "    & vbCr
       Response.Write "</Script>"                            & vbCr
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
            iStrNextkey = E1_Z_Co_Mnu(1,iLngRow)
            Exit For
        End if 

        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(E1_Z_Co_Mnu(ZC03_E1_Lang_Cd,iLngRow)))  
        iStrData = iStrData & Chr(11) & " "      'PopUp      
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(ZC03_E1_Mnu_Id,iLngRow))                            
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(ZC03_E1_Mnu_Nm,iLngRow))        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(ZC03_E1_Mnu_Type,iLngRow))        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(ZC03_E1_Sys_Lve,iLngRow))        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(ZC03_E1_Called_Frm_Id,iLngRow))    
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(ZC03_E1_Mnu_Seq,iLngRow))        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(ZC03_E1_Use_Yn,iLngRow))        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(ZC03_E1_Upper_Mnu_Id,iLngRow))    
        iStrData = iStrData & Chr(11) & " "     

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
    Response.Write "    .frm1.hMnuID.value = .frm1.txtMnuID.value " & vbCr        
    Response.Write "    .frm1.hMnuType.value = .frm1.cboMnuType.value" & vbCr
    Response.Write "    .frm1.hUseYN.value = .frm1.cboUseYN.value " & vbCr    
    Response.Write "    .DbQueryOk  "                               & vbCr    
    Response.Write "End With "                                        & vbCr
    Response.Write "</Script>"                                        & vbCr
    
    
End Sub

'=========================================================================================================
Sub SubBizSaveMulti()

    Dim iZC001
    Dim iErrorPosition
    Dim iStrSpread
    
    On Error Resume Next 
    Err.Clear
    
    
    Set iZC001 = Server.CreateObject("PZCG001.cCtrlCoMastMnu")

    If CheckSYSTEMError(Err,True) = True Then
       Set iZC001 = Nothing                                                       
       Exit Sub
    End If

    iStrSpread = Request("txtSpread")
    
    Call iZC001.ZC_CTRL_CO_MNU(gStrGlobalCollection,istrSpread,iErrorPosition)
    
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
