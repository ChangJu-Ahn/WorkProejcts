<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->

<%    

    Dim lgOpModeCRUD
    
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd    
    Call LoadBasisGlobalInf()
    
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

    Dim iZC033
    Dim I1_z_dc_ebname 
    Dim lgStrPrevKey
            
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    Dim iStrData
    Dim iMnuId
    Dim iMnuEbId
    Dim iMnuEbType
    
    
    
    Dim E1_z_dc_ebname
    Const C_SHEETMAXROWS_D = 100
    
    Const ZC33_E1_LangCd        = 0
    Const ZC33_E1_MnuEbId       = 1
    Const ZC33_E1_MnuId         = 2
    Const ZC33_E1_MnuNm         = 3
    Const ZC33_E1_MnuEbCallNm   = 4
    Const ZC33_E1_MnuFormWidth  = 5
    Const ZC33_E1_MnuFormHeight = 6
    Const ZC33_E1_MnuFormWidth2  = 7
    Const ZC33_E1_MnuFormHeight2 = 8
    Const ZC33_E1_MnuEbType     = 9

    Const ZC33_I1_LangCd        = 0    
    Const ZC33_I1_MnuId         = 1        
    Const ZC33_I1_MnuEbId       = 2    
    Const ZC33_I1_MnuEbCallNm   = 3    
    Const ZC33_I1_MnuEbType     = 4
        
    On Error Resume Next
    Err.Clear
        
    Redim I1_z_dc_ebname(ZC33_I1_MnuEbType)    

    I1_z_dc_ebname(ZC33_I1_LangCd)       = Request("txtLangCd")
    I1_z_dc_ebname(ZC33_I1_MnuId)        = Request("txtMnuId")
    I1_z_dc_ebname(ZC33_I1_MnuEbId)      = Request("txtMnuEbId")
    I1_z_dc_ebname(ZC33_I1_MnuEbCallNm)  = Request("txtMnuEbCallNm")
    I1_z_dc_ebname(ZC33_I1_MnuEbType)    = Request("txtMnuEbType")

    lgStrPrevKey = Request("lgStrPrevKey") 
    
    If lgStrPrevKey <>"" Then 
        I1_z_dc_ebname(ZC33_I1_MnuEbId) = lgStrPrevKey
    End if         
    
        
    Set iZC033 = Server.CreateObject("PZCG033.cListEBMast")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC033 = Nothing                                                       
       Exit Sub
    End If
    
    E1_z_dc_ebname = iZC033.Z_Read_Eb_Mnu(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_dc_ebname)

    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"            & vbCr       
       Response.Write "    Parent.frm1.txtLangCd.focus() "        & vbCr                       
       Response.Write " Parent.frm1.txtLangCd.Select() "    & vbCr
       Response.Write "</Script>"                            & vbCr
       Set iZC033 = Nothing                                                       
       Exit Sub
    End If
    
    If IsEmpty(E1_z_dc_ebname) Then
        Set iZC033 = Nothing                                                       
       Exit Sub
    End If
    
    Set iZC033 = Nothing 
            
    iLngMaxRow =CLng(Request("txtMaxRows"))
    
    For iLngRow = 0 To UBound(E1_z_dc_ebname,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            iStrNextkey  = E1_z_dc_ebname(1,iLngRow)
            Exit For
        End if 

        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(E1_z_dc_ebname(ZC33_E1_LangCd  ,iLngRow)))    
        iStrData = iStrData & Chr(11) & " "      'PopUp
        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(E1_z_dc_ebname(ZC33_E1_MnuEbId ,iLngRow)))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_dc_ebname(ZC33_E1_MnuId        ,iLngRow))                            
        iStrData = iStrData & Chr(11) & " "      'PopUp
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_dc_ebname(ZC33_E1_MnuNm        ,iLngRow))                
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_dc_ebname(ZC33_E1_MnuEbCallNm  ,iLngRow))       
        iStrData = iStrData & Chr(11) &   E1_z_dc_ebname(ZC33_E1_MnuFormWidth ,iLngRow)
        iStrData = iStrData & Chr(11) &   E1_z_dc_ebname(ZC33_E1_MnuFormHeight,iLngRow)
        iStrData = iStrData & Chr(11) &   E1_z_dc_ebname(ZC33_E1_MnuFormWidth2 ,iLngRow)
        iStrData = iStrData & Chr(11) &   E1_z_dc_ebname(ZC33_E1_MnuFormHeight2,iLngRow)
        iStrData = iStrData & Chr(11) &   E1_z_dc_ebname(ZC33_E1_MnuEbType    ,iLngRow)
        iStrData = iStrData & Chr(11) & iLngMaxRow +ConvSPChars(iLngRow)                        
        iStrData = iStrData & Chr(11) & Chr(12)
    Next 

    Response.Write "<Script Language=vbscript>"                                  & vbCr
    Response.Write "With Parent "                                                & vbCr
    Response.Write "    .ggoSpread.Source        =    .frm1.vspdData "           & vbCr
    Response.Write "    .ggoSpread.SSShowData        """ & iStrData      & """" & vbCr
    Response.Write "    .frm1.vspdData.ReDraw    =   True "                      & vbCr  
    Response.Write "    .lgStrPrevKey            =    """ & iStrNextkey & """"   & vbCr      
    Response.Write "    .frm1.hLangCd.value      =    .frm1.txtLangCd.value "    & vbCr  
    Response.Write "    .DbQueryOk  "                                            & vbCr
    Response.Write "End With "                                                   & vbCr
    Response.Write "</Script>"                                                   & vbCr

End Sub

'=========================================================================================================    
Sub SubBizSaveMulti()

    Dim iZC031
    Dim iErrorPosition
    Dim iStrSpread
    
    On Error Resume Next 
    Err.Clear

    
    Set iZC031 = Server.CreateObject("PZCG031.cCtrlEbMnu")
    If CheckSYSTEMError(Err,True) = True Then    
       Set iZC031 = Nothing                                                       
       Exit Sub
    End If

    iStrSpread = Request("txtSpread")
    
    
    Call iZC031.ZC_Eb_CTRL_MNU(gStrGlobalCollection,iStrSpread,iErrorPosition)
        
    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then          
       Set iZC031 = Nothing    
       Exit Sub                
    End If
    
    Set iZC031 = Nothing    

    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "Parent.DbSaveOk "                    & vbCr
    Response.Write "</Script>"    


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
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                
