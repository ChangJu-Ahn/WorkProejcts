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

    Dim iZC037
    Dim I1_z_pr_aspname 
    Dim E1_z_pr_aspname
    Dim lgStrPrevKey
            
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    Dim iStrData
    Dim iPgmId
    Dim iCalledUpperFid
    
    Const C_SHEETMAXROWS_D = 100
    
    Const ZC37_E1_LangCd            = 0
    Const ZC37_E1_PgmId                = 1
    Const ZC37_E1_PgmNm                = 2
    Const ZC37_E1_CalledId            = 3
    Const ZC37_E1_CalledUpperFid    = 4

    Const ZC37_I1_LangCd            = 0    
    Const ZC37_I1_PgmId                = 1        
    Const ZC37_I1_CalledUpperFid    = 2
        
    On Error Resume Next
    Err.Clear
        
    Redim I1_z_pr_aspname(ZC37_I1_CalledUpperFid)    

    I1_z_pr_aspname(ZC37_I1_LangCd)         = Request("txtLangCd")
    I1_z_pr_aspname(ZC37_I1_PgmId)          = Request("txtPgmId")
    I1_z_pr_aspname(ZC37_I1_CalledUpperFid) = Request("txtCalledUpperFid")
    
    lgStrPrevKey = Request("lgStrPrevKey") 
    
    If lgStrPrevKey <>"" Then 
        I1_z_pr_aspname(ZC37_I1_PgmId) = lgStrPrevKey
    End if         
    
        
    Set iZC037 = Server.CreateObject("PZCG037.cListAspMnu")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC033 = Nothing                                                       
       Exit Sub
    End If
    
    E1_z_pr_aspname = iZC037.Z_ASP_LIST_MNU(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_pr_aspname)

    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"            & vbCr       
       Response.Write "    Parent.frm1.txtLangCd.focus() "        & vbCr                       
       Response.Write " Parent.frm1.txtLangCd.Select() "    & vbCr
       Response.Write "</Script>"                            & vbCr
       Set iZC037 = Nothing                                                       
       Exit Sub
    End If
    
    If IsEmpty(E1_z_pr_aspname) Then
        Set iZC037 = Nothing                                                       
       Exit Sub
    End If
    
    Set iZC037 = Nothing 
            
    iLngMaxRow =CLng(Request("txtMaxRows"))
    
    For iLngRow = 0 To UBound(E1_z_pr_aspname,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            iStrNextkey         = E1_z_pr_aspname(1,iLngRow)
            iCalledUpperFid  = E1_z_pr_aspname(4,iLngRow)
            Exit For
        End if 

        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(E1_z_pr_aspname(ZC37_E1_LangCd,iLngRow)))   
        iStrData = iStrData & Chr(11) & " "      'PopUp 
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_pr_aspname(ZC37_E1_PgmId,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_pr_aspname(ZC37_E1_PgmNm,iLngRow))                            
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_pr_aspname(ZC37_E1_CalledId,iLngRow))                
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_pr_aspname(ZC37_E1_CalledUpperFid,iLngRow))          
        iStrData = iStrData & Chr(11) & iLngMaxRow +ConvSPChars(iLngRow)                        
        iStrData = iStrData & Chr(11) & Chr(12)
    Next 

    Response.Write "<Script Language=vbscript>"                                            & vbCr
    Response.Write "With Parent "                                                        & vbCr 
    Response.Write "    .ggoSpread.Source            =    .frm1.vspdData "                & vbCr
    Response.Write "    .ggoSpread.SSShowData            """ & iStrData          & """"    & vbCr
    Response.Write "    .frm1.vspdData.ReDraw        =   True "                           & vbCr
    Response.Write "    .lgStrPrevKey                =    """ & iStrNextkey     & """"    & vbCr         
    Response.Write "    .frm1.hLangCd.value            =   .frm1.txtLangCd.value"            & vbCr 
    Response.Write "    .frm1.hCalledUpperFid.value    =   """ & iCalledUpperFid & """"    & vbCr
    Response.Write "    .DbQueryOk  "                                                    & vbCr
    Response.Write "End With "                                                            & vbCr
    Response.Write "</Script>"                                                            & vbCr
    
    
End Sub

'=========================================================================================================
Sub SubBizSaveMulti()

    Dim iZC035
    Dim iErrorPosition
    Dim iStrSpread
    
    On Error Resume Next 
    Err.Clear

    
    Set iZC035 = Server.CreateObject("PZCG035.cCtrlAspMnu")
    If CheckSYSTEMError(Err,True) = True Then    
       Set iZC031 = Nothing                                                       
       Exit Sub
    End If

    iStrSpread = Request("txtSpread")
    
    
    Call iZC035.ZC_ASP_CTRL_MNU(gStrGlobalCollection,iStrSpread,iErrorPosition)
        
    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then          
       Set iZC035 = Nothing    
       Exit Sub                
    End If
    
    Set iZC035 = Nothing    

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
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                