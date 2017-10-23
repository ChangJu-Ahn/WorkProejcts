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

    Dim iZC036
    Dim I1_z_dc_docname 
    Dim lgStrPrevKey
            
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    Dim iStrData
    Dim iMnuId
    Dim iMnudocId
    'Dim iMnudocType
    
    
    
    Dim E1_z_dc_docname
    Const C_SHEETMAXROWS_D = 100
    
    Const zc36_E1_LangCd        = 0
    Const zc36_E1_MnuId         = 1
    Const zc36_E1_MnuNm         = 2
    Const zc36_E1_MnudocCallNm   = 3
    'Const zc36_E1_MnuFormWidth  = 5
    'Const zc36_E1_MnuFormHeight = 6
    'Const zc36_E1_MnudocType     = 7

    Const zc36_I1_LangCd        = 0    
    Const zc36_I1_MnuId         = 1        
    Const zc36_I1_MnudocCallNm   = 2    
    'Const zc36_I1_MnudocType     = 4
        
    On Error Resume Next
    Err.Clear
        
    Redim I1_z_dc_docname(zc36_I1_MnudocCallNm)    

    I1_z_dc_docname(zc36_I1_LangCd)       = Request("txtLangCd")
    I1_z_dc_docname(zc36_I1_MnuId)        = Request("txtMnuId")
    I1_z_dc_docname(zc36_I1_MnudocCallNm)  = Request("txtMnudocCallNm")
    'I1_z_dc_docname(zc36_I1_MnudocType)    = Request("txtMnudocType")

    lgStrPrevKey = Request("lgStrPrevKey") 
    
    If lgStrPrevKey <>"" Then 
        I1_z_dc_docname(zc36_I1_MnuId) = lgStrPrevKey
    End if         
    
        
    Set iZC036 = Server.CreateObject("PZCG036.cListdocMast")
    
   
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC036 = Nothing                                                       
       Exit Sub
    End If
    
    E1_z_dc_docname = iZC036.Z_Read_doc_Mnu(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_z_dc_docname)

    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"            & vbCr       
       Response.Write "    Parent.frm1.txtLangCd.focus() "        & vbCr                       
       Response.Write " Parent.frm1.txtLangCd.Select() "    & vbCr
       Response.Write "</Script>"                            & vbCr
       Set iZC036 = Nothing                                                       
       Exit Sub
    End If
    
    If IsEmpty(E1_z_dc_docname) Then
        Set iZC036 = Nothing                                                       
       Exit Sub
    End If
    
    Set iZC036 = Nothing 
            
    iLngMaxRow =CLng(Request("txtMaxRows"))
    
    For iLngRow = 0 To UBound(E1_z_dc_docname,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            iStrNextkey  = E1_z_dc_docname(1,iLngRow)
            Exit For
        End if 

        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(E1_z_dc_docname(zc36_E1_LangCd  ,iLngRow)))    
        iStrData = iStrData & Chr(11) & " "      'PopUp
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_dc_docname(zc36_E1_MnuId        ,iLngRow))                            
        iStrData = iStrData & Chr(11) & " "      'PopUp
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_dc_docname(zc36_E1_MnuNm        ,iLngRow))                
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_z_dc_docname(zc36_E1_MnudocCallNm  ,iLngRow))       
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

    Dim iZC036
    Dim iErrorPosition
    Dim iStrSpread
    
    On Error Resume Next 
    Err.Clear

    Set iZC036 = Server.CreateObject("PZCG036.cListdocMast")
    If CheckSYSTEMError(Err,True) = True Then    
       Set iZC036 = Nothing                                                       
       Exit Sub
    End If

    iStrSpread = Request("txtSpread")
    
    Call iZC036.ZC_doc_CTRL_MNU(gStrGlobalCollection,iStrSpread,iErrorPosition)
        
    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then          
       Set iZC036 = Nothing    
       Exit Sub                
    End If

    Set iZC036 = Nothing    

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
'    Call svrmsgbox("3", vbinformation, i_mkscript)
%>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                
