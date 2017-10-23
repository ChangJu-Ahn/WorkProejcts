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

    
    Dim iZC0016
    Dim istrCode 
    Dim lgStrPrevKey
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    Dim iStrData
    Dim iStrId
    
    Dim E1_Z_Usr_Mnu
    Const C_SHEETMAXROWS_D = 100
    
    Const ZC16_E1_UPPER_MNU_ID = 0
    Const ZC16_E1_MNU_ID       = 1
    Const ZC16_E1_MNU_NM       = 2
    Const ZC16_E1_MNU_TYPE     = 3
    Const ZC16_E1_SYS_LVL      = 4
    Const ZC16_E1_MNU_SEQ      = 5
    
    
    On Error Resume Next 
    Err.Clear 

    istrCode = Request("txtMnuID")
    
    lgStrPrevKey = Request("lgStrPrevKey")
        
    If lgStrPrevKey <> "" Then
         istrCode = lgStrPrevKey
    End If 
    

    Set iZC0016 = Server.CreateObject("PZCG016.cListUsrMnu")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC0016 = Nothing                                                       
       Exit Sub
    End If
    

    E1_Z_Usr_Mnu = iZC0016.ZC_LIST_USR_MNU (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)

    
    
    If CheckSYSTEMError(Err,True) = True Then
        Response.Write "<Script language=vbscript>" & vbCr        
        Response.Write "    parent.frm1.txtMnuId.Focus " & vbCr
        Response.Write "    parent.frm1.txtMnuId.Select " & vbCr        
        Response.Write "</script>" & vbCr
       Set iZC0016 = Nothing                                                       
       Exit Sub
    End If
    
    If IsEmpty(E1_Z_Usr_Mnu) Then
        Set iZC0016 = Nothing                                                       
       Exit Sub
    End If
    
        
    If istrCode = "" Then     
        Response.Write "<Script language=vbscript>" & vbCr        
        Response.Write "    parent.frm1.txtMnuNm.value = """"" & vbCr    
        Response.Write "</script>" & vbCr
    Else 
		iStrId = split(ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_MNU_ID,iLngRow)),"^")
        Response.Write "<Script language=vbscript>" & vbCr        
        Response.Write "    parent.frm1.txtMnuID.value = """ & iStrId & """" & vbCr    
        Response.Write "    parent.frm1.txtMnuNm.value = """ & E1_Z_Usr_Mnu(ZC16_E1_MNU_NM,0) & """" & vbCr    
        Response.Write "</script>" & vbCr    
    End If 
    
    Set iZC0016 = Nothing                                                       
    
    iLngMaxRow = CLng(Request("txtMaxRows"))
        
       
    For iLngRow = 0 To UBound(E1_Z_Usr_Mnu,2)
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else
            iStrNextKey = ConvSPChars(E1_Z_Usr_Mnu(1,iLngRow))
            Exit For
        End if 
        iStrId = split(ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_MNU_ID,iLngRow)),"^")
        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_UPPER_MNU_ID,iLngRow))
        'iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_MNU_ID,iLngRow))        
		iStrData = iStrData & Chr(11) & istrid(0)'ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_MNU_ID,iLngRow))
		iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_MNU_ID,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_MNU_NM,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_MNU_TYPE,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_SYS_LVL,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Mnu(ZC16_E1_MNU_SEQ,iLngRow))
        iStrData = iStrData & Chr(11) & iLngMaxRow + ConvSPChars(iLngRow) 
        iStrData = iStrData & Chr(11) & Chr(12)
    Next
    

    Response.Write "<Script Language=vbscript>"                        & vbCr
    Response.Write "With Parent "                                    & vbCr
    Response.Write "    .ggoSpread.Source = .frm1.vspdData "        & vbCr
    Response.Write "    .ggoSpread.SSShowData """ & iStrData        & """" & vbCr
    Response.Write "    .frm1.vspdData.ReDraw = True "               & vbCr
    Response.Write "    .frm1.hMnuID.value = .frm1.txtMnuID.value " & vbCr
    Response.Write "    .lgStrPrevKey = """ & iStrNextKey            & """" & vbCr
    Response.Write "    .DbQueryOk  "                               & vbCr
    Response.Write "End With "                                        & vbCr
    Response.Write "</Script>"                                        & vbCr

End Sub
'=========================================================================================================
Sub SubBizSaveMulti()

    
    Dim iZC015
    Dim iErrorPosition
    Dim iStrSpread
        
    On Error Resume Next 
    Err.Clear 
    
    Set iZC015 = Server.CreateObject("PZCG015.cCtrlUsrMnu")
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC015 = Nothing     
       Exit Sub
    End If    

    iStrSpread = Request("txtSpread")
    
    Call iZC015.ZC_CTRL_USR_MNU (gStrGlobalCollection,iStrSpread,iErrorPosition)
     
    If CheckSYSTEMError2(Err,True,iErrorPosition & "За","","","","") = True Then              
       Set iZC015 = Nothing    
       Exit Sub        
    End If
    
    Set iZC015 = Nothing    
        
    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "Parent.DbSaveOk "                    & vbCr
    Response.Write "</Script>"                            & vbCr
    
    

End Sub

%>
