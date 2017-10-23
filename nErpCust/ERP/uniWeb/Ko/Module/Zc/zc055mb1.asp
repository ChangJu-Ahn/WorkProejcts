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
    
    Dim iLoop
    
    
    Dim E1_Z_Co_Mnu
    
    Const C_SHEETMAXROWS_D     = 100
    
    Const ZC03_I1_Mnu_id       =  0
    Const ZC03_I1_Usr_id       =  1
    Const ZC03_I1_Next         =  2
        
    Const C_E1_Mnu_ID              = 0
    Const C_E1_Mnu_Nm              = 1 
    Const C_E1_USER_ID             = 2 
    Const C_E1_USER_NM             = 3 
    Const C_E1_BIZ_AREA_CD_ALL     = 4 
    Const C_E1_BIZ_AREA_CD         = 5 
    Const C_E1_BIZ_AREA_NM         = 6 
    Const C_E1_INTERNAL_CD_ALL     = 7 
    Const C_E1_INTERNAL_CD         = 8 
    Const C_E1_INTERNAL_NM         = 9 
    Const C_E1_SUB_INTERNAL_CD_ALL = 10
    Const C_E1_SUB_INTERNAL_CD     = 11
    Const C_E1_SUB_INTERNAL_NM     = 12
    Const C_E1_PERSONAL_ID_ALL     = 13
    Const C_E1_PERSONAL_ID         = 14
    Const C_E1_PERSONAL_NM         = 15
    Const C_E1_PLANT_CD_ALL        = 16
    Const C_E1_PLANT_CD            = 17
    Const C_E1_PLANT_NM            = 18
    Const C_E1_PUR_ORG_CD_ALL      = 19
    Const C_E1_PUR_ORG_CD          = 20
    Const C_E1_PUR_ORG_NM          = 21
    Const C_E1_PUR_GRP_CD_ALL      = 22
    Const C_E1_PUR_GRP_CD          = 23
    Const C_E1_PUR_GRP_NM          = 24
    Const C_E1_SALES_ORG_CD_ALL    = 25
    Const C_E1_SALES_ORG_CD        = 26
    Const C_E1_SALES_ORG_NM        = 27
    Const C_E1_SALES_GRP_CD_ALL    = 28
    Const C_E1_SALES_GRP_CD        = 29
    Const C_E1_SALES_GRP_NM        = 30
    Const C_E1_SL_CD_ALL           = 31
    Const C_E1_SL_CD               = 32
    Const C_E1_SL_NM               = 33
    Const C_E1_WC_CD_ALL           = 34
    Const C_E1_WC_CD               = 35
    Const C_E1_WC_NM               = 36
    Const C_E1_ALLOW_YN            = 37
    Const C_E1_BIZ_AREA_YN         = 38
    Const C_E1_INTERNAL_YN         = 39
    Const C_E1_SUB_INTERNAL_YN     = 40
    Const C_E1_PERSONAL_YN         = 41
    Const C_E1_PLANT_YN            = 42
    Const C_E1_PUR_ORG_YN          = 43
    Const C_E1_PUR_GRP_YN          = 44
    Const C_E1_SALES_ORG_YN        = 45
    Const C_E1_SALES_GRP_YN        = 46
    Const C_E1_SL_YN               = 47
    Const C_E1_WC_YN               = 48
    Const C_E1_DUMMY               = 49
    
    On Error Resume Next                                                             

    Redim istrCode(ZC03_I1_Next)
    
    istrCode(ZC03_I1_Mnu_id)  = Request("txtMnuid")
    istrCode(ZC03_I1_Usr_id)  = Request("txtUsrid")
    
    lgStrPrevKey = Request("lgStrPrevKey") 
    
    If lgStrPrevKey <>"" Then 
       istrCode(ZC03_I1_Next)  = lgStrPrevKey
    End if         
        
    Set iZC003 = Server.CreateObject("PZCG061.cListDataAuth")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC003 = Nothing                                                       
       Exit Sub
    End If

    
    E1_Z_Co_Mnu = iZC003.ZC_DATA_AUTH_LIST (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)

    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"          & vbCr       
       Response.Write " Parent.frm1.txtMnuID.focus() "     & vbCr                       
       Response.Write " Parent.frm1.txtMnuID.Select() "    & vbCr
       Response.Write "</Script>"                           & vbCr
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
            iStrNextkey = E1_Z_Co_Mnu(C_E1_Mnu_ID,iLngRow) & E1_Z_Co_Mnu(C_E1_USER_ID,iLngRow)
            Exit For
        End if 

        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_Mnu_ID              ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_Mnu_Nm              ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_USER_ID             ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_USER_NM             ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_BIZ_AREA_CD_ALL     ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_BIZ_AREA_CD         ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_BIZ_AREA_NM         ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_INTERNAL_CD_ALL     ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_INTERNAL_CD         ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_INTERNAL_NM         ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SUB_INTERNAL_CD_ALL ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SUB_INTERNAL_CD     ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SUB_INTERNAL_NM     ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PERSONAL_ID_ALL     ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PERSONAL_ID         ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PERSONAL_NM         ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PLANT_CD_ALL        ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PLANT_CD            ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PLANT_NM            ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PUR_ORG_CD_ALL      ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PUR_ORG_CD          ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PUR_ORG_NM          ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PUR_GRP_CD_ALL      ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PUR_GRP_CD          ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_PUR_GRP_NM          ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SALES_ORG_CD_ALL    ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SALES_ORG_CD        ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SALES_ORG_NM        ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SALES_GRP_CD_ALL    ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SALES_GRP_CD        ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SALES_GRP_NM        ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SL_CD_ALL           ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SL_CD               ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_SL_NM               ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_WC_CD_ALL           ,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_WC_CD               ,iLngRow))
        iStrData = iStrData & Chr(11) & ""
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_WC_NM               ,iLngRow))
        
        For iLoop = 1 To 12 
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(C_E1_WC_NM  + iLoop         ,iLngRow))
        Next
        
        iStrData = iStrData & Chr(11) & iLngMaxRow + iLngRow
        iStrData = iStrData & Chr(11) & Chr(12)
    Next 
    
    Response.Write "<Script Language=vbscript>"                      & vbCr
    Response.Write " Option Explicit "                                  & vbCr
    Response.Write "Dim iStartRow "                                  & vbCr
    Response.Write "Dim iEndRow   "                                  & vbCr
    Response.Write "Dim iLoop   "                                    & vbCr
    Response.Write "With Parent.frm1 "                               & vbCr
    Response.Write "    .vspdData.ReDraw  = False                  " & vbCrLf
    Response.Write "    Parent.ggoSpread.Source = .vspdData "        & vbCr

    Response.Write "     iStartRow = .vspdData.MaxRows    " & vbCrLf

    Response.Write "    Parent.ggoSpread.SSShowDataByClip  """ & iStrData  & """,""F""" & vbCrLf
    
    Response.Write "     iEndRow   = .vspdData.MaxRows    " & vbCrLf

    Response.Write "     For iLoop = iStartRow To  iEndRow           " & vbCrLf
    Response.Write "         If Parent.GetSpreadText(.vspdData, Parent.C_ALLOW_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_BIZ_AREA_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_BIZ_AREA_CD_ALL , iLoop , Parent.C_BIZ_AREA_CD_ALL   , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_BIZ_AREA_CD     , iLoop , Parent.C_BIZ_AREA_CD       , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_BIZ_AREA_POPUP  , iLoop , Parent.C_BIZ_AREA_POPUP    , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_INTERNAL_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_INTERNAL_CD_ALL , iLoop , Parent.C_INTERNAL_CD_ALL  , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_INTERNAL_CD     , iLoop , Parent.C_INTERNAL_CD      , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_INTERNAL_POPUP  , iLoop , Parent.C_INTERNAL_POPUP   , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_SUB_INTERNAL_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SUB_INTERNAL_CD_ALL , iLoop , Parent.C_SUB_INTERNAL_CD_ALL , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SUB_INTERNAL_CD     , iLoop , Parent.C_SUB_INTERNAL_CD     , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SUB_INTERNAL_POPUP  , iLoop , Parent.C_SUB_INTERNAL_POPUP  , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_PERSONAL_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PERSONAL_ID_ALL , iLoop , Parent.C_PERSONAL_ID_ALL , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PERSONAL_ID     , iLoop , Parent.C_PERSONAL_ID     , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PERSONAL_POPUP  , iLoop , Parent.C_PERSONAL_POPUP  , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_PLANT_YN  , iLoop, ""X"", ""X"") = ""1"" Then    " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PLANT_CD_ALL , iLoop , Parent.C_PLANT_CD_ALL , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PLANT_CD     , iLoop , Parent.C_PLANT_CD     , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PLANT_POPUP  , iLoop , Parent.C_PLANT_POPUP  , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_PUR_ORG_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PUR_ORG_CD_ALL , iLoop , Parent.C_PUR_ORG_CD_ALL , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PUR_ORG_CD     , iLoop , Parent.C_PUR_ORG_CD     , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PUR_ORG_POPUP  , iLoop , Parent.C_PUR_ORG_POPUP  , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_PUR_GRP_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PUR_GRP_CD_ALL , iLoop , Parent.C_PUR_GRP_CD_ALL  , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PUR_GRP_CD     , iLoop , Parent.C_PUR_GRP_CD      , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_PUR_GRP_POPUP  , iLoop , Parent.C_PUR_GRP_POPUP   , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                        " & vbCrLf
    Response.Write "                                                                                                          " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_SALES_ORG_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SALES_ORG_CD_ALL , iLoop , Parent.C_SALES_ORG_CD_ALL , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SALES_ORG_CD     , iLoop , Parent.C_SALES_ORG_CD     , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SALES_ORG_POPUP  , iLoop , Parent.C_SALES_ORG_POPUP  , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                        " & vbCrLf
    Response.Write "                                                                                                          " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_SALES_GRP_YN  , iLoop, ""X"", ""X"") = ""1"" Then " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SALES_GRP_CD_ALL , iLoop , Parent.C_SALES_GRP_CD_ALL , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SALES_GRP_CD     , iLoop , Parent.C_SALES_GRP_CD     , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SALES_GRP_POPUP  , iLoop , Parent.C_SALES_GRP_POPUP  , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_SL_YN  , iLoop, ""X"", ""X"") = ""1"" Then       " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SL_CD_ALL , iLoop , Parent.C_SL_CD_ALL , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SL_CD     , iLoop , Parent.C_SL_CD     , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_SL_POPUP  , iLoop , Parent.C_SL_POPUP  , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "            If Parent.GetSpreadText(.vspdData, Parent.C_WC_YN  , iLoop, ""X"", ""X"") = ""1"" Then       " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_WC_CD_ALL , iLoop , Parent.C_WC_CD_ALL , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_WC_CD     , iLoop , Parent.C_WC_CD     , iLoop                                                                                          " & vbCrLf
    Response.Write "               Parent.ggoSpread.SpreadUnLock Parent.C_WC_POPUP  , iLoop , Parent.C_WC_POPUP  , iLoop                                                                                          " & vbCrLf
    Response.Write "            End If                                                                                       " & vbCrLf
    Response.Write "                                                                                                         " & vbCrLf
    Response.Write "         End If                                                                                          " & vbCrLf
    Response.Write "     Next                                                                                                " & vbCrLf

    Response.Write "    .vspdData.ReDraw  = True                   " & vbCrLf
    Response.Write "End With "                                       & vbCr

    Response.Write "With Parent "                                      & vbCr
    Response.Write "    .lgStrPrevKey       = """ & iStrNextkey & """" & vbCr    
    Response.Write "    .frm1.hUsrID.value  = .frm1.txtUsrID.value "   & vbCr        
    Response.Write "    .frm1.hMnuID.value  = .frm1.txtMnuID.value "   & vbCr        
    Response.Write "    .DbQueryOk  "                                  & vbCr    
    Response.Write "End With "                                         & vbCr
    Response.Write "</Script>"                                         & vbCr
    

End Sub

'=========================================================================================================
Sub SubBizSaveMulti()

    Dim iZC001
    Dim iErrorPosition
    Dim iStrSpread
    
    On Error Resume Next
    
    Set iZC001 = Server.CreateObject("PZCG061.cCtrlDataAuth")

    If CheckSYSTEMError(Err,True) = True Then
       Set iZC001 = Nothing                                                       
       Exit Sub
    End If

    iStrSpread = Request("txtSpread")
    
    Call iZC001.ZC_CTRL_DATA_AUTH(gStrGlobalCollection,istrSpread,iErrorPosition)
    
    If CheckSYSTEMError2(Err,True,iErrorPosition & "За ","","","","") = True Then          
       Response.Write "<Script Language=vbscript>"               & vbCr
       Response.Write "Parent.SubSetErrPos(" & iErrorPosition & ")"  & vbCr
       Response.Write "</Script>"    
       Set iZC001 = Nothing    
       Exit Sub                
    End If
    
    Set iZC001 = Nothing    

    Response.Write "<Script Language=vbscript>"            & vbCr
    Response.Write "Parent.DbSaveOk "                    & vbCr
    Response.Write "</Script>"    


End Sub

%>
