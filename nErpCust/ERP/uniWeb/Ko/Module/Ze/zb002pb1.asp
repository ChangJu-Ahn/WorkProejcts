<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%                                                                            

    On Error Resume Next
    err.Clear 

    Call HideStatusWnd                                                               

    Call LoadBasisGlobalInf()        
    '------ Developer Coding part (Start ) ------------------------------------------------------------------

    '------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Dim iZB021
    Dim istrCode
    Dim E1_Z_Usr_Role    
    Dim iLngRow
    Dim StrNextKey        
    Dim iLngMaxRow
    
    Dim iStrData
    Dim istrMode
    
    Dim E1_Z_Usr_Role_Detail
    Const C_SHEETMAXROWS_D = 100 
        
    Const ZB21_I1_Role_Id = 0
    Const ZB21_I1_Mnu_Id = 1
    Const ZB21_I1_Mnu_Type = 2
    Const ZB21_I1_Lang_Cd = 3
        
    Const ZB21_E1_Mnu_Id = 0
    Const ZB21_E1_Mnu_Nm = 1
    Const ZB21_E1_Mnu_Type = 2
    Const ZB21_E1_Mnu_Action = 3
    Const ZB21_E1_Role_Nm = 4
    
    
    Redim istrCode(ZB21_I1_Lang_Cd)
    
    istrCode(ZB21_I1_Role_Id) = FilterVar(Request("txtroleid"),"","SNM")
    istrCode(ZB21_I1_Mnu_Id) = FilterVar(Request("txtmenuid"),"","SNM")
    istrCode(ZB21_I1_Mnu_Type) = Request("txtmenutype")
    istrCode(ZB21_I1_Lang_Cd) = gLang
    
    
    If Request("txtCode") <> "" Then        
        istrCode(ZB21_I1_Mnu_Id) = FilterVar(Request("txtCode"),"","SNM")
    End If        
    
    
    Set iZB021 = Server.CreateObject("PZBG021.cListUsrRoleDetail")
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZB021 = Nothing                                                         
       Response.End 
    End If
        
    E1_Z_Usr_Role_Detail = iZB021.ZB_LIST_ROLE_DETAIL (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)


    If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"                                    & vbCr
       Response.Write "        Parent.txtcd.focus    "                                    & vbCr
       Response.Write "        Parent.txtcd.select    "                                    & vbCr
       Response.Write "</Script>"                                                    & vbCr
       Set iZB021 = Nothing                                    
       Response.End 
    End If

    If IsEmpty(E1_Z_Usr_Role_Detail) Then
        Set iZB021 = Nothing    
       Response.End 
    End If
    
    Set iZB021 = nothing 
        
    iLngMaxRow = CLng(Request("txtMaxRows"))    
    
    For iLngRow = 0 To  UBound(E1_Z_Usr_Role_Detail,2)         
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else        
            StrNextKey = ConvSPChars(E1_Z_Usr_Role_Detail(0,iLngRow))
            Exit For
        End if                 
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role_Detail(ZB21_E1_Mnu_Id, iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role_Detail(ZB21_E1_Mnu_Nm, iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role_Detail(ZB21_E1_Mnu_Type, iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role_Detail(ZB21_E1_Mnu_Action, iLngRow))
        iStrData = iStrData & Chr(11) & iLngMaxRow + iLngRow + 1
        iStrData = iStrData & Chr(11) & Chr(12)        
     Next
  
    
    Response.Write "<Script Language=vbscript>"                                    & vbCr
    Response.Write "With Parent "                                                & vbCr
    Response.Write "    .ggoSpread.SSShowData """ & ConvSPChars(iStrData)        & """" & vbCr
    Response.Write "    .lgCode        = """ & StrNextKey        & """" & vbCr
    Response.Write "    .txtNm.value        = """ & ConvSPChars(E1_Z_Usr_Role_Detail(ZB21_E1_Role_Nm, iLngRow-1))    & """" & vbCr
    Response.Write "    .frm1.vspdData.focus            "                                & vbCr
    Response.Write "    .DbQueryOk()             "                                & vbCr
    Response.Write "End With"                                                    & vbCr
    Response.Write "</Script>"                                                    & vbCr

%>
