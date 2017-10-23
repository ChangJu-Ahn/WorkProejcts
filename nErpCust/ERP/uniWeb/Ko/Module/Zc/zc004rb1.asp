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

    Dim iZC020
    Dim istrCode
    Dim E1_Z_Usr_Role    
    Dim iLngMaxRow    
    Dim iLngRow
    Dim lgStrNextKey        
    Dim iStrData
    Dim istrMode
    Dim lgstrMenu
    Dim lgQueryFlag
    
    Const ZC20_I1_Mnu_Id = 0
    Const ZC20_I1_Lang_Cd = 1
    
    Const ZC20_E1_Mnu_Id = 0
    Const ZC20_E1_Mnu_Nm = 1
    Const ZC20_E1_Role_Id = 2    
    Const ZC20_E1_Action_Id = 3        
    
    
    Const C_SHEETMAXROWS_D = 100 
    
    
    Redim istrCode (ZC20_I1_Lang_Cd)
    
    if Request("NextCd") = "" then
        istrCode(ZC20_I1_Mnu_Id) = FilterVar(Request("txtCd"),"","SNM")
    else        
        istrCode(ZC20_I1_Mnu_Id) = FilterVar(Request("NextCd"))
    End if        
    
    istrCode (ZC20_I1_Lang_Cd) = gLang
    
    Set iZC020 = Server.CreateObject("PZCG020.cListMnuPerRole")
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC020 = Nothing                                                         
       Response.End 
    End If
    
    E1_Z_Usr_Role = iZC020.ZC_LIST_MNU_PER_ROLE (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)


    If CheckSYSTEMError(Err,True) = True Then
        Response.Write "<Script Language=vbscript>"            & vbCr
        Response.Write "    Parent.txtCd.focus() "            & vbCr                
        Response.Write "    Parent.txtCd.select() "            & vbCr                        
        Response.Write "    Parent.txtNm.value = """""    &vbCr        
        Response.Write "</Script>"    
       Set iZC020 = Nothing                                    
       Response.End 
    End If

    If IsEmpty(E1_Z_Usr_Role) Then
        Set iZC020 = Nothing    
        Response.End 
    End If
    Set iZC020 = nothing 
    
    Response.Write "<Script Language=vbscript>"                                                & vbCr
    Response.Write "With Parent "                                                            & vbCr     
    Response.Write "    .txtCd.value = """ & Trim(ConvSPChars(E1_Z_Usr_Role(ZC20_E1_Mnu_Id, 0)))    & """" & vbCr
    Response.Write "    .txtNm.value = """ & Trim(ConvSPChars(E1_Z_Usr_Role(ZC20_E1_Mnu_Nm, 0)))    & """" & vbCr
    Response.Write "End With "                                                                & vbCr
    Response.Write "</Script>"    

    iLngMaxRow = CLng(Request("txtMaxRows"))
         
    For iLngRow = 0 To  UBound(E1_Z_Usr_Role,2)         
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else        
            lgStrNextKey = ConvSPChars(E1_Z_Usr_Role(0,iLngRow))
            lgQueryFlag = "0"        
            Exit For
        End if                 
        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role(ZC20_E1_Role_Id, iLngRow))
        
        If ConvSPChars(E1_Z_Usr_Role(ZC20_E1_Action_Id, iLngRow))  = "A" Then 
            lgstrMenu = "All"
        ElseIf ConvSPChars(E1_Z_Usr_Role(ZC20_E1_Action_Id, iLngRow))  = "E" Then 
            lgstrmenu = "Excel / Print "
        ElseIf ConvSPChars(E1_Z_Usr_Role(ZC20_E1_Action_Id, iLngRow))  = "Q" Then 
            lgstrMenu="Query"
        Else
            lgstrMenu = "None"
        End if 
        
        iStrData = iStrData & Chr(11) & lgstrMenu  
        iStrData = iStrData & Chr(11) & iLngMaxRow + iLngRow + 1        
        iStrData = iStrData & Chr(11) & Chr(12)        
     Next
     
     Response.Write "<Script Language=vbscript>"                        & vbCr
     Response.Write "With Parent "                                        & vbCr     
     Response.Write "    .ggoSpread.SSShowData """ & iStrData            & """" & vbCr     
     Response.Write "    .lgCode = """ & lgStrNextKey                        & """" & vbCr
     Response.Write "    .lgQueryFlag = """ & lgQueryFlag                    & """" & vbCr     
     Response.Write "End With "                                            & vbCr
     Response.Write "</Script>"    
     

%>
