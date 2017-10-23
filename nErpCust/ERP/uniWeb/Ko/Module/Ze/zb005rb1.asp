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

    Dim iZB006
    Dim istrCode
    Dim E1_Z_Usr_Role    
    Dim iLngRow
    Dim lgStrNextKey        
    Dim iLngMaxRow            
    Dim iStrData
    Dim istrMode
    Dim lgQueryFlag
    
    Const ZB06_I1_USR_ROLE_ID = 0
    Const ZB06_I1_USR_COMPST_ROLE_ID = 1
    
    Const C_SHEETMAXROWS_D = 100 
    
    Redim istrCode (ZB06_I1_USR_COMPST_ROLE_ID)
    
    if Request("NextCd") = "" then
        istrCode(ZB06_I1_USR_ROLE_ID) = FilterVar(Request("txtCd"),"","SNM")
    else        
        istrCode(ZB06_I1_USR_ROLE_ID) = FilterVar(Request("NextCd"),"","SNM")
    End if        
       
    istrCode(ZB06_I1_USR_COMPST_ROLE_ID) = FilterVar(Request("txtCompstCd"),"","SNM")
       
    Set iZB006 = Server.CreateObject("PZBG006.cListUsrRole")
	
    If CheckSYSTEMError(Err,True) = True Then
       Set iZB006 = Nothing                                                         
       Response.End 
    End If
      
    istrMode="R"
    E1_Z_Usr_Role = iZB006.ZB_LIST_USR_ROLE (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode,istrMode)


    If CheckSYSTEMError(Err,True) = True Then
       Set iZB006 = Nothing                                    
       Response.End 
    End If

    If IsEmpty(E1_Z_Usr_Role) Then
        Set iZB006 = Nothing    
       Response.End 
    End If
    
    Set iZB006 = nothing 
    
    iLngMaxRow =CLng(Request("txtMaxRows"))    
    
    For iLngRow = 0 To  UBound(E1_Z_Usr_Role,2)         
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else        
            lgStrNextKey = ConvSPChars(E1_Z_Usr_Role(0,iLngRow))
            lgQueryFlag = "0"        
            Exit For
        End if                 
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role(0, iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Usr_Role(1, iLngRow))        
        iStrData = iStrData & Chr(11) & iLngMaxRow + iLngRow + 1                
        iStrData = iStrData & Chr(11) & Chr(12)        
     Next
     
     Response.Write "<Script Language=vbscript>"                                & vbCr
     Response.Write "With Parent "                                                & vbCr     
     Response.Write "    .ggoSpread.SSShowData """ & iStrData                    & """" & vbCr     
     Response.Write "    If .txtCd.value <> " & """""  Then  " & vbCr
     Response.Write "    .txtNm.value =""" & ConvSPChars(E1_Z_Usr_Role(1, 0))    & """" & vbCr
     Response.Write "   End If "    & vbCr
     Response.Write "    .lgCode =""" & lgStrNextKey                                & """" & vbCr
     Response.Write "    .lgQueryFlag = """ & lgQueryFlag                        & """" & vbCr     
     Response.Write "	 .DbQueryOk"												& vbCr
     Response.Write "End With "                                                    & vbCr
     Response.Write "</Script>"    
     

%>
