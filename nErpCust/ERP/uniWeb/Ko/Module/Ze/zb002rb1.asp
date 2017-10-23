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

    Dim iZB020
    Dim I1_Z_Used_Role
    Dim E1_Z_Usr_Role    
    Dim iLngRow
    Dim StrNextKey        
    Dim iLngMaxRow    
    Dim iStrData
    Dim istrMode
    
    Dim lgQueryFlag
    
    Dim E1_Z_Used_Role
    Const C_SHEETMAXROWS_D = 100 
    
    lgQueryFlag = "1"
    
    Redim I1_Z_Used_Role(1)
    
    I1_Z_Used_Role(0) = Request("txtCd")
    I1_Z_Used_Role(1) = Request("NextCd")
    
    Set iZB020 = Server.CreateObject("PZBG020.cListUsedRole")
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZB020 = Nothing                                                         
       Response.End 
    End If
        
    E1_Z_Used_Role = iZB020.ZB_LIST_USED_ROLE (gStrGlobalCollection, C_SHEETMAXROWS_D, I1_Z_Used_Role)


    If CheckSYSTEMError(Err,True) = True Then
       Set iZB020 = Nothing                                    
       Response.End 
    End If

    If IsEmpty(E1_Z_Used_Role) Then
        Set iZB020 = Nothing    
       Response.End 
    End If
    
    Set iZB020 = nothing 
    
    iLngMaxRow = CLng(Request("txtMaxRows"))        
    
    For iLngRow = 0 To  UBound(E1_Z_Used_Role,2)         
        If iLngRow < C_SHEETMAXROWS_D Then 
        Else        
            StrNextKey = ConvSPChars(E1_Z_Used_Role(0,iLngRow))
            lgQueryFlag = "0"
            Exit For
        End if                 
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Used_Role(0, iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Used_Role(1, iLngRow))
        iStrData = iStrData & Chr(11) & iLngMaxRow + iLngRow + 1        
        iStrData = iStrData & Chr(11) & Chr(12)        
     Next
    
     
     Response.Write "<Script Language=vbscript>"                        & vbCr
     Response.Write "With Parent "                                        & vbCr     
     Response.Write "    .ggoSpread.SSShowData """ & iStrData            & """" & vbCr  
     Response.Write "   .lgCode =""" & StrNextKey                    & """" & vbCr   
     Response.Write "    .lgQueryFlag = """ & lgQueryFlag            & """" & vbCr
     Response.Write "End With "                                            & vbCr
     Response.Write "</Script>"    
     

%>
