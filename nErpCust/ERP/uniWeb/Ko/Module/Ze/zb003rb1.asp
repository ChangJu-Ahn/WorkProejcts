<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%    
    
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                               

    Call LoadBasisGlobalInf()                
    '------ Developer Coding part (Start ) ------------------------------------------------------------------

    '------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Dim iZB14
    
    Dim istrCode 
    Dim istrNm
    
    Dim lgstrPrvNext 
    Dim lgStrNextkey    
    
    Dim lgQueryFlag
    
    Dim lgstrPrvKey
    Dim iLngMaxRow        
    Dim iStrData
    Dim iLngRow
    Dim lgstrMenu
    Dim iMaxRow
        
    Dim E1_ZB_LIST_USR_ROLE_MNU_AUTH
    
    lgQueryFlag = "1"
    
    Const C_SHEETMAXROWS_D = 30
    
    Const ZB14_I1_Mnu_Id = 0
    Const ZB14_I1_Mnu_Type = 1
    Const ZB14_I1_Action_Id = 2
    Const ZB14_I1_Role_Id = 3        
    Const ZB14_I1_Mnu_Nm = 4
    
    Redim istrCode(ZB14_I1_Mnu_Nm)
    
    istrCode(ZB14_I1_Mnu_Id) = FilterVar(Request("txtCd"),"","SNM")
    istrCode(ZB14_I1_Mnu_Type) = Request("MnuType")
    istrCode(ZB14_I1_Action_Id) = Request("ActType")    
    istrCode(ZB14_I1_Role_Id) = FilterVar(Request("txtRoleID"),"","SNM")
    istrCode(ZB14_I1_Mnu_Nm) = FilterVar(Request("txtNm"),"","SNM")

    if Request("NextCd") <> "" then
        istrCode(ZB14_I1_Mnu_Id) = FilterVar(Request("NextCd"),"","SNM")
    End if

    Set iZB14 = Server.CreateObject("PZBG014.cListMnuAuthztn")
        
    If CheckSYSTEMError(Err,True) = True Then
       Set iZB14 = Nothing                                                       
       Response.End 
    End If
        
    E1_ZB_LIST_USR_ROLE_MNU_AUTH= iZB14.ZB_LIST_MNU_AUTHZTN (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)
    
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZB14 = Nothing                                                       
       Response.End 
    End If
    
    If IsEmpty(E1_ZB_LIST_USR_ROLE_MNU_AUTH) Then
        Set iZB14 = Nothing                                                       
        Response.End 
    End If
    
    
    Set iZB14 = Nothing 
    
    iLngMaxRow =CLng(Request("txtMaxRows"))
    
    iMaxRow=UBound(E1_ZB_LIST_USR_ROLE_MNU_AUTH,2)

    For iLngRow = 0 To iMaxRow
        If iLngRow < C_SHEETMAXROWS_D Then 
        
        Else        
            lgStrNextKey=ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,iLngRow))
            'lgStrNextActId=ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(3,iLngRow))
            lgQueryFlag = "0"
            Exit For
        End if          
        
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(0,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(1,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(2,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(3,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(4,iLngRow))
        iStrData = iStrData & Chr(11) & ConvSPChars(E1_ZB_LIST_USR_ROLE_MNU_AUTH(5,iLngRow))        
        iStrData = iStrData & Chr(11) & iLngMaxRow + iLngRow + 1        
        iStrData = iStrData & Chr(11) & Chr(12)
    Next
        
    Response.Write "<Script Language=vbscript>"                            & vbCr
    Response.Write "With Parent "                                        & vbCr    
    Response.Write ".ggoSpread.SSShowData """ & iStrData                & """" & vbCr
    Response.Write ".lgCode =""" & lgStrNextKey                    & """" & vbCr
    Response.Write ".lgQueryFlag = """ & lgQueryFlag            & """" & vbCr
    Response.Write ".DbQueryOk"									& vbCr
    'Response.Write ".lgAct =""" & lgStrNextActId                & """" & vbCr            
    Response.Write "End With "                                            & vbCr
    Response.Write "</Script>"

%>


    
