f<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
Select Case CStr(lgOpModeCRUD)
   Case CStr(UID_M0001)   
        Call SubBizQuery()
   Case CStr(UID_M0002)                                                         '☜: Save,Update
        'Call SubBizSaveMulti()
End Select
    
Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

Sub SubBizQuery()
    On Error Resume Next                                                         '☜: Protect system from crashing
    Err.Clear                                                                    '☜: Clear Error status

    Dim iDx
    Dim iLoopMax
    Dim lgStrSQL
    Dim iSelCount
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'Call SVRMSGBOX(StrI & "-" & StrN ,0,1)
    iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1       
    
    lgStrSQL = "               SELECT TOP " & iSelCount  & " A.APP_REQ_NO, "
    lgStrSQL = lgStrSQL & "           A.APP_SEQ_NO, "
    lgStrSQL = lgStrSQL & "           A.APP_GBN, "
    lgStrSQL = lgStrSQL & "           A.APP_DT, "
    lgStrSQL = lgStrSQL & "           A.APP_PERSON, "
    lgStrSQL = lgStrSQL & "           A.APP_PERSON_NM, "
    lgStrSQL = lgStrSQL & "           A.APP_GRADE, "
    lgStrSQL = lgStrSQL & "           A.APP_GRADE_NM, "
    lgStrSQL = lgStrSQL & "           A.APP_DESC "
    lgStrSQL = lgStrSQL & "      FROM ( SELECT A.INTERNAL_CD, "
    lgStrSQL = lgStrSQL & "                    APP_REQ_NO   = A.REQ_NO, "
    lgStrSQL = lgStrSQL & "                    APP_SEQ_NO   = A.SEQ_NO, "
    lgStrSQL = lgStrSQL & "                    APP_SORT     = 1, "
    lgStrSQL = lgStrSQL & "                    APP_GBN      = '접수',  "
    lgStrSQL = lgStrSQL & "                    APP_DT       = A.R_DT,  "
    lgStrSQL = lgStrSQL & "                    APP_GRADE    = A.R_GRADE, " 
    lgStrSQL = lgStrSQL & "                    APP_GRADE_NM = B.MINOR_NM, "
    lgStrSQL = lgStrSQL & "                    APP_DESC     = A.R_DESC, "
    lgStrSQL = lgStrSQL & "                    APP_PERSON   = A.R_PERSON, "
    lgStrSQL = lgStrSQL & "                    APP_PERSON_NM= C.USR_NM "
    lgStrSQL = lgStrSQL & "               FROM B_CIS_NEW_ITEM_REQ_APP A(NOLOCK), "
    lgStrSQL = lgStrSQL & "                    B_MINOR                B(NOLOCK), "
    lgStrSQL = lgStrSQL & "                    Z_USR_MAST_REC         C(NOLOCK) "
    lgStrSQL = lgStrSQL & "              WHERE A.R_GRADE *=  B.MINOR_CD "
    lgStrSQL = lgStrSQL & "                AND B.MAJOR_CD= 'Y1007'  "
    lgStrSQL = lgStrSQL & "                AND C.USR_ID  =* A.R_PERSON "
    lgStrSQL = lgStrSQL & "              UNION "
    lgStrSQL = lgStrSQL & "             SELECT A.INTERNAL_CD, A.REQ_NO, A.SEQ_NO, 2, '기술',A.T_DT, A.T_GRADE, B.MINOR_NM, A.T_DESC, A.T_PERSON, C.USR_NM "
    lgStrSQL = lgStrSQL & "              FROM B_CIS_NEW_ITEM_REQ_APP A(NOLOCK), "
    lgStrSQL = lgStrSQL & "                   B_MINOR                B(NOLOCK), "
    lgStrSQL = lgStrSQL & "                   Z_USR_MAST_REC         C(NOLOCK) "
    lgStrSQL = lgStrSQL & "             WHERE A.T_GRADE *=  B.MINOR_CD "
    lgStrSQL = lgStrSQL & "               AND B.MAJOR_CD= 'Y1008'  "
    lgStrSQL = lgStrSQL & "               AND C.USR_ID  =* A.T_PERSON "
    lgStrSQL = lgStrSQL & "             UNION "
    lgStrSQL = lgStrSQL & "            SELECT A.INTERNAL_CD, A.REQ_NO, A.SEQ_NO, 3, '구매',A.P_DT, A.P_GRADE, B.MINOR_NM, A.P_DESC, A.P_PERSON, C.USR_NM "
    lgStrSQL = lgStrSQL & "              FROM B_CIS_NEW_ITEM_REQ_APP A(NOLOCK), "
    lgStrSQL = lgStrSQL & "                   B_MINOR                B(NOLOCK), "
    lgStrSQL = lgStrSQL & "                   Z_USR_MAST_REC         C(NOLOCK) "
    lgStrSQL = lgStrSQL & "             WHERE A.P_GRADE *=  B.MINOR_CD "
    lgStrSQL = lgStrSQL & "               AND B.MAJOR_CD= 'Y1008'  "
    lgStrSQL = lgStrSQL & "               AND C.USR_ID  =* A.P_PERSON " 
    lgStrSQL = lgStrSQL & "             UNION "
    lgStrSQL = lgStrSQL & "            SELECT A.INTERNAL_CD, A.REQ_NO, A.SEQ_NO, 4, '품질',A.Q_DT, A.Q_GRADE, B.MINOR_NM, A.Q_DESC, A.Q_PERSON, C.USR_NM "
    lgStrSQL = lgStrSQL & "              FROM B_CIS_NEW_ITEM_REQ_APP A(NOLOCK), "
    lgStrSQL = lgStrSQL & "                   B_MINOR                B(NOLOCK), "
    lgStrSQL = lgStrSQL & "                   Z_USR_MAST_REC         C(NOLOCK) "
    lgStrSQL = lgStrSQL & "             WHERE A.Q_GRADE *=  B.MINOR_CD "
    lgStrSQL = lgStrSQL & "               AND B.MAJOR_CD= 'Y1008'  "
    lgStrSQL = lgStrSQL & "               AND C.USR_ID  =* A.Q_PERSON "
    lgStrSQL = lgStrSQL & "            ) A "
    lgStrSQL = lgStrSQL & "       WHERE A.INTERNAL_CD = " & FilterVar(lgKeyStream(0),"","S") 
    lgStrSQL = lgStrSQL & "       ORDER BY A.APP_REQ_NO , A.APP_SEQ_NO, A.APP_SORT"
			
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

       Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

       lgstrData = ""        
       iDx       = 1        
       Do While Not lgObjRs.EOF
 
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APP_REQ_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APP_SEQ_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APP_GBN"))
          lgstrData = lgstrData & Chr(11) & UniConvDateDbToCompany(lgObjRs("APP_DT"),"")
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APP_PERSON"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APP_PERSON_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APP_GRADE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APP_GRADE_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("APP_DESC"))
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)
       
		  lgObjRs.MoveNext

          iDx =  iDx + 1
          
          If iDx > lgMaxCount Then
             lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
             Exit Do
          End If   
               
       Loop 
       
    End If

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs) 

End Sub    


'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                '.DBQueryOk        
	         End with
          End If       
    End Select    
       
</Script>	