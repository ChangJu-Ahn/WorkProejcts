f<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
On Error Resume Next                                                             '¢Ð: Protect system from crashing
Err.Clear                                                                        '¢Ð: Clear Error status

Call HideStatusWnd                                                               '¢Ð: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '¢Ð: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
lgMaxCount        = CInt(Request("lgMaxCount"))                                  '¢Ð: Fetch count at a time for VspdData
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
Select Case CStr(lgOpModeCRUD)
   Case CStr(UID_M0001)   
        Call SubBizQuery()
   Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
        'Call SubBizSaveMulti()
End Select
    
Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

Sub SubBizQuery()
    On Error Resume Next                                                         '¢Ð: Protect system from crashing
    Err.Clear                                                                    '¢Ð: Clear Error status

    Dim iDx
    Dim iLoopMax
    Dim lgStrSQL
    Dim iSelCount
      
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
       
    iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
    
    lgStrSQL = "               SELECT TOP " & iSelCount  & " A.ITEM_LVL1_CD, "
    lgStrSQL = lgStrSQL & "           dbo.ufn_s_CIS_GetParentNn(A.ITEM_ACCT,A.ITEM_KIND,A.ITEM_LVL1_CD,'*','L1') ITEM_LVL1_NM, "
    lgStrSQL = lgStrSQL & "           A.ITEM_LVL2_CD,              "
    lgStrSQL = lgStrSQL & "           dbo.ufn_s_CIS_GetParentNn(A.ITEM_ACCT,A.ITEM_KIND,A.ITEM_LVL2_CD,A.ITEM_LVL1_CD,'L2')  ITEM_LVL2_NM ,"
    lgStrSQL = lgStrSQL & "           A.ITEM_LVL3_CD,              "
    lgStrSQL = lgStrSQL & "           dbo.ufn_s_CIS_GetParentNn(A.ITEM_ACCT,A.ITEM_KIND,A.ITEM_LVL3_CD,A.ITEM_LVL2_CD,'L3') ITEM_LVL3_NM, "
    lgStrSQL = lgStrSQL & "           A.SPEC_ORDER,                "
    lgStrSQL = lgStrSQL & "           A.SPEC_NAME,                 "
    lgStrSQL = lgStrSQL & "           Spec_Value = '',   "
    lgStrSQL = lgStrSQL & "           Spec_Split = '',   "
    lgStrSQL = lgStrSQL & "           A.SPEC_UNIT,                 "
    lgStrSQL = lgStrSQL & "           A.SPEC_LENGTH,               "
    lgStrSQL = lgStrSQL & "           A.SPEC_EXAMPLE,              "
    lgStrSQL = lgStrSQL & "           A.REMARK                     "
    lgStrSQL = lgStrSQL & "      FROM B_CIS_ITEM_CLASS_CATEGORY A(NOLOCK) "

    lgStrSQL = lgStrSQL & "     WHERE 1=1"
	
	If Trim(lgKeyStream(0)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & FilterVar(lgKeyStream(0),"","S")
	End If
	If Trim(lgKeyStream(1)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_KIND = " & FilterVar(lgKeyStream(1),"","S")
	End If
	If Trim(lgKeyStream(2)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_LVL1_CD = " & FilterVar(lgKeyStream(2),"","S")
	End If
	If Trim(lgKeyStream(3)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_LVL2_CD = " & FilterVar(lgKeyStream(3),"","S")
	End If
	If Trim(lgKeyStream(4)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_LVL3_CD = " & FilterVar(lgKeyStream(4),"","S")
	End If
	
	lgStrSQL = lgStrSQL & " ORDER BY A.SPEC_ORDER ASC "
	
		
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

       Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

       lgstrData = ""        
       iDx       = 1        
       Do While Not lgObjRs.EOF
       
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL1_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL1_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL2_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL2_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL3_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL3_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC_ORDER"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC_NAME"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Spec_Value"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Spec_Split"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC_UNIT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC_LENGTH"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC_EXAMPLE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))		  
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
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
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