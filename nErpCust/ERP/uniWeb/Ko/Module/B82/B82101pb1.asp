f<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../B81/B81COMM.ASP" -->

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
    Err.Clear                                                                    '☜: Clear Error status

    Dim iDx
    Dim iLoopMax
    Dim lgStrSQL
    Dim iSelCount
    
    '---------- set code name  --------------------------------------------------------------------
      	call GetNameChk("MINOR_NM","B_MINOR","MINOR_CD="& FilterVar(lgKeyStream(5),"","S")& " AND MAJOR_CD=" & filterVar("Y1006","''","S") ,	lgKeyStream(5),"txtreq_user","의뢰자","N") '의뢰자 
        call GetNameChk("minor_nm","b_minor","major_cd='Y1001' and minor_cd="&FilterVar(lgKeyStream(4),"","S")&"",lgKeyStream(4),"txtItemKind","","N") '품목구분 
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
       
    iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1

    lgStrSQL = "               SELECT TOP " & iSelCount  & " A.REQ_NO,"
    lgStrSQL = lgStrSQL & "           A.REQ_ID,"
    lgStrSQL = lgStrSQL & "           REQ_NM = dbo.ufn_GetCodeName('Y1006' , A.REQ_ID ),"
    lgStrSQL = lgStrSQL & "           A.REQ_DT,"
    lgStrSQL = lgStrSQL & "           STATUS = (CASE A.STATUS WHEN 'R' THEN '의뢰' WHEN 'A' THEN '접수' WHEN 'D' THEN '반려' WHEN 'E' THEN '완료' WHEN 'S' THEN '중단' WHEN 'T' THEN '이관' END ),"
    lgStrSQL = lgStrSQL & "           A.ITEM_KIND, "
	lgStrSQL = lgStrSQL & "           ITEM_KIND_NM = dbo.ufn_GetCodeName('Y1001' , A.ITEM_KIND ), "
    lgStrSQL = lgStrSQL & "           A.ITEM_CD,"
    lgStrSQL = lgStrSQL & "           A.ITEM_NM,"
    lgStrSQL = lgStrSQL & "           A.ITEM_SPEC,"
    lgStrSQL = lgStrSQL & "           R_GRADE = dbo.ufn_GetCodeName('Y1007' , A.R_GRADE ),"
    lgStrSQL = lgStrSQL & "           T_GRADE = dbo.ufn_GetCodeName('Y1008' , A.T_GRADE ),"
    lgStrSQL = lgStrSQL & "           P_GRADE = dbo.ufn_GetCodeName('Y1008' , A.P_GRADE ),"
    lgStrSQL = lgStrSQL & "           Q_GRADE = dbo.ufn_GetCodeName('Y1008' , A.Q_GRADE ),"
    lgStrSQL = lgStrSQL & "           A.TRANS_DT, "
    lgStrSQL = lgStrSQL & "           A.REMARK "
    lgStrSQL = lgStrSQL & "      FROM B_CIS_NEW_ITEM_REQ A "
    lgStrSQL = lgStrSQL & "     WHERE A.REQ_DT >= " & FilterVar(UNIConvDate(lgKeyStream(0)),"","S")
    lgStrSQL = lgStrSQL & "       AND A.REQ_DT <= " & FilterVar(UNIConvDate(lgKeyStream(1)),"","S")
    
    If lgKeyStream(2) = "2" Then
       '진행       
       lgStrSQL = lgStrSQL & "       AND A.STATUS IN ('R','A','D') "
    ElseIf lgKeyStream(2) = "3" Then
       '완료 
       lgStrSQL = lgStrSQL & "       AND A.STATUS IN ('E','S','T') "
    End If
    
    If Trim(lgKeyStream(3)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & FilterVar(lgKeyStream(3),"","S")
	End If
    If Trim(lgKeyStream(4)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_KIND = " & FilterVar(lgKeyStream(4),"","S")
	End If
	If Trim(lgKeyStream(5)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.REQ_ID = " & FilterVar(lgKeyStream(5),"","S")
	End If	
	If Trim(lgKeyStream(6)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_SPEC LIKE " & FilterVar(lgKeyStream(6) & "%","","S")
	End If
			
	lgStrSQL = lgStrSQL & " ORDER BY A.REQ_NO ASC "
		
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       
        Call SetErrorStatus()
    Else

       Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

       lgstrData = ""        
       iDx       = 1        
       Do While Not lgObjRs.EOF
 
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REQ_NO"))
          'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REQ_ID"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REQ_NM"))
          lgstrData = lgstrData & Chr(11) & UniConvDateDbToCompany(lgObjRs("REQ_DT"),"")
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("STATUS"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_KIND"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_KIND_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_SPEC"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("R_GRADE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("T_GRADE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("P_GRADE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Q_GRADE"))
          lgstrData = lgstrData & Chr(11) & UniConvDateDbToCompany(lgObjRs("TRANS_DT"),"")
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
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

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