
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../inc/incSvrMain.asp" -->
<!-- #Include file="../inc/incSvrDate.inc" -->
<!-- #Include file="../inc/incSvrNumber.inc" -->

<!-- #Include file="../inc/lgsvrvariables.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
<% 

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")


DIM lgWhere, lgOrder
DIM strDeptcd
DIM StrDeptNm
Dim strFromDt, strToDt
Dim lgPageNo
Dim  iLoopCount

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

Call HideStatusWnd

strDeptcd	= trim(Request("txtCode"))
StrDeptNm	= trim(Request("txtName"))
strFromDt	= UNIConvDate(TRim(Request("txtFromDate")))
strToDt		= UNIConvDate(Trim(Request("txtToDate"))) 
lgMaxCount	= 30
lgPageNo	= UNICInt(Trim(Request("lgPageNo")),0)

' 권한관리 추가 
lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
lgInternalCd		= Trim(Request("lgInternalCd"))
lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
if strDeptcd = "" then
	if StrDeptNm <> "" then
		lgWhere = " AND Dept_NM >= " & FilterVar(StrDeptNm, "''", "S") 
	end if
else
	lgWhere = " AND Dept_Cd >= " & FilterVar(strDeptcd, "''", "S") 
end if


' 권한관리 추가 
If lgAuthBizAreaCd <> "" Then			
	lgBizAreaAuthSQL		= " AND COST_CD IN (select COST_CD from b_cost_center(nolock) where BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")	& " ) "
End If			

If lgInternalCd <> "" Then
	lgInternalCdAuthSQL		= " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
End If
	
If lgSubInternalCd <> "" Then
	lgSubInternalCdAuthSQL	= " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
End If
	
'If lgAuthUsrID <> "" Then
'	lgAuthUsrIDAuthSQL		= " AND INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
'End If
	
	
lgstrSQL = "SELECT DEPT_CD, DEPT_NM, ORG_CHANGE_ID, INTERNAL_CD "
lgstrSQL = lgstrSQL & "FROM B_ACCT_DEPT(nolock) "
lgstrSQL = lgstrSQL & " WHERE ORG_CHANGE_DT >=ISNULL( "
lgstrSQL = lgstrSQL & " (select max(org_change_dt) from B_ACCT_DEPT(nolock) where org_change_dt<='" & strFromDt & "'), ("
lgstrSQL = lgstrSQL & " select min(org_change_dt) from B_ACCT_DEPT(nolock))) "
lgstrSQL = lgstrSQL & " AND ORG_CHANGE_DT <= " 
lgstrSQL = lgstrSQL & " (select max(org_change_dt) from B_ACCT_DEPT(nolock) where org_change_dt<='" & strToDt & "') "

' 권한관리 추가 
lgstrSQL = lgstrSQL & lgWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

lgstrSQL = lgstrSQL & " Order by DEPT_CD asc, ORG_CHANGE_ID desc "

'Response.Write lgstrSQL

Call SubOpenDB(lgObjConn)
	
if 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
	If lgPrevNext = "" Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
	End If
else

    lgstrData      = ""
    
    If CDbl(lgPageNo) > 0 Then
       lgObjRs.Move     = CDbl(lgMaxCount) * CDbl(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1

	Do While Not lgObjRs.EOF
	    iLoopCount =  iLoopCount + 1
	    
        If  iLoopCount < lgMaxCount Then
			lgstrData = lgstrData & Chr(11) & lgObjRs("dept_cd")
			lgstrData = lgstrData & Chr(11) & lgObjRs("dept_nm")  
			lgstrData = lgstrData & Chr(11) & lgObjRs("ORG_CHANGE_ID")           
			lgstrData = lgstrData & Chr(11) & lgObjRs("internal_cd")
		    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow 
			lgstrData = lgstrData & Chr(11) & Chr(12)
        
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If

		lgObjRs.MoveNext
	        
	Loop 
	
	If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

end if
     
Call SubCloseDB(lgObjConn)    
    
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       'Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       'Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub
%>		    

<Script Language="vbscript">   
		
	With parent
	    .ggoSpread.Source = parent.vspdData
		.ggoSpread.SSShowData "<%=lgstrData%>"
		.lgPageNo      =  "<%=lgPageNo%>" 
		.vspdData.focus

		If .vspdData.MaxRows = 0 Then
		'	parent.UNIMsgBox "검색된 Data가 없습니다", 48, parent.top.document.title
		End If

	End With

</Script>
