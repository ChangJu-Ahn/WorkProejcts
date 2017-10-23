
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

Err.Clear
On Error Resume Next

Call LoadBasisGlobalInf()
'Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")
DIM lgWhere, lgOrder
DIM lgCode
DIM lgName

Call HideStatusWnd

lgCode = trim(Request("txtCode"))
lgName = trim(Request("txtName"))

if lgCode = "" then
	if lgName <> "" then
		lgWhere = " AND Dept_NM >= '" & filtervar(lgName,"","SNM") &"'"
		lgOrder = " Order by dept_NM "
	end if
else
	lgWhere = " AND A.Dept_Cd >= '" & filtervar(lgCode,"","SNM") &"'"
	lgOrder = " Order by A.dept_CD "
end if
If Trim(Request("txtIWhere")) = "2"  and Request("txtBizAreaCd") <> "" Then
	lgWhere = lgWhere & " AND B.BIZ_AREA_CD = '" & filtervar(Request("txtBizAreaCd"),"","SNM") &"'"
End If

lgstrSQL = "SELECT A.DEPT_CD, A.DEPT_NM, A.INTERNAL_CD , A.ORG_CHANGE_ID,B.BIZ_AREA_CD FROM B_ACCT_DEPT A, B_COST_CENTER B "
lgstrSQL = lgstrSQL & " where A.COST_CD = B.COST_CD "
lgstrSQL = lgstrSQL & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT "
lgstrSQL = lgstrSQL & "                        WHERE ORG_CHANGE_DT <= '" & UNIConvDate(Request("txtDate")) & " ') "
lgstrSQL = lgstrSQL & lgWhere
lgstrSQL = lgstrSQL & lgOrder

Call SubOpenDB(lgObjConn)
	
if 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
	If lgPrevNext = "" Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
	End If
else

	Do While Not lgObjRs.EOF
	            
		lgstrData = lgstrData & Chr(11) & lgObjRs("dept_cd")
		lgstrData = lgstrData & Chr(11) & lgObjRs("dept_nm")            
		lgstrData = lgstrData & Chr(11) & lgObjRs("biz_area_cd")
		lgstrData = lgstrData & Chr(11) & lgObjRs("org_change_id")
		lgstrData = lgstrData & Chr(11) & lgObjRs("internal_cd")
	    
'------ Developer Coding part (End   ) ------------------------------------------------------------------
		lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
		lgstrData = lgstrData & Chr(11) & Chr(12)

		lgObjRs.MoveNext
	        
	Loop 
	
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
		.vspdData.focus

		If .vspdData.MaxRows = 0 Then
		'	parent.UNIMsgBox "검색된 Data가 없습니다", 48, parent.top.document.title
		End If

	End With

</Script>
