
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

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

DIM lgWhere, lgOrder
DIM lgCode
DIM lgName

Call HideStatusWnd

lgCode = trim(Request("txtCode"))
lgName = trim(Request("txtName"))

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
if lgCode = "" then
	if lgName <> "" then
		lgWhere = " AND Dept_NM >= " & FilterVar(lgName, "''", "S") 
		lgOrder = " Order by dept_NM "
	end if
else
	lgWhere = " AND Dept_Cd >= " & FilterVar(lgCode, "''", "S")
	lgOrder = " Order by dept_CD "
end if


	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
	lgBizAreaAuthSQL		= " AND COST_CD IN (select COST_CD from b_cost_center(nolock) where BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")	& " ) "
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL  = " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL  = " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If
	
	'If lgAuthUsrID <> "" Then
	'	lgAuthUsrIDAuthSQL  = " AND INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	'End If
	
lgstrSQL = "SELECT DEPT_CD, DEPT_NM, INTERNAL_CD FROM B_ACCT_DEPT "
lgstrSQL = lgstrSQL & " where ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT "
lgstrSQL = lgstrSQL & "                        WHERE ORG_CHANGE_DT <= '" & UNIConvDate(Request("txtDate")) & " ') "
lgstrSQL = lgstrSQL & lgWhere

' 권한관리 추가 
lgstrSQL = lgstrSQL & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

lgstrSQL = lgstrSQL & lgOrder


Call WriteToLog (lgstrSQL)

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



Sub WriteToLog(pLogData)

    On Error Resume Next
    Dim objFSO
    Dim objFile
    Dim pPath
    
    pPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & gLang & "\Log"

    pPath = pPath & "\ADOCommonPopup" & "[" & gCompany & "][" & gDBServer & "][" & gDatabase & "][" & gUsrID & "][" & Request.ServerVariables("REMOTE_ADDR") & "].txt"

    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFile = objFSO.OpenTextFile( pPath,8,True)
       
    objFile.WriteLine pLogData
   
    If Not (objFSO Is Nothing) Then
       Set objFSO = Nothing
    End If
    
    If Not (objFile Is Nothing) Then
       objFile.Close
       Set objFile = Nothing
    End If

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
