<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7105b1
'*  4. Program Name         : 고정자산 부서별배분율등록 
'*  5. Program Desc         : 고정자산 부서별배분율을 등록,수정,삭제,조회 
'*  6. Comproxy List        : +As0061ManageSvr
'                             +As0068ListSvr
'*  7. Modified date(First) : 2000/09/19
'*  8. Modified date(Last)  : 2001/05/31
'*  9. Modifier (First)     : hersheys
'* 10. Modifier (Last)      : Kim Hee Jung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    

    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")                                                        '☜: Hide Processing message
'Dim lgCurrency, lgStrPrevKey_i, lgBlnFlgChgValue, plgStrPrevKey_i

	Dim lgOpModeCRUD
'	Dim lgPageNo, lgStrPrevKey

    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizCreate()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizCancel()
    End Select

    'Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizCreate()

    Dim iPAAG030
    'Dim import_String
    Dim import_Group
	Dim strYear, strMonth, strDay, stryyyymm, stDt
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 권한관리추가 
	Const A519_I2_a_data_auth_data_BizAreaCd = 0
	Const A519_I2_a_data_auth_data_internal_cd = 1
	Const A519_I2_a_data_auth_data_sub_internal_cd = 2
	Const A519_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A519_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
    ReDim import_Group(0)        
    
    stryyyymm = Trim(Request("txtYyyymm"))
	stDt = UniConvYYYYMMDDToDate(gDateFormat, Mid(stryyyymm,1,4), Mid(stryyyymm,6,2), "01")
    Call ExtractDateFrom(stDt, gDateFormat, gComDateType, strYear, strMonth, strDay)
    
    import_Group(0)	= strYear & Right("0" & strMonth, 2)
       
    Set iPAAG030 = Server.CreateObject("PAAG030.cAMngAsDptHistorySvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    Call iPAAG030.AS0061_MANAGE_ASSET_DEPT_HISTORY_SVR(gStrGloBalCollection, import_Group, "", I2_a_data_auth, "2")

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG030 = Nothing
       response.end
       Exit Sub
    End If    
    
    Set iPAAG030 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.fnButtonExecOk      " & vbCr
    Response.Write "</Script>                   " & vbCr    

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizCancel()

    Dim iPAAG030
    'Dim import_String
    Dim import_Group
	Dim strYear, strMonth, strDay, stryyyymm, stDt
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 권한관리추가 
	Const A519_I2_a_data_auth_data_BizAreaCd = 0
	Const A519_I2_a_data_auth_data_internal_cd = 1
	Const A519_I2_a_data_auth_data_sub_internal_cd = 2
	Const A519_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A519_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
    ReDim import_Group(0)        
    
    stryyyymm = Trim(Request("txtYyyymm"))
	stDt = UniConvYYYYMMDDToDate(gDateFormat, Mid(stryyyymm,1,4), Mid(stryyyymm,6,2), "01")
    Call ExtractDateFrom(stDt, gDateFormat, gComDateType, strYear, strMonth, strDay)
    
    import_Group(0)	= strYear & Right("0" & strMonth, 2)
       
    Set iPAAG030 = Server.CreateObject("PAAG030.cAMngAsDptHistorySvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    Call iPAAG030.AS0061_MANAGE_ASSET_DEPT_HISTORY_SVR(gStrGloBalCollection, import_Group, "", I2_a_data_auth, "3")

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG030 = Nothing
       response.end
       Exit Sub
    End If    
    
    Set iPAAG030 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.fnButtonExecOk      " & vbCr
    Response.Write "</Script>                   " & vbCr    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode)
    On Error Resume Next
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	
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
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
'    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>

<Script Language="VBScript">
	parent.fnButtonExecOk																		'☜: 화면 처리 ASP 를 지칭함 
</Script>	
<%					

	Response.End
%>
























