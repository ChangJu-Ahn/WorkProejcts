<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!--
======================================================================================================
*  1. Module Name		: 
*  2. Function Name		: Multi Sample
*  3. Program ID		: VB000MB1.ASP
*  4. Program Name		: 
*  5. Program Desc		: °æ¿µÁ¤º¸°èÈ¹ 
*  6. Comproxy List		:
*  7. Modified date(First)	: 2005/02/02
*  8. Modified date(Last)	: 
*  9. Modifier (First)	: Cho Ig Sung
* 10. Modifier (Last)	: 
* 11. Comment			: EIS
=======================================================================================================-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
Dim lgStrPrevKey
Const C_SHEETMAXROWS_D = 100
Call HideStatusWnd	'¢Ð: Hide Processing message
Call LoadBasisGlobalInf()  
Call LoadInfTB19029B("I", "A","NOCOOKIE","MB")

lgErrorStatus	= "NO"
lgErrorPos		= ""															'¢Ð: Set to space
lgOpModeCRUD	= Request("txtMode")											'¢Ð: Read Operation Mode (CRUD)
lgKeyStream		= Split(Request("txtKeyStream"),gColSep)

lgLngMaxRow		= Request("txtMaxRows")										'¢Ð: Read Operation Mode (CRUD)	
lgStrPrevKey	= UNICInt(Trim(Request("lgStrPrevKey")),0)				'¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
	
Call SubOpenDB(lgObjConn)			
	
Select Case lgOpModeCRUD
	Case CStr(UID_M0001)														'¢Ð: Query		
		Call SubBizQuery()
	Case CStr(UID_M0002)														'¢Ð: Save,Update
		Call SubBizSaveMulti()
	Case CStr(UID_M0003)														'¢Ð: Delete
		Call SubBizDelete()
End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next
	Err.Clear
	Call SubBizQueryMulti()
End Sub	
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
	On Error Resume Next
	Err.Clear			
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iDx
	Dim iLoopMax
	Dim iKey1
	Dim strWhere
	
	On Error Resume Next	
	Err.Clear
	
	strWhere = " WHERE YYYYMM BETWEEN " & FilterVar(Trim(lgKeyStream(0)),"''", "S") & " AND " & FilterVar(Trim(lgKeyStream(1)),"''", "S")
		
	strWhere = strWhere & " ORDER BY YYYYMM ASC, BS_PL_FLAG ASC "
	
	Call SubMakeSQLStatements("MR",strWhere,"X","X")							'¢Ð : Make sql statements

	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		lgStrPrevKey = ""
	
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'¢Ð : No data is found. 
		Call SetErrorStatus()
	Else
		Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		lgstrData = ""
		iDx		= 1
		Do While Not lgObjRs.EOF
		
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("YYYYMM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BS_PL_FLAG"))				
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUMMARY"))
			lgstrData = lgstrData & Chr(11) & uniNumClientFormat(lgObjRs("PLAN_AMT"), ggAmtOfMoney.DecPoint,0)
								
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
				lgObjRs.MoveNext
		
			iDx =  iDx + 1
			If iDx > C_SHEETMAXROWS_D Then
				lgStrPrevKey = lgStrPrevKey + 1
				Exit Do
			End If	
						
		Loop 
	End If

	If iDx <= C_SHEETMAXROWS_D Then
		lgStrPrevKey = ""
	End If	
				
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	Call SubCloseRs(lgObjRs)	

End Sub	

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
	Dim arrColVal
	Dim iDx

	On Error Resume Next
	Err.Clear 

	arrRowVal = Split(Request("txtSpread"), gRowSep)								'¢Ð: Split Row	data
	
	For iDx = 1 To lgLngMaxRow
		arrColVal = Split(arrRowVal(iDx-1), gColSep)								'¢Ð: Split Column data
		
		Select Case arrColVal(0)
			Case "C"
					Call SubBizSaveMultiCreate(arrColVal)							'¢Ð: Create
			Case "U"
					Call SubBizSaveMultiUpdate(arrColVal)							'¢Ð: Update
			Case "D"
					Call SubBizSaveMultiDelete(arrColVal)							'¢Ð: Delete
		End Select
		
		If lgErrorStatus	= "YES" Then
			lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
			Exit For
		End If
		
	Next
End Sub	

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next	
	Err.Clear

	lgStrSQL = "INSERT INTO VB000AT("
	lgStrSQL = lgStrSQL & " YYYYMM , BS_PL_FLAG , SUMMARY , PLAN_AMT ," 
	lgStrSQL = lgStrSQL & " INSRT_USER_ID , INSRT_DT , UPDT_USER_ID , UPDT_DT )" 
	lgStrSQL = lgStrSQL & " VALUES(" 
		
	lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(2)),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(4)),"''","S")			& ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5), 0)						& ","
		
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")						& "," 
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")				& "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")						& ","						
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")
	lgStrSQL = lgStrSQL & ")"
		
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

	On Error Resume Next 
	Err.Clear			'¢Ð: Clear Error status

	lgStrSQL = "UPDATE  VB000AT"
	lgStrSQL = lgStrSQL & " SET " 
	lgStrSQL = lgStrSQL & " SUMMARY			= " &  FilterVar(Trim(arrColVal(4)),"''","S")			& ","
	lgStrSQL = lgStrSQL & " PLAN_AMT		= " &  UNIConvNum(arrColVal(5), 0)						& ","
	lgStrSQL = lgStrSQL & " UPDT_USER_ID	= " &  FilterVar(gUsrId,"''","S")						& "," 
	lgStrSQL = lgStrSQL & " UPDT_DT			= " &  FilterVar(GetSvrDateTime,"''","S")
	lgStrSQL = lgStrSQL & " WHERE			"
	lgStrSQL = lgStrSQL & " YYYYMM			= " &  FilterVar(Trim(arrColVal(2)),"''","S")	& " AND "
	lgStrSQL = lgStrSQL & " BS_PL_FLAG		= " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
	On Error Resume Next	'¢Ð: Protect system from crashing
	Err.Clear			'¢Ð: Clear Error status

	lgStrSQL = "DELETE  VB000AT"
	lgStrSQL = lgStrSQL & " WHERE		"
	lgStrSQL = lgStrSQL & " YYYYMM			= " &  FilterVar(Trim(arrColVal(2)),"''","S")	& " AND "
	lgStrSQL = lgStrSQL & " BS_PL_FLAG		= " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)

	Dim iSelCount
		
	Select Case Mid(pDataType,1,1)
		Case "M"
			iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

		Select Case Mid(pDataType,2,1)
			Case "R"				
				lgStrSQL = "SELECT TOP " & iSelCount  
				lgStrSQL = lgStrSQL & "	a.YYYYMM, a.BS_PL_FLAG, isnull(b.MINOR_NM,'') as MINOR_NM, a.SUMMARY, a.PLAN_AMT "
				lgStrSQL = lgStrSQL & "  FROM VB000AT a(NOLOCK) left outer join B_MINOR b(NOLOCK) on a.BS_PL_FLAG = b.MINOR_CD AND b.MAJOR_CD = 'A1023' "
				lgStrSQL = lgStrSQL &  pCode 
			End Select			
	End Select
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	lgErrorStatus	= "YES"
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	lgErrorStatus	= "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	lgErrorStatus	= "YES"														'¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	On Error Resume Next	'¢Ð: Protect system from crashing
	Err.Clear			'¢Ð: Clear Error status

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
		Case "<%=UID_M0001%>"														'¢Ð : Query
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				With Parent
					.ggoSpread.Source	= .frm1.vspdData
					.lgStrPrevKey	= "<%=lgStrPrevKey%>"
					.ggoSpread.SSShowData "<%=lgstrData%>"				
					.DBQueryOk		
				End with		
				
			End If	
		Case "<%=UID_M0002%>"														'¢Ð : Save
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				Parent.DBSaveOk
			Else
				Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
			End If	
		Case "<%=UID_M0002%>"														'¢Ð : Delete
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				Parent.DbDeleteOk
			Else	
			End If	
	End Select		
		
</Script>	
