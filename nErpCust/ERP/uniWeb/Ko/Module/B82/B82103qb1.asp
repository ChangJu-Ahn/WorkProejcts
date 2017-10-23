<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         :  
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment
'* 12. Common Coding Guide  : 
'* 13. History              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim IntRetCD
Dim PvArr
Dim NextKey1
Dim strNextKey1

Const C_SHEETMAXROWS_D = 100

lgLngMaxRow     = Request("txtMaxRows") 
lgErrorStatus   = "NO"

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)
Call SubBizQuery()
Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)                                              '☜ : DB-Agent를 통한 ADO query
     
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
	Dim iDx
	
	'On Error Resume Next           
    Err.Clear
    
	Call SubMakeSQLStatements      
    
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then 
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()
		Call SubCloseRs(lgObjRs)
		Response.End 
	Else
		IntRetCD = 1

        iDx = 0
        ReDim PvArr(C_SHEETMAXROWS_D)
        
        Do While Not lgObjRs.EOF
 
            If iDx = C_SHEETMAXROWS_D Then
               NextKey1 = ConvSPChars(lgObjRs(0))
               Exit Do
            End If   
	    
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _						
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & ConvSPChars(lgObjRs(7)) & _
						Chr(11) & ConvSPChars(lgObjRs(8)) & _
						Chr(11) & ConvSPChars(lgObjRs(9)) & _
						Chr(11) & ConvSPChars(lgObjRs(10)) & _
						Chr(11) & ConvSPChars(lgObjRs(11)) & _
						Chr(11) & ConvSPChars(lgObjRs(12)) & _
						Chr(11) & ConvSPChars(lgObjRs(13)) & _						
						Chr(11) & UNIDateClientFormat(lgObjRs(14)) & _
						Chr(11) & ConvSPChars(lgObjRs(15)) & _
						Chr(11) & ConvSPChars(lgObjRs(16)) & _
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)
			
			PvArr(iDx) = lgstrData
			iDx = iDx + 1
		    lgObjRs.MoveNext
        Loop 
    End If

	lgstrData = Join(PvArr, "")

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                             
		
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements()

    On Error Resume Next
    Err.Clear           
	
	lgStrSQL = "               SELECT TOP " & C_SHEETMAXROWS_D + 1
	lgStrSQL = lgStrSQL & "           A.REQ_NO,"
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
    lgStrSQL = lgStrSQL & "           A.DOC_NO, "
    lgStrSQL = lgStrSQL & "           A.REMARK "
    lgStrSQL = lgStrSQL & "      FROM B_CIS_NEW_ITEM_REQ A "
    lgStrSQL = lgStrSQL & "     WHERE A.REQ_DT >= " & FilterVar(uniConvDate(Request("txtDtFr")),"","S")
    lgStrSQL = lgStrSQL & "       AND A.REQ_DT <= " & FilterVar(uniConvDate(Request("txtDtTo")),"","S")
        
    If Request("txtrdoStatus") = "2" Then
       '진행       
       lgStrSQL = lgStrSQL & "       AND A.STATUS IN ('R','A','D') "
    ElseIf Request("txtrdoStatus") = "3" Then
       '완료 
       lgStrSQL = lgStrSQL & "       AND A.STATUS IN ('E','S','T') "
    End If
    If Trim(Request("cboItemAcct")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & FilterVar(Request("cboItemAcct"),"","S")
	End If
    If Trim(Request("txtItem_Kind")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_KIND = " & FilterVar(Request("txtItem_Kind"),"","S")
	End If
	If Trim(Request("txtreq_user")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.REQ_ID = " & FilterVar(Request("txtreq_user"),"","S")
	End If	
	If Trim(Request("txtItemSpec")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_SPEC LIKE " & FilterVar(Request("txtItemSpec") & "%","","S")
	End If
	If Request("lgStrPrevKey") <> "" Then
		lgStrSQL = lgStrSQL & " AND A.REQ_NO  >= " & FilterVar(Request("lgStrPrevKey"),"","S")
	End If
		
	lgStrSQL = lgStrSQL & " ORDER BY A.REQ_NO ASC "

End Sub  

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"       
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                 
    Err.Clear                            

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
        Case "MB"
			ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub


Response.Write "<Script language=vbs> " & vbCr         
Response.Write " With Parent "      	& vbCr
Response.Write "	If """ & lgErrorStatus & """ = ""NO"" And """ & IntRetCd & """ <> -1 Then "	& vbCr
Response.Write "    .lgStrPrevKey  = """ & NextKey1 & """" & vbCr  
Response.Write "	.ggoSpread.Source	= .frm1.vspdData "				& vbCr
Response.Write "	.ggoSpread.SSShowDataByClip  """ & lgstrData  & """"        & vbCr
Response.Write "		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
Response.Write "			.DbQuery						"				& vbCr
Response.Write "		Else								"				& vbCr
Response.Write "			.DbQueryOK						"				& vbCr
Response.Write "		End If								"				& vbCr
Response.Write "		.frm1.vspdData.focus				"				& vbCr
Response.Write "    End If								"				& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 
Response.End     
												'☜: 비지니스 로직 처리를 종료함 
%>
