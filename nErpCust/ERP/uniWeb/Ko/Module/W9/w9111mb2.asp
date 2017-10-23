<%@ Language=VBScript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    'On Error Resume Next
    Err.Clear
   
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Const TYPE_1 = 0
	Const TYPE_2 = 1

	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W10

	Dim C_SEQ_NO
	Dim C_SEQ_NO_VIEW
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8
	Dim C_W9
	
	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn) 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_W1		= 0	' 콘트롤배열 순서(HTML기준)
	C_W2		= 1
	C_W3		= 2
	C_W4		= 3
	C_W10		= 4
	
	C_SEQ_NO	= 1	' 그리드배열 
	C_SEQ_NO_VIEW = 2
	C_W5		= 3
	C_W6		= 4
	C_W7		= 5
	C_W8		= 6
	C_W9		= 7
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngRow
    Dim iRow, iKey1, iKey2, iKey3
	Dim arrRow(2), iType, iStrData, iLngCol
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("RH",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
		iLngCol = lgObjRs.Fields.Count
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	.IsRunEvents = True " & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.
		Response.Write "	.frm1.txtData(" & C_W1 & ").value = """ & lgObjRs("CO_NM") & """" & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.
		Response.Write "	.frm1.txtData(" & C_W2 & ").value = """ & lgObjRs("OWN_RGST_NO") & """" & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.
		Response.Write "	.frm1.txtData(" & C_W3 & ").value = """ & lgObjRs("REPRE_NM") & """" & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.
		Response.Write "	.IsRunEvents = False " & vbCrLf	' 이벤트가 발생하게 한다.
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr

		lgObjRs.Close
		Set lgObjRs = Nothing
			
		iStrData = ""
		'TYPE_2 조회 
	    Call SubMakeSQLStatements("RD",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
		     lgStrPrevKey = ""
		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    Call SetErrorStatus()
			    
		Else
			lgstrData = ""
				
			Do While Not lgObjRs.EOF
				iStrData = iStrData & " .Row = .Row + 1 " & vbCrLf
				iStrData = iStrData & " .Col = " & C_W5 & " : .text = """ & ConvSPChars(lgObjRs("W18")) & """" & vbCrLf
				iStrData = iStrData & " .Col = " & C_W6 & " : ." & GetValueText(lgObjRs("W19")) & " = """ & ConvSPChars(lgObjRs("W19")) & """" & vbCrLf		
				iStrData = iStrData & " .Col = " & C_W9 & " : .value = """ & ConvSPChars(lgObjRs("W29")) & """" & vbCrLf	

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 

			lgObjRs.Close
			Set lgObjRs = Nothing
			
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " Dim iRow " & vbCr
			Response.Write " With parent.lgvspdData(" & TYPE_2 & ")" & vbCr
			Response.Write "	parent.ggoSpread.Source = parent.lgvspdData(" & TYPE_2 & ")        " & vbCr
			Response.Write "	If .MaxRows = 0 Then " & vbCrLf
			Response.Write "		iRow = .MaxRows + 1 " & vbCrLf
			Response.Write "	Else " & vbCrLf
			Response.Write "		iRow = .MaxRows " & vbCrLf
			Response.Write "	End If " & vbCrLf
			Response.Write "	Call parent.FncInsertRow(" & iLngRow & ")" & vbCr
			Response.Write "	.Row = iRow " & vbCrlf
			Response.Write iStrData 
			'Response.Write " Call parent.SetTotalLine                                  " & vbCr
			Response.Write "	Call parent.SetSpreadLock                                  " & vbCr
			Response.Write "	Call parent.vspdData_Change(" & C_W9 & ", 1)" & vbCr
			Response.Write "	parent.lgBlnFlgChgValue = true	                        " & vbCr
			Response.Write " End With	                        " & vbCr
			Response.Write " </Script>	                        " & vbCr
		
		End If
	End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


End Sub

Function GetValueText(Byval pData)
	If Instr(1, pData, "-") > 0 Then
		GetValueText = "text"
	Else
		GetValueText = "value"
	End If
End Function
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "RH"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.CO_NM, A.REPRE_NM, A.OWN_RGST_NO "
           lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "RD"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "   A.W18, A.W19, A.W29 "
            lgStrSQL = lgStrSQL & " FROM TB_54D2 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W17_1 = '3'" & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W17 IN ('01', '04', '05')" & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W29 > 0" & vbCrLf

			' -- 홈텍스: 주식등변동상황명세서_주식수변동상황(85A131)의 내용구분이 '3', 주주구분이 '1', '4', '5'이고, 감소_양도에 값이 있으면 입력합니다.
    End Select

	PrintLog "SubMakeSQLStatements = " & lgStrSQL
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
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"
End Sub

'========================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    'On Error Resume Next
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

       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             'Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>
<%
'   **************************************************************
'	1.4 Transaction 처러 이벤트 
'   **************************************************************

Sub	onTransactionCommit()
	' 트랜잭션 완료후 이벤트 처리 
End Sub

Sub onTransactionAbort()
	' 트랜잭선 실패(에러)후 이벤트 처리 
'PrintForm
'	' 에러 출력 
	'Call SaveErrorLog(Err)	' 에러로그를 남긴 
	
End Sub
%>
