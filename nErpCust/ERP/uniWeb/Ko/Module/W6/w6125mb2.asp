<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<%Option Explicit%> 
<% session.CodePage=949 %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
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
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8
	Dim C_W9
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14

	Dim C_W15
	Dim C_W15_NM
	Dim C_W16
	Dim C_W16_P
	Dim C_W17
	Dim C_W18
	Dim C_W19
	Dim C_W20
	Dim C_W21
	Dim C_W22
	Dim C_W23
	Dim C_W24
	Dim C_W25
	Dim C_W26
	Dim C_W27
	Dim C_W28
	Dim C_W29
	Dim C_W30
	Dim C_W31
	Dim C_W32
	Dim C_W33
	Dim C_W34
	Dim C_W35
	Dim C_W36
	Dim C_W37
	Dim C_W37_P
	
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
	
	C_W1		= 0
	C_W2		= 1
	C_W3		= 2
	C_W4		= 3
	C_W5		= 4
	C_W6		= 5
	C_W7		= 6
	C_W8		= 7
	C_W9		= 8
	C_W10		= 9
	C_W11		= 10
	C_W12		= 11
	C_W13		= 12
	C_W14		= 13
		
	C_W15		= 1
	C_W15_NM	= 2
	C_W16		= 3
	C_W16_P		= 4
	C_W17		= 5
	C_W18		= 6
	C_W19		= 7
	C_W20		= 8
	C_W21		= 9
	C_W22		= 10
	C_W23		= 11	
	C_W24		= 12
	C_W25		= 13
	C_W26		= 14
	C_W27		= 15
	C_W28		= 16
	C_W29		= 17
	C_W30		= 18
	C_W31		= 19
	C_W32		= 20
	C_W33		= 21
	C_W34		= 22
	C_W35		= 23
	C_W36		= 24
	C_W37		= 25
	C_W37_P		= 26
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
		Response.Write "	.frm1.txtData(" & C_W2 & ").value = """ & lgObjRs("REPRE_NM") & """" & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.
		Response.Write "	.frm1.txtData(" & C_W3 & ").value = """ & lgObjRs("CO_ADDR") & """" & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.
		Response.Write "	.frm1.txtData(" & C_W4 & ").value = """ & lgObjRs("OWN_RGST_NO") & """" & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.
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
				iStrData = iStrData & " .Col = " & C_W18 & " : .value = """ & ConvSPChars(lgObjRs("W33")) & """" & vbCrLf
				iStrData = iStrData & " .Col = " & C_W19 & " : .value = """ & ConvSPChars(lgObjRs("W34")) & """" & vbCrLf		
				iStrData = iStrData & " .Col = " & C_W20 & " : .value = """ & ConvSPChars(lgObjRs("W35")) & """" & vbCrLf	
				iStrData = iStrData & " .Col = " & C_W21 & " : .value = """ & ConvSPChars(lgObjRs("W37")) & """" & vbCrLf	
				iStrData = iStrData & " .Col = " & C_W22 & " : .value = """ & ConvSPChars(lgObjRs("W36")) & """" & vbCrLf	
				iStrData = iStrData & " .Col = " & C_W24 & " : .value = """ & ConvSPChars(lgObjRs("W45")) & """" & vbCrLf		
				iStrData = iStrData & " .Col = " & C_W30 & " : .value = """ & ConvSPChars(lgObjRs("W45")) & """" & vbCrLf		
				iStrData = iStrData & " .Col = " & C_W35 & " : .value = """ & ConvSPChars(lgObjRs("W50")) & """" & vbCrLf		
				iStrData = iStrData & " .Col = " & C_W37 & " : .value = """ & ConvSPChars(lgObjRs("W32")) & """" & vbCrLf		

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
			Response.Write " Call parent.SetSpreadLock                                  " & vbCr
			Response.Write " End With	                        " & vbCr
			Response.Write " </Script>	                        " & vbCr
		
		End If
	End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "RH"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.CO_NM, A.REPRE_NM, A.CO_ADDR, A.OWN_RGST_NO "
           lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "RD"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "   A.W33, A.W34, A.W35, A.W37, A.W36, A.W45, A.W50, A.W32 "
            lgStrSQL = lgStrSQL & " FROM TB_54AD A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W31 > '000002'" & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W45 > 0" & vbCrLf

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
