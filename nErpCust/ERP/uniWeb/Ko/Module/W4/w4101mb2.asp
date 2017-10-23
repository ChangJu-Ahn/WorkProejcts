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

    Call HideStatusWnd                                                               '��: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    'On Error Resume Next
    Err.Clear
   
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
	Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.
	Const TYPE_3	= 2		
		
	Dim C_SEQ_NO	
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14
	Dim C_W15
	Dim C_W16

	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '��: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' �׸��� ��ġ �ʱ�ȭ �Լ� 

    Call SubOpenDB(lgObjConn) 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)
    
    Response.End 
'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()	' ����Ÿ �Ѱ��ִ� �÷� ���� 
	C_SEQ_NO	= 1	' -- 1�� �׸��� 
    C_W10		= 2	' �������� 
    C_W11		= 3 ' �ݾ� 
    C_W12		= 4	' �����󰢴���� 
    C_W13		= 5	' �󰢺��δ���� 
    C_W14		= 6	' ������ 
    C_W15		= 7	' ���޼��񰡾� 
    C_W16		= 8	' �������񰡾� 
End Sub


'========================================================================================
Sub SubBizQuery()
	Dim iKey1, iKey2, iKey3
    Dim iDx, arrRs(2)
    Dim iRow, sW2, sW3, sW10, sW11, sW12, sW12_REF
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' ������� 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' �Ű��� 

	' 2�� �׸��� 
	Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
	     lgStrPrevKey = ""
	    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '�� : No data is found.
	    Call SetErrorStatus()
		    
	Else
	   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
	    lgstrData = "" : iLngRow = 1 : iRow = TYPE_2
		
		iRow = TYPE_1  
		arrRs(iRow) = ""
				
		Do While Not lgObjRs.EOF
			arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow	'  C_SESQ_NO
			arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))	' W10
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("DEBIT_SUM_AMT")	'W11
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("CREDIT_SUM_AMT")	'W12
			arrRs(iRow) = arrRs(iRow) & Chr(11) & "0"	' W13
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("DEBIT_SUM_AMT") - lgObjRs("CREDIT_SUM_AMT")	'W14
			arrRs(iRow) = arrRs(iRow) & Chr(11) & "0"	'W15
			arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("DEBIT_SUM_AMT") - lgObjRs("CREDIT_SUM_AMT")	'W16
			arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow
			arrRs(iRow) = arrRs(iRow) & Chr(11) & Chr(12)

			iLngRow = iLngRow + 1
			lgObjRs.MoveNext
		Loop 

	End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write "" & vbCrLf
	Response.Write "" & vbCrLf
	
	Response.Write " With parent.lgvspdData(" & TYPE_1 & ")" & vbCr   

	Response.Write "	parent.ggoSpread.Source = parent.lgvspdData(" & TYPE_1 & ")" & vbCr
	
	
	Response.Write "	parent.ggoSpread.SSShowData """ & arrRs(TYPE_1)       & """" & vbCr
		
	Response.Write "	parent.lgCurrGrid = " & TYPE_1 & vbCrLf
	Response.Write " End With"  & vbCrLf

    Response.Write " </Script>                                          " & vbCr

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)

    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.ACCT_NM, A.DEBIT_SUM_AMT, A.CREDIT_SUM_AMT "
            lgStrSQL = lgStrSQL & " FROM TB_WORK_2 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.ACCT_CD IN ( " & vbCrLf
			lgStrSQL = lgStrSQL & "			SELECT ACCT_CD FROM TB_ACCT_MATCH WITH (NOLOCK)  " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE MATCH_CD='22' ) " & vbCrLf
			'lgStrSQL = lgStrSQL & "	ORDER BY A.SEQ_NO ASC " & vbCrLf
			
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