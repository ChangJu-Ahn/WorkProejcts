<%@ Transaction=required Language=VBScript%>
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
   
	Const BIZ_MNU_ID = "W3133MA1"
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
	Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.
			
	Dim C_SEQ_NO
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8
	Dim C_W_TYPE
	Dim C_W_TYPE_NM

	Dim C_W9
	Dim C_W9_1
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14

	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '��: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' �׸��� ��ġ �ʱ�ȭ �Լ� 

    Call SubOpenDB(lgObjConn) 
    
    Call CheckVersion(sFISC_YEAR, sREP_TYPE)	' 2005-03-11 ����������� �߰� 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
        Case CStr(UID_M0004)                                                         '��: Delete
             Call SubBizDelete2()     
    End Select

    Call SubCloseDB(lgObjConn)
'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()	' ����Ÿ �Ѱ��ִ� �÷� ���� 
	C_SEQ_NO	= 1
	
	C_W1		= 2	' ���� 
	C_W2		= 3	' ȸ������ 
	C_W3		= 4	' �����̿��� 
	C_W4		= 5	' ������Ծ� 
	C_W5		= 6	' �� 
	C_W6		= 7	' 
	C_W7		= 8
	C_W8		= 9	' 
	C_W_TYPE	= 10
	C_W_TYPE_NM	= 11
	
	C_W9		= 2 ' 
	C_W9_1		= 3
	C_W10		= 4 ' �����ͱ� 
	C_W11		= 5 ' ȸ����ͱ� 
	C_W12		= 6 ' �������� 
	C_W13		= 7 ' �������� 
	C_W14		= 8 ' �����ݾ� 
End Sub

'========================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' �����Ϻ��� �����Ѵ�.
    lgStrSQL =            "	DELETE TB_40AD WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & "	WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	lgStrSQL = lgStrSQL & "	DELETE TB_40AH WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & "	WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
		
	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
 	' -- 15ȣ ���� 
 	Call TB_15_DeleData()
End Sub

Sub SubBizDelete2()
    'On Error Resume Next
    Err.Clear

	' �����Ϻ��� �����Ѵ�.
    lgStrSQL =            "	DELETE TB_40AD WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & "	WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	lgStrSQL = lgStrSQL & "	DELETE TB_40AH WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & "	WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
		
	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
 	' -- 15ȣ ���� 
 	Call TB_15_DeleData()
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngRow
    Dim iRow, iKey1, iKey2, iKey3
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' ������� 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' �Ű��� 

    Call SubMakeSQLStatements("RH",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '�� : No data is found.
        Call SetErrorStatus()
        
    Else
       ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = "" : iLngRow = 1
        
		Do While Not lgObjRs.EOF
			lgstrData = lgstrData & "	.Row = " & lgObjRs("SEQ_NO") & vbCrLf
			lgstrData = lgstrData & "	.Col = " & C_W10 & " : .value = " & lgObjRs("W10") & vbCrLf
			lgstrData = lgstrData & "	.Col = " & C_W11 & " : .value = " & lgObjRs("W11") & vbCrLf
			lgstrData = lgstrData & "	.Col = " & C_W12 & " : .value = " & lgObjRs("W12") & vbCrLf
			lgstrData = lgstrData & "	.Col = " & C_W13 & " : .value = " & lgObjRs("W13") & vbCrLf
			lgstrData = lgstrData & "	.Col = " & C_W14 & " : .value = " & lgObjRs("W14") & vbCrLf & vbCrLf

			iLngRow = iLngRow + 1
			lgObjRs.MoveNext
		Loop 

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent.lgvspdData(" & TYPE_2 & ")" & vbCr

		Response.Write "	parent.ggoSpread.Source = parent.lgvspdData(" & TYPE_2 & ")" & vbCr
		Response.Write lgstrData     & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr

		' 1�� �׸��� : ������ 
		Call SubMakeSQLStatements("RD",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

		   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = "" : iLngRow = 1
		    
			Do While Not lgObjRs.EOF
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W1"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W2"))	' -- ���� 
				lgstrData = lgstrData & Chr(11) & lgObjRs("W3")
				lgstrData = lgstrData & Chr(11) & lgObjRs("W4")
				lgstrData = lgstrData & Chr(11) & lgObjRs("W5")
				lgstrData = lgstrData & Chr(11) & lgObjRs("W6")
				lgstrData = lgstrData & Chr(11) & lgObjRs("W7")
				lgstrData = lgstrData & Chr(11) & lgObjRs("W8")
				lgstrData = lgstrData & Chr(11) & lgObjRs("W_TYPE")
				lgstrData = lgstrData & Chr(11) & lgObjRs("W_TYPE_NM")
				lgstrData = lgstrData & Chr(11) & iLngRow
				lgstrData = lgstrData & Chr(11) & Chr(12)

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent	" & vbCr

			Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_1 & ")" & vbCr
			Response.Write "	.ggoSpread.SSShowData """ &  lgstrData     & """" & vbCr
			Response.Write " End With                                           " & vbCr
			Response.Write " </Script>                                          " & vbCr
    
		End If

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write "	parent.DbQueryOk                                      " & vbCr
		Response.Write " </Script>  " & vbCr
		    
    End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


End Sub

Function PrtRate(Byval pRate)
	If pRate= 0 Then	
		PrtRate = ""
	Else
		PrtRate = pRate
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
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W10, A.W11, A.W12, A.W13, A.W14 "
            lgStrSQL = lgStrSQL & " FROM TB_40AH A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			
      Case "RD"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W1, A.W2, A.W3, A.W4, A.W5, A.W6, A.W7, A.W8, A.W_TYPE, dbo.ufn_GetCodeName('W1014', A.W_TYPE) W_TYPE_NM "
            lgStrSQL = lgStrSQL & " FROM TB_40AD A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
    End Select

	PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

    'On Error Resume Next
    Err.Clear 
 
	' 2�� �׸��� : ��� 
	PrintLog "txtSpread = " & Request("txtSpread" & CStr(TYPE_2))
		
	arrRowVal = Split(Request("txtSpread" & CStr(TYPE_2) ), gRowSep)                                 '��: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
	PrintLog "txtSpread = " & lgLngMaxRow
	For iDx = 1 To lgLngMaxRow

	    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
		    
	    Select Case arrColVal(0)
	        Case "C"
	                Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
	        Case "U"
	                Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
	    End Select
		    
	    If lgErrorStatus    = "YES" Then
	       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
	       Exit For
	    End If
		    
	Next
	   
	' 1�� �׸��� : ������ 
	PrintLog "txtSpread = " & Request("txtSpread" & CStr(TYPE_1))
		
	arrRowVal = Split(Request("txtSpread" & CStr(TYPE_1) ), gRowSep)                                 '��: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
	For iDx = 1 To lgLngMaxRow

	    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
		    
	    Select Case arrColVal(0)
	        Case "C"
	                Call SubBizSaveMultiCreate2(arrColVal)                            '��: Create
	        Case "U"
	                Call SubBizSaveMultiUpdate2(arrColVal)                            '��: Update
	        Case "D"
	                Call SubBizSaveMultiDelete2(arrColVal)                            '��: Delete
	    End Select
		    
	    If lgErrorStatus    = "YES" Then
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
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_40AH WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W10, W11, W12, W13, W14" & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")     & "," & vbCrLf

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate1 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	' -- 15�� Ǫ�� 
'����������		(�𣭨�)�� ����Ͽ� �Է��ϰ� �� �ݾ��� (+) ����� ��� 
'			15-1ȣ�� (1)���� "ȯ����������" (2)�ݾ׿� ���ݾ��� (3)ó�п��� "����(����)"�� �Է��ϰ� 
'			���������� "ȯ������������ �̻��ܾ��� ���� ȯ�Ծ��� �ͱݻ����ϰ� ����ó����"�� �Է��ϰ� 
'			(-)������ ���� 15-2ȣ�� (1)����"ȯ����������" (2)�ݾ��� ���ݾ��� ���밪�� (3)ó�п��� "����(����)"�� 
'			�Է��ϰ�, ���������� " ȯ������������ ��ȯ���ܾ��� ����ȯ�Ծ��� �ͱݺһ����ϰ� ����ó����"�� �Է��ϰ� �����.
	If arrColVal(C_SEQ_NO) = "1" Then
		PrintLog "C_W12=" & arrColVal(C_W12)
		If UNICDbl(arrColVal(C_W12), 0) > 0 Then
			Call TB_15_PushData("1", UNICDbl(arrColVal(C_W12), 0), 1, "4002", "400", "ȯ������������ �̻��ܾ��� ���� ȯ�Ծ��� �ͱݻ����ϰ� ����ó����")
		ElseIf UNICDbl(arrColVal(C_W12), 0) < 0 Then
			Call TB_15_PushData("2", ABS(UNICDbl(arrColVal(C_W12), 0)), 1, "4002", "100", "ȯ������������ ��ȯ���ܾ��� ����ȯ�Ծ��� �ͱݺһ����ϰ� ����ó����")
		End If
	End If
	
'����������		(�� - ��)�� ����Ͽ� �Է��ϰ� �� �ݾ��� (-) ������ ��� 
'			15-2ȣ�� (1)���� "ȯ����������" (2)�ݾ׿� ���ݾ��� (3)ó�п��� "����(����)"�� �Է��ϰ� 
'			���������� "ȯ���������� ��ȯ���ܾ��� ���� �󰢾��� �ձݻ����ϰ� ����ó����"�� �Է��ϰ� 
'			(+)����� ���� 15-1ȣ�� (1)����"ȯ����������" (2)�ݾ��� ���ݾ��� ���밪�� (3)ó�п��� "����(����)"�� 
'			�Է��ϰ�, ���������� " ȯ���������� �̻��ܾ��ǰ��ٻ󰢾��� �ձݺһ����ϰ� ����ó����"�� �Է��ϰ� �����.
	
	If arrColVal(C_SEQ_NO) = "2" Then
		PrintLog "C_W13=" & arrColVal(C_W13)
		If UNICDbl(arrColVal(C_W13), 0) < 0 Then
			Call TB_15_PushData("2", ABS(UNICDbl(arrColVal(C_W13), 0)), 2, "4002", "100", "�׽�Ʈ ȯ���������� ��ȯ���ܾ��� ���� �󰢾��� �ձݻ����ϰ� ����ó����")
		ElseIf UNICDbl(arrColVal(C_W12), 0) > 0 Then
			Call TB_15_PushData("1", UNICDbl(arrColVal(C_W13), 0), 2, "4002", "400", "�׽�Ʈ2 ȯ���������� �̻��ܾ��ǰ��ٻ󰢾��� �ձݺһ����ϰ� ����ó����")
		End If
	End If
'�����ݾ�		(�񣭨�)�� ����Ͽ� �Է��ϰ�, (+)����� ��� 
'			15-2ȣ�� (1)���� "��ȭȯ�����" (2)�ݾ��� ���ݾ��� (3)�ҵ�ó���� "����(����)"�� �Է��ϰ�,
'			���������� " ��ȭȯ������� �ձݻ����ϰ� ����ó����"�� �Է��ϰ� 
'			(-)������ ��� 15-1ȣ�� (1)���� "��ȭȯ�����" (2)�ݾ��� ���ݾ��� ���밪�� (3)�ҵ�ó���� "����(����)"�� 
'			���������� " ��ȭȯ������� �ͱݻ����ϰ� ����ó����"�� �Է��ϰ� �����.
	If arrColVal(C_SEQ_NO) = "3" Then
		PrintLog "C_W14=" & arrColVal(C_W14)
		If UNICDbl(arrColVal(C_W14), 0) > 0 Then
			Call TB_15_PushData("2", UNICDbl(arrColVal(C_W14), 0), 3, "4001", "100", "��ȭȯ������� �ձݻ����ϰ� ����ó����")
		ElseIf UNICDbl(arrColVal(C_W12), 0) < 0 Then
			Call TB_15_PushData("1", ABS(UNICDbl(arrColVal(C_W14), 0)), 3, "4001", "400", "��ȭȯ������� �ͱݻ����ϰ� ����ó����")
		End If
	End If
			
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	If Trim(UCase(arrColVal(C_W2))) = "" Then Exit Sub
	If UNICDbl(arrColVal(C_SEQ_NO), "0") >= 999999 Then	
		arrColVal(C_W1) = arrColVal(C_W2)
		arrColVal(C_W2) = ""
	End If
	
	lgStrSQL = "INSERT INTO TB_40AD WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W1, W2, W3, W4, W5, W6, W7, W8, W_TYPE" & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"NULL","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W6))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")     & "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "UPDATE  TB_40AH WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W10	    = " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & "," & vbCrLf
    'lgStrSQL = lgStrSQL & " W10		= " &  FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W11	    = " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W12     = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W13     = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W14     = " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

	' -- 15�� Ǫ�� 
'����������		(�𣭨�)�� ����Ͽ� �Է��ϰ� �� �ݾ��� (+) ����� ��� 
'			15-1ȣ�� (1)���� "ȯ����������" (2)�ݾ׿� ���ݾ��� (3)ó�п��� "����(����)"�� �Է��ϰ� 
'			���������� "ȯ������������ �̻��ܾ��� ���� ȯ�Ծ��� �ͱݻ����ϰ� ����ó����"�� �Է��ϰ� 
'			(-)������ ���� 15-2ȣ�� (1)����"ȯ����������" (2)�ݾ��� ���ݾ��� ���밪�� (3)ó�п��� "����(����)"�� 
'			�Է��ϰ�, ���������� " ȯ������������ ��ȯ���ܾ��� ����ȯ�Ծ��� �ͱݺһ����ϰ� ����ó����"�� �Է��ϰ� �����.
	If arrColVal(C_SEQ_NO) = "1" Then
		PrintLog "C_W12=" & arrColVal(C_W12)
		If UNICDbl(arrColVal(C_W12), 0) > 0 Then
			Call TB_15_PushData("1", UNICDbl(arrColVal(C_W12), 0), 1, "4002", "400", "ȯ������������ �̻��ܾ��� ���� ȯ�Ծ��� �ͱݻ����ϰ� ����ó����")
		ElseIf UNICDbl(arrColVal(C_W12), 0) < 0 Then
			Call TB_15_PushData("2", ABS(UNICDbl(arrColVal(C_W12), 0)), 1, "4002", "100", "ȯ������������ ��ȯ���ܾ��� ����ȯ�Ծ��� �ͱݺһ����ϰ� ����ó����")
		End If
	End If
	
'����������		(�� - ��)�� ����Ͽ� �Է��ϰ� �� �ݾ��� (-) ������ ��� 
'			15-2ȣ�� (1)���� "ȯ����������" (2)�ݾ׿� ���ݾ��� (3)ó�п��� "����(����)"�� �Է��ϰ� 
'			���������� "ȯ���������� ��ȯ���ܾ��� ���� �󰢾��� �ձݻ����ϰ� ����ó����"�� �Է��ϰ� 
'			(+)����� ���� 15-1ȣ�� (1)����"ȯ����������" (2)�ݾ��� ���ݾ��� ���밪�� (3)ó�п��� "����(����)"�� 
'			�Է��ϰ�, ���������� " ȯ���������� �̻��ܾ��ǰ��ٻ󰢾��� �ձݺһ����ϰ� ����ó����"�� �Է��ϰ� �����.
	If arrColVal(C_SEQ_NO) = "2" Then
		PrintLog "C_W13=" & arrColVal(C_W13)
		If UNICDbl(arrColVal(C_W13), 0) < 0 Then
			Call TB_15_PushData("2", ABS(UNICDbl(arrColVal(C_W13), 0)), 2, "4002", "100", "�׽�Ʈ ȯ���������� ��ȯ���ܾ��� ���� �󰢾��� �ձݻ����ϰ� ����ó����")
		ElseIf UNICDbl(arrColVal(C_W13), 0) > 0 Then
			Call TB_15_PushData("1", UNICDbl(arrColVal(C_W13), 0), 2, "4002", "400", "�׽�Ʈ2 ȯ���������� �̻��ܾ��ǰ��ٻ󰢾��� �ձݺһ����ϰ� ����ó����")
		End If
	End If
'�����ݾ�		(�񣭨�)�� ����Ͽ� �Է��ϰ�, (+)����� ��� 
'			15-2ȣ�� (1)���� "��ȭȯ�����" (2)�ݾ��� ���ݾ��� (3)�ҵ�ó���� "����(����)"�� �Է��ϰ�,
'			���������� " ��ȭȯ������� �ձݻ����ϰ� ����ó����"�� �Է��ϰ� 
'			(-)������ ��� 15-1ȣ�� (1)���� "��ȭȯ�����" (2)�ݾ��� ���ݾ��� ���밪�� (3)�ҵ�ó���� "����(����)"�� 
'			���������� " ��ȭȯ������� �ͱݻ����ϰ� ����ó����"�� �Է��ϰ� �����.
	If arrColVal(C_SEQ_NO) = "3" Then
		PrintLog "C_W14=" & arrColVal(C_W14) 
		If UNICDbl(arrColVal(C_W14), 0) > 0 Then
			Call TB_15_PushData("2", UNICDbl(arrColVal(C_W14), 0), 3, "4001", "100", "��ȭȯ������� �ձݻ����ϰ� ����ó����")
		ElseIf UNICDbl(arrColVal(C_W14), 0) < 0 Then
			Call TB_15_PushData("1", ABS(UNICDbl(arrColVal(C_W14), 0)), 3, "4001", "400", "��ȭȯ������� �ͱݻ����ϰ� ����ó����")
		End If
	End If
End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '��: Clear Error status

	If UNICDbl(arrColVal(C_SEQ_NO), "0") >= 999999 Then	
		arrColVal(C_W1) = arrColVal(C_W2)
		arrColVal(C_W2) = ""
	End If
		
    lgStrSQL = "UPDATE  TB_40AD WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2	    = " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"NULL","S")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6		= " &  FilterVar(Trim(UCase(arrColVal(C_W6))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W8		= " &  FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W_TYPE	= " &  FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  TB_40AD WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 
	
	PrintLog "SubBizSaveMultiDelete1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : 15ȣ���Ŀ� Ǫ�� 
' Desc :  
'============================================================================================================
Sub TB_15_PushData(Byval pType, Byval pAmt, Byval pSeqNo, Byval pAcctCd, Byval pCode, Byval pDesc)
	On Error Resume Next 
	Err.Clear  

	lgStrSQL = "EXEC usp_TB_15_PushData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' ���� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' ������� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' �Ű��� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' �������� ���α׷� 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "				' �������� ���� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pType)),"''","S") & ", "			' ��/�� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pAcctCd)),"''","S") & ", "		' ���� �ڵ� 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pAmt, "0"),"0","D")  & ", "			' �ݾ� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pCode)),"''","S") & ", "			' ó�� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pDesc)),"''","S") & ", "			' �������� 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "
	
		
	PrintLog "TB_15_PushData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

Sub TB_15_DeleData()
	On Error Resume Next 
	Err.Clear  

	lgStrSQL = "EXEC usp_TB_15_DeleData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' ���� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' ������� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' �Ű��� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' �������� ���α׷� 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(-1, "0"),"0","D") & ", "		' �������� ���� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1ȣ/2ȣ 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "


	PrintLog "TB_15_DeleData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
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
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>
<%
'   **************************************************************
'	1.4 Transaction ó�� �̺�Ʈ 
'   **************************************************************

Sub	onTransactionCommit()
	' Ʈ����� �Ϸ��� �̺�Ʈ ó�� 
End Sub

Sub onTransactionAbort()
	' Ʈ���輱 ����(����)�� �̺�Ʈ ó�� 
'PrintForm
'	' ���� ��� 
	'Call SaveErrorLog(Err)	' �����α׸� ���� 
	
End Sub
%>
