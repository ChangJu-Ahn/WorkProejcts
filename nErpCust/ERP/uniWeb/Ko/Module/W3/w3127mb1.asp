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
   
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
	Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.
	Const TYPE_3	= 2		
	Const TYPE_4A	= 3
	Const TYPE_4B	= 4
	Const TYPE_5	= 5
	Const TYPE_6	= 6
	
	Dim C_SEQ_NO
	Dim C_W_TYPE
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7

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
    End Select

    Call SubCloseDB(lgObjConn)
'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()	' ����Ÿ �Ѱ��ִ� �÷� ���� 
	C_SEQ_NO	= 1	' -- 1�� �׸��� 
    C_W_TYPE	= 2	' ���� 
    C_W1		= 3	' ������ 
    C_W2		= 4 ' ���� 
    C_W3		= 5	' ���� 
    C_W4		= 6	' �뺯 
    C_W5		= 7	' �ܾ� 
    C_W6		= 8	' �ϼ� 
    C_W7		= 9	' ���� 

	C_W10		= 5	' �ڻ��Ѱ� 
	C_W11		= 6	' ��ä�Ѱ� 
	C_W12		= 7	' �ڱ��ں� 
	C_W13		= 8 ' ���ϼ� 
	C_W14		= 9	' ���� 
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
    lgStrSQL =            "DELETE TB_26B WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(6), iIntMaxRows, iLngRow
    Dim iRow, iKey1, iKey2, iKey3
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' ������� 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' �Ű��� 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = False Then
  
         lgStrPrevKey = ""
       
        Response.Write "<script language=vbscript> " & vbcr
		Response.Write "	Call parent.FncNew() " & vbcr
        Response.Write "</script> " & vbcr
        
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '�� : No data is found.
        Call SetErrorStatus()
        
    Else
       ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = "" : iLngRow = 1
        
        For iRow = TYPE_1 To TYPE_6	' �����ڵ庰�� ����Ÿ�� ������ 
			arrRs(iRow) = ""
			lgObjRs.Filter = "W_TYPE='" & iRow & "'"	' �����ڵ�� ���͸���.
			
		    Do While Not lgObjRs.EOF
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W_TYPE"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & UNIDateClientFormat(lgObjRs("W1"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W2"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & UNINumClientFormat(lgObjRs("W3"), ggAmtOfMoney.DecPoint,"")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & UNINumClientFormat(lgObjRs("W4"), ggAmtOfMoney.DecPoint, "")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & UNINumClientFormat(lgObjRs("W5"), ggAmtOfMoney.DecPoint, "")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & UNINumClientFormat(lgObjRs("W6"), ggAmtOfMoney.DecPoint, "")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & UNINumClientFormat(lgObjRs("W7"), ggAmtOfMoney.DecPoint, 0)
				arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow
				arrRs(iRow) = arrRs(iRow) & Chr(11) & Chr(12)

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 
        
			iLngRow = 1
			lgObjRs.Filter = ""
		Next
		
    End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr

	If arrRs(TYPE_6) <> "" Then	
		Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_6 & ")" & vbCr
		Response.Write "	.ggoSpread.ClearSpreadData                     " & vbCr
	End If
	
	For iRow = TYPE_1 To TYPE_6	' �����ڵ庰�� ����Ÿ�� ������ 
	    
		Response.Write "	.ggoSpread.Source = .lgvspdData(" & iRow & ")" & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & arrRs(iRow)       & """" & vbCr
    'Response.Write "	.lgPageNo = """ & iIntQueryCount           & """" & vbCr
    'Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)          & """" & vbCr
    'Response.Write "	.frm1.hCtrlCd.value =	""" & ConvSPChars(E1_a_ctrl_item(i_a_ctrl_item_ctrl_cd))          & """" & vbCr
    'Response.Write "	.frm1.txtCtrlNM.value = """ & ConvSPChars(E1_a_ctrl_item(i_a_ctrl_item_ctrl_nm))          & """" & vbCr
	Next
	
    Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
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
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W_TYPE, A.W1, A.W2, "
            lgStrSQL = lgStrSQL & " case when A.W3=0 and a.seq_no=999999  then null else a.w3 end w3, "
            lgStrSQL = lgStrSQL & " case when A.W4=0 and a.seq_no=999999  then null else a.w4 end w4, "
            lgStrSQL = lgStrSQL & " case when A.W5=0 and a.seq_no=999999  then null else a.w5 end w5, "
            lgStrSQL = lgStrSQL & " case when A.W6=0 and a.seq_no=999999  then null else a.w6 end w6, A.W7 "
            lgStrSQL = lgStrSQL & " FROM TB_26B A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.W_TYPE ASC, A.W1 " 'A.SEQ_NO" & vbcrlf
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

	PrintLog "txtSpread = " & Request("txtSpread")
	
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
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
	'On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_26B WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W_TYPE, SEQ_NO  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , W1, W2, W3, W4, W5, W6, W7" & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"1","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W1), ""),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim((arrColVal(C_W2))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D")     & "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL
	
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

    lgStrSQL = "UPDATE  TB_26B WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1	   = " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W1), ""),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2     = " &  FilterVar(Trim((arrColVal(C_W2))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3     = " &  FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4     = " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5     = " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6     = " &  FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W7     = " &  FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W_TYPE = " & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  TB_26B WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W_TYPE = " & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  & vbCrLf 
	
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
	Call SaveErrorLog(Err)	' �����α׸� ���� 
	
End Sub
%>
