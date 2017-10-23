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

	Const BIZ_MNU_ID = "W4103MA1"
	Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
	Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.
			
	Dim C_SEQ_NO	

	Dim C_W9
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14
	Dim C_W15
	Dim C_W16
	Dim C_W17
	Dim C_W18

	Dim C_W19
	Dim C_W20
	Dim C_W21
	Dim C_W22
	Dim C_W23
	Dim C_W24
	Dim C_W25

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
	C_W9		= 2	' �ձݻ��Կ��� 
	C_W10		= 3	' ������ 
	C_W11		= 4	' ��λ��غ�� 
	C_W12		= 5 ' �����غ�� 
	C_W13		= 6	' �غ�� 
	C_W14		= 7	' ��ü�ҿ��ڱݻ��� 
	C_W15		= 8	' �̻��� 
	C_W16		= 9	' ��ü�ҿ��ڱݻ��� 
	C_W17		= 10 ' ��Ÿ 
	C_W18		= 11 ' �� 
	
	' C_SEQ_NO, C_W9 ���� 
	C_W19		= 3	' 1������ 
	C_W20		= 4	' 2������ 
	C_W21		= 5	' 3���⵵ 
	C_W22		= 6 ' �� 
	C_W23		= 7	' ȯ���ұݾ��հ� 
	C_W24		= 8	' ȸ��ȯ�Ծ� 
	C_W25		= 9	' ����ȯ�� 
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
    lgStrSQL =            "DELETE TB_31_2D2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgStrSQL = lgStrSQL & "DELETE TB_31_2D1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgStrSQL = lgStrSQL & "DELETE TB_31_2H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
 	Call TB_15_DeleData("", -1)
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

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '�� : No data is found.
        Call SetErrorStatus()
        
    Else
       ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = "" : iLngRow = 1
        
		arrRs(iRow) = ""

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent.frm1                                   " & vbCr
		Response.Write "	.txtW1.value = """ & lgObjRs("W1") & """" & vbCrLf
		Response.Write "	.cboW2.value = """ & lgObjRs("W2") & """"  & vbCrLf
		Response.Write "	.txtW2_VAL.value = """ & lgObjRs("W2_VAL") & """"  & vbCrLf
		Response.Write "	.txtW3.value = """ & lgObjRs("W3") & """"  &  vbCrLf
		Response.Write "	.txtW4.value = """ & lgObjRs("W4") & """"  &  vbCrLf
		Response.Write "	.txtW5.value = """ & lgObjRs("W5") & """"  &  vbCrLf
		Response.Write "	.txtW6.value = """ & lgObjRs("W6") & """"  &  vbCrLf
		Response.Write "	.txtW7.value = """ & lgObjRs("W7") & """"  &  vbCrLf
		Response.Write "	.txtW8.value = """ & lgObjRs("W8") & """"  &  vbCrLf
		Response.Write "	.txtDESC1.value = """ & lgObjRs("DESC1") & """"  &  vbCrLf
		Response.Write " End With                                  " & vbCr
		Response.Write " </Script>	                        " & vbCr
		
		' 1�� �׸��� 
	    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

		   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = "" : iLngRow = 1 : iRow = TYPE_1
		    
			arrRs(iRow) = ""
				
			Do While Not lgObjRs.EOF
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W9"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W10")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W11")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W12")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W13")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W14")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W15")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W16")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W17")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W18")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow
				arrRs(iRow) = arrRs(iRow) & Chr(11) & Chr(12)

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 

		End If
		
		' 2�� �׸��� 
	    Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

		   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = "" : iLngRow = 1 : iRow = TYPE_2
		    
			arrRs(iRow) = ""
				
			Do While Not lgObjRs.EOF
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & ConvSPChars(lgObjRs("W9"))
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W19")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W20")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W21")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W22")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W23")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W24")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & lgObjRs("W25")
				arrRs(iRow) = arrRs(iRow) & Chr(11) & iLngRow
				arrRs(iRow) = arrRs(iRow) & Chr(11) & Chr(12)

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 

		End If
		

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

		Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_1 & ")" & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & arrRs(TYPE_1)       & """" & vbCr

		Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2 & ")" & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & arrRs(TYPE_2)       & """" & vbCr

		Response.Write "	.DbQueryOk                                      " & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr
    
    End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  TOP 1 "
            lgStrSQL = lgStrSQL & " A.W1, A.W2, A.W2_VAL, A.W3, A.W4, A.W5, A.W6 "
            lgStrSQL = lgStrSQL & " , A.W7, A.W8, A.DESC1 "
            lgStrSQL = lgStrSQL & " FROM TB_31_2H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "R1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W9, A.W10, A.W11, A.W12, A.W13, A.W14, A.W15, A.W16, A.W17, A.W18 "
            lgStrSQL = lgStrSQL & " FROM TB_31_2D1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO" & vbcrlf
      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W9, A.W19, A.W20, A.W21, A.W22, A.W23, A.W24, A.W25 "
            lgStrSQL = lgStrSQL & " FROM TB_31_2D2 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf

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
    
    ' ��� ���� 
    If Request("txtHeadMode") = CStr(OPMD_UMODE) Then
		Call SubBizSaveUpdate
    Else
		Call SubBizSaveCreate
	End If
	
	' 1�� �׸��� 
	PrintLog "txtSpread = " & Request("txtSpread" & CStr(TYPE_1))
		
	arrRowVal = Split(Request("txtSpread" & CStr(TYPE_1) ), gRowSep)                                 '��: Split Row    data
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

	' 2�� �׸��� 
	PrintLog "txtSpread = " & Request("txtSpread" & CStr(TYPE_2))
		
	arrRowVal = Split(Request("txtSpread" & CStr(TYPE_2) ), gRowSep)                                 '��: Split Row    data
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
Sub SubBizSaveCreate()
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_31_2H WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , W1, W2, W2_VAL, W3, W4, W5, W6, W7, W8, DESC1 " & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("cboW2"))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW2_VAL"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtDESC1"))),"''","S")     & "," & vbCrLf	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreateH = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


	If UNICDbl(Request("txtW8"), 0) > 0 Then
		Call TB_15_PushData("2", UNICDbl(Request("txtW8"), 0), 1, "3102", "100", "�����η°��� �غ���� �ձݻ����ϰ� ����ó����")
	End If
	
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_31_2D1 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18 " & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")     & "," & vbCrLf

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate1 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_31_2D2 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W9, W19, W20, W21, W22, W23, W24, W25" & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D")     & "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

'	(+)����� ��� 15-1ȣ�� (1)����� "�����η°����غ��" (2)�ݾ׿��� ���� �ݾ��� 
'	(3)�ҵ�ó�п��� "����(����)"�� �Է��ϰ� ���������� " �����η°��� �غ�� ����ȯ�Ծ��� �ͱݻ����ϰ� 
'	����ó����."�� �Է��ϰ� ����Ͽ���.

'	(-)������ ��� 15-2ȣ�� (1)����� "�����ΰᰳ���غ��" (2)�ݾ׿��� ���� �ݾ��� ���밪�� 
'	(3)�ҵ�ó�п��� "����(����)"�� �Է��ϰ� ���������� " �����η°��� �غ�� ����ȯ�Ծ��� �ͱݺһ����ϰ� 
'	����ó����."�� �Է��ϰ� ����Ͽ���.
					
	If UNICDbl(arrColVal(C_SEQ_NO), 0) = 999999 Then
		Call TB_15_DeleData("", 999999)
		If UNICDbl(arrColVal(C_W25), 0) > 0 Then
			Call TB_15_PushData("1", UNICDbl(arrColVal(C_W25), 0), 999999, "3102", "400", "�����η°��� �غ�� ����ȯ�Ծ��� �ͱݻ����ϰ� ����ó����")
		ElseIf UNICDbl(arrColVal(C_W25), 0) < 0 Then
			Call TB_15_PushData("2", ABS(UNICDbl(arrColVal(C_W25), 0)), 999999, "3102", "100", "�����η°��� �غ�� ����ȯ�Ծ��� �ͱݺһ����ϰ� ����ó����")
		End If
	End If

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveUpdate()
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "UPDATE  TB_31_2H WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    'lgStrSQL = lgStrSQL & " W1     = " &  FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W1     = " &  FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2     = " &  FilterVar(Trim(UCase(Request("cboW2"))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2_VAL = " &  FilterVar(UNICDbl(Request("txtW2_VAL"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3     = " &  FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4     = " &  FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5     = " &  FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6     = " &  FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W7     = " &  FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W8     = " &  FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " DESC1  = " &  FilterVar(Request("txtDESC1"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

	Call TB_15_DeleData("", 1)
	If UNICDbl(Request("txtW8"), 0) > 0 Then
		Call TB_15_PushData("2", UNICDbl(Request("txtW8"), 0), 1, "3102", "100", "�����η°��� �غ���� �ձݻ����ϰ� ����ó����")
	End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "UPDATE  TB_31_2D1 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W9      = " &  FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W10     = " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W11     = " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W12     = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W13     = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W14     = " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W15     = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W16     = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W17     = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W18     = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "UPDATE  TB_31_2D2 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W9      = " &  FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W19     = " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W20     = " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W21     = " &  FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W22     = " &  FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W23     = " &  FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W24     = " &  FilterVar(UNICDbl(arrColVal(C_W24), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W25     = " &  FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate2 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

'	(+)����� ��� 15-1ȣ�� (1)����� "�����η°����غ��" (2)�ݾ׿��� ���� �ݾ��� 
'	(3)�ҵ�ó�п��� "����(����)"�� �Է��ϰ� ���������� " �����η°��� �غ�� ����ȯ�Ծ��� �ͱݻ����ϰ� 
'	����ó����."�� �Է��ϰ� ����Ͽ���.

'	(-)������ ��� 15-2ȣ�� (1)����� "�����ΰᰳ���غ��" (2)�ݾ׿��� ���� �ݾ��� ���밪�� 
'	(3)�ҵ�ó�п��� "����(����)"�� �Է��ϰ� ���������� " �����η°��� �غ�� ����ȯ�Ծ��� �ͱݺһ����ϰ� 
'	����ó����."�� �Է��ϰ� ����Ͽ���.
					
	If arrColVal(C_SEQ_NO) = "999999" Then
		Call TB_15_DeleData("", 999999)
		If UNICDbl(arrColVal(C_W25), 0) > 0 Then
			Call TB_15_PushData("1", UNICDbl(arrColVal(C_W25), 0), 999999, "3102", "400", "�����η°��� �غ�� ����ȯ�Ծ��� �ͱݻ����ϰ� ����ó����")
		ElseIf UNICDbl(arrColVal(C_W25), 0) < 0 Then
			Call TB_15_PushData("2", ABS(UNICDbl(arrColVal(C_W25), 0)), 999999, "3102", "100", "�����η°��� �غ�� ����ȯ�Ծ��� �ͱݺһ����ϰ� ����ó����")
		End If
	End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  TB_31_2D1 WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 
	
	PrintLog "SubBizSaveMultiDelete1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  TB_31_2D2 WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf 
	
	PrintLog "SubBizSaveMultiDelete1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

	' -- 15�� ���� 
	If UNICDbl(arrColVal(C_SEQ_NO), 0) = 999999 Then
 		Call TB_15_DeleData("", -1)
 	End If
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

Sub TB_15_DeleData(Byval pType, Byval pSeqNo)
	On Error Resume Next 
	Err.Clear  

	lgStrSQL = "EXEC usp_TB_15_DeleData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' ���� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' ������� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' �Ű��� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' �������� ���α׷� 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' �������� ���� 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pType)),"''","S") & ", "			' 1ȣ/2ȣ 
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
