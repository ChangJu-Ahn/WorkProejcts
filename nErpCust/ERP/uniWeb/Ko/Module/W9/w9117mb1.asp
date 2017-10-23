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
	Dim lgStrPrevKey, lgCurrGrid
		

	Dim C_W1
	Dim C_W1_NM
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8
	Dim C_W9
	Dim C_W10
	Dim C_W2_POP
	Dim C_W3_POP
	Dim C_W4_POP
	Dim C_W5_POP
	Dim C_W6_POP
	Dim C_W7_POP
	Dim C_W8_POP
	Dim C_W9_POP
	Dim C_W10_POP
	Dim C_W2_CD
	Dim C_W3_CD
	Dim C_W4_CD
	Dim C_W5_CD
	Dim C_W6_CD
	Dim C_W7_CD
	Dim C_W8_CD
	Dim C_W9_CD
	Dim C_W10_CD


	lgErrorStatus		= "NO"
    lgOpModeCRUD		= Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    sFISC_YEAR			= Request("txtFISC_YEAR")
    sREP_TYPE			= Request("cboREP_TYPE")
	
    lgPrevNext			= Request("txtPrevNext")                                       '��: "P"(Prev search) "N"(Next search)
    lgLngMaxRow			= Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey		= UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
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
    
    
Function  MinorQueryRs(byval strMajorcd, byval strMinorcd)
   Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
   Call CommonQueryRs("MINOR_NM"," B_MINOR "," MAJOR_CD = '"& strMajorcd &"' and minor_cd = '"& Trim(strMinorcd) &"' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   MinorQueryRs = replace(lgF0, chr(11),"")
   

End Function    
Function ChangevbCrLf(byval pData)
   ChangevbCrLf = Replace(Trim(pData), vbCrLf, "")


End Function

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()	' ����Ÿ �Ѱ��ִ� �÷� ���� 
  C_W1				= 1
    C_W1_NM				= 2
    C_W2				= 3
   
    C_W2_POP			= 4
    C_W2_CD				= 5
    C_W3				= 6
    
    C_W3_POP			= 7
      C_W3_CD			= 8
    C_W4				= 9
    
    C_W4_POP			= 10
    C_W4_CD				= 11
    C_W5				= 12
    
    C_W5_POP			= 13
    C_W5_CD				= 14
    C_W6				= 15
    
    C_W6_POP			= 16
    C_W6_CD				= 17
    C_W7				= 18
    
    C_W7_POP			= 19
    C_W7_CD				= 20
    C_W8				= 21
    
    C_W8_POP			= 22
    C_W8_CD				= 23
    C_W9				= 24
    C_W9_POP			= 25
    C_W9_CD				= 26
    C_W10				= 27
    C_W10_POP			= 28
    C_W10_CD				= 29
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
    On Error Resume Next
    Err.Clear

	' �����Ϻ��� �����Ѵ�.
    lgStrSQL =            "DELETE TB_KJ1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1), sData
    Dim iDx
    Dim iLoopMax,iLngCol,sW_TYPE,strMajor
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' ������� 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' �Ű��� 

	Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
	     lgStrPrevKey = ""
	    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '�� : No data is found.
	    Call SetErrorStatus()
		    
	Else
	    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
	    lgstrData = ""



		iLngCol = lgObjRs.Fields.Count
		sW_TYPE = "" : lgstrData = ""
		iDx = 1

				lgstrData = lgstrData & " With parent.frm1.vspdData " & vbCr
				lgstrData = lgstrData & "	.Redraw = false " & vbCr
				Do While Not lgObjRs.EOF
					
                    
                    
                    if iDx = 7   or iDx=9 or  idx = 10 then 
                       if idx = 7 then strMajor ="W1051"
                       if idx = 9 then strMajor ="W1089"
                       if idx = 10 then strMajor ="W1052"
						 lgstrData = lgstrData & "	.Row = " &iDx & "" & vbCrLf
						lgstrData = lgstrData & "	.Col = 0 : .value = """" " & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W1    & " : .text = """ & lgObjRs("W1") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W2    & " : .text = """ & MinorQueryRs(strMajor,lgObjRs("W2")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W2_CD & " : .text = """ & lgObjRs("W2") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W3    & " : .text = """ & MinorQueryRs(strMajor,lgObjRs("W3")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W3_CD & " : .text = """ & lgObjRs("W3") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W4    & " : .text = """ & MinorQueryRs(strMajor,lgObjRs("W4")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W4_CD & " : .text = """ & lgObjRs("W4") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W5    & " : .text = """ & MinorQueryRs(strMajor,lgObjRs("W5")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W5_CD & " : .text = """ & lgObjRs("W4") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W6    & " : .text = """ & MinorQueryRs(strMajor,lgObjRs("W6")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W6_CD & " : .text = """ & lgObjRs("W6") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W7    & " : .text = """ & MinorQueryRs(strMajor,lgObjRs("W7")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W7_CD & " : .text = """ & lgObjRs("W7") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W8    & " : .text = """ & MinorQueryRs(strMajor,lgObjRs("W8")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W8_CD & " : .text = """ & lgObjRs("W8") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W9    & " : .text = """ & MinorQueryRs(strMajor,lgObjRs("W9")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W9_CD & " : .text = """ & lgObjRs("W9") & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W10    & " :.text = """ & MinorQueryRs(strMajor,lgObjRs("W10")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W10_CD & " :.text = """ & lgObjRs("W10") & """" & vbCrLf
						
					else
						lgstrData = lgstrData & "	.Row = " &iDx & "" & vbCrLf
						lgstrData = lgstrData & "	.Col = 0 : .value = """" " & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W1    & " : .value = """ & ChangevbCrLf(lgObjRs("W1")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W2    & " : .value = """ & ChangevbCrLf(lgObjRs("W2")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W2_CD & " : .value = """ & ChangevbCrLf(lgObjRs("W2")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W3    & " : .value = """ & ChangevbCrLf(lgObjRs("W3")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W3_CD & " : .value = """ & ChangevbCrLf(lgObjRs("W3")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W4    & " : .value = """ & ChangevbCrLf(lgObjRs("W4")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W4_CD & " : .value = """ & ChangevbCrLf(lgObjRs("W4")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W5    & " : .value = """ & ChangevbCrLf(lgObjRs("W5")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W5_CD & " : .value = """ & ChangevbCrLf(lgObjRs("W4")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W6    & " : .value = """ & ChangevbCrLf(lgObjRs("W6")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W6_CD & " : .value = """ & ChangevbCrLf(lgObjRs("W6")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W7    & " : .value = """ & ChangevbCrLf(lgObjRs("W7")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W7_CD & " : .value = """ & ChangevbCrLf(lgObjRs("W7")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W8    & " : .value = """ & ChangevbCrLf(lgObjRs("W8")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W8_CD & " : .value = """ & ChangevbCrLf(lgObjRs("W8")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W9    & " : .value = """ & ChangevbCrLf(lgObjRs("W9")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W9_CD & " : .value = """ & ChangevbCrLf(lgObjRs("W9")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W10    & " :.value = """ & ChangevbCrLf(lgObjRs("W10")) & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W10_CD & " :.value = """ & ChangevbCrLf(lgObjRs("W10")) & """" & vbCrLf	
                    
                    end if
                   
				If Err.number <> 0 Then
					'PrintLog "iDx=" & iDx
					Exit Sub
				End If
		        iDx = iDx +1    
				lgObjRs.MoveNext
	

			Loop


		
		lgObjRs.Close
		Set lgObjRs = Nothing
			
		lgstrData = lgstrData & "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE" & vbCrLf
    End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write lgstrData  &  vbCrLf
		
	If lgstrData <> "" Then	
		Response.Write "	.Redraw = True " & vbCr
		Response.Write " End With " & vbCrLf	' With �� ���� 
		
		
	End If

'	Response.Write " Call parent.DbQueryOk                                      " & vbCr
	Response.Write " </Script>                                          " & vbCr
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	

     
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "  A.W1, A.W2, A.W3, A.W4, A.W5, A.W6  ,A.W7, A.W8, A.W9 ,  A.W10"
            lgStrSQL = lgStrSQL & " FROM TB_KJ1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.W1 ASC" & vbcrlf

    End Select
	'PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i, sData

    On Error Resume Next
    Err.Clear 

	sData = Request("txtSpread")
	PrintLog "1��° �׸���.. : " & sData & "lgLngMaxRow=" & lgLngMaxRow
	
	If sData <> "" Then
		arrRowVal = Split(sData, gRowSep)                                 '��: Split Row    data
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
		       Exit Sub
		    End If
		    
		Next
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
	
	lgStrSQL = "INSERT INTO TB_KJ1 WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE " 
	lgStrSQL = lgStrSQL & " , W1 , W2, W3, W4, W5, W6 , W7 ,  W8 , W9 , W10 "
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )" 
	lgStrSQL = lgStrSQL & " VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(1))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(6))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(7))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(8))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(9))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(10))),"''","S")     & ","
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"
 
	'PrintLog "SubBizSaveMultiCreate = " & lgStrSQL
	
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

    lgStrSQL = "UPDATE  TB_KJ1 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 

    lgStrSQL = lgStrSQL & " W2     = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")     & ","
    lgStrSQL = lgStrSQL & " W3     = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")     & ","
    lgStrSQL = lgStrSQL & " W4     = " &  FilterVar(Trim(UCase(arrColVal(4))),"''","S")     & ","
    lgStrSQL = lgStrSQL & " W5     = " &  FilterVar(Trim(UCase(arrColVal(5))),"''","S")     & ","
    lgStrSQL = lgStrSQL & " W6     = " &  FilterVar(Trim(UCase(arrColVal(6))),"''","S")     & ","
    lgStrSQL = lgStrSQL & " W7     = " &  FilterVar(Trim(UCase(arrColVal(7))),"''","S")     & ","
    lgStrSQL = lgStrSQL & " W8     = " &  FilterVar(Trim(UCase(arrColVal(8))),"''","S")     & ","
    lgStrSQL = lgStrSQL & " W9     = " &  FilterVar(Trim(UCase(arrColVal(9))),"''","S")     & ","
    lgStrSQL = lgStrSQL & " W10    = " &  FilterVar(Trim(UCase(arrColVal(10))),"''","S")     & ","
    



    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W1 = " & FilterVar(Trim(UCase(arrColVal(1))),"''","S") 	 & vbCrLf  

	'PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  TB_KJ1 WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W1 = " & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")  


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
    On Error Resume Next
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
       Case "<%=UID_M0001%>"

          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQueryOk        
          End If   

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