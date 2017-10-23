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

	Dim C_W1_1
	Dim C_W1_2
	Dim C_W1_CD
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5

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
	C_W1_1				= 1
    C_W1_2				= 2
    C_W1_CD				= 3
    C_W2				= 4
    C_W3				= 5
    C_W4				= 6
    C_W5				= 7
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
    lgStrSQL =            "DELETE TB_4 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
	'PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngRow
    Dim iRow, iKey1, iKey2, iKey3
	Dim arrRow(2), iType
	
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' ������� 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' �Ű��� 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '�� : No data is found.
        Call SetErrorStatus()
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write "Call Parent.FncNew"  &  vbCrLf
        Response.Write " </Script>"
    Else
       ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = "" : iLngRow = 1

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	.frm1.vspdData.ReDraw = False			" & vbCr
		' �׸��� �� �߰� 
		Response.Write "	.ggoSpread.Source = .frm1.vspdData		 " & vbCr
	
		' -- 2006-01-02 : 200603������ ���� (2�� �и�)
		Response.Write "	.ggoSpread.InsertRow , .C_ROW_23		 " & vbCr
        
		Do While Not lgObjRs.EOF
			
			If lgObjRs("W1") = "19" Then
				Response.Write "	Call .PutGrid(.C_W2, .C_ROW_" & lgObjRs("W1") & ", """ & MakePercent(lgObjRs("W2"))  & """)"  & vbCr
				Response.Write "	Call .PutGrid(.C_W3, .C_ROW_" & lgObjRs("W1") & ", """ & MakePercent(lgObjRs("W3"))  & """)"  & vbCr
				Response.Write "	Call .PutGrid(.C_W4, .C_ROW_" & lgObjRs("W1") & ", """ & MakePercent(lgObjRs("W4"))  & """)"  & vbCr
				Response.Write "	Call .PutGrid(.C_W5, .C_ROW_" & lgObjRs("W1") & ", """ & MakePercent(lgObjRs("W5"))  & """)"  & vbCr
			Else
				Response.Write "	Call .PutGrid(.C_W2, .C_ROW_" & lgObjRs("W1") & ", """ & lgObjRs("W2")  & """)"  & vbCr
				Response.Write "	Call .PutGrid(.C_W3, .C_ROW_" & lgObjRs("W1") & ", """ & lgObjRs("W3")  & """)"  & vbCr
				Response.Write "	Call .PutGrid(.C_W4, .C_ROW_" & lgObjRs("W1") & ", """ & lgObjRs("W4")  & """)"  & vbCr
				Response.Write "	Call .PutGrid(.C_W5, .C_ROW_" & lgObjRs("W1") & ", """ & lgObjRs("W5")  & """)"  & vbCr
			End If			
			lgObjRs.MoveNext
		Loop 

		lgObjRs.Close
		Set lgObjRs = Nothing

		Response.Write "	.frm1.vspdData.ReDraw = True			" & vbCr
		
		Response.Write "	.DbQueryOk                                      " & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr
    
    End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


End Sub

Function RemoveZero(Byval pVal)
	If CDbl(pVal) <> 0 Then
		RemoveZero = pVal
	Else
		RemoveZero = ""
	End If
End Function

Function MakePercent(Byval pVal)
	If CDbl(pVal) <> 0 Then
		MakePercent = "0." & pVal
	Else
		MakePercent = ""
	End If
End Function

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W1, A.W2, A.W3, A.W4, A.W5 "
            lgStrSQL = lgStrSQL & " FROM TB_4 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	

    End Select

	'PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , iType

    'On Error Resume Next
    Err.Clear 
    

	' �׸��� 
	'PrintLog "txtSpread = " & Request("txtSpread")
			
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
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

End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(Byref arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	

	lgStrSQL = "INSERT INTO TB_4 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , W1, W2, W3, W4, W5 " & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1_CD))),"''","S")     & "," & vbCrLf
	
	If arrColVal(C_W1_CD) = "19" Then
		lgStrSQL = lgStrSQL & FilterVar(RemovePercent(arrColVal(C_W2)),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(RemovePercent(arrColVal(C_W3)),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(RemovePercent(arrColVal(C_W4)),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(RemovePercent(arrColVal(C_W5)),"0","D")     & "," & vbCrLf
	Else
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W2), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")     & "," & vbCrLf
	End If
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	
	'PrintLog "SubBizSaveMultiCreate1 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(Byref arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '��: Clear Error status


	lgStrSQL = "UPDATE  TB_4 WITH (ROWLOCK) "
	lgStrSQL = lgStrSQL & " SET " 
	
	If arrColVal(C_W1_CD) = "19" Then
		lgStrSQL = lgStrSQL & " W2    = " &  FilterVar(RemovePercent(arrColVal(C_W2)),"0","D") & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W3    = " &  FilterVar(RemovePercent(arrColVal(C_W3)),"0","D") & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W4    = " &  FilterVar(RemovePercent(arrColVal(C_W4)),"0","D") & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W5    = " &  FilterVar(RemovePercent(arrColVal(C_W5)),"0","D") & "," & vbCrLf
	Else
		lgStrSQL = lgStrSQL & " W2    = " &  FilterVar(UNICDbl(arrColVal(C_W2), "0"),"0","D") & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W3    = " &  FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W4    = " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W5    = " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & "," & vbCrLf
	End If
	lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
	lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W1 = " & FilterVar(Trim(UCase(arrColVal(C_W1_CD))),"''","S") 	 & vbCrLf 
		
		
	'PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

Function RemovePercent(Byval pVal)
	RemovePercent = Replace(pVal, "%", "")
End Function

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
