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

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    'On Error Resume Next
    Err.Clear
   
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Const TYPE_1	= 0		' 그리드를 구분짓기 위한 상수 
	Const TYPE_2	= 1		

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
    
    Call CheckVersion(sFISC_YEAR, sREP_TYPE)	' 2005-03-11 버전관리기능 추가 
    	
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

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_54_BPD WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_54_BPH WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
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
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
		iLngCol = lgObjRs.Fields.Count
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	.IsRunEvents = True " & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.

		For iDx = 0 To iLngCol-1
			Select Case iDx
				Case C_W4
					lgstrData = lgstrData & "	.frm1.txtData(" & CStr(iDx) & ").value = """ & lgObjRs(iDx).value & """" & vbCrLf
					lgstrData = lgstrData & "	.frm1.txtW4_" & lgObjRs(iDx).value & ".checked = true" & vbCrLf
				Case Else
					lgstrData = lgstrData & "	.frm1.txtData(" & CStr(iDx) & ").value = """ & lgObjRs(iDx).value & """" & vbCrLf
			End Select
			
			If Err.number <> 0 Then
				PrintLog "iDx=" & iDx
				Exit Sub
			End If
		Next 
		
		Response.Write lgstrData  &  vbCrLf
		Response.Write "	.IsRunEvents = False " & vbCrLf	' 이벤트가 발생하게 한다.
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr

		lgObjRs.Close
		Set lgObjRs = Nothing
			
		iStrData = ""
		'TYPE_2 조회 
	    Call SubMakeSQLStatements("RD",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

			lgstrData = ""
				
			Do While Not lgObjRs.EOF
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				If CDbl(lgObjRs("SEQ_NO").value) = 1 Then
					iStrData = iStrData & Chr(11) & "계"
				Else
					iStrData = iStrData & Chr(11) & CDbl(lgObjRs("SEQ_NO").value) - 1
				End If
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W5"))			
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W6"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W7"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W8"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W9"))
				iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
				iStrData = iStrData & Chr(11) & Chr(12)
				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 

			lgObjRs.Close
			Set lgObjRs = Nothing
			
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent										" & vbCr
			Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2 & ")        " & vbCr
			Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
			Response.Write " Call .SetTotalLine                                  " & vbCr
			Response.Write " End With	                        " & vbCr
			Response.Write " </Script>	                        " & vbCr
		
		End If
	End If
				
	'Response.Write " <Script Language=vbscript>	                        " & vbCr
	'Response.Write " With parent                                        " & vbCr
	'Response.Write "	.DbQueryOk                                      " & vbCr
	'Response.Write " End With                                           " & vbCr
	'Response.Write " </Script>                                          " & vbCr
	    
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
            lgStrSQL = lgStrSQL & " A.W1, A.W2, A.W3, A.W4, A.W10 "	' HTML순서대로 읽어와야함 
            lgStrSQL = lgStrSQL & " FROM TB_54_BPH A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "RD"
			lgStrSQL =			  " SELECT "
            lgStrSQL = lgStrSQL & "   A.SEQ_NO, A.W5, A.W6, A.W7, A.W8, A.W9 "
            lgStrSQL = lgStrSQL & " FROM TB_54_BPD A WITH (NOLOCK) "
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
    Dim iDx , iType

    'On Error Resume Next
    Err.Clear 
    
 	' 그리드 
	PrintLog "txtSpread = " & Request("txtSpread" & CStr(TYPE_1))
			
    arrColVal = Split(Request("txtSpread" & CStr(TYPE_1)), gColSep)    
	
	PrintLog "txtHeadMode=" & Request("txtHeadMode") & ";" & OPMD_CMODE
	
	If CDbl(Request("txtHeadMode")) = OPMD_CMODE Then
	    Call SubBizSaveMultiCreate(TYPE_1, arrColVal)                            '☜: Create
	Else
	    Call SubBizSaveMultiUpdate(TYPE_1, arrColVal)                            '☜: Update
	End If
				    
	' 그리드 
	PrintLog "txtSpread = " & Request("txtSpread" & CStr(TYPE_2))
			
	arrRowVal = Split(Request("txtSpread" & CStr(TYPE_2) ), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
	For iDx = 1 To lgLngMaxRow

	    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
			    
	    Select Case arrColVal(0)
	        Case "C"
	                Call SubBizSaveMultiCreate(TYPE_2, arrColVal)                            '☜: Create
	        Case "U"
	                Call SubBizSaveMultiUpdate(TYPE_2, arrColVal)                            '☜: Update
	        Case "D"
	                Call SubBizSaveMultiDelete(TYPE_2, arrColVal)                            '☜: Update
	    End Select
			    
	    If lgErrorStatus    = "YES" Then
	       lgErrorPos = lgErrorPos & arrColVal(iDx) & gColSep
	       Exit For
	    End If
			    
	Next
	
End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(Byval pType, Byref arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	Select Case pType
		Case TYPE_1
	
			lgStrSQL = "INSERT INTO TB_54_BPH WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , W1, W2, W3, W4, W10" & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W4))),"''","S")     & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")     & "," & vbCrLf
			
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"

		Case TYPE_2

			lgStrSQL = "INSERT INTO TB_54_BPD WITH (ROWLOCK) (" & vbCrLf
			lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
			lgStrSQL = lgStrSQL & " , SEQ_NO, W5, W6, W7, W8, W9 " & vbCrLf
			lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
			lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
			    
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
			
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W5))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W6))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S")  & "," & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W8))),"''","S")  & "," & vbCrLf
	
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D")     & "," & vbCrLf

			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
			'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
			lgStrSQL = lgStrSQL & ")"
			
	End Select
	PrintLog "SubBizSaveMultiCreate1 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(Byval pType, Byref arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

	Select Case pType
		Case TYPE_1
		
			lgStrSQL = "UPDATE  TB_54_BPH WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(Trim(UCase(arrColVal(C_W4))),"''","S") & "," & vbCrLf

			lgStrSQL = lgStrSQL & " W10		= " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & "," & vbCrLf

			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
		
		Case TYPE_2
		
			lgStrSQL = "UPDATE  TB_54_BPD WITH (ROWLOCK) "
			lgStrSQL = lgStrSQL & " SET " 
			lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(Trim(UCase(arrColVal(C_W5))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W6		= " &  FilterVar(Trim(UCase(arrColVal(C_W6))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S") & "," & vbCrLf
			lgStrSQL = lgStrSQL & " W8		= " &  FilterVar(Trim(UCase(arrColVal(C_W8))),"''","S") & "," & vbCrLf
		
			lgStrSQL = lgStrSQL & " W9		= " &  FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D") & "," & vbCrLf
			
			lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
			lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 
									
	End Select
	
	PrintLog "SubBizSaveMultiUpdate1 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

Function RemovePercent(Byval pVal)
	RemovePercent = Replace(pVal, "%", "")
End Function

'========================================================================================
Sub SubBizSaveMultiDelete(Byval pType, Byref arrColVal)
    'On Error Resume Next
    Err.Clear


	Select Case pType

		Case TYPE_2

			lgStrSQL =            "DELETE TB_54_BPD WITH (ROWLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D") 	 & vbCrLf 

	End Select
	
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
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
