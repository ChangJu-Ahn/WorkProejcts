<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Dim lgStrPrevKey,lgStrPrevKey1
   	Const C_SHEETMAXROWS_D = 100
 
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
                                                                   '☜: Clear Error status
	
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)
	if lgCurrentSpd = "1" then
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	else		
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	end if
    
    Dim iCnt         '2개 Sheet모두 데이터가 없는지 체크후 메세지 
    iCnt=0

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubBizQueryMulti()
        
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iKey1,iKey2, iKey3, iKey4

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgstrData   = ""
    lgstrData1  = ""

    iKey1 = FilterVar(lgKeyStream(0) & "%", "''", "S")
    iKey3 = FilterVar(lgKeyStream(0) & "0101", "''", "S")
    iKey4 = FilterVar(lgKeyStream(0) & "1231", "''", "S")
    iKey2 = FilterVar(lgKeyStream(1), "''", "S")
        
    if lgCurrentSpd = 1 then
		Call SubMakeSQLStatements("MR",iKey1,iKey2,"")                                   '☆ : Make sql statements
		Call SubBizQueryMultiData(lgCurrentSpd)                                          '☆ : Save Array Data
	elseif lgCurrentSpd = 2  then
		Call SubMakeSQLStatements("MR",iKey2,iKey3,iKey4)                                '☆ : Make sql statements
		Call SubBizQueryMultiData(lgCurrentSpd)                                          '☆ : Save Array Data
	end if    

End Sub    

'============================================================================================================
' Name : SubBizQueryMultiData
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMultiData(pCurrentSpd)
     
    Dim istrData
             
    dim idx
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        iCnt = iCnt + 1           '2개 Sheet모두 데이터가 없는지 체크후 메세지 
        If iCnt = 2 Then
            lgStrPrevKey = ""
            lgStrPrevKey1 = ""
            Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
            Call SetErrorStatus()
        End If
    Else
		if lgCurrentSpd = "1" then
			Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		elseif lgCurrentSpd = "2" then
			Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)
		end if		
        iDx       = 1
        istrData = ""
        Do While Not lgObjRs.EOF
			If pCurrentSpd = "1" Then
		        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PROV_DT"))
			    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BONUS_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("NON_TOT_TAX"), ggAmtOfMoney.DecPoint,0)
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TAX_AMT"), ggAmtOfMoney.DecPoint,0)
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint,0)
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint,0)
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SAVE_FUND"), ggAmtOfMoney.DecPoint,0)
				lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)

			ElseIf pCurrentSpd = "2" Then
		        lgstrData1 = lgstrData1 & Chr(11) & UNIDateClientFormat(lgObjRs("PAY_YYMM"))
			    lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("PAY_TEXT"))
				lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
				lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint,0)
				lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint,0)
				lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + iDx
				lgstrData1 = lgstrData1 & Chr(11) & Chr(12)
				
			End If
            
	        lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
				if lgCurrentSpd = "1" then
					lgStrPrevKey = lgStrPrevKey + 1
				elseif lgCurrentSpd = "2" then
					lgStrPrevKey1 = lgStrPrevKey1 + 1
				end if		
               Exit Do
            End If                  
        Loop 
    End If
    If iDx <= C_SHEETMAXROWS_D Then
		if lgCurrentSpd = "1" then
			 lgStrPrevKey = ""
		elseif lgCurrentSpd = "2" then
			 lgStrPrevKey1= ""
		end if    
    End If   
    Call SubHandleError("MR",lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
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
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode1,pCode2,pCode3)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                '☜: Clear Error status
    lgStrSQL = ""

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           Select Case Mid(pDataType,2,1)
               Case "R"
                    Select Case lgCurrentSpd
                       Case "1"
							iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1                       
							
							lgStrSQL = "Select TOP " & iSelCount  & " PROV_DT, MINOR_NM, PAY_TOT_AMT, BONUS_TOT_AMT, "
							lgStrSQL = lgStrSQL & "       NON_TAX1 + NON_TAX2 + NON_TAX3 + NON_TAX4 + NON_TAX5 NON_TOT_TAX, "
							lgStrSQL = lgStrSQL & "       TAX_AMT, INCOME_TAX, RES_TAX, SAVE_FUND "
                            lgStrSQL = lgStrSQL & " From  HDF070T A, B_MINOR B"
			                lgStrSQL = lgStrSQL & " WHERE PAY_YYMM  LIKE " & pCode1 
                            lgStrSQL = lgStrSQL & "   AND EMP_NO    = "    & pCode2
                            lgStrSQL = lgStrSQL & "   AND PROV_TYPE NOT IN (" & FilterVar("P", "''", "S") & " ," & FilterVar("Q", "''", "S") & " ," & FilterVar("R", "''", "S") & " ," & FilterVar("S", "''", "S") & " ) "
                            lgStrSQL = lgStrSQL & "   AND MAJOR_CD  = " & FilterVar("H0040", "''", "S") & ""
                            lgStrSQL = lgStrSQL & "   AND PROV_TYPE = MINOR_CD"
                            lgStrSQL = lgStrSQL & " ORDER BY PROV_DT"
                       Case "2"
							iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1
                       
							lgStrSQL = "Select TOP " & iSelCount  & " PAY_YYMM, PAY_TEXT, PAY_TOT_AMT, INCOME_TAX, RES_TAX "
                            lgStrSQL = lgStrSQL & " From  HFA080T "
			                lgStrSQL = lgStrSQL & " WHERE PAY_YYMM  BETWEEN " & pCode2 & " AND " & pCode3
                            lgStrSQL = lgStrSQL & "   AND EMP_NO    = " & pCode1
                            lgStrSQL = lgStrSQL & " ORDER BY PAY_YYMM"
					End Select
           End Select 
    End Select
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             With Parent
				if .lgCurrentSpd = "1" then
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					.ggoSpread.Source     = .frm1.vspdData1
					.ggoSpread.SSShowData "<%=lgstrData%>"
					if .topleftOK then
						.DBQueryOk
					else
						.lgCurrentSpd = "2"						
						.DBQuery
					end if
                else
					.ggoSpread.Source     = .frm1.vspdData2
					.ggoSpread.SSShowData "<%=lgstrData1%>"
					.lgStrPrevKey1    = "<%=lgStrPrevKey1%>"
	                .DBQueryOk
                end if
	         End with
	      Else
             Parent.DBQueryNo
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	

