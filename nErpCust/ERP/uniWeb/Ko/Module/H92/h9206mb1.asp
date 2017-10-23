<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
 	Const C_SHEETMAXROWS_D = 100
	Dim lgStrPrevKey,lgStrPrevKey1
    Dim lgSpreadFlg
    	   
    Call LoadBasisGlobalInf()
    Call LoadinfTb19029B("Q", "H","NOCOOKIE","MB")                                                                     '☜: Clear Error status
	
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

    lgSpreadFlg       = Request("lgSpreadFlg")
	if lgSpreadFlg = "1" then
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	else		
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	end if

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
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

    Dim iYymm, iEmpNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iYymm     = FilterVar(lgKeyStream(0), "''", "S")
    iEmpNo    = FilterVar(lgKeyStream(1), "''", "S")
 	if lgSpreadFlg = "1" then   
		lgStrSQL = ""
		Call SubMakeSQLStatements("SR3",iYymm,iEmpNo,"")                               '☆: Make sql statements

		'HDF070T 조회 
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
		   If lgPrevNext = "" Then
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
				Call SetErrorStatus()
		   ElseIf lgPrevNext = "P" Then
		      Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the starting data. 
		      lgPrevNext = ""
		      Call SubBizQuery()
		   ElseIf lgPrevNext = "N" Then
		      Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the ending data.
		      lgPrevNext = ""
		      Call SubBizQuery()
		   End If
		   
		Else
		    %>
		    <Script Language=vbscript>
		        With Parent.Frm1
		              .txtProv_Dt.Text   = "<%=UNIConvDateDBToCompany(lgObjRs("PROV_DT"),"")%>"
		              .txtIncomeTaxAmt.text = "<%=UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint,0)%>"     
		              .txtResTaxAmt.text    = "<%=UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint,0)%>"     
		              .txtEmpInsurAmt.text= "<%=UNINumClientFormat(lgObjRs("EMP_INSUR"), ggAmtOfMoney.DecPoint,0)%>"     
		              .txtSubTotAmt.text    = "<%=UNINumClientFormat(lgObjRs("SUB_TOT_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
		              .txtNon_Tax5.text  = "<%=UNINumClientFormat(lgObjRs("NON_TAX5"), ggAmtOfMoney.DecPoint,0)%>"     
		              .txtTaxAmt.text    = "<%=UNINumClientFormat(lgObjRs("TAX_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
		              .txtTotAmt.text    = "<%=UNINumClientFormat(lgObjRs("PROV_TOT_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
		              .txtRealProvAmt.text= "<%=UNINumClientFormat(lgObjRs("REAL_PROV_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
		        End With
		    </Script>       
		    <%
		   Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
    
		    lgStrSQL = ""

		    '연차조회 
		    Call SubMakeSQLStatements("SR1",iYymm,iEmpNo,"")                               '☆: Make sql statements
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
		    Else
		        %>
		        <Script Language=vbscript>
		              With Parent.Frm1
		                     .txtDutyYy.value   = "<%=ConvSPChars(lgObjRs("DUTY_YY"))%>"            
		                     .txtDutyMm.value   = "<%=ConvSPChars(lgObjRs("DUTY_MM"))%>"            
		                     .txtDutyDd.value   = "<%=ConvSPChars(lgObjRs("DUTY_DD"))%>"        
		                     .txtYearBasAmt.text= "<%=UNINumClientFormat(lgObjRs("BAS_AMT"), ggAmtOfMoney.DecPoint,0)%>"         
		                     .txtYearSaveTot.text  = "<%=UNINumClientFormat(lgObjRs("YEAR_SAVE_TOT"), ggQty.DecPoint,0)%>"     
		                     .txtYearUse.text   = "<%=UNINumClientFormat(lgObjRs("YEAR_USE"), ggQty.DecPoint,0)%>"     
		                     .txtYearCnt.text   = "<%=UNINumClientFormat(lgObjRs("YEAR_CNT"), ggQty.DecPoint,0)%>"     
		                     .txtYearAmt.text   = "<%=UNINumClientFormat(lgObjRs("YEAR_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
		              End With          
		        </Script>       
		        <%     
		       Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
		    End If

		    lgStrSQL = ""

		    '월차사항 조회 
		    Call SubMakeSQLStatements("SR2",iYymm,iEmpNo,"")                               '☆: Make sql statements
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
		    Else
		        %>
		        <Script Language=vbscript>
		              With Parent.Frm1
		                     .txtDutyYy.value   = "<%=ConvSPChars(lgObjRs("DUTY_YY"))%>"            
		                     .txtDutyMm.value   = "<%=ConvSPChars(lgObjRs("DUTY_MM"))%>"            
		                     .txtDutyDd.value   = "<%=ConvSPChars(lgObjRs("DUTY_DD"))%>"        
		                     .txtMonthBasAmt.text= "<%=UNINumClientFormat(lgObjRs("BAS_AMT"), ggAmtOfMoney.DecPoint,0)%>"         
		                     .txtMonthSave.text = "<%=UNINumClientFormat(lgObjRs("MONTH_SAVE"), ggQty.DecPoint,0)%>"     
		                     .txtMonthUse.text  = "<%=UNINumClientFormat(lgObjRs("MONTH_USE"), ggQty.DecPoint,0)%>"     
		                     .txtMonthDutyCnt.text = "<%=UNINumClientFormat(lgObjRs("MONTH_DUTY_CNT"), ggQty.DecPoint,0)%>"     
		                     .txtMonthCnt.text  = "<%=UNINumClientFormat(lgObjRs("MONTH_CNT"), ggQty.DecPoint,0)%>"     
		                     .txtMonthAmt.text  = "<%=UNINumClientFormat(lgObjRs("MONTH_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
		              End With
		        </Script>       
		        <%     
		       Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
		    End If

		    Call SubBizQueryMultiA(iYymm,iEmpNo,"1")
		End If
	Else
		    Call SubBizQueryMultiB(iYymm,iEmpNo,"2")
	End if	
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
    
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '☜: Update
              Call SubBizSaveSingleUpdate()
    End Select

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HDF060T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " SUB_YYMM   = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO   = " & FilterVar(lgKeyStream(1), "''", "S")
    lgStrSQL = lgStrSQL & " AND SUB_TYPE = " & FilterVar("Z", "''", "S") & " "

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
    lgStrSQL = "DELETE  HDF070T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " PAY_YYMM   = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO   = " & FilterVar(lgKeyStream(1), "''", "S")
    lgStrSQL = lgStrSQL & " AND PROV_TYPE = " & FilterVar("Z", "''", "S") & " "
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
		
    lgStrSQL = "DELETE  HDF030T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " PAY_YYMM   = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO   = " & FilterVar(lgKeyStream(1), "''", "S")
    lgStrSQL = lgStrSQL & " AND ALLOW_CD = " & FilterVar("P13", "''", "S") & ""
                  
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMultiA(pKey1,pKey2,pKey3)
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    pKey3     = FilterVar(pKey3, "''", "S")

    Call SubMakeSQLStatements("MR",pKey1,pKey2,pKey3)                                   '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""        
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW"))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   

    Call SubHandleError("MR",lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMultiB(pKey1,pKey2,pKey3)
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    pKey3     = FilterVar(pKey3, "''", "S")

    Call SubMakeSQLStatements("MR",pKey1,pKey2,pKey3)                                   '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey1 = ""        
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)

        lgstrData1 = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NM"))
            lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("ALLOW"))
            lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + iDx
            lgstrData1 = lgstrData1 & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey1 = lgStrPrevKey1 + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey1 = ""
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
Sub SubBizSaveSingleCreate()
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate()
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

    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,2)
               Case "R1"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL =     "Select  DUTY_YY, DUTY_MM, DUTY_DD, BAS_AMT, " 
                                   lgStrSQL = lgStrSQL & " YEAR_SAVE_TOT, YEAR_USE, YEAR_CNT, " 
                                   lgStrSQL = lgStrSQL & " YEAR_AMT " 
                                   lgStrSQL = lgStrSQL & " From  HFB020T "
                                   lgStrSQL = lgStrSQL & " WHERE PROV_YYMM = " & pCode1
                                   lgStrSQL = lgStrSQL & "   AND PROV_TYPE = " & FilterVar("Z", "''", "S") & " "
                                   lgStrSQL = lgStrSQL & "   AND YEAR_TYPE = " & FilterVar("1", "''", "S") & " "
                                   lgStrSQL = lgStrSQL & "   AND EMP_NO    = " & pCode2                                   
                        End Select
               Case "R2"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL =     "Select  DUTY_YY, DUTY_MM, DUTY_DD, BAS_AMT, " 
                                   lgStrSQL = lgStrSQL & " MONTH_SAVE, MONTH_USE, MONTH_CNT, MONTH_AMT, MONTH_DUTY_CNT" 
                                   lgStrSQL = lgStrSQL & " From  HFB020T "
                                   lgStrSQL = lgStrSQL & " WHERE PROV_YYMM = " & pCode1
                                   lgStrSQL = lgStrSQL & "   AND PROV_TYPE = " & FilterVar("Z", "''", "S") & " "
                                   lgStrSQL = lgStrSQL & "   AND YEAR_TYPE = " & FilterVar("2", "''", "S") & ""
                                   lgStrSQL = lgStrSQL & "   AND EMP_NO    = " & pCode2
                        End Select
               Case "R3"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL =     "Select  BONUS_TOT_AMT, BONUS_TAX, PROV_TOT_AMT, SUB_TOT_AMT, " 
                                   lgStrSQL = lgStrSQL & " INCOME_TAX, RES_TAX, EMP_INSUR, REAL_PROV_AMT, PROV_DT, " 
                                   lgStrSQL = lgStrSQL & " TAX_AMT, NON_TAX5" 
                                   lgStrSQL = lgStrSQL & " From  HDF070T "
                                   lgStrSQL = lgStrSQL & " WHERE PAY_YYMM = " & pCode1
                                   lgStrSQL = lgStrSQL & "   AND PROV_TYPE = " & FilterVar("Z", "''", "S") & " "
                                   lgStrSQL = lgStrSQL & "   AND EMP_NO    = " & pCode2
                        End Select
           End Select             
        Case "M"           
			if lgSpreadFlg = "1" then        
				iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
			else
				iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1
			end if
           Select Case Mid(pDataType,2,1)
               Case "R"
                        lgStrSQL = "Select top " & iSelCount & " ALLOW_NM, ALLOW" 
                        lgStrSQL = lgStrSQL & " From  HFB030T A, HDA010T B, HFB020T C "
                        lgStrSQL = lgStrSQL & " WHERE C.PROV_YYMM = " & pCode1
                        lgStrSQL = lgStrSQL & "   AND C.PROV_TYPE = " & FilterVar("Z", "''", "S") & " "
                        lgStrSQL = lgStrSQL & "   AND C.YEAR_TYPE = " & pCode3
                        lgStrSQL = lgStrSQL & "   AND C.EMP_NO    = " & pCode2
                        lgStrSQL = lgStrSQL & "   AND A.YEAR_YYMM = C.YEAR_YYMM"
                        lgStrSQL = lgStrSQL & "   AND A.YEAR_TYPE = C.YEAR_TYPE"
                        lgStrSQL = lgStrSQL & "   AND A.EMP_NO    = C.EMP_NO"
                        lgStrSQL = lgStrSQL & "   AND A.ALLOW_CD  = B.ALLOW_CD"
                        lgStrSQL = lgStrSQL & "   AND CODE_TYPE = " & FilterVar("1", "''", "S") & " "
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
       Case "<%=UID_M0001%>"
              With Parent
				if .gSpreadFlg = "1" then              
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					.ggoSpread.Source     = .frm1.vspdData
					.ggoSpread.SSShowData "<%=lgstrData%>"                               '☜ : Display data
					if .topleftOK then
						.DBQueryOk
					else
						.gSpreadFlg = "2"						
						.DBQuery
					end if
                 else
					.ggoSpread.Source     = .frm1.vspdData2
					.ggoSpread.SSShowData "<%=lgstrData1%>"                               '☜ : Display data
					.lgStrPrevKey1    = "<%=lgStrPrevKey1%>"					
					.DBQueryOk
				end if
	         End with
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
