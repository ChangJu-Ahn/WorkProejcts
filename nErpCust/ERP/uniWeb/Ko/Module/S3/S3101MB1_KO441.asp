<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required %>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
Const C_IG_CMD = 0
Const C_IG_PROFORMA_NO = 1
Const C_IG_BP_CD = 2
Const C_IG_TEL_NO = 3
Const C_IG_CUST_USER = 4
Const C_IG_PROFORMA_DT = 5
Const C_IG_DOCUMENT = 6
Const C_IG_AMT = 7
Const C_IG_CONFIRM_YN = 8
Const C_IG_BILL_YN = 9
Const C_IG_TAX_DT = 10
Const C_IG_REMARK = 11
Const C_IG_ROW = 12

	Dim lgStrPrevKey
	Dim orgChangID
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    call LoadBasisGlobalInf()
    
    lgSvrDateTime = GetSvrDateTime    
    
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")   
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere, strFg

   Const C_BP_CD = 0
   Const C_TEL_NO = 1
   Const C_PRO_FR_DT = 2
   Const C_PRO_TO_DT = 3
   Const C_CONFIRM_FLAG = 4
   Const C_REQ_FR_DT = 5
   Const C_REQ_TO_DT = 6
   Const C_AR_FLAG = 7
   Const C_DOC = 8
  
     On Error Resume Next    
    Err.Clear                                                               '¢Ð: Clear Error status

	
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '¢Ð : Make sql statements

		If Trim(lgKeyStream(C_BP_CD)) <> "" then
				strWhere = " AND a.B_BIZ_PARTNER=" & FilterVar(lgKeyStream(C_BP_CD),"''","S")
		End If
		If Trim(lgKeyStream(C_TEL_NO)) <> "" then
				strWhere = strWhere & " AND a.TEL_NO like " & FilterVar(lgKeyStream(C_TEL_NO)&"%","'%'","S")
		End If
		If Trim(lgKeyStream(C_PRO_FR_DT)) <> "" then
				strWhere = strWhere & " AND a.PROFORMA_DT >=" & FilterVar(UNIConvDate(lgKeyStream(C_PRO_FR_DT)),"''","S")
		End If
		If Trim(lgKeyStream(C_PRO_TO_DT)) <> "" then
				strWhere = strWhere & " AND a.PROFORMA_DT <=" & FilterVar(UNIConvDate(lgKeyStream(C_PRO_TO_DT)),"''","S")
		End If
		If Trim(lgKeyStream(C_CONFIRM_FLAG)) <> "" then
				strWhere = strWhere & " AND a.CONFIRM_YN =" & FilterVar(lgKeyStream(C_CONFIRM_FLAG),"''","S")
		End If
		If Trim(lgKeyStream(C_REQ_FR_DT)) <> "" then
				strWhere = strWhere & " AND a.TAX_DT >=" & FilterVar(UNIConvDate(lgKeyStream(C_REQ_FR_DT)),"''","S")
		End If
		If Trim(lgKeyStream(C_REQ_TO_DT)) <> "" then
				strWhere = strWhere & " AND a.TAX_DT <=" & FilterVar(UNIConvDate(lgKeyStream(C_REQ_TO_DT)),"''","S")
		End If
		If Trim(lgKeyStream(C_AR_FLAG)) <> "" then
				strWhere = strWhere & " AND a.BILL_YN =" & FilterVar(lgKeyStream(C_AR_FLAG),"''","S")
		End If
		If Trim(lgKeyStream(C_DOC)) <> "" then
				strWhere = strWhere & " AND a.DOCUMENT like" & FilterVar(lgKeyStream(C_DOC)&"%","'%'","S")
		End If
'response.write strWhere
'response.end
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL&strWhere,"X","X") = False Then
       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROFORMA_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("B_BIZ_PARTNER"))
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEL_NO"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CUST_USER"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(ConvSPChars(lgObjRs("PROFORMA_DT")))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOCUMENT"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONFIRM_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BILL_YN"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(ConvSPChars(lgObjRs("TAX_DT")))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))
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

      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
         ObjectContext.SetAbort
      End If
            
      Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
      Call SubCloseRs(lgObjRs)    

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(C_IG_CMD)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(C_IG_ROW) & gColSep
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
    Err.Clear                                                                        '¢Ð: Clear Error status
  
	
	If CommonQueryRs(" BP_CD "," B_BIZ_PARTNER ", " BP_CD=" & FilterVar(arrColVal(C_IG_BP_CD), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
 		If lgF0="X" Then
			Call SubHandleError("M1",lgObjConn,lgObjRs,Err)
			Exit Sub
		End If
	End If
   	
		lgStrSQL = "INSERT INTO S_PROFORMA_KO441(PROFORMA_NO,B_BIZ_PARTNER,TEL_NO,CUST_USER,PROFORMA_DT,DOCUMENT,AMT,CONFIRM_YN,BILL_YN,TAX_DT,REMARK,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "
		lgStrSQL = lgStrSQL & " VALUES("     
		lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_IG_PROFORMA_NO)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_IG_BP_CD)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_TEL_NO)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_CUST_USER)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(C_IG_PROFORMA_DT)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_DOCUMENT)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(UniConvNum(arrColVal(C_IG_AMT),0), "''", "S")     & ","
		If arrColVal(C_IG_CONFIRM_YN)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_BILL_YN)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If
		lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(C_IG_TAX_DT)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_REMARK)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")		   & "," 
		lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   
		lgStrSQL = lgStrSQL & ")"  

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
 
 	If CommonQueryRs(" BP_CD "," B_BIZ_PARTNER ", " BP_CD=" & FilterVar(arrColVal(C_IG_BP_CD), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  Then
 		If lgF0="X" Then
			Call SubHandleError("M1",lgObjConn,lgObjRs,Err)
			Exit Sub
		End If
	End If

		lgStrSQL = "UPDATE S_PROFORMA_KO441 SET "
		lgStrSQL = lgStrSQL & " B_BIZ_PARTNER	=" & FilterVar(UCase(arrColVal(C_IG_BP_CD)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " TEL_NO				=" & FilterVar(Trim(arrColVal(C_IG_TEL_NO)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " CUST_USER			=" & FilterVar(Trim(arrColVal(C_IG_CUST_USER)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " PROFORMA_DT		=" & FilterVar(UniConvDate(arrColVal(C_IG_PROFORMA_DT)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " DOCUMENT			=" & FilterVar(Trim(arrColVal(C_IG_DOCUMENT)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " AMT						=" & FilterVar(UniConvNum(arrColVal(C_IG_AMT),0), "''", "S")     & ","
		If arrColVal(C_IG_CONFIRM_YN)="1" Then
			lgStrSQL = lgStrSQL & " CONFIRM_YN	=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CONFIRM_YN	=" & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_BILL_YN)="1" Then
			lgStrSQL = lgStrSQL & " BILL_YN			=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " BILL_YN			=" & FilterVar("N", "''", "S")     & ","
		End If
		lgStrSQL = lgStrSQL & " TAX_DT				=" & FilterVar(UniConvDate(arrColVal(C_IG_TAX_DT)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " REMARK				=" & FilterVar(Trim(arrColVal(C_IG_REMARK)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID	=" & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & " UPDT_DT				=" & FilterVar(lgSvrDateTime, "''", "S")   
		lgStrSQL = lgStrSQL & " WHERE PROFORMA_NO=" & FilterVar(UCase(arrColVal(C_IG_PROFORMA_NO)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  S_PROFORMA_KO441"
		lgStrSQL = lgStrSQL & " WHERE PROFORMA_NO=" & FilterVar(UCase(arrColVal(C_IG_PROFORMA_NO)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
     Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
      
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
           
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & " a.PROFORMA_NO, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.B_BIZ_PARTNER, " & vbCrLf
                       lgStrSQL = lgStrSQL & " b.BP_NM, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.TEL_NO, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.CUST_USER, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.PROFORMA_DT, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.DOCUMENT, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.AMT, " & vbCrLf
                       lgStrSQL = lgStrSQL & " c.USR_NM, " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CONFIRM_YN = 'Y' then 1 else 0 end CONFIRM_YN, " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.BILL_YN = 'Y' then 1 else 0 end BILL_YN, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.TAX_DT, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.REMARK " & vbCrLf
                       lgStrSQL = lgStrSQL & " from S_PROFORMA_KO441 a " & vbCrLf
                       lgStrSQL = lgStrSQL & " left outer join b_biz_partner b on (a.B_BIZ_PARTNER=b.BP_CD) " & vbCrLf
                       lgStrSQL = lgStrSQL & " left outer join Z_USR_MAST_REC c on (a.INSRT_USER_ID=c.usr_id) " & vbCrLf
                       lgStrSQL = lgStrSQL & " where 1=1 " & vbCrLf

          End Select
    End Select
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
    Response.Write "<BR> Commit Event occur"
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
    Response.Write "<BR> Abort Event occur"
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
        Case "MS"
                 Call DisplayMsgBox("800486", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)        
                 ObjectContext.SetAbort
                 Call SetErrorStatus
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
        Case "M1"
                 Call DisplayMsgBox("173132", vbInformation, "°í°´ÄÚµå", "", I_MKSCRIPT)     '
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "M2"
                 Call DisplayMsgBox("202404", vbInformation, "", "", I_MKSCRIPT)     '
                 ObjectContext.SetAbort
                 Call SetErrorStatus
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
