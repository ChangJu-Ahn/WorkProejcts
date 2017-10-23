<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1208mb3.asp
'*  4. Program Name         : Entry Manufacturing Instruction
'*  5. Program Desc         : use HR source (ADO SAVE)
'*  6. Modified date(First) : 2002/03/25
'*  7. Modified date(Last)  : 2002/11/20
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim lgGetSvrDateTime, lgGetSvrDate, lgFlgDataExists

lgGetSvrDateTime = GetSvrDateTime
lgGetSvrDate = GetSvrDate

'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
'Multi SpreadSheet

lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	                                                         
Call SubBizSaveMulti()															 '☜: Save,Update
       
Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

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
					Call SubBizSaveSelect(arrColVal)
						If lgFlgDataExists = True Then
							Call SubCloseDB(lgObjConn)
							Response.End
						End If
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
       Response.Write "IDX=" & IDX
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    Dim iclose_dt
    Dim strPay_dt
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "INSERT INTO P_MFG_INSTRUCTION_BY_ROUTING ("
    lgStrSQL = lgStrSQL & " PLANT_CD," 
    lgStrSQL = lgStrSQL & " ITEM_CD," 
    lgStrSQL = lgStrSQL & " ROUT_NO," 
    lgStrSQL = lgStrSQL & " OPR_NO,"
    lgStrSQL = lgStrSQL & " SEQ," 
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_DTL_CD," 
    lgStrSQL = lgStrSQL & " INSRT_USER_ID,"
    lgStrSQL = lgStrSQL & " INSRT_DT,"   
    lgStrSQL = lgStrSQL & " UPDT_USER_ID," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    Response.Write lgStrSQL
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "UPDATE  P_MFG_INSTRUCTION_BY_ROUTING"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_DTL_CD = " & FilterVar(arrColVal(7),"NULL", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(arrColVal(2), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " ITEM_CD = " & FilterVar(arrColVal(3), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " ROUT_NO = " & FilterVar(arrColVal(4), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " OPR_NO = " & FilterVar(arrColVal(5), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " SEQ = " & FilterVar(arrColVal(6), "NULL", "S") 

'    RESPONSE.WRITE  lgStrSQL
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  P_MFG_INSTRUCTION_BY_ROUTING"
    lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(arrColVal(2), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " ITEM_CD = " & FilterVar(arrColVal(3), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " ROUT_NO = " & FilterVar(arrColVal(4), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " OPR_NO = " & FilterVar(arrColVal(5), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " SEQ = " & FilterVar(arrColVal(6), "NULL", "S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub

Sub SubBizSaveSelect(arrColVal)
	
	Dim strTempMsg
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	lgStrSQL = ""
    lgStrSQL = "SELECT MFG_INSTRUCTION_DTL_CD,SEQ FROM P_MFG_INSTRUCTION_BY_ROUTING "
    lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(arrColVal(2), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " ITEM_CD = " & FilterVar(arrColVal(3), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " ROUT_NO = " & FilterVar(arrColVal(4), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " OPR_NO = " & FilterVar(arrColVal(5), "NULL", "S") & " AND "
    lgStrSQL = lgStrSQL & " SEQ = " & FilterVar(arrColVal(6), "NULL", "S")
    
	If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then
		
		strTempMsg = "공정:"
		strTempMsg = strTempMsg & arrColVal(5)
		strTempMsg =  strTempMsg & ",순서:"
		strTempMsg = strTempMsg & arrColVal(6)
		
		Call DisplayMsgBox("181420", VBInformation, strTempMsg , "" , I_MKSCRIPT)
		lgFlgDataExists = True
		
	End If
		
End Sub	 


'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
                    Call Displaymsgbox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
					   Call Displaymsgbox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call Displaymsgbox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call Displaymsgbox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

Response.Write "<Script Language = VBScript>" & vbCrLf
    If Trim(lgErrorStatus) = "NO" Then
       Response.Write "Parent.DBSaveOk" & vbCrLf
    End If   
Response.Write "</Script>" & vbCrLf
%>
