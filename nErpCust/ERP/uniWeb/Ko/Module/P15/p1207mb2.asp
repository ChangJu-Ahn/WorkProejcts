<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1207mb2.asp
'*  4. Program Name         : Manage Standard Manufacturing Instruction
'*  5. Program Desc         : use HR source (ADO SAVE)
'*  6. Modified date(First) : 2002/03/19
'*  7. Modified date(Last)  : 2002/11/21
'*  8. Modifier (First)     : Hong Chang Ho
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim lgGetSvrDateTime, lgGetSvrDate, lgFlgDataExists
Dim arrStdVal(4)

lgGetSvrDateTime = GetSvrDateTime
lgGetSvrDate = GetSvrDate
lgFlgDataExists = False

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

arrStdVal(0) = Request("txtStdMode")
arrStdVal(1) = Trim(UCase(Request("txtStdInstrCd1")))
arrStdVal(2) = Request("txtStdInstrNm1")
arrStdVal(3) = UNIConvDate(Request("txtValidFromDt"))
arrStdVal(4) = UNIConvDate(Request("txtValidToDt"))

'------ Developer Coding part (Start ) ------------------------------------------------------------------


'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

Select Case arrStdVal(0)
	Case "C"
		Call SubBizSaveSelect(arrStdVal)
		If lgFlgDataExists = True Then
			Call SubCloseDB(lgObjConn)
				
			Response.End
		End If
		Call SubBizSaveCreate(arrStdVal)

	Case "U"	
		Call SubBizSaveUpdate(arrStdVal)

	Case "D"	
		Call SubBizDelete(arrStdVal)
End Select
	
If arrStdVal(0) <> "D" Then
	Select Case lgOpModeCRUD
        Case CStr(UID_M0002)                                                         '☜: Save,Update
			Call SubBizSaveMulti()
	End Select
End If
    
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
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus  = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next
'       Response.Write "IDX=" & IDX
End Sub    


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "INSERT INTO P_MFG_INSTRUCTION_SET("
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_CD," 
    lgStrSQL = lgStrSQL & " SEQ," 
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_DTL_CD," 
    lgStrSQL = lgStrSQL & " INSRT_USER_ID,"
    lgStrSQL = lgStrSQL & " INSRT_DT,"   
    lgStrSQL = lgStrSQL & " UPDT_USER_ID," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(1), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
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

    lgStrSQL = "UPDATE P_MFG_INSTRUCTION_SET"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_DTL_CD = " & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE MFG_INSTRUCTION_CD = " & FilterVar(arrColVal(1), "''", "S")
    lgStrSQL = lgStrSQL & " AND SEQ = " & FilterVar(arrColVal(2), "''", "S")
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
    lgStrSQL = "DELETE P_MFG_INSTRUCTION_SET"
    lgStrSQL = lgStrSQL & " WHERE MFG_INSTRUCTION_CD = " & FilterVar(arrColVal(1), "''", "S")
    lgStrSQL = lgStrSQL & " AND SEQ = " & FilterVar(arrColVal(2), "''", "S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub


Sub SubBizSaveSelect(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "SELECT MFG_INSTRUCTION_CD FROM P_MFG_INSTRUCTION_HEADER WHERE MFG_INSTRUCTION_CD = "
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(1), "''", "S")
    
	If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "Parent.frm1.txtStdInstrCd1.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Call DisplayMsgBox("181420", VBInformation, arrColVal(1), "", I_MKSCRIPT)
		lgFlgDataExists = True
	End If
	
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "INSERT INTO P_MFG_INSTRUCTION_HEADER("
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_CD," 
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_NM," 
    lgStrSQL = lgStrSQL & " VALID_FROM_DT," 
    lgStrSQL = lgStrSQL & " VALID_TO_DT,"
	lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT,"   
    lgStrSQL = lgStrSQL & " UPDT_USER_ID," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(1), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(3)), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(4)), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Call SubHandleError("MC", lgObjConn , lgObjRs, Err)
End Sub

'============================================================================================================
' Name : SubBizSaveUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "UPDATE P_MFG_INSTRUCTION_HEADER"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_NM = " & FilterVar(arrColVal(2), "''", "S") & ","
    lgStrSQL = lgStrSQL & " VALID_FROM_DT = " & FilterVar(UNIConvDate(arrColVal(3)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " VALID_TO_DT = " & FilterVar(UNIConvDate(arrColVal(4)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE MFG_INSTRUCTION_CD = " & FilterVar(arrColVal(1), "''", "S")
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

Sub SubBizDelete(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "DELETE P_MFG_INSTRUCTION_HEADER"
    lgStrSQL = lgStrSQL & " WHERE MFG_INSTRUCTION_CD = " & FilterVar(arrColVal(1), "''", "S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    lgStrSQL = "DELETE P_MFG_INSTRUCTION_SET"
    lgStrSQL = lgStrSQL & " WHERE MFG_INSTRUCTION_CD = " & FilterVar(arrColVal(1), "''", "S")
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
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
                    Call Displaymsgbox("122918", vbInformation, "", "", I_MKSCRIPT)     'Can not Update (Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call Displaymsgbox("122918", vbInformation, "", "", I_MKSCRIPT)     'Can not Update (Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

Response.Write "<Script Language = VBScript>" & vbCrLf

    Select Case CStr(lgOpModeCRUD)
       Case CStr(UID_M0001)
          If Trim(lgErrorStatus) = "NO" Then
              Response.Write "With Parent" & vbCrLf
                Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
                Response.Write ".ggoSpread.SSShowDataByClip """ & lgstrData & """" & vbCrLf
                Response.Write ".lgStrPrevKey = """ & lgStrPrevKey & """" & vbCrLf
                Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
                Response.Write ".DBQueryOk()" & vbCrLf
             Response.Write "End with" & vbCrLf
          End If   
       Case CStr(UID_M0002)
          If Trim(lgErrorStatus) = "NO" Then
             Response.Write "Parent.DBSaveOk()" & vbCrLf
          End If   
       Case CStr(UID_M0003)
          If Trim(lgErrorStatus) = "NO" Then
             Response.Write "Parent.DbDeleteOk()" & vbCrLf
          End If   
    End Select    
       
Response.Write "</Script>" & vbCrLf
%>
