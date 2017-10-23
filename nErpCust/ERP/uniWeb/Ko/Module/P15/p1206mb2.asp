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
'*  3. Program ID           : p1206mb2.asp
'*  4. Program Name         : manage Manufacturing Instruction
'*  5. Program Desc         : use HR source (ADO SAVE)
'*  6. Modified date(First) : 2002/03/19
'*  7. Modified date(Last)  : 2002/11/21
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
lgKeyStream       = Split(Request("txtKeyStream"), gColSep)
lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
'Multi SpreadSheet

lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
         'Call SubBizQuery()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSave()
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)                                                         '☜: Delete
         'Call SubBizDelete()
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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
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
        arrColVal = Split(arrRowVal(iDx - 1), gColSep)                                 '☜: Split Column data
    
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
					Call SubBizDelSelect(arrColVal)
						If lgFlgDataExists = True Then
							Call SubCloseDB(lgObjConn)
							Response.End
						End If
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
    Dim iclose_dt
    Dim strPay_dt
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "INSERT INTO P_MFG_INSTRUCTION_DETAIL("
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_DTL_CD," 
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_DTL_DESC," 
    lgStrSQL = lgStrSQL & " VALID_START_DT," 
    lgStrSQL = lgStrSQL & " VALID_END_DT,"
    lgStrSQL = lgStrSQL & " INSRT_USER_ID,"
    lgStrSQL = lgStrSQL & " INSRT_DT,"   
    lgStrSQL = lgStrSQL & " UPDT_USER_ID," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(4)), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(5)), "''", "S") & ","
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

    lgStrSQL = "UPDATE  P_MFG_INSTRUCTION_DETAIL"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " MFG_INSTRUCTION_DTL_DESC = " & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & " VALID_END_DT = " & FilterVar(UNIConvDate(arrColVal(4)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE MFG_INSTRUCTION_DTL_CD = " & FilterVar(arrColVal(2), "NULL", "S")

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
    lgStrSQL = "DELETE  P_MFG_INSTRUCTION_DETAIL"
    lgStrSQL = lgStrSQL & " WHERE MFG_INSTRUCTION_DTL_CD = " & FilterVar(arrColVal(2), "NULL", "S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub


Sub SubBizDelSelect(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	lgStrSQL = ""
    lgStrSQL = "SELECT MFG_INSTRUCTION_DTL_CD FROM P_MFG_INSTRUCTION_SET WHERE MFG_INSTRUCTION_DTL_CD = "
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")
    
	If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then
		
		Call DisplayMsgBox("181417", VBInformation, arrColVal(2), "", I_MKSCRIPT)
		lgFlgDataExists = True
		
	End If
	
	If lgFlgDataExists = False Then
		lgStrSQL = ""
		lgStrSQL = "SELECT MFG_INSTRUCTION_DTL_CD FROM P_MFG_INSTRUCTION_BY_ROUTING WHERE MFG_INSTRUCTION_DTL_CD = "
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")
    
		If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then
			
			Call DisplayMsgBox("181418", VBInformation, arrColVal(2), "", I_MKSCRIPT)
			lgFlgDataExists = True
		End If
	End If
	
End Sub	    

Sub SubBizSaveSelect(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	lgStrSQL = ""
    lgStrSQL = "SELECT MFG_INSTRUCTION_DTL_CD FROM P_MFG_INSTRUCTION_DETAIL WHERE MFG_INSTRUCTION_DTL_CD = "
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")
    
	If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then
		
		Call DisplayMsgBox("181419", VBInformation, arrColVal(2), "", I_MKSCRIPT)
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
