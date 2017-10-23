<%@ LANGUAGE=VBSCript %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
	Call LoadBasisGlobalInf()
	Call loadInfTB19029B("B", "A", "NOCOOKIE", "MB")
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Const C_SHEETMAXROWS_D  = 100
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

 
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query

             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case CStr(UID_M0006)
                                                              '☜: Batch
'            Call SubCreateCommandObject(lgObjComm)
            Call SubBizBatch()
'            Call SubCloseCommandObject(lgObjComm)
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data

        Call SubCreateCommandObject(lgObjComm)
        Call SubBizBatchMulti(arrColVal)                            '☜: Run Batch
        Call SubCloseCommandObject(lgObjComm)


        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next
End Sub

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
' Name : SubBizSave
' Desc : Save Data
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
    Dim arrVal
    Dim strTemp
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    strWhere = " AND C.YYYYMM = " & FilterVar(lgKeyStream(0), "''", "S")   

	    If Trim(lgKeyStream(1)) <> "" Then
		Call CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA ","  BIZ_AREA_CD = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtBizArea = ""
		else   
		  txtBizArea = Trim(Replace(lgF0,Chr(11),""))
		end if
	else 
		txtBizArea = ""

	End If
 
	  
	  
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("800414", vbInformation, "Database Error", "", I_MKSCRIPT)
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""

        iDx       = 1

        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & "0"
            lgstrData = lgstrData & Chr(11) & "0"
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROGRESS_FG"))
            If lgObjRs("BIZ_AREA_CD") = "" Then
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgKeyStream(1))
            Else
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))
            End If
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_CD"))
            Select Case CStr(lgObjRs("MINOR_CD"))
                Case "01"
                    lgstrData = lgstrData & Chr(11) & "A_USP_A5955BA1_07_KO441"		'20080422. 월수계산 ->일수계산 변경 >>AIR
                Case "02"
                    lgstrData = lgstrData & Chr(11) & "A_USP_A5955BA1_08"
                Case "03"
                   lgstrData = lgstrData & Chr(11) & ""'"A_USP_A5955BA1_12"
                Case "04"
                   lgstrData = lgstrData & Chr(11) & ""'"A_USP_A5955BA1_13"
                Case "05"
                    lgstrData = lgstrData & Chr(11) & ""'"A_USP_A5955BA1_14"
                Case "06"
                    lgstrData = lgstrData & Chr(11) & "A_USP_A5955BA1_11"
                Case "07"
                    lgstrData = lgstrData & Chr(11) & "A_USP_A5955BA1_09"
                Case "08"
                    lgstrData = lgstrData & Chr(11) & "A_USP_A5955BA1_17"
                Case "09"
                    lgstrData = lgstrData & Chr(11) & "A_USP_A5955BA1_10"
				Case "10"
                    lgstrData = lgstrData & Chr(11) & "A_USP_A5955BA1_15_KO441"		'20080422. 선급비용관리항목 표시 >>AIR
				Case "11"
                    lgstrData = lgstrData & Chr(11) & "A_USP_A5955BA1_16"
				
            
            
            End Select
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(strWhere)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("NUM_OF_ERROR"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
            lgstrData = lgstrData & Chr(11) & ConvSPChars("")
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If

        Loop
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
    
%>
<SCRIPT LANGUAGE=vbscript>
	With Parent
		.frm1.txtBizArea.value = "<%=ConvSPChars(txtBizArea)%>"
	END With                                   
</SCRIPT>
<%    

End Sub

'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti(arrColVal)
    on error resume next
    Dim IntRetCD
    Dim strMsg_cd, strMsg_text
    Dim strYYYYMMDD

    strYYYYMMDD =Trim(lgKeyStream(2))   
    strBizAreaCd = UCase(Trim(lgKeyStream(1)))
	
    if arrColVal(0) = "R" then
    ' 실행 선택시 
        With lgObjComm
            .CommandText = arrColVal(1)
            .CommandType = adCmdStoredProc
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@yyyymmdd"     ,adVarWChar,adParamInput,Len(Trim(strYYYYMMDD)), strYYYYMMDD)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@biz_area"     ,adVarWChar,adParamInput,Len(Trim(strBizAreaCd)), strBizAreaCd)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cancel_yn"     ,adVarWChar,adParamInput,Len("1"), "1")            
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id"     ,adVarWChar,adParamInput,Len(Trim(gUsrID)), gUsrID)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarWChar,adParamOutput,6)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@num_of_error"   ,adSmallInt ,adParamOutput,2)

            lgObjComm.Execute ,, adExecuteNoRecords
        End With
    Else
    ' 취소 선택시 
        With lgObjComm
            .CommandText = arrColVal(1)
            .CommandType = adCmdStoredProc
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@yyyymmdd"     ,adVarWChar,adParamInput,Len(Trim(strYYYYMMDD)), strYYYYMMDD)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@biz_area"     ,adVarWChar,adParamInput,Len(Trim(strBizAreaCd)), strBizAreaCd)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cancel_yn"     ,adVarWChar,adParamInput,Len("2"), "2")            
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id"     ,adVarWChar,adParamInput,Len(Trim(gUsrID)), gUsrID)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarWChar,adParamOutput,6)
            lgObjComm.Parameters.Append lgObjComm.CreateParameter("@num_of_error"   ,adSmallInt ,adParamOutput,2)

            lgObjComm.Execute ,, adExecuteNoRecords
        End With
    End If

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        if  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value    
            if Trim(strMsg_cd) = ""       then
				Call DisplayMsgBox("800407",vbInformation,"","",I_MKSCRIPT)
            else
				Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
			end if
%>
<Script Language="VBScript">
    Parent.ExeReflectNo
</Script>
<%
            Response.end
        end if
    Else    
        lgErrorStatus     = "YES"                                                         '☜: Set error status
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End if
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
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
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

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
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

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1
           Select Case Mid(pDataType,2,1)
               Case "C"
               Case "D"
               Case "R"
                       lgStrSQL = " SELECT ISNULL(C.BIZ_AREA_CD," & FilterVar(lgKeyStream(1), "''", "S") & ") as BIZ_AREA_CD, A.MINOR_CD, A.MINOR_NM, UPPER(ISNULL(C.PROGRESS_FG," & FilterVar("N", "''", "S") & " )) AS PROGRESS_FG1, "
                       lgStrSQL = lgStrSQL & " case when C.JOB_FG = " & FilterVar("1", "''", "S") & "  and C.PROGRESS_FG = " & FilterVar("Y", "''", "S") & "  then " & FilterVar("Y", "''", "S") & "  else " & FilterVar("N", "''", "S") & "  end AS PROGRESS_FG, "                                                  
                       lgStrSQL = lgStrSQL & " ISNULL(C.NUM_OF_ERROR," & FilterVar("0", "''", "S") & " ) AS NUM_OF_ERROR, "
                       lgStrSQL = lgStrSQL & " case when a.minor_cd in (" & FilterVar("10", "''", "S") & " , " & FilterVar("11", "''", "S") & " ) then " & FilterVar("Y", "''", "S") & "  else (select use_yn from a_monthly_base where reg_cd = a.minor_cd) end "
                       lgStrSQL = lgStrSQL & " FROM B_MINOR A, A_JOB_RESULT C "
                       lgStrSQL = lgStrSQL & " WHERE A.MINOR_CD *= C.JOB_CD "
                       lgStrSQL = lgStrSQL & " AND A.MAJOR_CD = " & FilterVar("A1032", "''", "S") & "  "
                       lgStrSQL = lgStrSQL & pCode
                       lgStrSQL = lgStrSQL & " AND C.BIZ_AREA_CD = " & FilterVar(lgKeyStream(1), "''", "S")
                       lgStrSQL = lgStrSQL & "ORDER BY A.MINOR_CD "
               Case "U"
           End Select
     
       
    End Select
    'Response.write lgStrSQL
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub

%>

<Script Language="VBScript">

    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .DBQueryOk
	          End with
          End If
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "<%=UID_M0006%>"                                                         '☜ : Batch
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.ExeReflectOk
          Else
             Parent.ExeReflectNo
          End If
    End Select
</Script>
