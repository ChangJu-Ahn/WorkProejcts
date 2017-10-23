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
	Dim lgSpreadFlg	
	Const C_SHEETMAXROWS_D = 100

    Dim lgLngMaxRow1
	Dim txtCheck	' 회수 구간을 체크하기 위한 변수 

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
	
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgLngMaxRow1      = Request("txtMaxRows1")                                       '☜: Read Operation Mode (CRUD)
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

    Dim txtAllow_cd

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    txtAllow_cd = FilterVar(lgKeyStream(0), "''", "S")
    
    Call SubMakeSQLStatements("SR",txtAllow_cd,"X",C_EQ)                             '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then               'If data not exists

          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '☜: No data is found. 
          Call SetErrorStatus()
    Else    ' Single Query 정상일경우 
%>
<Script Language=vbscript>
        With Parent.Frm1
            .txtcrt_strt_mm.Value  = "<%=ConvSPChars(lgObjRs("crt_strt_mm"))%>"
            .txtcrt_strt_dd.Value  = "<%=ConvSPChars(lgObjRs("crt_strt_dd"))%>"
            .txtcrt_end_mm.Value   = "<%=ConvSPChars(lgObjRs("crt_end_mm"))%>"
            .txtcrt_end_dd.Value   = "<%=ConvSPChars(lgObjRs("crt_end_dd"))%>"
            .txtday_calcu.Value    = "<%=ConvSPChars(lgObjRs("day_calcu"))%>"
            .txtcalcu_bas_dd.Value = "<%=ConvSPChars(lgObjRs("calcu_bas_dd"))%>"
            
        End With          
</Script>       
<%     
    End If    ' Single Query 정상일경우 

	if lgSpreadFlg = "1" then
		Call SubBizQueryMulti(txtAllow_cd)
	else
		Call SubBizQueryMulti1(txtAllow_cd)
	end if

End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Call CommonQueryRs(" count(*) "," HDA010T "," code_type =" & FilterVar("0", "''", "S") & "  and allow_cd = " & FilterVar(lgKeyStream(0), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If  Replace(lgF0, Chr(11), "") = "0" then
        Call SubBizSaveSingleCreate()
    Else
        Call SubBizSaveSingleUpdate()
    End if

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
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pKey1)
    Dim iDx
    Dim iLoopMax
	
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MC",pKey1,"X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cd"))
            lgstrData = lgstrData & Chr(11) & ""

            ' 근태코드명 조회 
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("dilig_nm")))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("dilig_cnt"),    ggAmtOfMoney.DecPoint,0)
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

    Call SubHandleError("MC",lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub

Sub SubBizQueryMulti1(pKey1)
    Dim iDx
    Dim iLoopMax
	
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("NC",pKey1,"X",C_EQ)                                   '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey1 = ""
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)    
        lgstrData1 = ""
        txtCheck   = ""
        iDx        = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("dilig_strt"), 3,0)
            lgstrData1 = lgstrData1 & Chr(11) & "~"
            lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("dilig_end"), 3,0)

            lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("minus_rate"), 3,0)
            lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("minus_amt"),  ggAmtOfMoney.DecPoint,0)
            lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + iDx
            lgstrData1 = lgstrData1 & Chr(11) & Chr(12)
			
			txtCheck   = CheckPeriod(Cint(ConvSPChars(lgObjRs("dilig_strt"))), Cint(ConvSPChars(lgObjRs("dilig_end"))), txtCheck)
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
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    


'============================================================================================================
' Name : CheckPeriod
' Desc : 근태 회수에 대한 체크 스트림 작성 
'============================================================================================================
Function CheckPeriod(Scnt, Ecnt, str)
	DIM I
	DIM temp
	
	temp = str
	for i = len(str) + 1 to Scnt -1
		temp = temp & "0"
	next
	
	for i = Scnt to Ecnt -1
		temp = temp & "1"
	next

	CheckPeriod = temp	
end function

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim arrRowVal1
    Dim arrColVal1
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
'** Multi-1
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

'** Multi-2
	arrRowVal1 = Split(Request("txtSpread1"), gRowSep)                               '☜: Split Row    data
    For iDx = 1 To lgLngMaxRow1
        arrColVal1 = Split(arrRowVal1(iDx-1), gColSep)                                 '☜: Split Column data
        Select Case arrColVal1(0)
            Case "C"
                    Call SubBizSaveMultiCreate1(arrColVal1)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate1(arrColVal1)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete1(arrColVal1)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal1(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    Dim txtGlNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call CommonQueryRs(" allow_cd "," HDA010T "," code_type =" & FilterVar("0", "''", "S") & "  and allow_cd = " & FilterVar(lgKeyStream(0), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    if  Replace(lgF0, Chr(11), "") <> "X" then
        exit sub
    end if

    lgStrSQL = "INSERT INTO HDA010T("
    lgStrSQL = lgStrSQL & " code_type,"
    lgStrSQL = lgStrSQL & " allow_cd,"
    lgStrSQL = lgStrSQL & " crt_strt_mm,"
    lgStrSQL = lgStrSQL & " crt_strt_dd,"
    lgStrSQL = lgStrSQL & " crt_end_mm,"
    lgStrSQL = lgStrSQL & " crt_end_dd,"
    lgStrSQL = lgStrSQL & " day_calcu,"
    lgStrSQL = lgStrSQL & " calcu_bas_dd,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & "" & FilterVar("0", "''", "S") & " ,"                                        ' code_type(PK)
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","     ' allow_cd(PK)
    lgStrSQL = lgStrSQL & UniConvNum(Request("txtcrt_strt_mm"),0)  & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtcrt_strt_dd"), "''", "S") & ","
    lgStrSQL = lgStrSQL & UniConvNum(Request("txtcrt_end_mm"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtcrt_end_dd"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtday_calcu"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtcalcu_bas_dd"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDA010T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " crt_strt_mm = " &  UniConvNum(Request("txtcrt_strt_mm"),0) & ","
    lgStrSQL = lgStrSQL & " crt_strt_dd = " & FilterVar(Request("txtcrt_strt_dd"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " crt_end_mm = " &  UniConvNum(Request("txtcrt_end_mm"),0) & ","
    lgStrSQL = lgStrSQL & " crt_end_dd = " & FilterVar(Request("txtcrt_end_dd"), "''","S") & ","
    lgStrSQL = lgStrSQL & " day_calcu = " & FilterVar(Request("txtday_calcu"), "''","S") & ","
    lgStrSQL = lgStrSQL & " calcu_bas_dd = " & FilterVar(Request("txtcalcu_bas_dd"), "''","S") & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE code_type = " & FilterVar("0", "''", "S") & " "
    lgStrSQL = lgStrSQL & "   AND allow_cd = " & FilterVar(lgKeyStream(0), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HDA100T("
    lgStrSQL = lgStrSQL & " allow_cd," 
    lgStrSQL = lgStrSQL & " dilig_cd," 
    lgStrSQL = lgStrSQL & " dilig_cnt," 
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO," 
    lgStrSQL = lgStrSQL & " UPDT_DT )" 
    lgStrSQL = lgStrSQL & " VALUES( "
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(4)),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

Sub SubBizSaveMultiCreate1(arrColVal1)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HEA010T("
    lgStrSQL = lgStrSQL & " dilig_strt," 
    lgStrSQL = lgStrSQL & " dilig_end," 
    lgStrSQL = lgStrSQL & " minus_rate,"
    lgStrSQL = lgStrSQL & " minus_amt,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO," 
    lgStrSQL = lgStrSQL & " UPDT_DT )" 
    lgStrSQL = lgStrSQL & " VALUES( "
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal1(2)),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal1(3)),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal1(4)),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal1(5)),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"

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

    lgStrSQL = "UPDATE  HDA100T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " dilig_cnt = " &  UNIConvNum(arrColVal(4),0) & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

Sub SubBizSaveMultiUpdate1(arrColVal1)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HEA010T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " dilig_end = " &  UNIConvNum(arrColVal1(3),0) & ","
    lgStrSQL = lgStrSQL & " minus_rate = " &  UNIConvNum(arrColVal1(4),0) & ","
    lgStrSQL = lgStrSQL & " minus_amt = " &  UNIConvNum(arrColVal1(5),0) & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE dilig_strt = " &  UNIConvNum(arrColVal1(2),0)

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

    lgStrSQL = "DELETE  HDA100T"
    lgStrSQL = lgStrSQL & " WHERE allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub

Sub SubBizSaveMultiDelete1(arrColVal1)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HEA010T"
    lgStrSQL = lgStrSQL & " WHERE dilig_strt = " & FilterVar(UCase(arrColVal1(2)), "''", "S")

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
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL = "Select crt_strt_mm, crt_strt_dd, crt_end_mm, crt_end_dd, day_calcu, calcu_bas_dd " 
                                   lgStrSQL = lgStrSQL & " From  HDA010T "
                                   lgStrSQL = lgStrSQL & " WHERE allow_cd " & pComp & pCode
                                   lgStrSQL = lgStrSQL & "   AND code_type = " & FilterVar("0", "''", "S") & "  "
                        End Select
           End Select             
        Case "M"           '                  0               1                 2              3                4                5
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           Select Case Mid(pDataType,2,1)
               Case "C"
                       lgStrSQL = "Select top " & iSelCount & " ALLOW_CD, DILIG_CD, DILIG_CNT, dbo.ufn_H_GetCodeName(" & FilterVar("HCA010T", "''", "S") & ", DILIG_CD, null)  DILIG_NM " 
                       lgStrSQL = lgStrSQL & " From  HDA100T "
                       lgStrSQL = lgStrSQL & " WHERE allow_cd " & pComp & pCode
           End Select             
        Case "N"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1        
           Select Case Mid(pDataType,2,1)
               Case "C"
                    lgStrSQL = "Select  top " & iSelCount & " DILIG_STRT, DILIG_END, MINUS_RATE, MINUS_AMT  " 
                    lgStrSQL = lgStrSQL & " From  HEA010T "
                    lgStrSQL = lgStrSQL & " Order by DILIG_STRT "
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
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .frm1.txtPeriod.value   = "<%=txtCheck%>"              
				if .gSpreadFlg = "1" then              
					.ggoSpread.Source     = .frm1.vspdData
					.ggoSpread.SSShowData "<%=lgstrData%>"                               '☜ : Display data
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					if .topleftOK then
						.DBQueryOk
					else
						.gSpreadFlg = "2"
						.DBQuery
					end if
                else
					.lgStrPrevKey1    = "<%=lgStrPrevKey1%>"                
					.ggoSpread.Source     = .frm1.vspdData1
					.ggoSpread.SSShowData "<%=lgstrData1%>"
					.DBQueryOk        
                end if					
	         End with
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
