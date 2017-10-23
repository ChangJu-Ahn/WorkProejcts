<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%

	Dim lgStrPrevKey ,lgStrPrevKey1 
	Dim lgLngMaxRow1
    Dim lgSpreadFlg
    
	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
   
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

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then               'If data not exists
%>
<Script Language=vbscript>
        With Parent.Frm1
            ' 연차지급방법 조회-

            .txtcrt_strt_yy.Value = "<%=ConvSPChars(lgObjRs("crt_strt_yy"))%>"
            .txtcrt_strt_mm.Value = "<%=ConvSPChars(lgObjRs("crt_strt_mm"))%>"
            .txtcrt_strt_dd.Value = "<%=ConvSPChars(lgObjRs("crt_strt_dd"))%>"
            .txtcrt_end_yy.Value = "<%=ConvSPChars(lgObjRs("crt_end_yy"))%>"
            .txtcrt_end_mm.Value = "<%=ConvSPChars(lgObjRs("crt_end_mm"))%>"
            .txtcrt_end_dd.Value = "<%=ConvSPChars(lgObjRs("crt_end_dd"))%>"

            .txtuse_strt_yy.Value = "<%=ConvSPChars(lgObjRs("use_strt_yy"))%>"
            .txtuse_strt_mm.Value = "<%=ConvSPChars(lgObjRs("use_strt_mm"))%>"
            .txtuse_strt_dd.Value = "<%=ConvSPChars(lgObjRs("use_strt_dd"))%>"
            .txtuse_end_yy.Value = "<%=ConvSPChars(lgObjRs("use_end_yy"))%>"
            .txtuse_end_mm.Value = "<%=ConvSPChars(lgObjRs("use_end_mm"))%>"
            .txtuse_end_dd.Value = "<%=ConvSPChars(lgObjRs("use_end_dd"))%>"

            .txtorg_bas_cnt.Value = "<%=ConvSPChars(lgObjRs("org_bas_cnt"))%>"
            .txtabsnt_bas_cnt.Value = "<%=ConvSPChars(lgObjRs("absnt_bas_cnt"))%>"
            .txtprov_year_cnt1.Value = "<%=ConvSPChars(lgObjRs("prov_year_cnt1"))%>"
            .txtprov_year_cnt2.Value = "<%=ConvSPChars(lgObjRs("prov_year_cnt2"))%>"
            
            .txtyear_part_type.Value = "<%=ConvSPChars(lgObjRs("year_part_type"))%>"
            .txtyear_part.Value = "<%=ConvSPChars(lgObjRs("year_part"))%>"

            .txtServ_Add_Basis_Over.Value = "<%=ConvSPChars(lgObjRs("Serv_Add_Basis_Over"))%>"
            .txtServ_Add_Basis_Prov.Value = "<%=ConvSPChars(lgObjRs("Serv_Add_Basis_Prov"))%>"
            .txtServ_Add_Per.Value = "<%=ConvSPChars(lgObjRs("Serv_Add_Per"))%>"
            .txtServ_Add_Cnt.Value = "<%=ConvSPChars(lgObjRs("Serv_Add_Cnt"))%>"
            .txtMaxCnt.Value = "<%=ConvSPChars(lgObjRs("MAX_YEAR_CNT"))%>"
            .txtSingleQ.value = "OK"
        End With          
</Script>       
<%     
    End If    ' Single Query 정상일경우 

    Call SubBizQueryMulti(txtAllow_cd)

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

    Call CommonQueryRs(" count(*) "," hda150t ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜: Create
              If Replace(lgF0,Chr(11),"")="0" Then 
                Call SubBizSaveSingleCreate()
              Else
                Call SubBizSaveSingleUpdate()
              End If
        Case  OPMD_UMODE                                                             '☜: Update
              Call SubBizSaveSingleUpdate()
    End Select

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
   
    Call SubMakeSQLStatements("MR",pKey1,"X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cnt"))
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
' Name : SubBizQueryMulti1
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti1(pKey1)
    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    Call SubMakeSQLStatements("MR",pKey1,"X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey1 = ""
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)

        lgstrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_nm"))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow1 + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

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
                    Call SubBizSaveMultiCreate("1",arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate("1",arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete("1",arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next
    
    arrRowVal = Split(Request("txtSpread1"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow1
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
       
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate("2",arrColVal)                            '☜: Create
            Case "D"
                    Call SubBizSaveMultiDelete("2",arrColVal)                            '☜: Delete
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
    Dim txtGlNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HDA150T( allow_cd, crt_strt_yy, crt_strt_mm, crt_strt_dd,"
    lgStrSQL = lgStrSQL & " crt_end_yy, crt_end_mm, crt_end_dd,"
    lgStrSQL = lgStrSQL & " use_strt_yy, use_strt_mm, use_strt_dd, use_end_yy, use_end_mm, use_end_dd,"
    lgStrSQL = lgStrSQL & " absnt_bas_cnt, prov_year_cnt1, prov_year_cnt2,"
    lgStrSQL = lgStrSQL & " year_part_type, year_part, org_bas_cnt,"
    lgStrSQL = lgStrSQL & " serv_add_basis_over, serv_add_basis_prov, serv_add_per, serv_add_cnt,MAX_YEAR_CNT,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO, ISRT_DT , UPDT_EMP_NO , UPDT_DT  )" 
    lgStrSQL = lgStrSQL & " VALUES("  
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtcrt_strt_yy"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtcrt_strt_mm")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtcrt_strt_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtcrt_end_yy"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtcrt_end_mm")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtcrt_end_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtuse_strt_yy"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtuse_strt_mm")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtuse_strt_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtuse_end_yy"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtuse_end_mm")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtuse_end_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtabsnt_bas_cnt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtprov_year_cnt1"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtprov_year_cnt2"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtyear_part_type")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtyear_part")), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtorg_bas_cnt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtserv_add_basis_over"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtserv_add_basis_prov"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtserv_add_per"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtserv_add_cnt"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtMaxCnt"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDA150T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " crt_strt_yy = " & UNIConvNum(Request("txtcrt_strt_yy"),0) & ","
    lgStrSQL = lgStrSQL & " crt_strt_mm = " & FilterVar(Request("txtcrt_strt_mm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " crt_strt_dd = " & FilterVar(Request("txtcrt_strt_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    lgStrSQL = lgStrSQL & " crt_end_yy = " & UNIConvNum(Request("txtcrt_end_yy"),0) & ","
    lgStrSQL = lgStrSQL & " crt_end_mm = " & FilterVar(Request("txtcrt_end_mm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " crt_end_dd = " & FilterVar(Request("txtcrt_end_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    lgStrSQL = lgStrSQL & " use_strt_yy = " & UNIConvNum(Request("txtuse_strt_yy"),0) & ","
    lgStrSQL = lgStrSQL & " use_strt_mm = " & FilterVar(Request("txtuse_strt_mm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " use_strt_dd = " & FilterVar(Request("txtuse_strt_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    lgStrSQL = lgStrSQL & " use_end_yy = " & UNIConvNum(Request("txtuse_end_yy"),0) & ","
    lgStrSQL = lgStrSQL & " use_end_mm = " & FilterVar(Request("txtuse_end_mm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " use_end_dd = " & FilterVar(Request("txtuse_end_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    lgStrSQL = lgStrSQL & " absnt_bas_cnt =	" & UNIConvNum(Request("txtabsnt_bas_cnt"),0) & ","
    lgStrSQL = lgStrSQL & " prov_year_cnt1 = " & UNIConvNum(Request("txtprov_year_cnt1"),0) & ","
    lgStrSQL = lgStrSQL & " prov_year_cnt2 = " & UNIConvNum(Request("txtprov_year_cnt2"),0) & ","
    lgStrSQL = lgStrSQL & " year_part_type = " & FilterVar(Request("txtyear_part_type"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " year_part = " & FilterVar(Request("txtyear_part"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " org_bas_cnt = " & UNIConvNum(Request("txtorg_bas_cnt"),0) & ","
    lgStrSQL = lgStrSQL & " serv_add_basis_over = " & UNIConvNum(Request("txtserv_add_basis_over"),0) & ","
    lgStrSQL = lgStrSQL & " serv_add_basis_prov = " & UNIConvNum(Request("txtserv_add_basis_prov"),0) & ","
    lgStrSQL = lgStrSQL & " serv_add_per = " & UNIConvNum(Request("txtserv_add_per"),0) & ","
    lgStrSQL = lgStrSQL & " serv_add_cnt = " & UNIConvNum(Request("txtserv_add_cnt"),0) & ","
    lgStrSQL = lgStrSQL & " MAX_YEAR_CNT = " & UNIConvNum(Request("txtMaxCnt"),0) & ","

    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE           "
    lgStrSQL = lgStrSQL & " allow_cd = " & FilterVar(lgKeyStream(0), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(spreadNo, arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HDA100T( allow_cd, dilig_cd, dilig_cnt, flag" 
    lgStrSQL = lgStrSQL & " ,ISRT_EMP_NO,ISRT_DT, UPDT_EMP_NO, UPDT_DT )" 
    lgStrSQL = lgStrSQL & " VALUES( "

    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(4),0)     & ","

	if spreadNo = "1" then
		lgStrSQL = lgStrSQL & " '1',"
	else
		lgStrSQL = lgStrSQL & " '2',"
	end if
	
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(spreadNo,arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDA100T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " dilig_cnt = " & UNIConvNum(arrColVal(4),0) & ","

    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")
	if spreadNo = "1" then
		lgStrSQL = lgStrSQL & " AND  flag = '1'"
	else
		lgStrSQL = lgStrSQL & " AND  flag = '2'"
	end if
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(spreadNo,arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HDA100T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(3)), "''", "S")
	if spreadNo = "1" then
		lgStrSQL = lgStrSQL & " AND  flag = '1'"
	else
		lgStrSQL = lgStrSQL & " AND  flag = '2'"
	end if
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
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
                                   lgStrSQL = "Select * " 
                                   lgStrSQL = lgStrSQL & " From  HDA150T "
                                   lgStrSQL = lgStrSQL & " WHERE allow_cd " & pComp & pCode 	
                        End Select
           End Select             
        Case "M"           '                  0               1                 2              3                4                5
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount & "  dilig_cd, " 
                       lgStrSQL = lgStrSQL & "                  dbo.ufn_H_GetCodeName(" & FilterVar("HCA010T", "''", "S") & ",dilig_cd,'') dilig_nm, " 
                       lgStrSQL = lgStrSQL & "                  dilig_cnt " 
                       lgStrSQL = lgStrSQL & " From  HDA100T "
                       lgStrSQL = lgStrSQL & " WHERE allow_cd " & pComp & pCode

						if lgSpreadFlg = "1" then
							lgStrSQL = lgStrSQL & " and flag = '1'"
						else
							lgStrSQL = lgStrSQL & " and flag = '2'"
						end if						
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
					.ggoSpread.Source     = .frm1.vspdData1
					.ggoSpread.SSShowData "<%=lgstrData%>"          
					.lgStrPrevKey1    = "<%=lgStrPrevKey1%>"
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
