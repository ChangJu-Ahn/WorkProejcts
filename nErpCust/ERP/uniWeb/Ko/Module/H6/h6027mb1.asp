<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%

    Call HideStatusWnd                                                               '☜: Hide Processing message

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
 
	Dim lgGetSvrDateTime
    Dim sChk	

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
	
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    lgGetSvrDateTime = GetSvrDateTime
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")     '사업장으로 조회 
    iKey1 = iKey1 & " AND PROV_YYMM = " & FilterVar(lgKeyStream(1), "''", "S")
	
    Call SubMakeSQLStatements("R",iKey1)                                     '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
Response.Write "<Script Language=vbscript>" & vbCrLf
Response.Write "   '  Call Parent.DisableToolBar(Parent.Parent.TBC_NEW) " & vbCrLf

Response.Write "   ' Call Parent.CancelRestoreToolBar() " & vbCrLf
Response.Write "  '  Call Parent.DisableToolBar( Parent.Parent.TBC_DELETE) " & vbCrLf
Response.Write "</Script>" & vbCrLf
      
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the starting data. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else


	    Call SetFixSrting(lgObjRs("AMT_LIST"),"/") 
%>

<Script Language=vbscript>
       With Parent.Frm1

			.txtrevert_yymm.text   = "<%= UNIConvDateDBToCompany(lgObjRs("revert_yymm"),Null)%>"
			.txtsubmit_yymm.text   = "<%= UNIConvDateDBToCompany(lgObjRs("submit_yymm"),Null)%>"

			.txtRetireFr_dt.text   = "<%= UNIConvDateDBToCompany(lgObjRs("RETIRE_FR_DT"),Null)%>"
			.txtRetireTo_dt.text   = "<%= UNIConvDateDBToCompany(lgObjRs("RETIRE_TO_DT"),Null)%>"
			.txtYearEnd_yymm.text   = "<%= UNIConvDateDBToCompany(lgObjRs("YEAREND_YYMM"),Null)%>"
			
			.txt_i_A011.text       = "<%=UNINumClientFormat(lgKeyStream(0), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A012.text       = "<%=UNINumClientFormat(lgKeyStream(1), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A013.text       = "<%=UNINumClientFormat(lgKeyStream(2), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A014.text       = "<%=UNINumClientFormat(lgKeyStream(3), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A015.text       = "<%=UNINumClientFormat(lgKeyStream(4), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A021.text       = "<%=UNINumClientFormat(lgKeyStream(5), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A022.text       = "<%=UNINumClientFormat(lgKeyStream(6), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A023.text       = "<%=UNINumClientFormat(lgKeyStream(7), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A024.text       = "<%=UNINumClientFormat(lgKeyStream(8), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A025.text       = "<%=UNINumClientFormat(lgKeyStream(9), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A031.text       = "<%=UNINumClientFormat(lgKeyStream(10), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A032.text       = "<%=UNINumClientFormat(lgKeyStream(11), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A033.text       = "<%=UNINumClientFormat(lgKeyStream(12), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A035.text       = "<%=UNINumClientFormat(lgKeyStream(13), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A041.text       = "<%=UNINumClientFormat(lgKeyStream(14), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A042.text       = "<%=UNINumClientFormat(lgKeyStream(15), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A043.text       = "<%=UNINumClientFormat(lgKeyStream(16), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A044.text       = "<%=UNINumClientFormat(lgKeyStream(17), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A045.text       = "<%=UNINumClientFormat(lgKeyStream(18), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A101.text       = "<%=UNINumClientFormat(lgKeyStream(19), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A102.text       = "<%=UNINumClientFormat(lgKeyStream(20), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A103.text       = "<%=UNINumClientFormat(lgKeyStream(21), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A104.text       = "<%=UNINumClientFormat(lgKeyStream(22), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A105.text       = "<%=UNINumClientFormat(lgKeyStream(23), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A106.text       = "<%=UNINumClientFormat(lgKeyStream(24), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A107.text       = "<%=UNINumClientFormat(lgKeyStream(25), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A108.text       = "<%=UNINumClientFormat(lgKeyStream(26), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A201.text       = "<%=UNINumClientFormat(lgKeyStream(27), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A202.text       = "<%=UNINumClientFormat(lgKeyStream(28), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A203.text       = "<%=UNINumClientFormat(lgKeyStream(29), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A205.text       = "<%=UNINumClientFormat(lgKeyStream(30), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A206.text       = "<%=UNINumClientFormat(lgKeyStream(31), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A207.text       = "<%=UNINumClientFormat(lgKeyStream(32), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A251.text       = "<%=UNINumClientFormat(lgKeyStream(33), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A252.text       = "<%=UNINumClientFormat(lgKeyStream(34), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A253.text       = "<%=UNINumClientFormat(lgKeyStream(35), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A255.text       = "<%=UNINumClientFormat(lgKeyStream(36), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A261.text       = "<%=UNINumClientFormat(lgKeyStream(37), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A262.text       = "<%=UNINumClientFormat(lgKeyStream(38), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A263.text       = "<%=UNINumClientFormat(lgKeyStream(39), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A264.text       = "<%=UNINumClientFormat(lgKeyStream(40), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A265.text       = "<%=UNINumClientFormat(lgKeyStream(41), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A301.text       = "<%=UNINumClientFormat(lgKeyStream(42), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A302.text       = "<%=UNINumClientFormat(lgKeyStream(43), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A303.text       = "<%=UNINumClientFormat(lgKeyStream(44), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A304.text       = "<%=UNINumClientFormat(lgKeyStream(45), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A305.text       = "<%=UNINumClientFormat(lgKeyStream(46), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A306.text       = "<%=UNINumClientFormat(lgKeyStream(47), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A307.text       = "<%=UNINumClientFormat(lgKeyStream(48), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A308.text       = "<%=UNINumClientFormat(lgKeyStream(49), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A401.text       = "<%=UNINumClientFormat(lgKeyStream(50), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A402.text       = "<%=UNINumClientFormat(lgKeyStream(51), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A403.text       = "<%=UNINumClientFormat(lgKeyStream(52), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A405.text       = "<%=UNINumClientFormat(lgKeyStream(53), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A406.text       = "<%=UNINumClientFormat(lgKeyStream(54), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A407.text       = "<%=UNINumClientFormat(lgKeyStream(55), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A451.text       = "<%=UNINumClientFormat(lgKeyStream(56), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A452.text       = "<%=UNINumClientFormat(lgKeyStream(57), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A453.text       = "<%=UNINumClientFormat(lgKeyStream(58), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A455.text       = "<%=UNINumClientFormat(lgKeyStream(59), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A456.text       = "<%=UNINumClientFormat(lgKeyStream(60), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A457.text       = "<%=UNINumClientFormat(lgKeyStream(61), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A501.text       = "<%=UNINumClientFormat(lgKeyStream(62), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A502.text       = "<%=UNINumClientFormat(lgKeyStream(63), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A503.text       = "<%=UNINumClientFormat(lgKeyStream(64), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A504.text       = "<%=UNINumClientFormat(lgKeyStream(65), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A505.text       = "<%=UNINumClientFormat(lgKeyStream(66), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A506.text       = "<%=UNINumClientFormat(lgKeyStream(67), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A507.text       = "<%=UNINumClientFormat(lgKeyStream(68), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A508.text       = "<%=UNINumClientFormat(lgKeyStream(69), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A601.text       = "<%=UNINumClientFormat(lgKeyStream(70), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A602.text       = "<%=UNINumClientFormat(lgKeyStream(71), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A603.text       = "<%=UNINumClientFormat(lgKeyStream(72), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A604.text       = "<%=UNINumClientFormat(lgKeyStream(73), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A605.text       = "<%=UNINumClientFormat(lgKeyStream(74), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A606.text       = "<%=UNINumClientFormat(lgKeyStream(75), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A607.text       = "<%=UNINumClientFormat(lgKeyStream(76), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A608.text       = "<%=UNINumClientFormat(lgKeyStream(77), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A691.text       = "<%=UNINumClientFormat(lgKeyStream(78), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A693.text       = "<%=UNINumClientFormat(lgKeyStream(79), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A694.text       = "<%=UNINumClientFormat(lgKeyStream(80), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A695.text       = "<%=UNINumClientFormat(lgKeyStream(81), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A696.text       = "<%=UNINumClientFormat(lgKeyStream(82), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A697.text       = "<%=UNINumClientFormat(lgKeyStream(83), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A698.text       = "<%=UNINumClientFormat(lgKeyStream(84), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A801.text       = "<%=UNINumClientFormat(lgKeyStream(85), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A802.text       = "<%=UNINumClientFormat(lgKeyStream(86), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A803.text       = "<%=UNINumClientFormat(lgKeyStream(87), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A805.text       = "<%=UNINumClientFormat(lgKeyStream(88), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A806.text       = "<%=UNINumClientFormat(lgKeyStream(89), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A807.text       = "<%=UNINumClientFormat(lgKeyStream(90), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A903.text       = "<%=UNINumClientFormat(lgKeyStream(91), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A904.text       = "<%=UNINumClientFormat(lgKeyStream(92), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A905.text       = "<%=UNINumClientFormat(lgKeyStream(93), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A906.text       = "<%=UNINumClientFormat(lgKeyStream(94), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A907.text       = "<%=UNINumClientFormat(lgKeyStream(95), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A908.text       = "<%=UNINumClientFormat(lgKeyStream(96), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A991.text       = "<%=UNINumClientFormat(lgKeyStream(97), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A992.text       = "<%=UNINumClientFormat(lgKeyStream(98), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A993.text       = "<%=UNINumClientFormat(lgKeyStream(99), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A994.text       = "<%=UNINumClientFormat(lgKeyStream(100), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A995.text       = "<%=UNINumClientFormat(lgKeyStream(101), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A996.text       = "<%=UNINumClientFormat(lgKeyStream(102), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A997.text       = "<%=UNINumClientFormat(lgKeyStream(103), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A998.text       = "<%=UNINumClientFormat(lgKeyStream(104), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_ii_A001.text       = "<%=UNINumClientFormat(lgKeyStream(105), ggAmtOfMoney.DecPoint,0)%>"
			.txt_ii_A002.text       = "<%=UNINumClientFormat(lgKeyStream(106), ggAmtOfMoney.DecPoint,0)%>"
			.txt_ii_A003.text       = "<%=UNINumClientFormat(lgKeyStream(107), ggAmtOfMoney.DecPoint,0)%>"
			.txt_ii_A004.text       = "<%=UNINumClientFormat(lgKeyStream(108), ggAmtOfMoney.DecPoint,0)%>"
			.txt_ii_A005.text       = "<%=UNINumClientFormat(lgKeyStream(109), ggAmtOfMoney.DecPoint,0)%>"
			.txt_ii_A006.text       = "<%=UNINumClientFormat(lgKeyStream(110), ggAmtOfMoney.DecPoint,0)%>"
			.txt_ii_A007.text       = "<%=UNINumClientFormat(lgKeyStream(111), ggAmtOfMoney.DecPoint,0)%>"
			.txt_ii_A008.text       = "<%=UNINumClientFormat(lgKeyStream(112), ggAmtOfMoney.DecPoint,0)%>"
			.txt_ii_A009.text       = "<%=UNINumClientFormat(lgKeyStream(113), ggAmtOfMoney.DecPoint,0)%>"
       End With          
</Script>       
<%     
    End If
    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
    
	
End Sub	


'============================================================================================================
' Name : SetFixSrting(입력값,비교문자,대체문자,고정길이,문자정렬방향)
' Desc : This Function return srting
'============================================================================================================
	Function SetFixSrting(InValue, ComSymbol)
		lgKeyStream  = Split(InValue , ComSymbol)
				
	End Function	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    lgIntFlgMode = CInt(Request("txtFlgMode")) 
    

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '☜ : Update
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

    lgStrSQL = "DELETE  HDF500T"
    lgStrSQL = lgStrSQL & " WHERE BIZ_AREA_CD = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND PROV_YYMM = " & FilterVar(lgKeyStream(1), "''", "S")    

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HDF500T("
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD, "
    lgStrSQL = lgStrSQL & " PROV_YYMM, "
    lgStrSQL = lgStrSQL & " REVERT_YYMM, "
    lgStrSQL = lgStrSQL & " SUBMIT_YYMM, "
    lgStrSQL = lgStrSQL & " RETIRE_FR_DT, "
    lgStrSQL = lgStrSQL & " RETIRE_TO_DT, "
    lgStrSQL = lgStrSQL & " YEAREND_YYMM, "

    lgStrSQL = lgStrSQL & " AMT_LIST, "
    lgStrSQL = lgStrSQL & " ISRT_DT, "
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO, "
    lgStrSQL = lgStrSQL & " UPDT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO) "
    lgStrSQL = lgStrSQL & " VALUES ("
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(2), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(5), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(6), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(7), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvNum(Request("txt_i_A011"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A012"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A013"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A014"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A015"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A021"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A022"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A023"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A024"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A025"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A031"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A032"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A033"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A035"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A041"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A042"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A043"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A044"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A045"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A101"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A102"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A103"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A104"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A105"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A106"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A107"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A108"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A201"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A202"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A203"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A205"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A206"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A207"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A251"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A252"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A253"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A255"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A261"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A262"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A263"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A264"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A265"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A301"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A302"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A303"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A304"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A305"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A306"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A307"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A308"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A401"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A402"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A403"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A405"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A406"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A407"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A451"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A452"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A453"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A455"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A456"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A457"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A501"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A502"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A503"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A504"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A505"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A506"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A507"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A508"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A601"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A602"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A603"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A604"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A605"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A606"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A607"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A608"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A691"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A693"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A694"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A695"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A696"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A697"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A698"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A801"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A802"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A803"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A805"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A806"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A807"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A903"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A904"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A905"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A906"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A907"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A908"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A991"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A992"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A993"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A994"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A995"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A996"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A997"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A998"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A001"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A002"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A003"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A004"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A005"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A006"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A007"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A008"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A009"),0) & "/", "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & ","                          
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


    lgStrSQL = "UPDATE  HDF500T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " REVERT_YYMM = " & FilterVar(lgKeyStream(2), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " SUBMIT_YYMM = " & FilterVar(lgKeyStream(3), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " RETIRE_FR_DT = " & FilterVar(lgKeyStream(5), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " RETIRE_TO_DT = " & FilterVar(lgKeyStream(6), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " YEAREND_YYMM = " & FilterVar(lgKeyStream(7), "''", "S")  & ","
    
    lgStrSQL = lgStrSQL & " AMT_LIST = "
    lgStrSQL = lgStrSQL & FilterVar(UNIConvNum(Request("txt_i_A011"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A012"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A013"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A014"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A015"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A021"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A022"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A023"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A024"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A025"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A031"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A032"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A033"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A035"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A041"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A042"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A043"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A044"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A045"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A101"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A102"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A103"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A104"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A105"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A106"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A107"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A108"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A201"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A202"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A203"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A205"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A206"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A207"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A251"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A252"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A253"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A255"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A261"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A262"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A263"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A264"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A265"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A301"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A302"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A303"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A304"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A305"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A306"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A307"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A308"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A401"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A402"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A403"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A405"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A406"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A407"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A451"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A452"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A453"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A455"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A456"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A457"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A501"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A502"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A503"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A504"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A505"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A506"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A507"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A508"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A601"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A602"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A603"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A604"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A605"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A606"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A607"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A608"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A691"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A693"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A694"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A695"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A696"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A697"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A698"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A801"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A802"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A803"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A805"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A806"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A807"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A903"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A904"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A905"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A906"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A907"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A908"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A991"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A992"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A993"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A994"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A995"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A996"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A997"),0) & "/" &_
                                     UNIConvNum(Request("txt_i_A998"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A001"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A002"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A003"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A004"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A005"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A006"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A007"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A008"),0) & "/" &_
                                     UNIConvNum(Request("txt_ii_A009"),0) & "/", "''", "S")  & ","    
                                     
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
    lgStrSQL = lgStrSQL & " ISRT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S") & ","                ' datetime
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S")                    ' datetime
    lgStrSQL = lgStrSQL & " WHERE BIZ_AREA_CD = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND PROV_YYMM = " & FilterVar(lgKeyStream(1), "''", "S")    

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)

    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
	                 Case ""
                           lgStrSQL = "Select convert(datetime , convert(varchar(8),revert_yymm + " & FilterVar("01", "''", "S") & ")) revert_yymm ,"
                           lgStrSQL = lgStrSQL & "convert(datetime , submit_yymm) submit_yymm ,RETIRE_TO_DT,RETIRE_FR_DT,"
                           lgStrSQL = lgStrSQL & "convert(datetime , convert(varchar(8),YEAREND_YYMM + " & FilterVar("01", "''", "S") & ")) YEAREND_YYMM,"
                           lgStrSQL = lgStrSQL & "amt_list " 
                           lgStrSQL = lgStrSQL & " From  HDF500T  WHERE biz_area_cd = " & pCode
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
        Case "SC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "SD"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "SU"
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
		  Else	
             Parent.DBQueryNG
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
