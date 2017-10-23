<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    
    Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")    
                                                                           '☜: Clear Error status
    Dim startDate
    Dim endDate
    Dim prevStartDate
    Dim prevEndDate

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = UniConvNumStringToDouble(Request("lgMaxCount"),0)                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UniConvNumStringToDouble(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)


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
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim txtGlNo
    Dim iLcNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

                                                        '☜ : Release RecordSSet
    Call SubBizQueryMulti()

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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    dim pKey1
    Dim txtDeptNm
    Dim strFromToWhere

    Dim b_amt, b_amt2    ' 매출액 
    Dim c_amt, c_amt2    ' 매출원가 
    Dim c01_amt, c01_amt2    ' 재료비 
    Dim c02_amt, c02_amt2    ' 변동노무비 
    Dim c03_amt, c03_amt2    ' 변동경비 
    Dim c04_amt, c04_amt2    ' 관세환급 
    Dim c11_amt, c11_amt2    ' 고정노무비 
    Dim c12_amt, c12_amt2    ' 고정경비 
    Dim d_amt, d_amt2    ' 매출이익 
    Dim e_amt, e_amt2    ' 판매비 
    Dim e01_amt, e01_amt2    ' 판매변동비 
    Dim e02_amt, e02_amt2    ' 판매고정비 
    Dim f_amt, f_amt2    ' 일반관리비 
    Dim g_amt, g_amt2    ' 영업이익 
    Dim h_amt, h_amt2    ' 영업외수익 
    Dim i_amt, i_amt2    ' 영업외비용 
    Dim j_amt, j_amt2    ' 총원가 
    Dim k_amt, k_amt2    ' 경상이익 
    Dim l_amt, l_amt2    ' 특별이익 
    Dim m_amt, m_amt2    ' 특별손실 
    Dim n_amt, n_amt2    ' 세전이익 
    Dim o_amt, o_amt2    ' 공헌원가 
    Dim p_amt, p_amt2    ' 공헌이익 
    Dim q_amt, q_amt2    ' 공통경비 
    Dim q01_amt, q01_amt2    ' 본사공통비 
    Dim q02_amt, q02_amt2    ' 사업부공통비 
    Dim r_amt, r_amt2    ' 통제이익 

    Dim calcPcnt          '매출대비 비율 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    If Trim(lgKeyStream(2)) <> "" Then
		Call CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP "," ITEM_GROUP_CD = " & FilterVar(lgKeyStream(2), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtDeptNm = ""
		else
		  txtDeptNm = Trim(Replace(lgF0,Chr(11),""))
		end if
	else
		txtDeptNm = ""
	End If

    Call SubMakeSQLStatements("MR","X","X","X")                                   '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""

        iDx = 1

        Do While Not lgObjRs.EOF

            Select Case lgObjRs("minor_cd")
                Case "B"
                    b_amt  = lgObjRs("amt")
                    b_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(b_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(lgObjRs("amt")) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(b_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(b_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(b_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(b_amt) - CDbl(b_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(b_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(b_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "C"
                    c_amt  = lgObjRs("amt")
                    c_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(c_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(c_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(c_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(c_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(c_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(c_amt) - CDbl(c_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(c_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(c_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "C01"
                    c01_amt  = lgObjRs("amt")
                    c01_amt2 = lgObjRs("prevamt")
                Case "C02"
                    c02_amt  = lgObjRs("amt")
                    c02_amt2 = lgObjRs("prevamt")
                Case "C03"
                    c03_amt  = lgObjRs("amt")
                    c03_amt2 = lgObjRs("prevamt")
                Case "C04"
                    c04_amt  = lgObjRs("amt")
                    c04_amt2 = lgObjRs("prevamt")
                Case "C11"
                    c11_amt  = lgObjRs("amt")
                    c11_amt2 = lgObjRs("prevamt")
                Case "C12"
                    c12_amt  = lgObjRs("amt")
                    c12_amt2 = lgObjRs("prevamt")
                Case "D"
                    d_amt  = CDbl(b_amt) - CDbl(c_amt)
                    d_amt2 = CDbl(b_amt2) - CDbl(c_amt2)

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(d_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(d_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(d_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(d_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(d_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(d_amt) - CDbl(d_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(d_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(d_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "E"
                    e_amt  = lgObjRs("amt")
                    e_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(e_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(e_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(e_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(e_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(e_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(e_amt) - CDbl(e_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(e_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(e_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "E01"
                    e01_amt  = lgObjRs("amt")
                    e01_amt2 = lgObjRs("prevamt")
                Case "F"
                    f_amt  = lgObjRs("amt")
                    f_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(f_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(f_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(f_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(f_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(f_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(f_amt) - CDbl(f_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(f_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(f_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "G"
                    g_amt  = CDbl(d_amt) - CDbl(e_amt) - CDbl(f_amt)
                    g_amt2 = CDbl(d_amt2) - CDbl(e_amt2) - CDbl(f_amt2)

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(g_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(g_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(g_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(g_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(g_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(g_amt) - CDbl(g_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(g_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(g_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "H"
                    h_amt  = lgObjRs("amt")
                    h_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(h_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(h_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(h_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(h_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(h_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(h_amt) - CDbl(h_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(h_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(h_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "I"
                    i_amt  = lgObjRs("amt")
                    i_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(i_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(i_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(i_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(i_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(i_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(i_amt) - CDbl(i_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(i_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(i_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "J"
                    j_amt  = CDbl(c_amt) + CDbl(e_amt) + CDbl(f_amt) - CDbl(h_amt) + CDbl(i_amt)
                    j_amt2 = CDbl(c_amt2) + CDbl(e_amt2) + CDbl(f_amt2) - CDbl(h_amt2) + CDbl(i_amt2)

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(j_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(j_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(j_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(j_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(j_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(j_amt) - CDbl(j_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(j_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(j_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "K"
                    k_amt  = CDbl(b_amt) - CDbl(j_amt)
                    k_amt2 = CDbl(b_amt2) - CDbl(j_amt2)

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(k_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(k_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(k_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(k_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(k_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(k_amt) - CDbl(k_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(k_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(k_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "L"
                    l_amt  = lgObjRs("amt")
                    l_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(l_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(l_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(l_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(l_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(l_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(l_amt) - CDbl(l_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(l_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(l_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "M"
                    m_amt  = lgObjRs("amt")
                    m_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(m_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(m_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(m_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(m_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(m_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(m_amt) - CDbl(m_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(m_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(m_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "N"
                    n_amt  = CDbl(k_amt) + CDbl(l_amt) - CDbl(m_amt)
                    n_amt2 = CDbl(k_amt2) + CDbl(l_amt2) - CDbl(m_amt2)

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(n_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(n_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(n_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(n_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(n_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(n_amt) - CDbl(n_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(n_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(n_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "O"
                    o_amt  = CDbl(c01_amt) + CDbl(c02_amt) + CDbl(c03_amt) + CDbl(c04_amt) + CDbl(e01_amt)
                    o_amt2 = CDbl(c01_amt2) + CDbl(c02_amt2) + CDbl(c03_amt2) + CDbl(c04_amt2) + CDbl(e01_amt2)

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(o_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(o_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(o_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(o_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(o_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(o_amt) - CDbl(o_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(o_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(o_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "P"
                    p_amt  = CDbl(b_amt) - CDbl(o_amt)
                    p_amt2 = CDbl(b_amt2) - CDbl(o_amt2)

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(p_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(p_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(p_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(p_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(p_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(p_amt) - CDbl(p_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(p_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(p_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "Q"
                    q_amt  = lgObjRs("amt")
                    q_amt2 = lgObjRs("prevamt")

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(q_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(q_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(q_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(q_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(q_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(q_amt) - CDbl(q_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(q_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(q_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
                Case "Q01"
                    q01_amt  = lgObjRs("amt")
                    q01_amt2 = lgObjRs("prevamt")
                Case "Q02"
                    q02_amt  = lgObjRs("amt")
                    q02_amt2 = lgObjRs("prevamt")
                Case "R"
                    r_amt  = CDbl(n_amt) + CDbl(q_amt)
                    r_amt2 = CDbl(n_amt2) + CDbl(q_amt2)

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(r_amt, ggAmtOfMoney.DecPoint, 0)
                    calcPcnt =  CDbl(r_amt) / CDbl(b_amt) * 100
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(r_amt2, ggAmtOfMoney.DecPoint, 0)
                    If CLng(r_amt2) = 0 then
                        calcPcnt = 0
                    else
                        calcPcnt =  CDbl(r_amt2) / CDbl(b_amt2) * 100
                    end if
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)

                    calcPcnt = CDbl(r_amt) - CDbl(r_amt2)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggAmtOfMoney.DecPoint, 0)

                    If CLng(r_amt2) = 0 then
                        lgstrData = lgstrData & Chr(11) & "100"
                    else
                        calcPcnt = CDbl(calcPcnt) / Abs(CDbl(r_amt2)) * 100
                        lgstrData = lgstrData & Chr(11) & UNINumClientFormat(calcPcnt, ggExchRate.DecPoint, 0)
                    end if

                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
            End Select

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If

        Loop
    End If

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If

    Call SubHandleError("MR",lgObjRs,Err)

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

%>
<SCRIPT LANGUAGE=vbscript>
	With Parent
		.frm1.txtDeptNm.value = "<%=ConvSPChars(txtDeptNm)%>"
	END With
</SCRIPT>
<%

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
    Dim txtGlNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
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
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount
    Dim strYear, strMonth, strDay

    startDate = Replace(lgKeyStream(0), gServerDateType ,"")
    endDate = Replace(lgKeyStream(1), gServerDateType ,"")

    prevStartDate = UNIDateAdd("yyyy",-1,lgKeyStream(0) & gServerDateType & "01",gServerDateFormat)
    Call ExtractDateFrom(prevStartDate,gServerDateFormat,gServerDateType,strYear,strMonth,strDay)
    prevStartDate = strYear & strMonth
    prevEndDate = UNIDateAdd("yyyy",-1,lgKeyStream(1) & gServerDateType & "01",gServerDateFormat)
    Call ExtractDateFrom(prevEndDate,gServerDateFormat,gServerDateType,strYear,strMonth,strDay)
    prevEndDate = strYear & strMonth

    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)

        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
               Case "D"
               Case "U"
               Case "C"
           End Select
        Case "M"
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           Select Case Mid(pDataType,2,1)
               Case "C"
               Case "D"
               Case "R"
                    lgStrSQL = "Select TOP " & iSelCount  & "A.MINOR_CD, A.MINOR_NM, ISNULL(A.AMT,0) AS AMT, ISNULL(B.AMT,0) AS PREVAMT "
                    lgStrSQL = lgStrSQL & " FROM (   SELECT  B.MINOR_CD, B.MINOR_NM, "
                    lgStrSQL = lgStrSQL & "                  SUM(ISNULL(A.AMOUNT,0)) AS AMT "
                    lgStrSQL = lgStrSQL & "            FROM  G_ITEM_PROFIT A, B_MINOR B "
                    lgStrSQL = lgStrSQL & "           WHERE  A.YYYYMM BETWEEN " & FilterVar(startDate, "''", "S") & "AND " & FilterVar(endDate, "''", "S")
                    lgStrSQL = lgStrSQL & "             AND  A.GAIN_GROUP =* B.MINOR_CD "
                    lgStrSQL = lgStrSQL & "             AND  B.MAJOR_CD = " & FilterVar("G1005", "''", "S") & " "
                    lgStrSQL = lgStrSQL & "             AND  B.MINOR_CD IN (" & FilterVar("B", "''", "S") & " ," & FilterVar("C", "''", "S") & " ," & FilterVar("D", "''", "S") & " ," & FilterVar("E", "''", "S") & " ," & FilterVar("F", "''", "S") & " ," & FilterVar("G", "''", "S") & " ," & FilterVar("H", "''", "S") & " ," & FilterVar("I", "''", "S") & " ," & FilterVar("J", "''", "S") & " ," & FilterVar("K", "''", "S") & " ," & FilterVar("L", "''", "S") & " ," & FilterVar("M", "''", "S") & " ," & FilterVar("N", "''", "S") & " ," & FilterVar("O", "''", "S") & " ," & FilterVar("P", "''", "S") & " ," & FilterVar("Q", "''", "S") & " ," & FilterVar("R", "''", "S") & " ) "
                    If FilterVar(lgKeyStream(2),"N","S") <> "N" then
                        lgStrSQL = lgStrSQL & "             AND  A.ITEM_GROUP_CD =  " & FilterVar(lgKeyStream(2), "''", "S")
                    End if
                    lgStrSQL = lgStrSQL & "        GROUP BY  B.MINOR_CD, B.MINOR_NM UNION "
                    lgStrSQL = lgStrSQL & "          SELECT  B.MINOR_CD, B.MINOR_NM, "
                    lgStrSQL = lgStrSQL & "                  SUM(ISNULL(A.AMOUNT,0)) AS AMT "
                    lgStrSQL = lgStrSQL & "            FROM  G_ITEM_PROFIT A, B_MINOR B "
                    lgStrSQL = lgStrSQL & "           WHERE  A.YYYYMM BETWEEN " & FilterVar(startDate, "''", "S") & "AND " & FilterVar(endDate, "''", "S")
                    lgStrSQL = lgStrSQL & "             AND  A.GAIN_CD =* B.MINOR_CD "
                    lgStrSQL = lgStrSQL & "             AND  B.MAJOR_CD = " & FilterVar("G1006", "''", "S") & " "
                    lgStrSQL = lgStrSQL & "             AND  B.MINOR_CD IN (" & FilterVar("C01", "''", "S") & "," & FilterVar("C02", "''", "S") & "," & FilterVar("C03", "''", "S") & "," & FilterVar("C04", "''", "S") & "," & FilterVar("C11", "''", "S") & "," & FilterVar("C12", "''", "S") & "," & FilterVar("E01", "''", "S") & "," & FilterVar("E02", "''", "S") & "," & FilterVar("Q01", "''", "S") & "," & FilterVar("Q02", "''", "S") & ") "
                    If FilterVar(lgKeyStream(2),"N","S") <> "N" then
                        lgStrSQL = lgStrSQL & "             AND  A.ITEM_GROUP_CD =  " & FilterVar(lgKeyStream(2), "''", "S")
                    End if
                    lgStrSQL = lgStrSQL & "        GROUP BY  B.MINOR_CD, B.MINOR_NM ) A, "
                    lgStrSQL = lgStrSQL & "      (   SELECT  B.MINOR_CD, B.MINOR_NM, "
                    lgStrSQL = lgStrSQL & "                  SUM(ISNULL(A.AMOUNT,0)) AS AMT "
                    lgStrSQL = lgStrSQL & "            FROM  G_ITEM_PROFIT A, B_MINOR B "
                    lgStrSQL = lgStrSQL & "           WHERE  A.YYYYMM BETWEEN " & FilterVar(prevStartDate, "''", "S") & "AND " & FilterVar(prevEndDate, "''", "S")
                    lgStrSQL = lgStrSQL & "             AND  A.GAIN_GROUP =* B.MINOR_CD "
                    lgStrSQL = lgStrSQL & "             AND  B.MAJOR_CD = " & FilterVar("G1005", "''", "S") & " "
                    lgStrSQL = lgStrSQL & "             AND  B.MINOR_CD IN (" & FilterVar("B", "''", "S") & " ," & FilterVar("C", "''", "S") & " ," & FilterVar("D", "''", "S") & " ," & FilterVar("E", "''", "S") & " ," & FilterVar("F", "''", "S") & " ," & FilterVar("G", "''", "S") & " ," & FilterVar("H", "''", "S") & " ," & FilterVar("I", "''", "S") & " ," & FilterVar("J", "''", "S") & " ," & FilterVar("K", "''", "S") & " ," & FilterVar("L", "''", "S") & " ," & FilterVar("M", "''", "S") & " ," & FilterVar("N", "''", "S") & " ," & FilterVar("O", "''", "S") & " ," & FilterVar("P", "''", "S") & " ," & FilterVar("Q", "''", "S") & " ," & FilterVar("R", "''", "S") & " ) "
                    If FilterVar(lgKeyStream(2),"N","S") <> "N" then
                        lgStrSQL = lgStrSQL & "             AND  A.ITEM_GROUP_CD =  " & FilterVar(lgKeyStream(2), "''", "S")
                    End if
                    lgStrSQL = lgStrSQL & "        GROUP BY  B.MINOR_CD, B.MINOR_NM UNION "
                    lgStrSQL = lgStrSQL & "          SELECT  B.MINOR_CD, B.MINOR_NM, "
                    lgStrSQL = lgStrSQL & "                  SUM(ISNULL(A.AMOUNT,0)) AS AMT "
                    lgStrSQL = lgStrSQL & "            FROM  G_ITEM_PROFIT A, B_MINOR B "
                    lgStrSQL = lgStrSQL & "           WHERE  A.YYYYMM BETWEEN " & FilterVar(prevStartDate, "''", "S") & "AND " & FilterVar(prevEndDate, "''", "S")
                    lgStrSQL = lgStrSQL & "             AND  A.GAIN_CD =* B.MINOR_CD "
                    lgStrSQL = lgStrSQL & "             AND  B.MAJOR_CD = " & FilterVar("G1006", "''", "S") & " "
                    lgStrSQL = lgStrSQL & "             AND  B.MINOR_CD IN (" & FilterVar("C01", "''", "S") & "," & FilterVar("C02", "''", "S") & "," & FilterVar("C03", "''", "S") & "," & FilterVar("C04", "''", "S") & "," & FilterVar("C11", "''", "S") & "," & FilterVar("C12", "''", "S") & "," & FilterVar("E01", "''", "S") & "," & FilterVar("E02", "''", "S") & "," & FilterVar("Q01", "''", "S") & "," & FilterVar("Q02", "''", "S") & ") "
                    If FilterVar(lgKeyStream(2),"N","S") <> "N" then
                        lgStrSQL = lgStrSQL & "             AND  A.ITEM_GROUP_CD =  " & FilterVar(lgKeyStream(2), "''", "S")
                    End if
                    lgStrSQL = lgStrSQL & "        GROUP BY  B.MINOR_CD, B.MINOR_NM ) B "
                    lgStrSQL = lgStrSQL & "WHERE  A.MINOR_CD = B.MINOR_CD "
               Case "U"
           End Select
    End Select
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
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '☜ : Display data
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"                          '☜ : Next next data tag
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .DBQueryOk
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
