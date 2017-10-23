<%@ LANGUAGE="VBSCript"%>
<%Option Explicit%>

<!--
======================================================================================================
*  1. Module Name          : 생산
*  2. Function Name        : 생산실적조회및출력
*  3. Program ID           : P4422OA1_LKO391
*  4. Program Name         : P4422OA1_LKO391
*  5. Program Desc         : 생산실적조회및출력
*  6. Comproxy List        :
*  7. Modified date(First) : 2007/01/24
*  8. Modified date(Last)  :
*  9. Modifier (First)     : Lim, JaeBon
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
    Dim lgStrPrevKey
    Const C_SHEETMAXROWS_D = 500
    Dim lgSvrDateTime
    lgSvrDateTime = GetSvrDateTime

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "P","NOCOOKIE","MB")

    Call HideStatusWnd

    lgErrorStatus = "NO"
    lgErrorPos    = ""                                                           '☜: Set to space
    lgOpModeCRUD  = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream   = Split(Request("txtKeyStream"), gColSep)
    lgStrPrevKey  = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Dim totalAmtSum, itemQtySum
    totalAmtSum = 0
    itemQtySum  = 0

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
            Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
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
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
    
    Dim query1, query2
    query2 = Request("txtQuery2")
    Call SubBizSaveMultiUpdate(query2)

    query1 = Request("txtQuery1")
    Call SubBizSaveMultiUpdate(query1)
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim iDx
    dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR", lgKeyStream, "X", C_EQ)                                 '☆ : Make sql statements

    If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs, C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""

        iDx = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("prodt_order_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("report_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("report_type_nm"))
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("prod_qty_in_order_unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("insp_good_qty_in_order_unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("insp_bad_qty_in_order_unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("rcpt_qty_in_order_unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("remark"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("cur_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("seq"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("opr_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("insrt_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_no_ko441"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("report_type"))

            lgstrData = lgstrData & Chr(11) & C_SHEETMAXROWS_D + iDx
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

    Call SubHandleError("MR", lgObjConn, lgObjRs, Err)
    Call SubCloseRs(lgObjRs)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType, pCode, pCode1, pComp)
    Dim iSelCount
    Dim whereQuery
    whereQuery = ""


    If pCode(0) <> "" Then
        whereQuery = whereQuery & " AND a.plant_cd = " & FilterVar(pCode(0), "''", "S")
    End If
    If pCode(1) <> "" Then
        whereQuery = whereQuery & " AND a.report_dt >= " & FilterVar(UniConvDate(pCode(1)), "''", "S")
    End If
    If pCode(2) <> "" Then
        whereQuery = whereQuery & " AND a.report_dt <= " & FilterVar(UniConvDate(pCode(2)), "''", "S")
    End If
    
    If pCode(3) <> "" Then
        whereQuery = whereQuery & " AND c.wc_cd = " & FilterVar(pCode(3), "''", "S")    
    End If
    

    whereQuery = whereQuery & " AND D.COST_CD LIKE " & FilterVar(pCode(10) & "%", "''", "S")    
  
    
    If pCode(4) = "" Then
        whereQuery = whereQuery & " AND b.item_cd >= " & FilterVar("0", "''", "S")
    Else
        whereQuery = whereQuery & " AND b.item_cd >= " & FilterVar(pCode(4), "''", "S")
    End If
    If pCode(5) = "" Then
        whereQuery = whereQuery & " AND b.item_cd <= " & FilterVar("zzzzzzzzzzzzzzzzzz", "''", "S")
    Else
        whereQuery = whereQuery & " AND b.item_cd <= " & FilterVar(pCode(5), "''", "S")
    End If
    If pCode(6) = "" Then
        whereQuery = whereQuery & " AND a.prodt_order_no >= " & FilterVar("0", "''", "S")
    Else
        whereQuery = whereQuery & " AND a.prodt_order_no >= " & FilterVar(pCode(6), "''", "S")
    End If
    If pCode(7) = "" Then
        whereQuery = whereQuery & " AND a.prodt_order_no <= " & FilterVar("zzzzzzzzzzzzzzzzzz", "''", "S")
    Else
        whereQuery = whereQuery & " AND a.prodt_order_no <= " & FilterVar(pCode(7), "''", "S")
    End If
    If pCode(8) <> "" Then
        whereQuery = whereQuery & " AND a.report_type = " & FilterVar(pCode(8), "''", "S")
    End If
    
    If pCode(9) <> "" Then
        whereQuery = whereQuery & " AND e.base_item_cd = " & FilterVar(pCode(9), "''", "S")
    End If
    

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

           Select Case Mid(pDataType,2,1)
               Case "R"
                    lgStrSQL = " SELECT TOP " & iSelCount & " * from " & vbcrlf
                    lgStrSQL = lgStrSQL & " ( select c.wc_cd    , d.wc_nm                 , a.prodt_order_no             , b.item_cd                   , e.item_nm               , " & vbcrlf
                    lgStrSQL = lgStrSQL & " a.report_dt, A.REPORT_TYPE, " & vbcrlf
                    lgStrSQL = lgStrSQL & " CASE	A.REPORT_TYPE " & vbcrlf
                    lgStrSQL = lgStrSQL & "			WHEN	'G'	THEN	'양품' " & vbcrlf
                    lgStrSQL = lgStrSQL & "			WHEN	'B'	THEN	'불량' " & vbcrlf
                    lgStrSQL = lgStrSQL & " END	REPORT_TYPE_NM, " & vbcrlf
                    lgStrSQL = lgStrSQL & " a.prod_qty_in_order_unit, a.insp_good_qty_in_order_unit, a.insp_bad_qty_in_order_unit, a.rcpt_qty_in_order_unit, " & vbcrlf
                    lgStrSQL = lgStrSQL & " a.remark   , a.cur_cd                , a.seq                        , a.opr_no                   , convert(varchar(20),a.insrt_dt,120)  AS   insrt_dt, a.lot_no_ko441                        " & vbcrlf
                    lgStrSQL = lgStrSQL & " FROM p_production_results a                                                                           " & vbcrlf
                    lgStrSQL = lgStrSQL & " INNER JOIN p_production_order_header b ON a.prodt_order_no = b.prodt_order_no                         " & vbcrlf
                    lgStrSQL = lgStrSQL & " INNER JOIN p_production_order_detail c ON a.prodt_order_no = c.prodt_order_no AND a.opr_no = c.opr_no " & vbcrlf
                    lgStrSQL = lgStrSQL & " LEFT OUTER JOIN p_work_center d ON a.plant_cd = d.plant_cd AND c.wc_cd = d.wc_cd                      " & vbcrlf
                    lgStrSQL = lgStrSQL & " LEFT OUTER JOIN b_item e ON b.item_cd = e.item_cd                                                     " & vbcrlf
                    lgStrSQL = lgStrSQL & " WHERE a.del_flg ='N' " & vbcrlf
                    lgStrSQL = lgStrSQL & whereQuery                     & vbcrlf
                                        
                    lgStrSQL = lgStrSQL & " union all  SELECT  "   & vbcrlf
                    lgStrSQL = lgStrSQL & " '' wc_cd,'총계' wc_nm, '' prodt_order_no, '' item_cd, '' item_nm, " & vbcrlf
                    lgStrSQL = lgStrSQL & " null report_dt, '' REPORT_TYPE, " & vbcrlf
                    lgStrSQL = lgStrSQL & " '' REPORT_TYPE_NM, " & vbcrlf
                    lgStrSQL = lgStrSQL & " sum(a.prod_qty_in_order_unit), sum(a.insp_good_qty_in_order_unit), sum(a.insp_bad_qty_in_order_unit), sum(a.rcpt_qty_in_order_unit), " & vbcrlf
                    lgStrSQL = lgStrSQL & " '' remark, '' cur_cd,  null seq, null opr_no, null insrt_dt, null lot_no_ko441                        " & vbcrlf
                    lgStrSQL = lgStrSQL & " FROM p_production_results a                                                                           " & vbcrlf
                    lgStrSQL = lgStrSQL & " INNER JOIN p_production_order_header b ON a.prodt_order_no = b.prodt_order_no                         " & vbcrlf
                    lgStrSQL = lgStrSQL & " INNER JOIN p_production_order_detail c ON a.prodt_order_no = c.prodt_order_no AND a.opr_no = c.opr_no " & vbcrlf
                    lgStrSQL = lgStrSQL & " LEFT OUTER JOIN p_work_center d ON a.plant_cd = d.plant_cd AND c.wc_cd = d.wc_cd                      " & vbcrlf
                    lgStrSQL = lgStrSQL & " LEFT OUTER JOIN b_item e ON b.item_cd = e.item_cd                                                     " & vbcrlf
                    lgStrSQL = lgStrSQL & " WHERE a.del_flg ='N' " & vbcrlf
                    lgStrSQL = lgStrSQL & whereQuery & vbcrlf
                    lgStrSQL = lgStrSQL & " ) a order by prodt_order_no,wc_cd,item_cd,report_dt" & vbcrlf
                    
                    'response.write lgStrSQL
                    'response.write whereQuery
           End Select
    End Select
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgObjConn.Execute arrColVal,,adCmdText + adExecuteNoRecords

    Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
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
    Case "<%=UID_M0001%>"                                                         '☜ : Query
        If Trim("<%=lgErrorStatus%>") = "NO" Then
            With Parent
                .ggoSpread.Source = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk
            End with
        End If
    Case "<%=UID_M0002%>"                                                         '☜ : Save
        If Trim("<%=lgErrorStatus%>") = "NO" Then
            Parent.DBSaveOk
        Else
            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
        End If
    Case "<%=UID_M0002%>"                                                         '☜ : Delete
        If Trim("<%=lgErrorStatus%>") = "NO" Then
            Parent.DbDeleteOk
        Else
        End If
End Select
</Script>