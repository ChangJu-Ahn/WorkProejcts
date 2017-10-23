<%@ LANGUAGE="VBSCript"%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
    Dim lgStrPrevKey
    Const C_SHEETMAXROWS_D = 10000
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

    'Dim queryFlag
'    queryFlag = Request("queryFlag")

    Dim whereQuery

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

    arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

    For iDx = 1 To C_SHEETMAXROWS_D
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data

        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select

        If lgErrorStatus = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    Err.Clear                                                                        '☜: Clear Error status

        Call SubMakeSQLStatements("MR", lgKeyStream, "X", C_EQ)                                 '☆ : Make sql statements

        If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
            lgStrPrevKey = ""
            Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
            Call SetErrorStatus()
        Else

            Call SubSkipRs(lgObjRs, C_SHEETMAXROWS_D * lgStrPrevKey)

            lgstrData = ""

            iDx = 1

            Do While Not lgObjRs.EOF

                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tracking_no"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("type"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("type"))

'                If UniConvNumberDBToCompany(lgObjRs("qty01"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) = 0 Then
'                lgstrData = lgstrData & Chr(11) & ""
'                Else
'                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty01"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
'                End If
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty01"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty02"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty03"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty04"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty05"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty06"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty07"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty08"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty09"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)


                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty10"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty11"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty12"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty13"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty14"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty15"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty16"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty17"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty18"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty19"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty20"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty21"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty22"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty23"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty24"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty25"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty26"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty27"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty28"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty29"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty30"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty31"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty32"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::32
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty33"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::33
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty34"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::34
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty35"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::35
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty36"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::36

                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty37"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::37
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty38"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::38
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty39"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::39
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty40"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::40
                
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("month_1"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("month_2"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("month_3"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & ""	'ConvSPChars(lgObjRs("plant_cd"))
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty01"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty02"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty03"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty04"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty05"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty06"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty07"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty08"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty09"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty10"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty11"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty12"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty13"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty14"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty15"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty16"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty17"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty18"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty19"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty20"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty21"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty22"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty23"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty24"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty25"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty26"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty27"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty28"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty29"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty30"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty31"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)

                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty32"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::32
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty33"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::33
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty34"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::34
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty35"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::35
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty36"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::36

                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty37"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::37
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty38"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::38
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty39"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::39
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty40"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::40

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
    whereQuery = ""

    If pCode(1) <> "" Then
        whereQuery = whereQuery & " AND x.item_cd = " & FilterVar(pCode(1), "''", "S")
    End If
    
    If pCode(2) <> "" Then
        whereQuery = whereQuery & " AND x.tracking_no = " & FilterVar(pCode(2), "''", "S")
    End If

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

           Select Case Mid(pDataType,2,1)
               Case "R"

					lgStrSQL = ""
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @ProdPlanMonth DATETIME                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @iPeriod       NUMERIC(10,0)                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @PlantCd       Varchar(20)                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @sales_group       Varchar(20)                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " set @ProdPlanMonth = " & FilterVar(pCode(0), "''", "S")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @iPeriod       = " & FilterVar(pCode(4), "''", "D")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @PlantCd       = " & FilterVar(pCode(5), "''", "S")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @sales_group       = " & FilterVar(pCode(6), "''", "S")
'                    lgStrSQL = lgStrSQL & vbCrLf & "  set @ProdPlanMonth = '2007-12-01'                                                                                                                                                "
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " SELECT	DISTINCT A.ITEM_CD,                                                                                                                                                        "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B.ITEM_NM,                                                                                                                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		'' tracking_no,                                                                                                                                                            "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		'수주계획(수주일)' TYPE,                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 0, @ProdPlanMonth)) QTY01,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 1, @ProdPlanMonth)) QTY02,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 2, @ProdPlanMonth)) QTY03,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 3, @ProdPlanMonth)) QTY04,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 4, @ProdPlanMonth)) QTY05,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 5, @ProdPlanMonth)) QTY06,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 6, @ProdPlanMonth)) QTY07,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 7, @ProdPlanMonth)) QTY08,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 8, @ProdPlanMonth)) QTY09,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 9, @ProdPlanMonth)) QTY10,    "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 10, @ProdPlanMonth)) QTY11,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 11, @ProdPlanMonth)) QTY12,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 12, @ProdPlanMonth)) QTY13,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 13, @ProdPlanMonth)) QTY14,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 14, @ProdPlanMonth)) QTY15,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 15, @ProdPlanMonth)) QTY16,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 16, @ProdPlanMonth)) QTY17,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 17, @ProdPlanMonth)) QTY18,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 18, @ProdPlanMonth)) QTY19,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 19, @ProdPlanMonth)) QTY20,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 20, @ProdPlanMonth)) QTY21,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 21, @ProdPlanMonth)) QTY22,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 22, @ProdPlanMonth)) QTY23,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 23, @ProdPlanMonth)) QTY24,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 24, @ProdPlanMonth)) QTY25,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 25, @ProdPlanMonth)) QTY26,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 26, @ProdPlanMonth)) QTY27,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 27, @ProdPlanMonth)) QTY28,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 28, @ProdPlanMonth)) QTY29,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 29, @ProdPlanMonth)) QTY30,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 30, @ProdPlanMonth)) QTY31,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 31, @ProdPlanMonth)) QTY32,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 32, @ProdPlanMonth)) QTY33,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 33, @ProdPlanMonth)) QTY34,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 34, @ProdPlanMonth)) QTY35,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 35, @ProdPlanMonth)) QTY36,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 36, @ProdPlanMonth)) QTY37,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 37, @ProdPlanMonth)) QTY38,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 38, @ProdPlanMonth)) QTY39,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE  AND	sales_group = @sales_group    AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 39, @ProdPlanMonth)) QTY40,   "
                    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_1,                                                                                                                                                                "
                    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_2,                                                                                                                                                                "
                    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_3                                                                                                                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " FROM	SALES_ITEM_REQ_PLAN_KO441 A,                                                                                                                                               "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B_ITEM B                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " WHERE	A.ITEM_CD = B.ITEM_CD                                                                                                                                                      "
                    lgStrSQL = lgStrSQL & vbCrLf & " AND	A.PROJECT_CODE = @PlantCd                                                                                                                                                      "
                    lgStrSQL = lgStrSQL & vbCrLf & " AND	A.sales_group = @sales_group                                                                                                                                                      "
                    lgStrSQL = lgStrSQL & vbCrLf & "  AND    CONVERT(CHAR(10), A.DLVY_PLAN_DT, 120) BETWEEN @ProdPlanMonth AND DATEADD(day, @iPeriod, @ProdPlanMonth)   "

'					lgStrSQL = ""
'                    lgStrSQL = lgStrSQL & vbCrLf & " declare @ProdPlanMonth varchar(7)"
'                    lgStrSQL = lgStrSQL & vbCrLf & " "
'                    lgStrSQL = lgStrSQL & vbCrLf & " set @ProdPlanMonth = " & FilterVar(pCode(0), "''", "S")
'                    lgStrSQL = lgStrSQL & vbCrLf & " "
'                    lgStrSQL = lgStrSQL & vbCrLf & " select     x.item_cd, x.item_nm, x.tracking_no, x.type,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			sum(x.qty01) qty01, sum(x.qty02) qty02, sum(x.qty03) qty03, sum(x.qty04) qty04, sum(x.qty05) qty05, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			sum(x.qty06) qty06, sum(x.qty07) qty07, sum(x.qty08) qty08, sum(x.qty09) qty09, sum(x.qty10) qty10, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			sum(x.qty11) qty11, sum(x.qty12) qty12, sum(x.qty13) qty13, sum(x.qty14) qty14, sum(x.qty15) qty15, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			sum(x.qty16) qty16, sum(x.qty17) qty17, sum(x.qty18) qty18, sum(x.qty19) qty19, sum(x.qty20) qty20, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			sum(x.qty21) qty21, sum(x.qty22) qty22, sum(x.qty23) qty23, sum(x.qty24) qty24, sum(x.qty25) qty25, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			sum(x.qty26) qty26, sum(x.qty27) qty27, sum(x.qty28) qty28, sum(x.qty29) qty29, sum(x.qty30) qty30, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			sum(x.qty31) qty31, sum(x.month_1) month_1, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			sum(x.month_2) month_2, sum(x.month_3) month_3"
'                    lgStrSQL = lgStrSQL & vbCrLf & " from"
'                    lgStrSQL = lgStrSQL & vbCrLf & " "
'                    lgStrSQL = lgStrSQL & vbCrLf & " (  "
'                    'lgStrSQL = lgStrSQL & vbCrLf & "    -- (1) 생판요청(실수량)"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    select  a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			a.project_code as tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			'생판요청' as type,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '01' then sum(a.plan_qty) else 0 end) as qty01,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '02' then sum(a.plan_qty) else 0 end) as qty02,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '03' then sum(a.plan_qty) else 0 end) as qty03,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '04' then sum(a.plan_qty) else 0 end) as qty04,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '05' then sum(a.plan_qty) else 0 end) as qty05,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '06' then sum(a.plan_qty) else 0 end) as qty06,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '07' then sum(a.plan_qty) else 0 end) as qty07,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '08' then sum(a.plan_qty) else 0 end) as qty08,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '09' then sum(a.plan_qty) else 0 end) as qty09,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '10' then sum(a.plan_qty) else 0 end) as qty10,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '11' then sum(a.plan_qty) else 0 end) as qty11,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '12' then sum(a.plan_qty) else 0 end) as qty12,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '13' then sum(a.plan_qty) else 0 end) as qty13,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '14' then sum(a.plan_qty) else 0 end) as qty14,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '15' then sum(a.plan_qty) else 0 end) as qty15,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '16' then sum(a.plan_qty) else 0 end) as qty16,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '17' then sum(a.plan_qty) else 0 end) as qty17,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '18' then sum(a.plan_qty) else 0 end) as qty18,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '19' then sum(a.plan_qty) else 0 end) as qty19,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '20' then sum(a.plan_qty) else 0 end) as qty20,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '21' then sum(a.plan_qty) else 0 end) as qty21,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '22' then sum(a.plan_qty) else 0 end) as qty22,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '23' then sum(a.plan_qty) else 0 end) as qty23,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '24' then sum(a.plan_qty) else 0 end) as qty24,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '25' then sum(a.plan_qty) else 0 end) as qty25,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt)"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '26' then sum(a.plan_qty) else 0 end) as qty26,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '27' then sum(a.plan_qty) else 0 end) as qty27,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '28' then sum(a.plan_qty) else 0 end) as qty28,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '29' then sum(a.plan_qty) else 0 end) as qty29,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '30' then sum(a.plan_qty) else 0 end) as qty30,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case day(a.dlvy_plan_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when '31' then sum(a.plan_qty) else 0 end) as qty31,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			(select case convert(varchar(07), a.dlvy_plan_dt, 120) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			when @ProdPlanMonth then sum(a.plan_qty) else 0 end) as month_1,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as month_2,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as month_3"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    from sales_item_req_plan_ko441 a (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    inner join b_item b (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    on a.item_cd = b.item_cd"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    where convert(varchar(7), a.dlvy_plan_dt ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    group by a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.project_code,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.dlvy_plan_dt"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    union all"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    'lgStrSQL = lgStrSQL & vbCrLf & "    -- (2) 생판요청(폼)"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    select  a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			a.project_code as tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			'생산계획' as type,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as qty01, 0 as qty02, 0 as qty03, 0 as qty04, 0 as qty05,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as qty06, 0 as qty07, 0 as qty08, 0 as qty09, 0 as qty10,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as qty11, 0 as qty12, 0 as qty13, 0 as qty14, 0 as qty15,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as qty16, 0 as qty17, 0 as qty18, 0 as qty19, 0 as qty20,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as qty21, 0 as qty22, 0 as qty23, 0 as qty24, 0 as qty25,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as qty26, 0 as qty27, 0 as qty28, 0 as qty29, 0 as qty30,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			0 as qty31, 0 as month_1, 0 as month_2, 0 as month_3"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    from sales_item_req_plan_ko441 a (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    inner join b_item b (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    on a.item_cd = b.item_cd"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    where convert(varchar(7), dateadd(m,-1,a.dlvy_plan_dt) ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    or convert(varchar(7), dateadd(m,-2,a.dlvy_plan_dt) ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    or convert(varchar(7), a.dlvy_plan_dt ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    group by a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "			b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			a.project_code,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "			a.dlvy_plan_dt"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    union all"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    'lgStrSQL = lgStrSQL & vbCrLf & "    -- (2) 생판요청(두달)"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    select  a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.project_code as tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        '생판요청' as type,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty01, 0 as qty02, 0 as qty03, 0 as qty04, 0 as qty05,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty06, 0 as qty07, 0 as qty08, 0 as qty09, 0 as qty10,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty11, 0 as qty12, 0 as qty13, 0 as qty14, 0 as qty15,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty16, 0 as qty17, 0 as qty18, 0 as qty19, 0 as qty20,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty21, 0 as qty22, 0 as qty23, 0 as qty24, 0 as qty25,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty26, 0 as qty27, 0 as qty28, 0 as qty29, 0 as qty30,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty31, 0 as month_1 ,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case convert(varchar(07), dateadd(m,-1,a.dlvy_plan_dt), 120) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when @ProdPlanMonth then sum(a.plan_qty) else 0 end) as month_2,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case convert(varchar(07), dateadd(m,-2,a.dlvy_plan_dt), 120) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when @ProdPlanMonth then sum(a.plan_qty) else 0 end) as month_3"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    from sales_item_req_plan_ko441 a (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    inner join b_item b (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    on a.item_cd = b.item_cd"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    where convert(varchar(7), dateadd(m,-1,a.dlvy_plan_dt) ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    or convert(varchar(7), dateadd(m,-2,a.dlvy_plan_dt) ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    group by a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.project_code,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.dlvy_plan_dt"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    union all"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    'lgStrSQL = lgStrSQL & vbCrLf & "    -- (4) 생산계획(실수량)"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    select  a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        '생산계획' as type,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '01' then sum(a.mps_qty) else 0 end) as qty01,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '02' then sum(a.mps_qty) else 0 end) as qty02,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '03' then sum(a.mps_qty) else 0 end) as qty03,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '04' then sum(a.mps_qty) else 0 end) as qty04,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '05' then sum(a.mps_qty) else 0 end) as qty05,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '06' then sum(a.mps_qty) else 0 end) as qty06,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '07' then sum(a.mps_qty) else 0 end) as qty07,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '08' then sum(a.mps_qty) else 0 end) as qty08,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '09' then sum(a.mps_qty) else 0 end) as qty09,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '10' then sum(a.mps_qty) else 0 end) as qty10,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '11' then sum(a.mps_qty) else 0 end) as qty11,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '12' then sum(a.mps_qty) else 0 end) as qty12,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '13' then sum(a.mps_qty) else 0 end) as qty13,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '14' then sum(a.mps_qty) else 0 end) as qty14,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '15' then sum(a.mps_qty) else 0 end) as qty15,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '16' then sum(a.mps_qty) else 0 end) as qty16,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '17' then sum(a.mps_qty) else 0 end) as qty17,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '18' then sum(a.mps_qty) else 0 end) as qty18,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '19' then sum(a.mps_qty) else 0 end) as qty19,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '20' then sum(a.mps_qty) else 0 end) as qty20,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '21' then sum(a.mps_qty) else 0 end) as qty21,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '22' then sum(a.mps_qty) else 0 end) as qty22,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '23' then sum(a.mps_qty) else 0 end) as qty23,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '24' then sum(a.mps_qty) else 0 end) as qty24,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '25' then sum(a.mps_qty) else 0 end) as qty25,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt)"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '26' then sum(a.mps_qty) else 0 end) as qty26,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '27' then sum(a.mps_qty) else 0 end) as qty27,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '28' then sum(a.mps_qty) else 0 end) as qty28,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '29' then sum(a.mps_qty) else 0 end) as qty29,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '30' then sum(a.mps_qty) else 0 end) as qty30,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case day(a.mps_dt) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when '31' then sum(a.mps_qty) else 0 end) as qty31,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case convert(varchar(07), a.mps_dt, 120) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when @ProdPlanMonth then sum(a.mps_qty) else 0 end) as month_1,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as month_2,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as month_3"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    from p_mps a (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    inner join b_item b (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    on a.item_cd = b.item_cd"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    where convert(varchar(7), a.mps_dt ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    group by a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.mps_dt"
'                    lgStrSQL = lgStrSQL & vbCrLf & " "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    union all"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    'lgStrSQL = lgStrSQL & vbCrLf & "    -- (3) 생산계획(폼)"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    select  a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        '생판요청' as type,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty01, 0 as qty02, 0 as qty03, 0 as qty04, 0 as qty05,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty06, 0 as qty07, 0 as qty08, 0 as qty09, 0 as qty10,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty11, 0 as qty12, 0 as qty13, 0 as qty14, 0 as qty15,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty16, 0 as qty17, 0 as qty18, 0 as qty19, 0 as qty20,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty21, 0 as qty22, 0 as qty23, 0 as qty24, 0 as qty25,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty26, 0 as qty27, 0 as qty28, 0 as qty29, 0 as qty30,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty31, 0 as month_1, 0 as month_2, 0 as month_3"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    from p_mps a (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    inner join b_item b (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    on a.item_cd = b.item_cd"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    where convert(varchar(7), dateadd(m,-1,a.mps_dt) ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    or convert(varchar(7), dateadd(m,-2,a.mps_dt) ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    or convert(varchar(7), a.mps_dt ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    group by a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.mps_dt"
'                    lgStrSQL = lgStrSQL & vbCrLf & " "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    union all"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    'lgStrSQL = lgStrSQL & vbCrLf & "    -- (2) 생산계획(두달)"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    select  a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        '생산계획' as type,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty01, 0 as qty02, 0 as qty03, 0 as qty04, 0 as qty05,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty06, 0 as qty07, 0 as qty08, 0 as qty09, 0 as qty10,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty11, 0 as qty12, 0 as qty13, 0 as qty14, 0 as qty15,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty16, 0 as qty17, 0 as qty18, 0 as qty19, 0 as qty20,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty21, 0 as qty22, 0 as qty23, 0 as qty24, 0 as qty25,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty26, 0 as qty27, 0 as qty28, 0 as qty29, 0 as qty30,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        0 as qty31, 0 as month_1, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case convert(varchar(07), dateadd(m,-1,a.mps_dt), 120) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when @ProdPlanMonth then sum(a.mps_qty) else 0 end) as month_2,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        (select case convert(varchar(07), dateadd(m,-2,a.mps_dt), 120) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        when @ProdPlanMonth then sum(a.mps_qty) else 0 end) as month_3"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    from p_mps a (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    inner join b_item b (nolock) "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    on a.item_cd = b.item_cd"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    where convert(varchar(7), dateadd(m,-1,a.mps_dt) ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    or convert(varchar(7), dateadd(m,-2,a.mps_dt) ,120) = @ProdPlanMonth"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    group by a.item_cd, "
'                    lgStrSQL = lgStrSQL & vbCrLf & "        b.item_nm,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.tracking_no,"
'                    lgStrSQL = lgStrSQL & vbCrLf & "        a.mps_dt"
'                    lgStrSQL = lgStrSQL & vbCrLf & "    "
'                    lgStrSQL = lgStrSQL & vbCrLf & "    ) x"
'                    lgStrSQL = lgStrSQL & vbCrLf & " WHERE 1 = 1 " & whereQuery
'                    lgStrSQL = lgStrSQL & vbCrLf & " group by x.item_cd, x.item_nm, x.tracking_no, x.type "
'                    lgStrSQL = lgStrSQL & vbCrLf & " order by x.item_cd, x.item_nm, x.tracking_no, x.type desc "

           End Select
    End Select
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
'    On Error Resume Next                                                             '☜: Protect system from crashing
'    Err.Clear                                                                        '☜: Clear Error status

    Dim i
    Dim plantCd, itemCd, TrackingNo, mpsDt, mpsQty, mpsType, entUser

    plantCd = FilterVar(UCase(arrColVal(2)), "''", "S")
    itemCd = FilterVar(UCase(arrColVal(3)), "''", "S")
    TrackingNo = FilterVar(UCase(arrColVal(4)), "''", "S")
    entUser = FilterVar(gUsrId, "''", "S")

'    If UCase(arrColVal(4)) = "5MMPS" Then
'        mpsType = FilterVar("M", "''", "S")
'    ElseIf UCase(arrColVal(4)) = "6pmps" Then
'        mpsType = FilterVar("O", "''", "S")
'    Else
'        mpsType = FilterVar("", "''", "S")
'    End If

    For i = 0 To 30
    
        If UNIConvNum(arrColVal(3 * i + 6), 0) <> UNIConvNum(arrColVal(3 * i + 7), 0) Then
			If len(Replace(arrColVal(3 * i + 5), "-", "")) < 2 Then
				strdt = "0" + Replace(arrColVal(3 * i + 5), "-", "")
			End If
            
            mpsDt = FilterVar(Replace(arrColVal(3 * i + 5), "-", ""), "''", "S")
            mpsQty = UNIConvNum(arrColVal(3 * i + 6), 0)
            mpsType = FilterVar("", "''", "S")
            
            ' 확정여부 체크 ------------------------------------------------------------------------------------
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & vbCrLf & " select isnull(count(*),0) as nvnum from p_mps (nolock) "
			lgStrSQL = lgStrSQL & vbCrLf & " where plant_cd = " & plantCd
			lgStrSQL = lgStrSQL & vbCrLf & " and item_cd = " & itemCd
			lgStrSQL = lgStrSQL & vbCrLf & " and tracking_no = " & TrackingNo
			lgStrSQL = lgStrSQL & vbCrLf & " and convert(varchar(8), mps_dt, 112) = " & mpsDt
			lgStrSQL = lgStrSQL & vbCrLf & " and isnull(mps_confirm_flg,'N') = 'Y'"
			
			If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
			    Call DisplayMsgBox("P43002", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
			    Call SetErrorStatus()
			Else
				If UNIConvNum(lgObjRs("nvnum"),0) > 0 Then
					Call DisplayMsgBox("P43002", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
					Call SetErrorStatus()
				Else
					'미확정된 자료라면
			        Call SubBizSaveMultiUpdateReal(plantCd, itemCd, TrackingNo, mpsDt, mpsQty, mpsType, entUser)
			    End If
			End If
			
			Call SubHandleError("MR", lgObjConn, lgObjRs, Err)
			Call SubCloseRs(lgObjRs)
			'----------------------------------------------------------------------------------------------------
        End If
    Next
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdateReal(plantCd, itemCd, TrackingNo, mpsDt, mpsQty, mpsType, entUser)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = " EXEC usp_mps_ko441 " & plantCd & "," & itemCd & "," & TrackingNo & "," & mpsDt & "," & mpsQty & "," & mpsType & "," & entUser

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

<OBJECT RUNAT=server PROGID=ADODB.Recordset id=adoRec></OBJECT>