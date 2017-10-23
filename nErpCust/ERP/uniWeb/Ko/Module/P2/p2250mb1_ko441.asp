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
    Dim lgstrData_header
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

                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bp_cd"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bp_nm"))

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

'------------------------------ add :: begin
        Call SubMakeSQLStatementsh("MR", lgKeyStream, "X", C_EQ)                                 '☆ : Make sql statements

        If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
            lgStrPrevKey = ""
            Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
            Call SetErrorStatus()
        Else

            Call SubSkipRs(lgObjRs, C_SHEETMAXROWS_D * lgStrPrevKey)

            lgstrData_header = ""

            iDx = 1

            Do While Not lgObjRs.EOF

                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt1"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt2"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt3"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt4"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt5"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt6"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt7"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt8"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt9"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt10"))

                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt11"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt12"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt13"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt14"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt15"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt16"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt17"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt18"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt19"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt20"))

                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt21"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt22"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt23"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt24"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt25"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt26"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt27"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt28"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt29"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt30"))

                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt31"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt32"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt33"))
                lgstrData_header = lgstrData_header & Chr(11) & ConvSPChars(lgObjRs("dt34"))


                lgstrData_header = lgstrData_header & Chr(11) & C_SHEETMAXROWS_D + iDx
		        lgstrData_header = lgstrData_header & Chr(11) & Chr(12)



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

'-----------------------------------add :: end


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
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @version       Varchar(14)                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " set @ProdPlanMonth = " & FilterVar(pCode(0), "''", "S")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @iPeriod       = " & FilterVar(pCode(4), "''", "D")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @PlantCd       = " & FilterVar(pCode(5), "''", "S")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @version       = " & FilterVar(pCode(6), "''", "S")
'                    lgStrSQL = lgStrSQL & vbCrLf & "  set @ProdPlanMonth = '2007-12-01'                                                                                                                                                "
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " SELECT	DISTINCT                                                                                                                                                         "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.bp_CD,                                                                                                                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(select bp_nm from b_biz_partner where bp_cd = a.bp_cd ) bp_nm,                                                                                                                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.ITEM_CD,                                                                                                                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B.ITEM_NM,                                                                                                                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		'' tracking_no,                                                                                                                                                            "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		'수주계획(수주일)' TYPE,                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 1 and item_cd = a.item_cd and version = a.version) QTY01,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 2 and item_cd = a.item_cd and version = a.version) QTY02,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 3 and item_cd = a.item_cd and version = a.version) QTY03,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 4 and item_cd = a.item_cd and version = a.version) QTY04,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 5 and item_cd = a.item_cd and version = a.version) QTY05,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 6 and item_cd = a.item_cd and version = a.version) QTY06,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 7 and item_cd = a.item_cd and version = a.version) QTY07,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 8 and item_cd = a.item_cd and version = a.version) QTY08,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 9 and item_cd = a.item_cd and version = a.version) QTY09,     "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 10 and item_cd = a.item_cd and version = a.version) QTY10,    "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 11 and item_cd = a.item_cd and version = a.version) QTY11,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 12 and item_cd = a.item_cd and version = a.version) QTY12,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 13 and item_cd = a.item_cd and version = a.version) QTY13,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 14 and item_cd = a.item_cd and version = a.version) QTY14,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 15 and item_cd = a.item_cd and version = a.version) QTY15,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 16 and item_cd = a.item_cd and version = a.version) QTY16,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 17 and item_cd = a.item_cd and version = a.version) QTY17,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 18 and item_cd = a.item_cd and version = a.version) QTY18,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 19 and item_cd = a.item_cd and version = a.version) QTY19,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 20 and item_cd = a.item_cd and version = a.version) QTY20,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 21 and item_cd = a.item_cd and version = a.version) QTY21,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 22 and item_cd = a.item_cd and version = a.version) QTY22,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 23 and item_cd = a.item_cd and version = a.version) QTY23,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 24 and item_cd = a.item_cd and version = a.version) QTY24,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 25 and item_cd = a.item_cd and version = a.version) QTY25,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 26 and item_cd = a.item_cd and version = a.version) QTY26,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 27 and item_cd = a.item_cd and version = a.version) QTY27,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 28 and item_cd = a.item_cd and version = a.version) QTY28,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 29 and item_cd = a.item_cd and version = a.version) QTY29,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 30 and item_cd = a.item_cd and version = a.version) QTY30,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 31 and item_cd = a.item_cd and version = a.version) QTY31,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 32 and item_cd = a.item_cd and version = a.version) QTY32,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 33 and item_cd = a.item_cd and version = a.version) QTY33,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 34 and item_cd = a.item_cd and version = a.version) QTY34,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 35 and item_cd = a.item_cd and version = a.version) QTY35,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 36 and item_cd = a.item_cd and version = a.version) QTY36,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 37 and item_cd = a.item_cd and version = a.version) QTY37,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 38 and item_cd = a.item_cd and version = a.version) QTY38,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 39 and item_cd = a.item_cd and version = a.version) QTY39,   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT isnull(PLAN_QTY,0) FROM prod_Item_PLAN_KO441 where seq_no = 40 and item_cd = a.item_cd and version = a.version) QTY40,   "
                    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_1,                                                                                                                                                                "
                    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_2,                                                                                                                                                                "
                    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_3                                                                                                                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " FROM	prod_Item_PLAN_KO441 A,                                                                                                                                               "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B_ITEM B                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " WHERE	A.ITEM_CD = B.ITEM_CD                                                                                                                                                      "
                    lgStrSQL = lgStrSQL & vbCrLf & " AND	A.PROJECT_CODE = @PlantCd                                                                                                                                                      "
                    lgStrSQL = lgStrSQL & vbCrLf & " AND	A.version = @version                                                                                                                                                      "
                    lgStrSQL = lgStrSQL & vbCrLf & "  AND    CONVERT(CHAR(10), A.DLVY_PLAN_DT, 120) BETWEEN @ProdPlanMonth AND DATEADD(day, @iPeriod, @ProdPlanMonth)   "


           End Select
    End Select
End Sub

Sub SubMakeSQLStatementsh(pDataType, pCode, pCode1, pComp)
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

           Select Case Mid(pDataType,2,1)
               Case "R"


					lgStrSQL = ""
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @ProdPlanMonth DATETIME                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @iPeriod       NUMERIC(10,0)                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @PlantCd       Varchar(20)                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "  declare @version       Varchar(14)                                                                                                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " set @ProdPlanMonth = " & FilterVar(pCode(0), "''", "S")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @iPeriod       = " & FilterVar(pCode(4), "''", "D")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @PlantCd       = " & FilterVar(pCode(5), "''", "S")
                    lgStrSQL = lgStrSQL & vbCrLf & " set @version       = " & FilterVar(pCode(6), "''", "S")
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & "  select "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 1 )  dt1 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 2 )  dt2 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 3 )  dt3 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 4 )  dt4 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 5 )  dt5 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 6 )  dt6 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 7 )  dt7 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 8 )  dt8 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 9 )  dt9 ,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 10)  dt10,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                         "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 11 )  dt11 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 12 )  dt12 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 13 )  dt13 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 14 )  dt14 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 15 )  dt15 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 16 )  dt16 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 17 )  dt17 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 18 )  dt18 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 19 )  dt19 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 20)  dt20,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                         "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 21 )  dt21 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 22 )  dt22 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 23 )  dt23 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 24 )  dt24 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 25 )  dt25 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 26 )  dt26 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 27 )  dt27 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 28 )  dt28 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 29 )  dt29 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 30)  dt30,    "
                    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                         "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 31 )  dt31 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 32 )  dt32 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 33 )  dt33 ,  "
                    lgStrSQL = lgStrSQL & vbCrLf & "        (select distinct convert(char(10), dlvy_plan_dt, 120)    from prod_Item_Plan_Ko441   where version = @version and seq_no = 34 )  dt34    "

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

    lgStrSQL = "" '20080624::hanc EXEC usp_mps_ko441 " & plantCd & "," & itemCd & "," & TrackingNo & "," & mpsDt & "," & mpsQty & "," & mpsType & "," & entUser

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
                    .SetHeader("<%=lgstrData_header%>")            '20080626::hanc    

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