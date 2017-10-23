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
        Case CStr(UID_M0022)                                                         '☜: Save,Update
            Call SubBizSaveMulti1()
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
Sub SubBizSaveMultiUpdate1(arrColVal)

    Dim lgStrSQL
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL =              " UPDATE	M_ISSUE_REQ_DTL_KO441 "
    lgStrSQL = lgStrSQL &   " SET		REMARK  		= " & FilterVar(            arrColVal(2) ,"","S")  & ", "
    lgStrSQL = lgStrSQL &   " 	    	REQ_QTY  		= " & FilterVar(            arrColVal(6) ,"","D")  & ", "
    lgStrSQL = lgStrSQL &   "   		ISSUE_QTY  		= " & FilterVar(            arrColVal(7) ,"","D")  & " "
    lgStrSQL = lgStrSQL &   " WHERE	    PLANT_CD        = " & FilterVar(            arrColVal(4) ,"","S")  & " "
    lgStrSQL = lgStrSQL &   " AND		ISSUE_REQ_NO    = " & FilterVar(            arrColVal(3) ,"","S")  & " "
    lgStrSQL = lgStrSQL &   " AND		ITEM_SEQ        = " & FilterVar(            arrColVal(5) ,"","D")  & " "


    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords


    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
    

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
' Name : SubBizSaveMulti1
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti1()

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
                    Call SubBizSaveMultiCreate1(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate1(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete1(arrColVal)                            '☜: Delete
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
            Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
       Response.Write  " <Script Language=vbscript>                                  " & vbCr
       Response.Write  "    Parent.DBsaveDel   " & vbCr      
       Response.Write  " </Script>             " & vbCr
            Call SetErrorStatus()
        Else

            Call SubSkipRs(lgObjRs, C_SHEETMAXROWS_D * lgStrPrevKey)

            lgstrData = ""

            iDx = 1

            Do While Not lgObjRs.EOF

                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
        		lgstrData = lgstrData & Chr(11) & ""								'2                   
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_UNIT"))
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("REQ_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("ISSUE_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRODT_ORDER_NO"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK1"))
                
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_REQ_NO"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRNS_TYPE"))   
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REQ_DT"))      
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_TYPE"))  
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))     
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))      
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_SEQ"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONFIRM_FLAG"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PoTypeCdNm"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NM"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
				lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("LIMIT_DT"))									'☆: Planned Start Date
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOC"))

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
                    lgStrSQL = lgStrSQL & vbCrLf & " SELECT	B.ITEM_SEQ,                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B.ITEM_CD,                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = B.ITEM_CD) ITEM_NM,             "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT BASIC_UNIT FROM B_ITEM WHERE ITEM_CD = B.ITEM_CD) BASIC_UNIT,       "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B.REQ_QTY,                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B.ISSUE_QTY,                                                                "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B.PRODT_ORDER_NO,                                                           "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		B.REMARK REMARK1,                                                           "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.ISSUE_REQ_NO,                                                             "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.TRNS_TYPE,                                                                "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		CONVERT(CHAR(10), A.REQ_DT, 120) REQ_DT,                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.ISSUE_TYPE,                                                               "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.MOV_TYPE,                                                                 "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.DEPT_CD,                                                                  "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.EMP_NO,                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.REMARK ,                                                                   "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.CONFIRM_FLAG,                                                                    "
                    lgStrSQL = lgStrSQL & vbCrLf & "	    [DBO].[ufn_GetMovType](TRNS_TYPE, MOV_TYPE ) PoTypeCdNm , "
                    lgStrSQL = lgStrSQL & vbCrLf & "	    (select name from haa010t where emp_no = A.EMP_NO) EMP_NM, "
                    lgStrSQL = lgStrSQL & vbCrLf & "	    dbo.ufn_getDeptName(dbo.ufn_H_get_dept_cd(A.EMP_NO, getdate()), getdate()) DEPT_NM, "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		b.LIMIT_DT   ,                                                             "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		A.LOC                                                                    "
                    lgStrSQL = lgStrSQL & vbCrLf & " FROM	M_ISSUE_REQ_HDR_KO441 A,                                                    "
                    lgStrSQL = lgStrSQL & vbCrLf & " 		M_ISSUE_REQ_DTL_KO441 B                                                     "
                    lgStrSQL = lgStrSQL & vbCrLf & " WHERE	A.PLANT_CD		=	B.PLANT_CD                                              "
                    lgStrSQL = lgStrSQL & vbCrLf & " AND		A.ISSUE_REQ_NO	=	B.ISSUE_REQ_NO                                      "
                    lgStrSQL = lgStrSQL & vbCrLf & " AND		A.PLANT_CD		=	" & FilterVar(Trim(Request("txtPlantCd")) ,"","S")  & " "
                    lgStrSQL = lgStrSQL & vbCrLf & " AND		A.ISSUE_REQ_NO	=	" & FilterVar(Trim(Request("txtPoNo")) ,"","S")  & " "


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

    Dim lgStrSQL
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL =              " UPDATE	M_ISSUE_REQ_DTL_KO441 "
    lgStrSQL = lgStrSQL &   " SET		REMARK  		= " & FilterVar(            arrColVal(2) ,"","S")  & ", "
    lgStrSQL = lgStrSQL &   " 	    	REQ_QTY  		= " & FilterVar(            arrColVal(6) ,"","D")  & ", "
    lgStrSQL = lgStrSQL &   "   		ISSUE_QTY  		= " & FilterVar(            arrColVal(7) ,"","D")  & " "
    lgStrSQL = lgStrSQL &   " WHERE	    PLANT_CD        = " & FilterVar(            arrColVal(4) ,"","S")  & " "
    lgStrSQL = lgStrSQL &   " AND		ISSUE_REQ_NO    = " & FilterVar(            arrColVal(3) ,"","S")  & " "
    lgStrSQL = lgStrSQL &   " AND		ITEM_SEQ        = " & FilterVar(            arrColVal(5) ,"","D")  & " "


    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords


    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
    
''    On Error Resume Next                                                             '☜: Protect system from crashing
''    Err.Clear                                                                        '☜: Clear Error status
'
'    Dim i
'    Dim plantCd, itemCd, TrackingNo, mpsDt, mpsQty, mpsType, entUser
'
'    plantCd = FilterVar(UCase(arrColVal(2)), "''", "S")
'    itemCd = FilterVar(UCase(arrColVal(3)), "''", "S")
'    TrackingNo = FilterVar(UCase(arrColVal(4)), "''", "S")
'    entUser = FilterVar(gUsrId, "''", "S")
'
''    If UCase(arrColVal(4)) = "5MMPS" Then
''        mpsType = FilterVar("M", "''", "S")
''    ElseIf UCase(arrColVal(4)) = "6pmps" Then
''        mpsType = FilterVar("O", "''", "S")
''    Else
''        mpsType = FilterVar("", "''", "S")
''    End If
'
'    For i = 0 To 30
'    
'        If UNIConvNum(arrColVal(3 * i + 6), 0) <> UNIConvNum(arrColVal(3 * i + 7), 0) Then
'			If len(Replace(arrColVal(3 * i + 5), "-", "")) < 2 Then
'				strdt = "0" + Replace(arrColVal(3 * i + 5), "-", "")
'			End If
'            
'            mpsDt = FilterVar(Replace(arrColVal(3 * i + 5), "-", ""), "''", "S")
'            mpsQty = UNIConvNum(arrColVal(3 * i + 6), 0)
'            mpsType = FilterVar("", "''", "S")
'            
'            ' 확정여부 체크 ------------------------------------------------------------------------------------
'			lgStrSQL = ""
'			lgStrSQL = lgStrSQL & vbCrLf & " select isnull(count(*),0) as nvnum from p_mps (nolock) "
'			lgStrSQL = lgStrSQL & vbCrLf & " where plant_cd = " & plantCd
'			lgStrSQL = lgStrSQL & vbCrLf & " and item_cd = " & itemCd
'			lgStrSQL = lgStrSQL & vbCrLf & " and tracking_no = " & TrackingNo
'			lgStrSQL = lgStrSQL & vbCrLf & " and convert(varchar(8), mps_dt, 112) = " & mpsDt
'			lgStrSQL = lgStrSQL & vbCrLf & " and isnull(mps_confirm_flg,'N') = 'Y'"
'			
'			If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
'			    Call DisplayMsgBox("P43002", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
'			    Call SetErrorStatus()
'			Else
'				If UNIConvNum(lgObjRs("nvnum"),0) > 0 Then
'					Call DisplayMsgBox("P43002", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
'					Call SetErrorStatus()
'				Else
'					'미확정된 자료라면
'			        Call SubBizSaveMultiUpdateReal(plantCd, itemCd, TrackingNo, mpsDt, mpsQty, mpsType, entUser)
'			    End If
'			End If
'			
'			Call SubHandleError("MR", lgObjConn, lgObjRs, Err)
'			Call SubCloseRs(lgObjRs)
'			'----------------------------------------------------------------------------------------------------
'        End If
'    Next
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