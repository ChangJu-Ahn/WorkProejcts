<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)
    
    Dim i               '6개 Sheet모두 데이터가 없는지 체크후 메세지 
    i=0

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

    Dim iDx
    Dim iKey1, iKey2, iKey3, iKey4, iKey5

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    iKey2 = FilterVar(lgKeyStream(1), "''", "S")
    iKey3 = FilterVar(lgKeyStream(2), "''", "S")

	If lgKeyStream(1) = "1" Then
		iKey4 = "" & FilterVar("P", "''", "S") & " "
	ElseIf (lgKeyStream(1) >= "2" And lgKeyStream(1) <= "9") Then
		iKey4 = "" & FilterVar("Q", "''", "S") & " "
	End If

	If (UCase(lgKeyStream(1)) = "P") OR (UCase(lgKeyStream(1)) = "Q") Then 
		iKey4 = FilterVar(UCase(lgKeyStream(1)), "''", "S")
	End if

	If lgKeyStream(1) = "1" OR UCase(lgKeyStream(1)) = "P" Then
        iKey5 = " From  HDF040T a, HDA010T b"
    Else
        iKey5 = " From  HDF041T a, HDA010T b"
	End if

    Call SubBizQueryMulti(iKey1,iKey2,iKey3,iKey4,iKey5)

End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(iKey1,iKey2,iKey3,iKey4,iKey5)
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgstrData   = ""
    lgstrData1  = ""
    lgstrData2  = ""
    lgstrData3  = ""
    lgstrData4  = ""
    lgstrData5  = ""

	For iDx = 1 To 6
	    lgStrSQL = ""
		lgCurrentSpd = Cstr(iDx)

		Call SubMakeSQLStatements("MR",iKey1,C_EQ,iKey2,C_EQ,iKey3,C_EQ,iKey4,C_EQ,iKey5)  '☆ : Make sql statements

		Call SubBizQueryMultiData(lgCurrentSpd)                                          '☆ : Save Array Data
	Next

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizQueryMultiData
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMultiData(pCurrentSpd)
     
    Dim iDx, istrData
    On Error Resume Next                                                             '☜: Protect system from crashing
             
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        i=i+1           '6개 Sheet모두 데이터가 없는지 체크후 메세지 
        If i=6 Then
            Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
            Call SetErrorStatus()
        End If
    Else
        iDx = 1
        
        Do While Not lgObjRs.EOF

            istrData = istrData & Chr(11) & ConvSPChars(lgObjRs(0))
            istrData = istrData & Chr(11) & UNINumClientFormat(lgObjRs(1), ggAmtOfMoney.DecPoint,0)

            istrData = istrData & Chr(11) & lgLngMaxRow + iDx
            istrData = istrData & Chr(11) & Chr(12)
     
	        lgObjRs.MoveNext

            iDx =  iDx + 1
        Loop 

		If pCurrentSpd = "1" Then
		    lgstrData  = iStrData
		ElseIf pCurrentSpd = "2" Then
		    lgstrData1 = iStrData
		ElseIf pCurrentSpd = "3" Then
		    lgstrData2 = iStrData
		ElseIf pCurrentSpd = "4" Then
		    lgstrData3 = iStrData
		ElseIf pCurrentSpd = "5" Then
		    lgstrData4 = iStrData
		ElseIf pCurrentSpd = "6" Then
		    lgstrData5 = iStrData
		End If

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
Sub SubBizSaveMultiCreate(arrColVal)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode1,pComp1,pCode2,pComp2,pCode3,pComp3,pCode4,pComp4,pCode5)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                '☜: Clear Error status
    lgStrSQL = ""

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           Select Case Mid(pDataType,2,1)
               Case "R"               
                    Select Case lgCurrentSpd
                       Case "1"
							lgStrSQL = "Select  ALLOW_NM, ALLOW"
                            lgStrSQL = lgStrSQL & pCode5
			                lgStrSQL = lgStrSQL & " WHERE PAY_YYMM  " & pComp1 & pCode1
                            lgStrSQL = lgStrSQL & "   AND PROV_TYPE " & pComp2 & pCode2
                            lgStrSQL = lgStrSQL & "   AND EMP_NO    " & pComp3 & pCode3
                            lgStrSQL = lgStrSQL & "   AND b.CODE_TYPE = " & FilterVar("1", "''", "S") & "  "
                            lgStrSQL = lgStrSQL & "   AND b.ALLOW_CD = a.ALLOW_CD "
                            lgStrSQL = lgStrSQL & " ORDER BY ALLOW_SEQ"

                       Case "2"
							lgStrSQL = "Select  ALLOW_NM, ALLOW"
                            lgStrSQL = lgStrSQL & pCode5
			                lgStrSQL = lgStrSQL & " WHERE PAY_YYMM  " & pComp1 & pCode1
                            lgStrSQL = lgStrSQL & "   AND PROV_TYPE " & pComp4 & pCode4
                            lgStrSQL = lgStrSQL & "   AND EMP_NO    " & pComp3 & pCode3
                            lgStrSQL = lgStrSQL & "   AND b.CODE_TYPE = " & FilterVar("1", "''", "S") & "  "
                            lgStrSQL = lgStrSQL & "   AND b.ALLOW_CD = a.ALLOW_CD "
                            lgStrSQL = lgStrSQL & " ORDER BY ALLOW_SEQ"
						Case "3"
							lgStrSQL = "Select IsNull(X.ALLOW_NM, Y.ALLOW_NM), IsNull(Y.ALLOW,0)-IsNull(X.ALLOW,0)"
                            lgStrSQL = lgStrSQL & " From  (SELECT a.ALLOW_CD ALLOW_CD, ALLOW_NM, ALLOW, ALLOW_SEQ"
                            lgStrSQL = lgStrSQL & pCode5
			                lgStrSQL = lgStrSQL & "         WHERE PAY_YYMM  " & pComp1 & pCode1
                            lgStrSQL = lgStrSQL & "           AND PROV_TYPE " & pComp2 & pCode2
                            lgStrSQL = lgStrSQL & "           AND EMP_NO    " & pComp3 & pCode3 
                            lgStrSQL = lgStrSQL & "           AND b.CODE_TYPE = " & FilterVar("1", "''", "S") & "  "
                            lgStrSQL = lgStrSQL & "           AND b.ALLOW_CD = a.ALLOW_CD) AS X"
                            lgStrSQL = lgStrSQL & "       FULL OUTER JOIN"
                            lgStrSQL = lgStrSQL & "       (SELECT a.ALLOW_CD ALLOW_CD, ALLOW_NM, ALLOW, ALLOW_SEQ"
                            lgStrSQL = lgStrSQL & pCode5
			                lgStrSQL = lgStrSQL & "         WHERE PAY_YYMM  " & pComp1 & pCode1
                            lgStrSQL = lgStrSQL & "           AND PROV_TYPE " & pComp4 & pCode4
                            lgStrSQL = lgStrSQL & "           AND EMP_NO    " & pComp3 & pCode3 
                            lgStrSQL = lgStrSQL & "           AND b.CODE_TYPE = " & FilterVar("1", "''", "S") & "  "
                            lgStrSQL = lgStrSQL & "           AND b.ALLOW_CD = a.ALLOW_CD) AS Y"
                            lgStrSQL = lgStrSQL & "   ON (X.ALLOW_CD = Y.ALLOW_CD) "
                            lgStrSQL = lgStrSQL & "  ORDER BY IsNull(X.ALLOW_SEQ,Y.ALLOW_SEQ) "
                       Case "4"
                            lgStrSQL = "Select  C.ALLOW_NM, A.SUB_AMT"
                            lgStrSQL = lgStrSQL & " From  HDF060T A, HDF060T B, HDA010T C"
                            lgStrSQL = lgStrSQL & " WHERE A.SUB_YYMM  " & pComp1 & pCode1
                            lgStrSQL = lgStrSQL & "   AND A.SUB_TYPE  " & pComp2 & pCode2
                            lgStrSQL = lgStrSQL & "   AND A.EMP_NO    " & pComp3 & pCode3
                            lgStrSQL = lgStrSQL & "   AND A.SUB_YYMM  = B.SUB_YYMM"
                            lgStrSQL = lgStrSQL & "   AND B.SUB_TYPE  " & pComp4 & pCode4
                            lgStrSQL = lgStrSQL & "   AND A.EMP_NO    = B.EMP_NO"
                            lgStrSQL = lgStrSQL & "   AND A.SUB_CD    = B.SUB_CD"
                            lgStrSQL = lgStrSQL & "   AND C.CODE_TYPE = " & FilterVar("2", "''", "S") & " "
                            lgStrSQL = lgStrSQL & "   AND C.ALLOW_CD  = A.SUB_CD "
                            lgStrSQL = lgStrSQL & " ORDER BY ALLOW_SEQ"

                       Case "5"
                            lgStrSQL = "Select  C.ALLOW_NM, B.SUB_AMT"
                            lgStrSQL = lgStrSQL & " From  HDF060T A, HDF060T B, HDA010T C"
                            lgStrSQL = lgStrSQL & " WHERE A.SUB_YYMM  " & pComp1 & pCode1
                            lgStrSQL = lgStrSQL & "   AND A.SUB_TYPE  " & pComp2 & pCode2
                            lgStrSQL = lgStrSQL & "   AND A.EMP_NO    " & pComp3 & pCode3
                            lgStrSQL = lgStrSQL & "   AND A.SUB_YYMM  = B.SUB_YYMM"
                            lgStrSQL = lgStrSQL & "   AND B.SUB_TYPE  " & pComp4 & pCode4
                            lgStrSQL = lgStrSQL & "   AND A.EMP_NO    = B.EMP_NO"
                            lgStrSQL = lgStrSQL & "   AND A.SUB_CD    = B.SUB_CD"
                            lgStrSQL = lgStrSQL & "   AND C.CODE_TYPE = " & FilterVar("2", "''", "S") & " "
                            lgStrSQL = lgStrSQL & "   AND C.ALLOW_CD  = B.SUB_CD "
                            lgStrSQL = lgStrSQL & " ORDER BY ALLOW_SEQ"

                       Case "6"
                            lgStrSQL = "Select  C.ALLOW_NM, IsNull(B.SUB_AMT,0) - IsNull(A.SUB_AMT,0)"
                            lgStrSQL = lgStrSQL & " From  HDF060T A, HDF060T B, HDA010T C"
                            lgStrSQL = lgStrSQL & " WHERE A.SUB_YYMM  " & pComp1 & pCode1
                            lgStrSQL = lgStrSQL & "   AND A.SUB_TYPE  " & pComp2 & pCode2
                            lgStrSQL = lgStrSQL & "   AND A.EMP_NO    " & pComp3 & pCode3
                            lgStrSQL = lgStrSQL & "   AND A.SUB_YYMM  = B.SUB_YYMM"
                            lgStrSQL = lgStrSQL & "   AND B.SUB_TYPE  " & pComp4 & pCode4
                            lgStrSQL = lgStrSQL & "   AND A.EMP_NO    = B.EMP_NO"
                            lgStrSQL = lgStrSQL & "   AND A.SUB_CD    = B.SUB_CD"
                            lgStrSQL = lgStrSQL & "   AND C.CODE_TYPE = " & FilterVar("2", "''", "S") & " "
                            lgStrSQL = lgStrSQL & "   AND C.ALLOW_CD  = A.SUB_CD "
                            lgStrSQL = lgStrSQL & " ORDER BY ALLOW_SEQ"

					End Select
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData1
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .ggoSpread.Source     = .frm1.vspdData2
                .ggoSpread.SSShowData "<%=lgstrData1%>"
                .ggoSpread.Source     = .frm1.vspdData3
                .ggoSpread.SSShowData "<%=lgstrData2%>"
                .ggoSpread.Source     = .frm1.vspdData4
                .ggoSpread.SSShowData "<%=lgstrData3%>"
                .ggoSpread.Source     = .frm1.vspdData5
                .ggoSpread.SSShowData "<%=lgstrData4%>"
                .ggoSpread.Source     = .frm1.vspdData6
                .ggoSpread.SSShowData "<%=lgstrData5%>"
                .DBQueryOk        
	         End with
	      Else
             Parent.DBQueryNo
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
