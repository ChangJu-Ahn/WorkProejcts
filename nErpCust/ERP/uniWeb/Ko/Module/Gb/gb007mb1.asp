<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear  
    
   Call LoadBasisGlobalInf() 
   Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")                                                                                '��: Clear Error status

	Dim lgstrDataTotal
	Dim lgDataAmt
	Dim lgDataAmt1
	Dim lgStrPrevKey

    Call HideStatusWnd                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '��: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)


    'Multi Multi SpreadSheet
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '��: "M"(Spread #1) "S"(Spread #2)
    
    Const C_SHEETMAXROWS_D  = 400                      									
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time


	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

   If lgCurrentSpd = "M" Then
        Call SubBizQueryMulti()
    Else
        Call SubBizQueryMulti1()
    End if

End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Date data
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim YYYYMM

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

  If Trim(lgKeyStream(0)) <> "" Then

 
	YYYYMM =  FilterVar(lgKeyStream(0), "''", "S")

	strWhere =  " and G.YYYYMM = " & YYYYMM
	END IF

    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                                 '��: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        If lgCurrentSpd = "M" Then
           %>
           <Script Language = VBScript>
				parent.frm1.txtDataAmt1.text = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint, 0)%>"
				parent.frm1.hYYYYMM.Value = "<%=ConvSPChars(Trim(lgKeyStream(0)))%>"
			</Script>
           <%

           Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found.
        End If
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""
        iDx = 1

        Do While Not lgObjRs.EOF
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("from_alloc"))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("cost_nm"))
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TOTAL"), ggAmtOfMoney.DecPoint, 0)


	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

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
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubBizQueryMulti3()                               '�հ踦 ���� query
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet

End Sub
'============================================================================================================
' Name : SubBizQueryMulti1    �ι�° dbqueryok()���� ȣ��� �ι�° 
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti1()
    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim YYYYMM
    Dim cost_cd

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------


  If Trim(lgKeyStream(0)) <> "" Then
	    YYYYMM =  FilterVar(lgKeyStream(0), "''", "S")

	    strWhere =  " and C.YYYYMM = " & YYYYMM
	End If

  If Trim(lgKeyStream(1)) <> "" Then
	    cost_cd = FilterVar(lgKeyStream(1), "''", "S")
	    strWhere = strWhere & " and C.from_alloc = " & cost_cd
	End If
	
	

    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                                 '��: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""

           Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found.
           Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""
        iDx = 1

        Do While Not lgObjRs.EOF
            'Select Case lgCurrentSpd

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("to_alloc"))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("cost_nm"))
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B"), ggAmtOfMoney.DecPoint, 0)
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_unit_nm"))


	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
		Loop

    End If

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubBizQueryMulti2()
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubBizQueryMulti2
' Desc : Query Data from Db       ������  �հ踦 ���ϱ�����..
'============================================================================================================
Sub SubBizQueryMulti2()
    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim YYYYMM
    dim cost_Cd

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

  If Trim(lgKeyStream(0)) <> "" Then
	    YYYYMM =  FilterVar(lgKeyStream(0), "''", "S")

	    strWhere =  " and C.YYYYMM = " & YYYYMM
	End If

  If Trim(lgKeyStream(1)) <> "" Then
	    cost_cd = FilterVar(lgKeyStream(1), "''", "S")
	    strWhere = strWhere & " and C.from_alloc = " & cost_cd
	End If
    Call SubMakeSQLStatements("MR",strWhere,"22",C_LIKE)                                 '��: Make sql statements

      If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
 
        Call SetErrorStatus()
    Else

%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
     Parent.Frm1.txtDataAmt1.text	= "<%=UNINumClientFormat(lgObjRs(0), ggAmtOfMoney.DecPoint, 0)%>"

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
</Script>
<%
    End If

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubBizQueryMulti3
' Desc : Query Data from Db         �հ踦 ���ϱ�����..
'============================================================================================================
Sub SubBizQueryMulti3()
    Dim iDx
    Dim iKey1
    Dim strFromToWhere
    Dim strWhere
    Dim YYYYMM

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    If Trim(lgKeyStream(0)) <> "" Then
	    YYYYMM =  FilterVar(lgKeyStream(0), "''", "S")
	    strWhere =  " and YYYYMM = " & YYYYMM
	End If


    Call SubMakeSQLStatements("MR",strWhere,"5",C_EQGT)                                 '��: Make sql statements

	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
		lgDataAmt = 0 
        Call SetErrorStatus()
    Else
		lgDataAmt = lgObjRs(0)
    End If

	

%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
              Parent.Frm1.txtDataAmt.text		=	"<%=UNINumClientFormat(lgDataAmt, ggAmtOfMoney.DecPoint, 0)%>"
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
</Script>
<%

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '��: Protect system from crashing

    Err.Clear                                                                        '��: Clear Error status

End Sub

'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

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

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

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

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

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

   On Error Resume Next                                                             '��: Protect system from crashing
   Err.Clear                                                                        '��: Clear Error status

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"

           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1

           Select Case Mid(pDataType,2,1)
			   Case "C"
                       lgStrSQL = "Select *   "
                       lgStrSQL = lgStrSQL & " From  B_MINOR "
                       lgStrSQL = lgStrSQL & " Where MAJOR_CD " & pComp & pCode
                       lgStrSQL = lgStrSQL & " And   MINOR_CD " & pComp & pCode1
               Case "D"
                       lgStrSQL = "Select *   "
                       lgStrSQL = lgStrSQL & " From  B_MINOR "
                       lgStrSQL = lgStrSQL & " WHERE MAJOR_CD " & pComp & pCode
                       lgStrSQL = lgStrSQL & " AND   MINOR_CD " & pComp & pCode1
               Case "R"
                       If lgCurrentSpd = "M" Then
                       select case pCode1
                       case "X"

                          lgStrSQL = "Select  TOP " & iSelCount  & " G.from_alloc, B.cost_nm, sum(case c.bal_fg when " & FilterVar("DR", "''", "S") & " then G.to_amount else -1*G.to_amount end) AS TOTAL"
                          lgStrSQL = lgStrSQL & " From  g_alloc_result G, b_cost_center B, a_acct C"
                          lgStrSQL = lgStrSQL & " Where G.from_alloc  = B.cost_cd						   "
                          lgStrSQL = lgStrSQL & pCode
                          lgStrSQL = lgStrSQL & " AND  G.acct_cd = C.acct_cd							   "
                          lgStrSQL = lgStrSQL & " AND  alloc_kinds   = " & FilterVar(1, "''", "S")
	                      lgStrSQL = lgStrSQL & " AND  G.from_type   = " & FilterVar("c", "''", "S") & " and g.ctrl_cd = " & FilterVar("*", "''", "S") & "  "
	                      lgStrSQL = lgStrSQL & " GROUP BY G.from_alloc, B.cost_nm "
						  lgStrSQL = lgStrSQL & " ORDER BY G.from_alloc "

                       case "5"
                          lgStrSQL = "Select  sum(case c.bal_fg when " & FilterVar("DR", "''", "S") & " then G.to_amount else -1*G.to_amount end) AS TOTAL"
                          lgStrSQL = lgStrSQL & " From  g_alloc_result G, b_cost_center B, a_acct C"
                          lgStrSQL = lgStrSQL & " Where G.from_alloc  = B.cost_cd						   "
                          lgStrSQL = lgStrSQL & pCode
                          lgStrSQL = lgStrSQL & " AND  G.acct_cd = C.acct_cd							   "
                          lgStrSQL = lgStrSQL & " AND  alloc_kinds   = " & FilterVar(1, "''", "S")
	                      lgStrSQL = lgStrSQL & " AND  G.from_type   = " & FilterVar("c", "''", "S") & " and g.ctrl_cd = " & FilterVar("*", "''", "S") & " "
							
						END SELECT


				  else
				     Select Case pCode1
					 Case "X"
					    lgStrSQL = "select C.to_alloc, A.cost_nm, B.biz_unit_nm, "
					    lgStrSQL = lgStrSQL & " sum(case d.bal_fg when " & FilterVar("DR", "''", "S") & " then c.to_amount else -1*c.to_amount end) as B "
						lgStrSQL = lgStrSQL & " FROM  b_cost_center A, b_biz_unit B, g_alloc_result C , a_acct D"
						lgStrSQL = lgStrSQL & " where C.to_alloc = A.cost_cd AND A.biz_unit_cd = B.biz_unit_cd and C.acct_cd = D.acct_cd "
                        lgStrSQL = lgStrSQL & " AND   alloc_kinds   = " & FilterVar(1, "''", "S")
                        lgStrSQL = lgStrSQL & " AND   C.from_type   = " & FilterVar("c", "''", "S") & " and c.ctrl_cd = " & FilterVar("*", "''", "S") & "  "
						lgStrSQL = lgStrSQL & pcode
						lgStrSQL = lgStrSQL & "group by c.to_alloc, A.cost_nm , B.biz_unit_nm "
						lgStrSQL = lgStrSQL & " order by c.to_alloc "


					 Case "22"

					    lgStrSQL = "select sum(case d.bal_fg when " & FilterVar("DR", "''", "S") & " then c.to_amount else -1*c.to_amount end) as B "
						lgStrSQL = lgStrSQL & " FROM  b_cost_center A, b_biz_unit B, g_alloc_result C , a_acct D"
					lgStrSQL = lgStrSQL & " where C.to_alloc = A.cost_cd AND A.biz_unit_cd = B.biz_unit_cd and C.acct_cd = D.acct_cd "
                        lgStrSQL = lgStrSQL & " AND   alloc_kinds   = " & FilterVar(1, "''", "S")
                        lgStrSQL = lgStrSQL & " AND  C.from_type   = " & FilterVar("c", "''", "S") & " and c.ctrl_cd = " & FilterVar("*", "''", "S") & "  "
						lgStrSQL = lgStrSQL & pcode
				      End Select
                       End If

               Case "U"
                       lgStrSQL = "Select *   "
                       lgStrSQL = lgStrSQL & " From  B_MINOR "
                       lgStrSQL = lgStrSQL & " Where MAJOR_CD " & pComp & pCode
                       lgStrSQL = lgStrSQL & " And   MINOR_CD " & pComp & pCode1
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                If Trim("<%=lgCurrentSpd%>") = "M" Then
                   .ggoSpread.Source     = .frm1.vspdData
					.ggoSpread.SSShowData "<%=lgstrData%>"
                   .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                    .DBQueryOk
                Else
                   .ggoSpread.Source     = .frm1.vspdData1
                   .lgStrPrevKeyIndex1    = "<%=lgStrPrevKeyIndex%>"
                   .ggoSpread.SSShowData "<%=lgstrData%>"
				   .lgStrPrevKey         = "<%=lgStrPrevKey%>"
				   .DBQueryOk2

                End If

	         End with
          Else
	          With Parent
                If Trim("<%=lgCurrentSpd%>") = "M" Then
                   .DBQueryOk
                Else
				   .DBQueryOk2
                End If
	         End with

          End If
    End Select

</Script>