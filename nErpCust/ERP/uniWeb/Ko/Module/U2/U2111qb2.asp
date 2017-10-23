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
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)							 '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

     strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PO_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRACKING_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_UNIT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))
            
            If ConvSPChars(lgObjRs("CLS_FLG")) = "N" Then
				lgstrData = lgstrData & Chr(11) & "미마감"
			ElseIf ConvSPChars(lgObjRs("CLS_FLG")) = "Y" Then	
				lgstrData = lgstrData & Chr(11) & "마감"
			Else
				lgstrData = lgstrData & Chr(11) & ""	
            End If
            
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PLAN_DVRY_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("UNRCPT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("FIRM_DVRY_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("REMAIN_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("PO_PRC"), 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("PO_DOC_AMT"), 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))
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

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
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
        CALL SVRMSGBOX(arrColVal ,0,1)
        
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                    lgStrSQL	= " SELECT	TOP " & iSelCount  & " c.po_dt, d.item_cd, e.item_nm, e.spec, D.TRACKING_NO , d.po_unit, a.po_no, a.po_seq_no, " _
								& "	a.plan_dvry_qty, a.plan_dvry_dt, d.rcpt_qty, (a.plan_dvry_qty - d.rcpt_qty) unrcpt_qty, " _
								& "	isNULL(b.firm_dvry_qty,0) firm_dvry_qty, " _
								& "	(a.plan_dvry_qty - d.rcpt_qty - isNULL(b.firm_dvry_qty,0)) remain_qty, a.ret_flg, c.remark, " _
								& "	d.po_prc, d.po_prc_flg, d.po_doc_amt , D.SL_CD , G.SL_NM , D.CLS_FLG" _
								& " FROM	M_SCM_PLAN_PUR_RCPT a, " _
								& "	(SELECT	po_no, po_seq_no, SUM(CASE WHEN rcpt_qty = 0 AND rcpt_dt is NULL THEN confirm_qty ELSE 0 END) firm_dvry_qty " _
								& "	FROM M_SCM_FIRM_PUR_RCPT WHERE	(plan_dvry_qty - rcpt_qty) > 0 GROUP BY po_no, po_seq_no) b, " _
								& "	m_pur_ord_hdr c, m_pur_ord_dtl d, b_item e, b_biz_partner f , B_STORAGE_LOCATION G  " _
								& " WHERE	a.split_seq_no = 0 and a.po_no *= b.po_no " _
								& "  and	a.po_seq_no *= b.po_seq_no and a.po_no = c.po_no and a.po_no = d.po_no " _
								& "  and	a.po_seq_no = d.po_seq_no and c.bp_cd = f.bp_cd and d.item_cd = e.item_cd and c.release_flg = " & FilterVar("Y","''","S") _
								& "	 and	d.cls_flg = " & FilterVar("N","''","S") _
								& "	 and    D.SL_CD = G.SL_CD "
					
					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "  and d.plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
					End If
					
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "  and d.item_cd = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
					End If

					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "  and c.bp_cd = " & FilterVar(lgKeyStream(2) & "" ,"''", "S")
					End If
					
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "  and a.plan_dvry_dt >= " & FilterVar(lgKeyStream(3),"''", "S")
					End If
					
					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "  and a.plan_dvry_dt <= " & FilterVar(lgKeyStream(4),"''", "S")
					End If
					
					If lgkeystream(5) <> "N"  Then
						lgStrSQL = lgStrSQL & "  and a.ret_flg = " & FilterVar("Y","''","S")
					Else 	
						lgStrSQL = lgStrSQL & "  and a.ret_flg = " & FilterVar("N","''","S")
					End If
					
					If lgkeystream(6) <> "" Then
						lgStrSQL = lgStrSQL & "  and c.po_dt = " & FilterVar(lgKeyStream(6),"''", "S")
					End If
					
					If lgkeystream(7) <> "" Then
						lgStrSQL = lgStrSQL & "  and c.po_dt >= " & FilterVar(lgKeyStream(7),"''", "S")
					End If
					
					If lgkeystream(8) <> "" Then
						lgStrSQL = lgStrSQL & "  and c.po_dt <= " & FilterVar(lgKeyStream(8),"''", "S")
					End If
					
					If lgkeystream(9) <> "" Then
						lgStrSQL = lgStrSQL & " and 	D.SL_cd = " & FilterVar(lgKeyStream(9),"''", "S")
					End If
					
					If lgkeystream(10) <> "" Then
						lgStrSQL = lgStrSQL & " and 	D.tracking_no = " & FilterVar(lgKeyStream(10),"''", "S")
					End If
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
	Dim lsMsg
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
        Case "MD"
        Case "MR"
        Case "MU"
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData2
                .lgStrPrevKey1    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBDtlQueryOk        
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
