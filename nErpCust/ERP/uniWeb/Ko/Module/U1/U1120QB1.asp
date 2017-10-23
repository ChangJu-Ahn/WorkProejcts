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
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)							 '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryCond()
    If lgErrorStatus <> "YES" Then
		Call SubBizQueryMulti()
	End If
End Sub    

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQueryCond
' Desc : Verify Condition Value from Db
'============================================================================================================
Sub SubBizQueryCond()

    On Error Resume Next
    Err.Clear

	If lgKeyStream(0) <> "" Then
   
		Call SubMakeSQLStatements("CP")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("SCM002", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(lgObjRs("PLANT_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(1) <> "" THEN 'AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CI")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(lgObjRs("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(2) <> "" THEN 'AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CB")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("SCM003", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """ & ConvSPChars(lgObjRs("BP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

     strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR")                                 '¡Ù : Make sql statements
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))   
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("DLVY_DT"))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PO_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("UNRCPT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("UNRCPT_WAIT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("UNRCPT_FIRM_QTY"),ggQty.DecPoint,0)
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
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        CALL SVRMSGBOX(arrColVal ,0,1)
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
Sub SubMakeSQLStatements(pDataType)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
					lgStrSQL =                   "	select	TOP " & iSelCount  & " a.po_no,	b.po_seq_no, b.dlvy_dt, "
					lgStrSQL = lgStrSQL & "             	a.ret_flg, b.po_unit, b.po_qty, b.rcpt_qty, b.po_qty - b.rcpt_qty unrcpt_qty, "
					lgStrSQL = lgStrSQL & "             	b.item_cd, c.item_nm, c.spec, b.plant_cd, e.plant_nm, "
					lgStrSQL = lgStrSQL & "					isNULL(f.unrcpt_firm_qty,0) unrcpt_wait_qty, "
					lgStrSQL = lgStrSQL & "					b.po_qty - b.rcpt_qty - isNULL(f.unrcpt_firm_qty,0) unrcpt_firm_qty , B.SL_CD , G.SL_NM  "
					lgStrSQL = lgStrSQL & "			from	m_pur_ord_hdr a, m_pur_ord_dtl b, b_item c, M_SCM_PLAN_PUR_RCPT d, b_plant e, "
					lgStrSQL = lgStrSQL & "					(select	d.po_no, d.po_seq_no, sum(case when d.rcpt_qty = 0 and rcpt_dt is null then d.confirm_qty else 0 end) unrcpt_firm_qty "
					lgStrSQL = lgStrSQL & "					from	m_pur_ord_hdr a, m_pur_ord_dtl b, M_SCM_PLAN_PUR_RCPT c, M_SCM_FIRM_PUR_RCPT d "
					lgStrSQL = lgStrSQL & "					where	a.po_no = b.po_no and a.release_flg = 'Y' and b.cls_flg= 'N' and b.po_no = c.po_no "

					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "             and 	b.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					End If

					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If

					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.bp_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If

					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.dlvy_dt >= " & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S")
					End If

					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.dlvy_dt <= " & FilterVar(UNIConvDate(lgKeyStream(4)),"''", "S")
					End If
					
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.TRACKING_NO = " & FilterVar(lgKeyStream(5),"''", "S")
					End If
					
					lgStrSQL = lgStrSQL & "					and		b.po_seq_no = c.po_seq_no and c.split_seq_no = 0 and c.po_no = d.po_no and c.po_seq_no = d.po_seq_no "
					lgStrSQL = lgStrSQL & "					group by d.po_no, d.po_seq_no) f , B_STORAGE_LOCATION G "
					lgStrSQL = lgStrSQL & "			where 	a.po_no = b.po_no and (b.po_qty - b.rcpt_qty) > 0 and b.plant_cd = e.plant_cd"
					lgStrSQL = lgStrSQL & "			and		b.cls_flg = 'N' and a.release_flg = 'Y' and b.item_cd *= c.item_cd "
					lgStrSQL = lgStrSQL & "			and		b.po_no = d.po_no and b.po_seq_no = d.po_seq_no and d.split_seq_no = 0 "
					lgStrSQL = lgStrSQL & "			and		d.po_no *= f.po_no and d.po_seq_no *= f.po_seq_no "					
					lgStrSQL = lgStrSQL & "			and		B.SL_CD = G.SL_CD "
					
					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "             and 	b.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					End If
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.bp_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.dlvy_dt >= " & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S")
					End If
					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.dlvy_dt <= " & FilterVar(UNIConvDate(lgKeyStream(4)),"''", "S")
					End If
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.TRACKING_NO = " & FilterVar(lgKeyStream(5),"''", "S")
					End If
           End Select
           
        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "P"
                    lgStrSQL =            " select plant_nm from b_plant where plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
               Case "I"
                    lgStrSQL =            " select item_nm from b_item where item_cd = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
               Case "B"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(2) & "" ,"''", "S")
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	Dim lsMsg
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	