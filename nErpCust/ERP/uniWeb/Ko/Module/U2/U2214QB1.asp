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
	Dim lgStrColorFlag
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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

	If lgKeyStream(1) <> "" AND lgErrorStatus <> "YES" Then
   
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

	If lgKeyStream(2) <> "" AND lgErrorStatus <> "YES" Then
   
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

     strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR")                                 '☆ : Make sql statements
 
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            
            If lgObjRs("group_flag") = 0 Then 
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "1" & gRowSep
			ElseIf lgObjRs("group_flag") = 3 Then 
				lgstrData = lgstrData & Chr(11) & "[품목별 소계]"
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "2" & gRowSep
			Else
				lgstrData = lgstrData & Chr(11) & "[총계]"
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "5" & gRowSep
			End If
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PO_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_UNIT"))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PLAN_DVRY_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("UNRCPT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("FIRM_DVRY_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("REMAIN_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("PO_LOC_AMT"), 0)
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
Sub SubMakeSQLStatements(pDataType)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

           Select Case Mid(pDataType,2,1)
               Case "R"
               

		  lgStrSQL =            " SELECT  TOP " & iSelCount  & " A.* , B.ITEM_NM , B.SPEC , C.PLANT_NM FROM " 
		  lgStrSQL = lgStrSQL & " (SELECT (grouping(A.item_cd) + grouping(E.plant_cd) + grouping(D.po_dt) + grouping(E.po_unit)) group_flag,  " 
		  lgStrSQL = lgStrSQL & " 	      A.item_cd, E.plant_cd, D.po_dt, E.po_unit, sum(a.plan_dvry_qty) plan_dvry_qty,  " 
		  lgStrSQL = lgStrSQL & " 	      sum(isNULL(E.rcpt_qty,0)) rcpt_qty, sum(a.plan_dvry_qty - E.rcpt_qty) unrcpt_qty,  " 
		  lgStrSQL = lgStrSQL & " 	      sum(isNULL(b.firm_dvry_qty,0)) firm_dvry_qty,  " 
		  lgStrSQL = lgStrSQL & " 	      sum(a.plan_dvry_qty - E.rcpt_qty - isNULL(b.firm_dvry_qty,0)) remain_qty,  " 
		  lgStrSQL = lgStrSQL & " 	      sum(E.po_loc_amt) po_loc_amt  " 
		  lgStrSQL = lgStrSQL & "    FROM M_SCM_PLAN_PUR_RCPT A , " 
		  lgStrSQL = lgStrSQL & "         (select a.po_no, a.po_seq_no, d_bp_cd , e.bp_cd ,  " 
		  lgStrSQL = lgStrSQL & "                 sum(case when a.rcpt_qty = 0 then a.confirm_qty else 0 end) firm_dvry_qty  " 
		  lgStrSQL = lgStrSQL & "            from M_SCM_FIRM_PUR_RCPT a , m_pur_ord_dtl d , b_storage_location e  " 
		  lgStrSQL = lgStrSQL & "           where A.po_no = d.po_no  " 
		  lgStrSQL = lgStrSQL & "             AND A.po_SEQ_no = d.po_sEQ_no  " 
		  lgStrSQL = lgStrSQL & "             and a.d_bp_cd = e.sl_cd  " 
		  lgStrSQL = lgStrSQL & "           group by a.po_no , a.po_seq_no , a.d_bp_cd , e.bp_cd )B , B_STORAGE_LOCATION C , M_PUR_ORD_HDR D , M_PUR_ORD_DTL E " 
		  lgStrSQL = lgStrSQL & "   WHERE A.PO_NO = B.PO_NO   AND A.PO_SEQ_NO = B.PO_SEQ_NO " 
		  lgStrSQL = lgStrSQL & "     AND B.D_BP_CD = C.SL_CD AND A.PO_NO = D.PO_NO  "   
		  lgStrSQL = lgStrSQL & "     AND A.PO_NO = E.PO_NO AND A.PO_SEQ_NO = E.PO_SEQ_NO " 
					
					If lgkeystream(0) <> "" Then			
		  lgStrSQL = lgStrSQL & "     AND E.PLANT_CD =  " & FilterVar(lgKeyStream(0),"''", "S") 
					End If
					If lgkeystream(1) <> "" Then			
		  lgStrSQL = lgStrSQL & "     AND A.ITEM_CD = " & FilterVar(lgKeyStream(1),"''", "S")   
					End If
					If lgkeystream(2) <> "" Then			
		  lgStrSQL = lgStrSQL & "     AND C.BP_CD = " & FilterVar(lgKeyStream(2),"''", "S") 
					End If
					If lgkeystream(3) <> "" Then			
		  lgStrSQL = lgStrSQL & "     AND A.plan_dvry_dt >= " & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S")
					End If
					If lgkeystream(4) <> "" Then			
		  lgStrSQL = lgStrSQL & "     AND A.plan_dvry_dt <= " & FilterVar(UNIConvDate(lgKeyStream(4)),"''", "S")
					End If
					If lgkeystream(5) <> "N" Then			
		  lgStrSQL = lgStrSQL & "     AND D.RET_FLG = 'Y' "  
					Else
		  lgStrSQL = lgStrSQL & "     AND D.RET_FLG = 'N' "  
					End If
					If lgkeystream(6) <> "" Then
		  lgStrSQL = lgStrSQL & "     AND D.PO_DT >= " & FilterVar(UNIConvDate(lgKeyStream(6)),"''", "S")
					End If
					If lgkeystream(7) <> "" Then			
		  lgStrSQL = lgStrSQL & "     AND D.PO_DT <= " & FilterVar(UNIConvDate(lgKeyStream(7)),"''", "S")
					End If
					
					If lgkeystream(8) <> "" Then			
		  lgStrSQL = lgStrSQL & "     AND B.D_BP_CD = 'S1453' "  
					End If
								
								
		  lgStrSQL = lgStrSQL & "   GROUP BY A.item_cd, E.plant_cd, D.po_dt, E.po_unit  WITH ROLLUP )A , b_item b, b_plant c  " 
		  lgStrSQL = lgStrSQL & "   WHERE A.ITEM_CD *= B.ITEM_CD "   
		  lgStrSQL = lgStrSQL & "     AND A.PLANT_CD *= C.PLANT_CD " 
		  lgStrSQL = lgStrSQL & "     AND A.group_flag in (0,3,4) "  
		  lgStrSQL = lgStrSQL & "   ORDER BY isNULL(a.item_cd,'ZZZZZZZZZZZZZZZZZZ'),  isNULL(po_dt,'2999-12-31') "
					
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
                .ggoSpread.Source     = .frm1.vspdData1
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrColorFlag = "<%=lgStrColorFlag%>"
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
      
</Script> 