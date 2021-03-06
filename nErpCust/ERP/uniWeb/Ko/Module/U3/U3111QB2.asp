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
    Call HideStatusWnd                                                               '��: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)							 '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
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
    Call SubBizQueryMulti()
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

     strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                                 '�� : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
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
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRACKING_NO"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("MVMT_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MVMT_RCPT_SL_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))            
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("PO_PRC"), 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("RCPT_AMT"), 0)
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
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data
        CALL SVRMSGBOX(arrColVal ,0,1)
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
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
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
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
					lgStrSQL =            "	SELECT	TOP " & iSelCount  & " A.ITEM_CD, D.ITEM_NM, D.SPEC, A.TRACKING_NO , A.PLANT_CD, F.PLANT_NM, A.MVMT_DT, A.MVMT_UNIT, A.PO_NO, A.PO_SEQ_NO,   "           	
					lgStrSQL = lgStrSQL & "	                CASE WHEN C.RET_FLG = 'N' AND C.RCPT_FLG = 'Y' THEN A.MVMT_QTY ELSE - A.MVMT_QTY END RCPT_QTY ,   "
					lgStrSQL = lgStrSQL & "	            	CASE WHEN C.RET_FLG = 'N' AND C.RCPT_FLG = 'Y' THEN  "
					lgStrSQL = lgStrSQL & "	                          (A.MVMT_QTY * (CASE WHEN A.PO_NO IS NULL AND A.PO_SEQ_NO IS NULL THEN A.MVMT_PRC ELSE I.PO_PRC END))  "
					lgStrSQL = lgStrSQL & "	                     ELSE - (A.MVMT_QTY * (CASE WHEN A.PO_NO IS NULL AND A.PO_SEQ_NO IS NULL THEN A.MVMT_PRC ELSE I.PO_PRC END))  "
					lgStrSQL = lgStrSQL & "	                 END RCPT_AMT ,          "
					lgStrSQL = lgStrSQL & "	             	(CASE WHEN A.PO_NO IS NULL AND A.PO_SEQ_NO IS NULL THEN A.MVMT_PRC ELSE I.PO_PRC END) PO_PRC, MVMT_RCPT_SL_CD , H.SL_NM 			 "
					lgStrSQL = lgStrSQL & "	  FROM	M_PUR_GOODS_MVMT A, B_BIZ_PARTNER B, M_MVMT_TYPE C, B_ITEM D, B_PLANT F,  "
					lgStrSQL = lgStrSQL & "	        B_STORAGE_LOCATION H , M_PUR_ORD_DTL I 	, B_ITEM_BY_PLANT J		 "
					lgStrSQL = lgStrSQL & "	 WHERE 	A.BP_CD = B.BP_CD  "
					lgStrSQL = lgStrSQL & "	   AND	A.ITEM_CD = D.ITEM_CD  "
					lgStrSQL = lgStrSQL & "	   AND	A.PLANT_CD = F.PLANT_CD 			 "
					lgStrSQL = lgStrSQL & "	   AND	B.BP_TYPE <> 'C'  "
					lgStrSQL = lgStrSQL & "	   AND  A.IO_TYPE_CD = C.IO_TYPE_CD   "
					lgStrSQL = lgStrSQL & "	   AND  A.PO_NO *= I.PO_NO 	 "		
					lgStrSQL = lgStrSQL & "	   AND  A.PO_SEQ_NO *= I.PO_SEQ_NO 		 "	
					lgStrSQL = lgStrSQL & "	   AND  A.MVMT_RCPT_SL_CD = H.SL_CD               "
					lgStrSQL = lgStrSQL & "	   AND  A.ITEM_CD = J.ITEM_CD "
					lgStrSQL = lgStrSQL & "	   AND  A.PLANT_CD = J.PLANT_CD "
					
					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "             and 	a.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					End If
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.bp_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.mvmt_dt >= " & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S")
					End If
					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.mvmt_dt <= " & FilterVar(UNIConvDate(lgKeyStream(4)),"''", "S")
					End If
					If lgkeystream(5) = "Y" Then
						lgStrSQL = lgStrSQL & "            	and 	J.PROCUR_TYPE = 'P' "
                    ElseIf lgkeystream(5) = "N" Then
 						lgStrSQL = lgStrSQL & "            	and 	J.PROCUR_TYPE <> 'P' "
					End If	
					If lgkeystream(6) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.mvmt_sl_cd = " & FilterVar(lgKeyStream(6),"''", "S")
					End If
					
					If lgkeystream(7) <> "A" Then
						lgStrSQL = lgStrSQL & "            	and 	c.RET_FLG = " & FilterVar(lgKeyStream(7),"''", "S")
					End If
					
					If lgkeystream(8) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	A.tracking_no = " & FilterVar(lgKeyStream(8),"''", "S")
					End If
					
					lgStrSQL = lgStrSQL & "			order by 		a.po_no, a.po_seq_no "
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	Dim lsMsg
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData2
                .lgStrPrevKey1    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrColorFlag = "<%=lgStrColorFlag%>"
                .DBDtlQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
