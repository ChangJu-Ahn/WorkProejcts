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
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
	Call LoadInfTB19029B("Q", "*","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE", "MB")
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)					 '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

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

	If lgKeyStream(1) <> "" then 'AND lgErrorStatus <> "YES" Then
   
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

	If lgKeyStream(2) <> "" then 'AND lgErrorStatus <> "YES" Then
   
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

	If lgKeyStream(7) <> "" then 'AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CS")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("SCM016", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """ & ConvSPChars(lgObjRs("SL_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtSlNm.value = """"" & vbCrLf	
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
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            If lgObjRs("group_flag") = 0 Then 
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "1" & gRowSep
			ElseIf lgObjRs("group_flag") = 3 Then 
				lgstrData = lgstrData & Chr(11) & "[Ç°¸ñ¼Ò°è]"
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "2" & gRowSep
			Else
				lgstrData = lgstrData & Chr(11) & "[ÃÑ°è]"
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "5" & gRowSep
			End If
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRACKING_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MVMT_UNIT"))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("RCPT_AMT"), 0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RET_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("RET_AMT"), 0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("TOT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("TOT_AMT"), 0)
            lgstrData = lgstrData & Chr(11) & lgObjRs("group_flag")
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
					lgStrSQL =            " SELECT TOP " & iSelCount & " A.GROUP_FLAG, A.ITEM_CD, B.ITEM_NM, B.SPEC,A.TRACKING_NO , A.PLANT_CD, C.PLANT_NM,  A.MVMT_UNIT,    	"
					lgStrSQL = lgStrSQL & "           	A.RCPT_QTY, A.RCPT_AMT , A.RET_QTY, A.RET_AMT , 	"
					lgStrSQL = lgStrSQL & "                 (A.RCPT_QTY - RET_QTY)TOT_QTY , (A.RCPT_AMT - RET_AMT)TOT_AMT  				"
					lgStrSQL = lgStrSQL & "   FROM	(SELECT	(GROUPING(A.ITEM_CD) + GROUPING(A.TRACKING_NO) + GROUPING(A.PLANT_CD) + GROUPING(A.MVMT_UNIT) + GROUPING(A.TRACKING_NO)) GROUP_FLAG, 								"
					lgStrSQL = lgStrSQL & "                 A.ITEM_CD, A.PLANT_CD,  A.MVMT_UNIT, A.TRACKING_NO , "
					lgStrSQL = lgStrSQL & "                 SUM(CASE WHEN C.RET_FLG = 'N' AND C.RCPT_FLG = 'Y' THEN A.MVMT_QTY ELSE 0 END)RCPT_QTY,  					"
					lgStrSQL = lgStrSQL & " 		SUM(CASE WHEN C.RET_FLG = 'N' AND C.RCPT_FLG = 'Y' THEN A.MVMT_QTY * (CASE WHEN A.PO_NO IS NULL AND A.PO_SEQ_NO IS NULL THEN A.MVMT_PRC ELSE E.PO_PRC END) ELSE 0 END)RCPT_AMT,  								"
					lgStrSQL = lgStrSQL & "                 SUM(CASE WHEN C.RET_FLG = 'Y' AND C.RCPT_FLG = 'N' THEN A.MVMT_QTY ELSE 0 END)RET_QTY,  								"
					lgStrSQL = lgStrSQL & "                 SUM(CASE WHEN C.RET_FLG = 'Y' AND C.RCPT_FLG = 'N' THEN A.MVMT_QTY * (CASE WHEN A.PO_NO IS NULL AND A.PO_SEQ_NO IS NULL THEN A.MVMT_PRC ELSE E.PO_PRC END) ELSE 0 END)RET_AMT               		"
					lgStrSQL = lgStrSQL & "            FROM	M_PUR_GOODS_MVMT A, B_BIZ_PARTNER B, M_MVMT_TYPE C, 	"
					lgStrSQL = lgStrSQL & "                 M_PUR_ORD_DTL E , B_ITEM_BY_PLANT F            		"
					lgStrSQL = lgStrSQL & "           WHERE A.BP_CD = B.BP_CD               		"
					lgStrSQL = lgStrSQL & "             AND	B.BP_TYPE <> 'C'               		"
					lgStrSQL = lgStrSQL & "             AND	A.IO_TYPE_CD = C.IO_TYPE_CD               		"
					lgStrSQL = lgStrSQL & "             AND A.PO_NO *= E.PO_NO               		"
					lgStrSQL = lgStrSQL & "             AND A.PO_SEQ_NO *= E.PO_SEQ_NO              	"
					lgStrSQL = lgStrSQL & "             AND A.ITEM_CD = F.ITEM_CD              	"
					lgStrSQL = lgStrSQL & "             AND A.PLANT_CD = F.PLANT_CD              	"
					
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
						lgStrSQL = lgStrSQL & "            	and 	F.PROCUR_TYPE = 'P' "
                    ElseIf lgkeystream(5) = "N" Then
 						lgStrSQL = lgStrSQL & "            	and 	F.PROCUR_TYPE <> 'P' "
					End If	
					
					If lgkeystream(6) <> "A" Then
						lgStrSQL = lgStrSQL & "            	and 	c.RET_FLG = " & FilterVar(lgKeyStream(6),"''", "S")
					End If
					
					If lgkeystream(7) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.mvmt_rcpt_sl_cd = " & FilterVar(lgKeyStream(7),"''", "S")
					End If
					
					If lgkeystream(8) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	A.tracking_no = " & FilterVar(lgKeyStream(8) ,"''", "S")
					End If
					
					lgStrSQL = lgStrSQL & "             	group by a.item_cd, a.plant_cd,  a.mvmt_unit , A.TRACKING_NO "
					lgStrSQL = lgStrSQL & "             	with rollup) a,  "
					lgStrSQL = lgStrSQL & "             	b_item b, b_plant c  "
					lgStrSQL = lgStrSQL & "        where	a.item_cd *= b.item_cd  "
					lgStrSQL = lgStrSQL & "        and		a.plant_cd *= c.plant_cd  "
					lgStrSQL = lgStrSQL & "        and		a.group_flag in (0,3,5)  "
					lgStrSQL = lgStrSQL & "        order by isNULL(a.item_cd,'ZZZZZZZZZZZZZZZZZZ'),  a.group_flag  "
           End Select
           
        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "P"
                    lgStrSQL =            " select plant_nm from b_plant where plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
               Case "I"
                    lgStrSQL =            " select item_nm from b_item where item_cd = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
               Case "B"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(2) & "" ,"''", "S")
               Case "S"
                    lgStrSQL =            " select sl_nm from b_storage_location where sl_cd = " & FilterVar(lgKeyStream(7) & "" ,"''", "S")                         
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
                .ggoSpread.Source     = .frm1.vspdData1
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrColorFlag = "<%=lgStrColorFlag%>"
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
