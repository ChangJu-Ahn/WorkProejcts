<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Dim lgSeq
	Dim lgSumQty
	Dim lgQty
	Dim lgOrgQty
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("Q", "M","NOCOOKIE","MB")

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

	If lgKeyStream(4) <> "" Then
   
		Call SubMakeSQLStatements("CP")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(5) <> "" AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CI")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("169961", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """ & ConvSPChars(lgObjRs("SL_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtSLNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(7) <> "" AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CB")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
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
    
    If lgKeyStream(0) = "Y" Then
	    Call SubMakeSQLStatements("MY")
    Else
	    Call SubMakeSQLStatements("MN")
    End IF
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("d_bp_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRACKING_NO"))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PO_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_UNIT"))
			lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("PO_PRC"), 0)
			lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("PO_DOC_AMT"), 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_CUR"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("DLVY_DT"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("LC_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PRE_IV_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("INSPECT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("IV_QTY"),ggQty.DecPoint,0)
			
			If ConvSPChars(lgObjRs("RECV_INSPEC_FLG")) = "N" Then
				lgstrData = lgstrData & Chr(11) & ""
			else
				lgstrData = lgstrData & Chr(11) & "Y"
			End If	
			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INSPECT_METHOD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_GRP"))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("LC_RCPT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_FLG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_GEN_MTHD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SERIAL_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SERIAL_SUB_NO"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("remain_rcpt_qty"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPLIT_SEQ_NO"))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
'call svrmsgbox(lgstrData ,0,1)
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
    Call SubCloseRs(lgObjRs)

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next
    Err.Clear
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
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
           
               Case "Y"
                    lgStrSQL	= " SELECT TOP " & iSelCount & " A.PO_NO, A.PO_SEQ_NO, A.PLANT_CD, X.D_BP_CD, A.ITEM_CD, " _
								& " B.ITEM_NM, B.SPEC, A.TRACKING_NO, A.PO_QTY, A.PO_UNIT, " _
								& " A.PO_PRC, A.PO_DOC_AMT, F.PO_CUR, A.DLVY_DT, ISNULL(A.RCPT_QTY,0) RCPT_QTY, " _
								& " ISNULL(A.LC_QTY,0) LC_QTY, ISNULL(A.PRE_IV_QTY,0) PRE_IV_QTY, ISNULL(A.INSPECT_QTY,0) INSPECT_QTY, ISNULL(A.IV_QTY,0) IV_QTY, L.RECV_INSPEC_FLG, " _
								& "	I.MINOR_NM, F.INSPECT_METHOD, C.PLANT_NM, E.SL_NM, F.PUR_GRP, " _
								& " ISNULL(A.LC_RCPT_QTY,0) LC_RCPT_QTY, L.LOT_FLG, N.LOT_GEN_MTHD, (X.CONFIRM_QTY - X.RCPT_QTY) REMAIN_RCPT_QTY, X.LOT_NO SERIAL_NO, " _
								& " X.LOT_SUB_NO SERIAL_SUB_NO, X.PLAN_DVRY_DT, X.SPLIT_SEQ_NO " _
								& " FROM M_PUR_ORD_DTL A LEFT OUTER JOIN B_LOT_CONTROL N " _
								& " ON A.PLANT_CD = N.PLANT_CD AND A.ITEM_CD = N.ITEM_CD " _
								& " INNER JOIN B_ITEM B ON A.ITEM_CD = B.ITEM_CD " _
								& " INNER JOIN B_PLANT C ON A.PLANT_CD = C.PLANT_CD " _
								& " INNER JOIN B_ITEM_BY_PLANT L ON A.PLANT_CD = L.PLANT_CD AND A.ITEM_CD = L.ITEM_CD " _
								& " INNER JOIN M_PUR_ORD_HDR F ON A.PO_NO = F.PO_NO " _
								& " INNER JOIN B_PUR_GRP J ON F.PUR_GRP = J.PUR_GRP " _
								& " INNER JOIN M_MVMT_TYPE K ON K.IO_TYPE_CD = F.RCPT_TYPE " _
								& " INNER JOIN P_PLANT_CONFIGURATION M ON C.PLANT_CD = M.PLANT_CD " _
								& " INNER JOIN M_SCM_FIRM_PUR_RCPT X ON A.PO_NO = X.PO_NO AND A.PO_SEQ_NO = X.PO_SEQ_NO AND X.CONFIRM_QTY > X.RCPT_QTY " _
								& " INNER JOIN M_SCM_DLVY_PUR_RCPT Y ON X.DLVY_NO = Y.DLVY_NO AND F.BP_CD = Y.BP_CD " _
								& " INNER JOIN B_STORAGE_LOCATION E ON X.D_BP_CD = E.SL_CD " _
								& " INNER JOIN B_MINOR I ON F.INSPECT_METHOD = I.MINOR_CD AND I.MAJOR_CD = " & FilterVar("B9016","''","S")
					
					lgStrSQL	= lgStrSQL & " WHERE A.CLS_FLG = " & FilterVar("N","''","S") _
								& " AND F.RELEASE_FLG = " & FilterVar("Y","''","S") _
								& " AND A.PO_QTY > A.RCPT_QTY AND A.PO_QTY > A.IV_QTY  " _
								& " AND (L.MATERIAL_TYPE <> " & FilterVar("20","''","S") & " OR M.DELIVERY_ORDER_FLG <> " & FilterVar("Y","''","S") & ") " 
'								& " AND X.RCPT_QTY = 0 " _
					lgStrSQL	= lgStrSQL & " AND A.PO_QTY > (A.RCPT_QTY + A.PRE_IV_QTY + A.INSPECT_QTY + A.LC_QTY - A.LC_RCPT_QTY) " _
								& " AND F.STO_FLG = " & FilterVar("N","''","S") _
								& " AND F.RET_FLG = " & FilterVar("Y","''","S")
					
'								& " AND A.PO_QTY > A.RCPT_QTY AND A.PO_QTY > A.IV_QTY AND A.RCPT_QTY = 0 " _

					If lgkeystream(1) <> "" Then					
						lgStrSQL = lgStrSQL & " AND X.PLAN_DVRY_DT >= " & FilterVar(lgkeystream(1),"''","S")
					End If

					If lgkeystream(2) <> "" Then					
						lgStrSQL = lgStrSQL & " AND X.PLAN_DVRY_DT <= " & FilterVar(lgkeystream(2),"''","S")
					End If

					If lgkeystream(3) <> "" Then					
						lgStrSQL = lgStrSQL & " AND A.PLANT_CD = " & FilterVar(lgkeystream(3),"''","S")
					End If

					If lgkeystream(4) <> "" Then					
						lgStrSQL = lgStrSQL & " AND X.DLVY_NO = " & FilterVar(lgkeystream(4),"''","S")
					End If

					If lgkeystream(5) <> "" Then					
						lgStrSQL = lgStrSQL & " AND X.D_BP_CD = " & FilterVar(lgkeystream(5),"''","S")
					End If

					If lgkeystream(6) <> "" Then					
						lgStrSQL = lgStrSQL & " AND X.CONFIRM_YN = " & FilterVar(lgkeystream(6),"''","S")
					End If

					If lgkeystream(7) <> "" Then					
						lgStrSQL = lgStrSQL & " AND F.BP_CD = " & FilterVar(lgkeystream(7),"''","S")
					End If

					If lgkeystream(12) <> "" Then					
						lgStrSQL = lgStrSQL & " AND F.RCPT_TYPE = " & FilterVar(lgkeystream(12),"''","S")
					End If

					If lgkeystream(16) <> "" Then					
						lgStrSQL = lgStrSQL & lgKeyStream(16)
					End If

					lgStrSQL = lgStrSQL & "  ORDER BY A.PO_NO, A.PO_SEQ_NO "
					
               Case "N"
                    lgStrSQL	= " SELECT TOP " & iSelCount & " A.PO_NO, A.PO_SEQ_NO, A.PLANT_CD, X.D_BP_CD, A.ITEM_CD, " _
								& " B.ITEM_NM, B.SPEC, A.TRACKING_NO, A.PO_QTY, A.PO_UNIT, " _
								& " A.PO_PRC, A.PO_DOC_AMT, F.PO_CUR, A.DLVY_DT, ISNULL(A.RCPT_QTY,0) RCPT_QTY, " _
								& " ISNULL(A.LC_QTY,0) LC_QTY, ISNULL(A.PRE_IV_QTY,0) PRE_IV_QTY, ISNULL(A.INSPECT_QTY,0) INSPECT_QTY, ISNULL(A.IV_QTY,0) IV_QTY, L.RECV_INSPEC_FLG, " _
								& "	I.MINOR_NM, F.INSPECT_METHOD, C.PLANT_NM, E.SL_NM, F.PUR_GRP, " _
								& " ISNULL(A.LC_RCPT_QTY,0) LC_RCPT_QTY, L.LOT_FLG, N.LOT_GEN_MTHD, (X.CONFIRM_QTY - X.RCPT_QTY) REMAIN_RCPT_QTY, X.LOT_NO SERIAL_NO, " _
								& " X.LOT_SUB_NO SERIAL_SUB_NO, X.PLAN_DVRY_DT, X.SPLIT_SEQ_NO " _
								& " FROM M_PUR_ORD_DTL A INNER JOIN B_ITEM B ON A.ITEM_CD = B.ITEM_CD " _
								& " INNER JOIN B_PLANT C ON A.PLANT_CD = C.PLANT_CD " _
								& " INNER JOIN M_PUR_ORD_HDR F ON A.PO_NO = F.PO_NO " _
								& " INNER JOIN B_PUR_GRP J ON F.PUR_GRP = J.PUR_GRP " _
								& " INNER JOIN M_MVMT_TYPE K ON F.RCPT_TYPE = K.IO_TYPE_CD " _
								& " INNER JOIN B_ITEM_BY_PLANT L ON A.PLANT_CD = L.PLANT_CD AND A.ITEM_CD = L.ITEM_CD " _
								& " INNER JOIN M_SCM_FIRM_PUR_RCPT X ON A.PO_NO = X.PO_NO AND A.PO_SEQ_NO = X.PO_SEQ_NO AND X.CONFIRM_QTY > X.RCPT_QTY " _
								& " INNER JOIN M_SCM_DLVY_PUR_RCPT Y ON X.DLVY_NO = Y.DLVY_NO AND F.BP_CD = Y.BP_CD " _
								& " INNER JOIN B_STORAGE_LOCATION E ON X.D_BP_CD = E.SL_CD " _
								& " INNER JOIN B_MINOR I ON F.INSPECT_METHOD = I.MINOR_CD AND I.MAJOR_CD = " & FilterVar("B9016","''","S") _
								& " LEFT OUTER JOIN B_LOT_CONTROL N ON A.PLANT_CD = N.PLANT_CD AND A.ITEM_CD = N.ITEM_CD "
								
					lgStrSQL	= lgStrSQL & " WHERE A.CLS_FLG = " & FilterVar("N","''","S") _
								& " AND F.RELEASE_FLG = " & FilterVar("Y","''","S") _
								& " AND (A.PO_QTY - A.RCPT_QTY) > 0 AND A.PO_QTY > A.IV_QTY " _
								& " AND A.PO_QTY > (A.RCPT_QTY + A.PRE_IV_QTY + A.INSPECT_QTY) " _
								& " AND X.RCPT_QTY = 0  " _
								& " AND F.STO_FLG = " & FilterVar("N","''","S") _
								& " AND F.RET_FLG = " & FilterVar("Y","''","S")
					
'								& " AND X.RCPT_QTY = 0 AND A.RCPT_QTY = 0 " _

					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "  AND F.PO_DT >= " & FilterVar(lgKeyStream(1),"''", "S")
					End If
										
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "  AND F.PO_DT <= " & FilterVar(lgKeyStream(2),"''", "S")
					End If
										
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "  AND A.PLANT_CD = " & FilterVar(lgKeyStream(3),"''", "S")
					End If

					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "  AND X.DLVY_NO = " & FilterVar(lgKeyStream(4),"''", "S")
					End If
										
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "  AND X.D_BP_CD = " & FilterVar(lgKeyStream(5),"''", "S")
					End If
										
					If lgkeystream(6) <> "" Then
						lgStrSQL = lgStrSQL & "  AND X.CONFIRM_YN = " & FilterVar(lgKeyStream(6),"''", "S") 
					End If					

					If lgkeystream(7) <> "" Then
						lgStrSQL = lgStrSQL & "  AND F.BP_CD = " & FilterVar(lgKeyStream(7),"''", "S")
					End If
										
					If lgkeystream(12) <> "" Then
						lgStrSQL = lgStrSQL & "  AND F.RCPT_TYPE = " & FilterVar(lgKeyStream(12),"''", "S")
					End If	
										
					If lgkeystream(16) <> "" Then
						lgStrSQL = lgStrSQL & lgKeyStream(16) 
					End If	
										
					lgStrSQL = lgStrSQL & " ORDER BY A.PO_NO, A.PO_SEQ_NO "
				
           End Select      
           
        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "P"
                    lgStrSQL =            " select DLVY_NO from M_SCM_DLVY_PUR_RCPT where DLVY_NO = " & FilterVar(lgKeyStream(4) & "" ,"''", "S")
               Case "I"
                    lgStrSQL =            " select SL_NM from b_storage_location where SL_CD = " & FilterVar(lgKeyStream(5) & "" ,"''", "S")
               Case "B"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(7) & "" ,"''", "S")
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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgPageNo    = "<%=lgStrPrevKey%>"
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
       Case "<%=UID_M0003%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	                                                                                                      