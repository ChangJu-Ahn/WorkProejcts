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
	Dim lgSeq
	Dim lgQty
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

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
			 Call SubBizSave()
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
    
    Call SubMakeSQLStatements("S")                                                  '¢Ð : Make sql statements ,SR : Single Read
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then               'R(Read) X(CursorType) X(LockType) 
      
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '¢Ð : No data is found. 
      
    Else
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	
       Response.Write  " <Script Language=vbscript>" & vbCr
       Response.Write  " With Parent               " & vbCr
       
       Response.Write  "       .Frm1.txtDlvyNo2.Value   = """ & ConvSPChars(lgObjRs("DLVY_NO"))          & """" & vbCr
       Response.Write  "       .Frm1.txtDlvyTime.Value	= """ & ConvSPChars(lgObjRs("DLVY_TIME"))        & """" & vbCr
       Response.Write  "       .Frm1.txtTransCo.Value	= """ & ConvSPChars(lgObjRs("TRANS_CO"))         & """" & vbCr
       Response.Write  "       .Frm1.txtVehicleNo.Value	= """ & ConvSPChars(lgObjRs("VEHICLE_NO"))       & """" & vbCr
       Response.Write  "       .Frm1.txtDriver.Value	= """ & ConvSPChars(lgObjRs("DRIVER"))           & """" & vbCr
       Response.Write  "       .Frm1.txtTelNo1.Value	= """ & ConvSPChars(lgObjRs("TEL_NO1"))          & """" & vbCr
       Response.Write  "       .Frm1.txtTelNo2.Value	= """ & ConvSPChars(lgObjRs("TEL_NO2"))          & """" & vbCr
       Response.Write  "       .Frm1.txtDlvyPlace.Value	= """ & ConvSPChars(lgObjRs("DLVY_PLACE"))       & """" & vbCr
       Response.Write  "       .Frm1.txtREMARK.Value	= """ & ConvSPChars(lgObjRs("REMARK"))           & """" & vbCr
       
       'Response.Write  "       .DBQueryOk " & vbCr
       Response.Write  " End With         " & vbCr
       Response.Write  " </Script>        " & vbCr
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
    End If

    Call SubCloseRs(lgObjRs)                                                    '¢Ð : Release RecordSSet
    
    Call SubBizQueryCond()
    If lgErrorStatus <> "YES" Then
		Call SubBizQueryMulti()
	End If
End Sub    

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    Dim lgIntFlgMode
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE  : Call SubBizSaveSingleCreate()                            '¢Ð : Create
        Case  OPMD_UMODE  : Call SubBizSaveSingleUpdate()                            '¢Ð : Update
    End Select
	
End Sub	

'============================================================================================================
' Name : SubBizQueryCond
' Desc : Verify Condition Value from Db
'============================================================================================================
Sub SubBizQueryCond()

    On Error Resume Next
    Err.Clear

	If lgKeyStream(0) <> "" AND lgErrorStatus <> "YES" Then
   
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
    On Error Resume Next
    Err.Clear

    strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR")
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
    '    Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
    '    Call SetErrorStatus()
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
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_UNIT"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PLAN_DVRY_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("D_BP_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPLIT_SEQ_NO"))
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
    Call SubCloseRs(lgObjRs)

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)
		
		Select Case arrColVal(0)
            Case "U"                            '¢Ð: Create 
				Call SubBizQuerySeq(arrColVal)
				Call SubBizSaveMultiUpdate(arrColVal)
			Case "D"
				Call SubBizSaveMultiDelete(arrColVal)
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
Sub SubBizQuerySeq(arrColVal)
    On Error Resume Next
    Err.Clear

	lgStrSQL =            " SELECT isNULL(MAX(DLVY_SEQ_NO),0) DLVY_SEQ_NO "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_FIRM_PUR_RCPT "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO = " & FilterVar(Trim(UCase(arrColVal(2))),"","S") 
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		If lgObjRs("DLVY_SEQ_NO") = "" OR lgObjRs("DLVY_SEQ_NO") = 0 Then
			lgSeq = 1
		Else
			lgSeq = lgObjRs("DLVY_SEQ_NO") + 1
		End If
    End If
    
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
					lgStrSQL =            "	SELECT TOP " & iSelCount  & " A.PO_NO , A.PO_SEQ_NO , A.SPLIT_SEQ_NO , B.ITEM_CD , B.TRACKING_NO , C.ITEM_NM , C.SPEC , "
					lgStrSQL = lgStrSQL & "	       C.BASIC_UNIT , A.PLAN_DVRY_DT , A.PLAN_DVRY_QTY , A.D_BP_CD ,  D.SL_NM "
					lgStrSQL = lgStrSQL & "	  FROM M_SCM_FIRM_PUR_RCPT A, m_pur_ord_dtl B, B_ITEM C , B_STORAGE_LOCATION D "
					lgStrSQL = lgStrSQL & "	 WHERE A.PO_NO     = B.PO_NO  "
					lgStrSQL = lgStrSQL & "	   AND A.PO_SEQ_NO = B.PO_SEQ_NO  "
					lgStrSQL = lgStrSQL & "	   AND B.ITEM_CD = C.ITEM_CD "
					lgStrSQL = lgStrSQL & "	   AND A.D_BP_CD = D.SL_CD "
					
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "  AND A.DLVY_NO = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
					End If
					
           End Select             

        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "B"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
           End Select 
           
		Case "S"
			
			lgStrSQL =            " SELECT * " 
			lgStrSQL = lgStrSQL & "   FROM M_SCM_DLVY_PUR_RCPT "
			lgStrSQL = lgStrSQL & "  WHERE BP_CD   =  " & FilterVar(lgKeyStream(0),"''", "S")
			lgStrSQL = lgStrSQL & "    AND DLVY_NO =  " & FilterVar(lgKeyStream(1),"''", "S")
			
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
                .ggoSpread.Source     = .frm1.vspdData1
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk("<%=lgStrPrevKey%>")
	         End with
	      Else
				Parent.DBQueryNotOk()
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
      
</Script>	                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               