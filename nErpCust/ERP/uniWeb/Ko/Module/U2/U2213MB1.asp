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
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)					 '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001) 
             Call SubBizQueryCond()                                                        '☜: Query
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
	
	If lgKeyStream(5) <> "" AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CP")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPLANTNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    'Call DisplayMsgBox("SCM003", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPLANTNm.value = """ & ConvSPChars(lgObjRs("PLANT_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPLANTNm.value = """"" & vbCrLf	
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
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF

			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("RET_DT"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
			lgstrData = lgstrData & Chr(11)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_UNIT"))
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RET_TYPE"))
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
            
            If ConvSPChars(lgObjRs("status")) = "O" Then
				lgstrData = lgstrData & Chr(11) & "실행"
			ElseIf ConvSPChars(lgObjRs("status")) = "E" Then
				lgstrData = lgstrData & Chr(11) & "확정"	
			Else
				lgstrData = lgstrData & Chr(11) & "계획"		
			End If
			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("STATUS"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
			lgstrData = lgstrData & Chr(11)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("D_BP_CD"))
			lgstrData = lgstrData & Chr(11)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))
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
            Case "C"                            '☜: Create 
				Call SubBizSaveMultiCreate(arrColVal)
            Case "U"                            '☜: Create 
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
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear

    lgStrSQL = "INSERT INTO M_SCM_RET ( "
    lgStrSQL = lgStrSQL & vbCrLf & " BP_CD,D_BP_CD,RET_DT,ITEM_CD,QTY,UNIT, PLANT_CD, "
    lgStrSQL = lgStrSQL & vbCrLf & " STATUS , RET_TYPE , "
	lgStrSQL = lgStrSQL & vbCrLf & " INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT ) "
    lgStrSQL = lgStrSQL & vbCrLf & " VALUES			( "
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(6))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(7))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(2))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(3))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(4))),"","D")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(5))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(REQUEST("txtPlantCd"))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & " 'P',  "
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(8))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(gUsrId,"","S")                      & "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(GetSvrDateTime,"''","S")			 & "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(gUsrId,"","S")                      & "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(GetSvrDateTime,"''","S")			 & ")"     
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
    
    lgStrSQL =                     " UPDATE M_SCM_RET SET "
	lgStrSQL = lgStrSQL & vbCrLf & "        QTY			  = " & FilterVar(Trim(UCase(arrColVal(04))),"","D") & "," 
	lgStrSQL = lgStrSQL & vbCrLf & "        RET_TYPE	  = " & FilterVar(Trim(UCase(arrColVal(08))),"","S")  
	lgStrSQL = lgStrSQL & vbCrLf & "  WHERE BP_CD         = " & FilterVar(Trim(UCase(arrColVal(06))),"","S")  
	lgStrSQL = lgStrSQL & vbCrLf & "    AND D_BP_CD       = " & FilterVar(Trim(UCase(arrColVal(07))),"","S")
	lgStrSQL = lgStrSQL & vbCrLf & "    AND RET_DT		  = " & FilterVar(Trim(UCase(arrColVal(02))),"","S")
	lgStrSQL = lgStrSQL & vbCrLf & "    AND ITEM_CD		  = " & FilterVar(Trim(UCase(arrColVal(03))),"","S")
	

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
    
    lgStrSQL =                     " DELETE FROM M_SCM_RET "		
	lgStrSQL = lgStrSQL & vbCrLf & "  WHERE BP_CD         = " & FilterVar(Trim(UCase(arrColVal(04))),"","S")  
	lgStrSQL = lgStrSQL & vbCrLf & "    AND D_BP_CD       = " & FilterVar(Trim(UCase(arrColVal(05))),"","S")
	lgStrSQL = lgStrSQL & vbCrLf & "    AND RET_DT        = " & FilterVar(Trim(UCase(arrColVal(02))),"","S")
	lgStrSQL = lgStrSQL & vbCrLf & "    AND ITEM_CD       = " & FilterVar(Trim(UCase(arrColVal(03))),"","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
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
					lgStrSQL =            "	SELECT TOP " & iSelCount  & " A.* , (B.BP_NM) AS PLANT_NM ,C.BP_NM , D.ITEM_NM , D.SPEC , D.BASIC_UNIT , STATUS , A.RET_TYPE , E.MINOR_NM "
					lgStrSQL = lgStrSQL & "	  FROM M_SCM_RET A , B_BIZ_PARTNER B , B_BIZ_PARTNER C , B_ITEM D , B_MINOR E "
					lgStrSQL = lgStrSQL & "	 WHERE A.BP_CD     = B.BP_CD "
					lgStrSQL = lgStrSQL & "	   AND A.D_BP_CD   = C.BP_CD "
					lgStrSQL = lgStrSQL & "	   AND A.ITEM_CD = D.ITEM_CD "
					lgStrSQL = lgStrSQL & "	   AND E.MAJOR_CD = 'B9017' "
					lgStrSQL = lgStrSQL & "	   AND A.RET_TYPE *= E.MINOR_CD "
					
					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "  AND A.D_BP_CD = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
					End If
					
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "  AND A.RET_DT >= " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
					End If
					
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "  AND A.RET_DT <= " & FilterVar(lgKeyStream(2) & "" ,"''", "S")
					End If
					
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "  AND A.BP_CD = " & FilterVar(lgKeyStream(3) & "" ,"''", "S")
					End If
					
					If lgkeystream(4) = "P" Then
						lgStrSQL = lgStrSQL & "  AND A.STATUS = 'P' "
					ElseIf lgkeystream(4) = "O" Then	
						lgStrSQL = lgStrSQL & "  AND A.STATUS = 'O' "
					ElseIf lgkeystream(4) = "E" Then		
						lgStrSQL = lgStrSQL & "  AND A.STATUS = 'E' "	
					End If	
					
					lgStrSQL = lgStrSQL & "  ORDER BY A.RET_DT ASC , A.ITEM_CD ASC "
					
           End Select             

        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "B"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
               Case "P"
                    lgStrSQL =            " select PLANT_nm from b_PLANT where PLANT_cd = " & FilterVar(lgKeyStream(5) & "" ,"''", "S")     
                    
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
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk("<%=lgStrPrevKey%>")
	         End with
	      Else
				Parent.DBQueryNotOk()
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