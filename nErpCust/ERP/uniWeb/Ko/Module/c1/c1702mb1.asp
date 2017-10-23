<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%'======================================================================================================
'*  1. Module Name          : COST
'*  2. Function Name        : C_COST_BASIC_DATA_FOR_COSTING
'*  3. Program ID           : c1702mb1.asp
'*  4. Program Name         : 배부규칙 등록 
'*  5. Program Desc         : 배부규칙 조회, 등록, 수정, 삭제 
'*  6. Modified date(First) : 2000/11/03
'*  7. Modified date(Last)  : 2002/06/18
'*  8. Modifier (First)     : Cho Ig Sung / Park, Joon-Won
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
	Call LoadBasisGlobalInf()

'	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
'    Dim lgLngMaxRow
    
	On Error Resume Next								'☜: 
	Err.Clear

	Call HideStatusWnd


'---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'   lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

'------ Developer Coding part (Start ) ------------------------------------------------------------------
	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizCopy()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	Dim PC1G040Data		
    Dim iStrData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim importArray
    Dim iIntLoopCount
	Dim iVerCd 
	Dim lgMaxCount

	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
	Const C_VerCd      = 2
     
    Const C_SHEETMAXROWS_D  = 100                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
          
   lgMaxCount = CInt(C_SHEETMAXROWS_D)                  '☜: Max fetched data at a time
     
	'Key 값을 읽어온다 
	iVerCd      = Trim(Request("txtVerCd"))
	iStrPrevKey = Trim(Request("lgStrPrevKey"))
	
  	  	
    'Component 입력변수        
    ReDim importArray(2)
   
    importArray(C_MaxFetchRc)	= lgMaxCount        
	importArray(C_NextKey)		= iStrPrevKey
	importArray(C_VerCd)		= iVerCd
   
    Set PC1G040Data = Server.CreateObject("PC1G040.cCListDstRlSvr")
    
	If CheckSYSTEMError(Err, True) = True Then	
	   Call SetErrorStatus									
       Exit Sub
    End If    
    
    Call PC1G040Data.C_LIST_DSTB_RULE_SVR(gStrGlobalCollection,importArray, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then
	   Call SetErrorStatus										
       Set PC1G040Data = Nothing
       Exit Sub
    End If    
  
    Set PC1G040Data = nothing 
    
    Const E_WorkStep = 0	
	Const E_MinorNm = 1
	Const E_DstbFctrCd = 2
	Const E_DstbFctrNm = 3   
	
    iStrData = ""
    iIntLoopCount = 0	

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
		
   	    If  iIntLoopCount < (lgMaxCount + 1) Then
   	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_WorkStep)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_MinorNm)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_DstbFctrCd)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_DstbFctrNm)))
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
        Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_WorkStep)
			Exit For
		
		End If
 	Next

	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey  = ""
	End If

	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData			& """" & vbCr
    Response.Write " .frm1.hVerCd.value = """ & iVerCd        & """" & vbCr
    Response.Write " .lgStrPrevKey          = """ & ConvSPChars(iStrPrevKey)	& """" & vbCr
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr  

End Sub    	 



'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
  
    Dim PC1G040Data
    Dim importString
    Dim txtSpread
    Dim iErrPosition 
    
    importString = Trim(Request("txtVerCd"))
    txtSpread    = Trim(Request("txtSpread"))
 
    Set PC1G040Data = Server.CreateObject("PC1G040.cCMngDstRlSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC1G040Data.C_MANAGE_DSTB_RULE_SVR(gStrGlobalCollection, importString,  txtSpread, iErrPosition)			
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then															
       Call SetErrorStatus
       Set PC1G040Data = Nothing
       Exit Sub
    End If    
    
    Set PC1G040Data = Nothing
End Sub    

'============================================================================================================
' Name : SubBizCopy
' Desc : 
'============================================================================================================
Sub SubBizCopy()
	Dim iOldVerCd
	Dim iNewVerCd
	Dim IntRetCD
	Dim strMsg_cd
	
	iOldVerCd = Trim(Request("txtVerCd"))
	iNewVerCd = Trim(Request("txtNewVerCd"))

	Call SubCreateCommandObject(lgObjComm)	
    
	With lgObjComm
		.CommandTimeOut = 0
	    .CommandText = "usp_c_copy_dstb_rule"
	    .CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    .Parameters.Append .CreateParameter("@old_ver_cd"  ,adVarXChar,adParamInput,6,iOldVerCd)
		.Parameters.Append .CreateParameter("@new_ver_cd"  ,adVarXChar,adParamInput,6,iNewVerCd)
	    .Parameters.Append .CreateParameter("@usrid"       ,adVarXChar,adParamInput,13, gUsrID)
	    .Parameters.Append .CreateParameter("@msgcd"       ,adVarXChar,adParamOutput,6)

	    .Execute ,, adExecuteNoRecords
	End With

	If  Err.number = 0 Then
	    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

	    If  IntRetCD <> 1 Then
	        strMsg_cd = lgObjComm.Parameters("@msgcd").Value
	        Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
	    End If
	Else
	    lgErrorStatus     = "YES"                                                         '☜: Set error status
	    Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End If

	Call SubCloseCommandObject(lgObjComm)
	
	Response.Write " <Script Language=vbscript>	                    " & vbCr
	Response.Write " With parent                                    " & vbCr
    Response.Write "	.frm1.txtVerCd.value = """ & iNewVerCd & """" & vbCr
    Response.Write "	.frm1.txtNewVerCd.value = """ & """         " & vbCr
    Response.Write " End With										" & vbCr
    Response.Write " </Script>										" & vbCr 	
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
'          Else
'            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
</Script>	
