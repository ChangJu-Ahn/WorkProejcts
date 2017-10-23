<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		: 공정별원가 
'*  3. Program ID			: C4005MA1.asp
'*  4. Program Name			:배부요소DATA등록 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4005Mb1.asp
'*						
'*  7. Modified date(First)	: 2005/09/05
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: 
'* 11. Comment				: 
'* 12. History              : 
'*                          : 
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%	

call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 

    Dim lgOpModeCRUD
    Dim iStrCode
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	 
	lgOpModeCRUD  = Request("txtMode") 

	
	iStrCode=request("txtYYYYMM") & gColSep	
	iStrCode=iStrCode &  request("txtCode") & gColSep	
	iStrCode=iStrCode &  request("txtFctrCd")  & gColSep
	iStrCode =iStrCode &  request("txtGubun")  & gColSep	

								                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
         Case CStr("btnCopyPrev")														'☜: copy    
         	Call SubCopyPrev()			
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	Dim iPC4G005																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iMax
	Dim PvArr
	
	
	Dim EG1_export_group
	Const c_i1_code = 0
	Const c_i1_code_nm =1		
	Const c_i1_fctr_nm =2
	Const c_i1_flag_nm =3
	Const c_i1_fctr_cd =4		
	Const c_i1_alloc_data = 5

  	Dim E2_con_code(1)
    Const con_E2_code = 0
    Const con_E2_code_nm = 1

  	Dim E3_c_dstb_fctr_s(1)
    Const c_E3_dstb_fctr_cd = 0
    Const c_E3_dstb_fctr_nm = 1  
    	
	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 

	Dim intARows
	Dim intTRows
	intARows=0
	intTRows=0
	
	Const C_SHEETMAXROWS_D  = 100    

	StrNextKey = Trim(Request("lgStrPrevKey"))	

   Set iPC4G005 = Server.CreateObject("PC4G005.CcListMFCAlloc")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------    
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPC4G005 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함	
	End if


    '-----------------------
    'Com action area
    '-----------------------    
      Call iPC4G005.C_LIST_MFC(gStrGlobalCollection, C_SHEETMAXROWS_D, iStrCode,EG1_export_group, _
				E2_con_code, E3_c_dstb_fctr_s,StrNextKey)	
	
	
	If request("txtGubun")="C" Then	'cc
		If CheckSYSTEMError(Err,True) = True Then
			Set iPC4G005 = Nothing
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData2.MaxRows = 0" & vbCr
			Response.Write "parent.lgStrPrevKey2 = """"" & vbCr 	
			Response.Write " parent.frm1.hGubun.value    = """ & ConvSPChars(Request("txtGubun")) & """" & vbCr
			Response.Write "parent.DbQueryOk " & intARows & ",iMaxRow"   & vbCr 
			Response.Write "</Script>"
			Exit Sub
		End If
	Else										'wp
		If CheckSYSTEMError(Err,True) = True Then
			Set iPC4G005 = Nothing
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write " parent.frm1.vspdData.MaxRows = 0" & vbCr
			Response.Write "parent.lgStrPrevKey = """"" & vbCr 	
			Response.Write " parent.frm1.hGubun.value    = """ & ConvSPChars(Request("txtGubun")) & """" & vbCr
			Response.Write "parent.DbQueryOk " & intARows & ",iMaxRow"   & vbCr 
			Response.Write "</Script>"
			Exit Sub
		End If
	End If

	
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "with parent" & vbCr
		Response.Write "	.frm1.txtFctrNm.value = """ & ConvSPChars(E3_c_dstb_fctr_s(c_E3_dstb_fctr_nm))      & """" & vbCr
		Response.Write "	.frm1.txtWPNm.value  = """ & ConvSPChars(E2_con_code(con_E2_code_nm))      & """" & vbCr
		Response.Write "	.frm1.hGubun.value    = """ & ConvSPChars(Request("txtGubun")) & """" & vbCr
		Response.Write "End With "   & vbCr
		Response.Write "</Script>"                  & vbCr
			
		iLngMaxRow = CLng(Request("txtMaxRows"))											'Save previous Maxrow                                                
		GroupCount = UBound(EG1_export_group,1)

		'-----------------------
		'Result data display area
		'----------------------- 
	
		iMax = UBound(EG1_export_group,1)
		ReDim PvArr(iMax)
		strNextKey= ubound(EG1_export_group(imax,c_i1_alloc_data+1))
	If request("txtGubun")="C" Then	'cc 	
		For iLngRow = 0 To UBound(EG1_export_group,1)
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,c_i1_code))		   
			iStrData = iStrData & chr(11) & ""                                                                      
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_code_nm))  
		    istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,c_i1_flag_nm))   
		    istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,c_i1_fctr_cd))	
		    iStrData = iStrData & chr(11) & ""
		    istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_fctr_nm)) 
		    istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, c_i1_alloc_data),ggExchRate.DecPoint,0)
		   istrData = istrData & Chr(11) &  EG1_export_group(iLngRow, c_i1_alloc_data+1)		'max
		    'istrData = istrData & Chr(11) & iLngMaxRow + iLngRow + 1                             
		    istrData = istrData & Chr(11) & Chr(12)               
			PvArr(iLngRow) = istrData
			istrData=""
		Next    
		istrData = Join(PvArr, "")
		intARows=iLngRow+1	

		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "Dim iMaxRow " & vbCr
		Response.Write " iMaxRow = .frm1.vspdData2.maxrows" & vbCr
		Response.Write "	.ggoSpread.Source    = .frm1.vspdData2 " & vbCr
		Response.Write "	.ggoSpread.SSShowData  """ & istrData	& """" & vbCr	
		Response.Write "	.lgStrPrevKey2        = """ & StrNextKey & """" & vbCr  
		Response.Write " .frm1.hCode.value    = """ & ConvSPChars(Request("txtCode"))   & """" & vbCr
		Response.Write " .frm1.hFctrCd.value     = """ & ConvSPChars(Request("txtFctrCd"))    & """" & vbCr
		Response.Write " .frm1.hGubun.value    = """ & ConvSPChars(Request("txtGubun")) & """" & vbCr
		Response.Write " .frm1.hYYYYMM.value     = """ & ConvSPChars(Request("txtYYYYMM"))      & """" & vbCr
		Response.Write " .DbQueryOk " & intARows & ",iMaxRow"   & vbCr 
		Response.Write " .frm1.vspdData2.focus "		   & vbCr 
		Response.Write "End With"   & vbCr
		Response.Write "</Script>" & vbCr    
	Else	
		For iLngRow = 0 To UBound(EG1_export_group,1)
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,c_i1_code))                                                            
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_code_nm))  
			iStrData = iStrData & chr(11) & ""          
		    istrData = istrData & Chr(11)  & ConvSPChars(EG1_export_group(iLngRow,c_i1_flag_nm))      
		    istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,c_i1_fctr_cd))	
		    iStrData = iStrData & chr(11) & ""
		    istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_fctr_nm)) 
		    istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, c_i1_alloc_data),ggExchRate.DecPoint,0)
		    istrData = istrData & Chr(11) &  EG1_export_group(iLngRow, c_i1_alloc_data+1)		'max
		    
		    istrData = istrData & Chr(11) & Chr(12)               
			PvArr(iLngRow) = istrData
			istrData=""
		Next    
		istrData = Join(PvArr, "")
			
		intARows=iLngRow+1
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "Dim iMaxRow " & vbCr
		Response.Write " iMaxRow = .frm1.vspdData.maxrows" & vbCr
		Response.Write "	.ggoSpread.Source    = .frm1.vspdData " & vbCr
		Response.Write "	.ggoSpread.SSShowData  """ & istrData	& """" & vbCr	
		Response.Write "	.lgStrPrevKey        = """ & StrNextKey & """" & vbCr  
		Response.Write " .frm1.hCode.value    = """ & ConvSPChars(Request("txtCode"))   & """" & vbCr
		Response.Write " .frm1.hFctrCd.value     = """ & ConvSPChars(Request("txtFctrCd"))    & """" & vbCr
		Response.Write " .frm1.hGubun.value    = """ & ConvSPChars(Request("txtGubun")) & """" & vbCr
		Response.Write " .frm1.hYYYYMM.value     = """ & ConvSPChars(Request("txtYYYYMM"))      & """" & vbCr
		Response.Write " .DbQueryOk " & intARows & ",iMaxRow"   & vbCr 
		Response.Write " .frm1.vspdData.focus "		   & vbCr 
		Response.Write "End With"   & vbCr
		Response.Write "</Script>" & vbCr    	
	End If
		
    Set iPC4G005 = Nothing											'☜: Unload Comproxy
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()


On Error Resume Next																'☜: 

Dim PC4G005																			'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim strCode																			'☆ : Lookup 용 코드 저장 변수 
Dim strProdOrderNo																	'☆ : Lookup 용 코드 저장 변수									
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

Dim iStrCodeL

Err.Clear																		'☜: Protect system from crashing

strMode = Request("txtMode")														'☜ : 현재 상태를 받음 

Err.Clear																		'☜: Protect system from crashing


itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
iDCount  = Request.Form("txtDSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount + iDCount)
             
For ii = 1 To iDCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
Next
For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

Set PC4G005 = Server.CreateObject("PC4G005.cCMFCAllocSvr")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

If request("hGubun")="C" Then	'cc 
	iStrCodeL=request("txtYYYYMM") & gColSep	
	iStrCodeL=iStrCodeL &  request("txtCCCd") & gColSep	
	iStrCodeL=iStrCodeL &  request("txtFctrCd")  & gColSep
	iStrCodeL =iStrCodeL &  request("hGubun")  & gColSep	
	
	Call PC4G005.C_MFC_ALLOC_BASIS_SVR(gStrGlobalCollection, iStrCodeL, itxtSpread, iErrorPosition)

	If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
		Set PC4G005 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "Call parent.SheetFocus(""B""," & iErrorPosition & ",1)" & vbCrLF
		Response.Write "</Script>" & vbCrLF	

		Response.End
	End If
Else
	iStrCodeL=request("txtYYYYMM") & gColSep	
	iStrCodeL=iStrCodeL &  request("txtWPCd") & gColSep	
	iStrCodeL=iStrCodeL &  request("txtFctrCd")  & gColSep
	iStrCodeL =iStrCodeL &  request("hGubun")  & gColSep	
	
	Call PC4G005.C_MFC_ALLOC_BASIS_SVR(gStrGlobalCollection, iStrCodeL, itxtSpread, iErrorPosition)


	If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
		Set PC4G005 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "Call parent.SheetFocus(""A""," & iErrorPosition & ",1)" & vbCrLF
		Response.Write "</Script>" & vbCrLF	

		Response.End
	End If

End If

Set PC4G005 = Nothing	

Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
        
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
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    
    On Error Resume Next

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub
'============================================================================================================
' Name : SubCopyPrev
' Desc : This method copy data from previous month  to current month in DB
'============================================================================================================
Sub SubCopyPrev()
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	Dim cPC4G005																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iMax
	Dim PvArr
	
	Dim strYYYYMM1,strGubun
	Dim strYYYYMM2
	Dim iErrorPosition
 
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 

	Dim intARows
	Dim intTRows
	intARows=0
	intTRows=0
	
	strGubun = request("txtGubun")
	strYYYYMM1=Request("txtYYYYMM1")
	strYYYYMM2=Request("txtYYYYMM2")
	
	
   Set cPC4G005 = Server.CreateObject("PC4G005.cCMFCAllocCopy")   
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------    
    
    If CheckSYSTEMError(Err,True) = true Then 		
		Set cPC4G005 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if

    '-----------------------
    'Com action area
    '-----------------------
	Call cPC4G005.C_MFC_ALLOC_COPY(gStrGlobalCollection,strGubun, strYYYYMM1,strYYYYMM2,  iErrorPosition)

	If Trim(iErrorPosition) <>"" Then	
		If CheckSYSTEMError2(Err, True, iErrorPosition , "", "", "", "") = True Then
			Set PC4G005 = Nothing
			Response.End
		End If	
	Else
		If CheckSYSTEMError(Err,True) = True Then
			Response.End 
		End If
	End If
	
	Set cPC4G005 = Nothing	

	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End    
		
End Sub    

%>
