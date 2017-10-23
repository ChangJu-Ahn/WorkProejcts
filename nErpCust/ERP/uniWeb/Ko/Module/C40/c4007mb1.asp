<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		: 공정별원가 
'*  3. Program ID			: C4007MA1.asp
'*  4. Program Name			:원부자재그룹별원가요소등록 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4007Mb1.asp
'*						
'*  7. Modified date(First)	: 2005/09/12
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

	
	iStrCode = request("txtItemGroup") & gColSep	

	

								                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	Dim iPC4G007																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iMax
	Dim PvArr
	
	Dim EG1_export_group

	Const c_i1_group_lvl		= 0
	Const c_i1_item_group		= 1
	Const c_i1_item_group_nm	= 2
	Const c_i1_cost_elmt_cd		= 3
	Const c_i1_cost_elmt_nm		= 4
	Const c_i1_com_cost_elmt_cd = 5
	Const c_i1_com_cost_elmt_nm = 6


  	Dim E3_b_item_group(1)
    Const b_E3_item_group		= 0
    Const b_E3_item_group_nm	= 1  
    	
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

	lgStrPrevKey = Trim(Request("lgStrPrevKey"))	
   Set iPC4G007 = Server.CreateObject("PC4G007.CcListCostElmt")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------    
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPC4G007 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함	
	End if

    '-----------------------
    'Com action area
    '-----------------------    
      Call iPC4G007.C_LIST_COST_ELMT(gStrGlobalCollection, C_SHEETMAXROWS_D, iStrCode,EG1_export_group, E3_b_item_group)
	
	If CheckSYSTEMError(Err,True) = True Then
		Set iPC4G007 = Nothing
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.write "parent.frm1.vspdData.MaxRows = 0" & vbCr
		'Response.Write "parent.DbQueryOk " & intARows & ",iMaxRow"   & vbCr
		Response.Write "parent.DbQueryOk " & vbCr  
		Response.Write "</Script>"
		Exit Sub
	End If
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtItemGroupNm.value = """ & ConvSPChars(E3_b_item_group(b_E3_item_group_nm))      & """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr
		
	iLngMaxRow = CLng(Request("txtMaxRows"))											'Save previous Maxrow                                                


	'-----------------------
	'Result data display area
	'----------------------- 
	
	iMax = UBound(EG1_export_group,1)
	ReDim PvArr(iMax)
	
	For iLngRow = 0 To iMax 'UBound(EG1_export_group,1)
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_group_lvl))  
		iStrData = iStrData & chr(11) & ""        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_item_group))      
        iStrData = iStrData & chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_item_group_nm))   
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_cost_elmt_cd))       
        iStrData = iStrData & chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_cost_elmt_nm))    
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_com_cost_elmt_cd))       
        iStrData = iStrData & chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, c_i1_com_cost_elmt_nm))    
       
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow + 1                             
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
	Response.Write " .frm1.hItemGroup.value     = """ & ConvSPChars(Request("txtItemGroup"))    & """" & vbCr	    
    Response.Write " .DbQueryOk "   & vbCr  
    Response.Write " .frm1.vspdData.focus "		   & vbCr 
    Response.Write "End With"   & vbCr
    Response.Write "</Script>" & vbCr    

    Set iPC4G007 = Nothing											'☜: Unload Comproxy
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
Response.Write "AA"

On Error Resume Next																'☜: 

Dim cPC4G007																			'☆ : 입력/수정용 ComProxy Dll 사용 변수 
'Dim strCode																			'☆ : Lookup 용 코드 저장 변수 
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount,ii


Err.Clear																		'☜: Protect system from crashing

strMode = Request("txtMode")														'☜ : 현재 상태를 받음 

Err.Clear																		'☜: Protect system from crashing


itxtSpread = ""
             
iCUCount = Request.Form("txtSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)
      

For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")


Set cPC4G007 = Server.CreateObject("PC4G007.cCostElmtByRawSvr")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Call cPC4G007.C_COST_ELMT_BY_RAW_S_SVR(gStrGlobalCollection, iStrCode, itxtSpread, iErrorPosition)


If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set cPC4G007 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",2)" & vbCrLF
	Response.Write "</Script>" & vbCrLF	
	Response.End
End If

Set cPC4G007 = Nothing	

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

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)

	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
%>
