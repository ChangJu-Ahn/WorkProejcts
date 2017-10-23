<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		: 공정별원가 
'*  3. Program ID			: C4006MA1.asp
'*  4. Program Name			:완성품환산율등록 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4006Mb1.asp
'*						
'*  7. Modified date(First)	: 2005/08/29
'*  8. Modified date(Last)	: 2005/11/03
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: HJO
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
	iStrCode=iStrCode &  request("txtWcCd") & gColSep	
	iStrCode=iStrCode &  request("txtProdOrderNo")  & gColSep
	iStrCode =iStrCode &  request("txtPlantCd")  & gColSep
	

								                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             'Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
         Case CStr("btnCopyPrev")														'☜: copy    
         	Call SubCopyPrev()			
    End Select

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()


On Error Resume Next																'☜: 

Dim pC4G006																			'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim strCode																			'☆ : Lookup 용 코드 저장 변수 
Dim strProdOrderNo																	'☆ : Lookup 용 코드 저장 변수									
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii
const i_yyyymm=0
const i_wc_cd=1

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

Set pC4G006 = Server.CreateObject("pC4G006.cCProdRateByOrsSvr")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Call pC4G006.C_PROD_RATE_BY_ORS_S_SVR(gStrGlobalCollection, iStrCode, itxtSpread, iErrorPosition)


If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set pC4G006 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",1)" & vbCrLF
	Response.Write "</Script>" & vbCrLF	
	'Exit Sub
	Response.End
End If

Set pC4G006 = Nothing	

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

	Dim cPC4G006																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iMax
	Dim PvArr
	
	Dim strYYYYMM1
	Dim strYYYYMM2
	Dim iErrorPosition
 
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 

	Dim intARows
	Dim intTRows
	intARows=0
	intTRows=0
	
	
	strYYYYMM1=Request("txtYYYYMM1")
	strYYYYMM2=Request("txtYYYYMM2")
	
	
   Set cPC4G006 = Server.CreateObject("PC4G006.cCProdRateCopy")   
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------    
    
    If CheckSYSTEMError(Err,True) = true Then 		
		Set cPC4G006 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if

    '-----------------------
    'Com action area
    '-----------------------
	Call cPC4G006.C_PROD_COPY(gStrGlobalCollection, strYYYYMM1,strYYYYMM2,  iErrorPosition)

	If Trim(iErrorPosition) <>"" Then	
		If CheckSYSTEMError2(Err, True, iErrorPosition , "", "", "", "") = True Then
			Set pC4G006 = Nothing
			Response.End
		End If	
	Else
		If CheckSYSTEMError(Err,True) = True Then
			Response.End 
		End If
	End If
	
	Set cPC4G006 = Nothing	

	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End    
		
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
