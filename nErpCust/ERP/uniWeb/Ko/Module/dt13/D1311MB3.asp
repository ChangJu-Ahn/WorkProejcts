<%@ LANGUAGE= VBScript %>
<% Option Explicit%>

<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4114mb3.asp
'*  4. Program Name			: Operation Management
'*  5. Program Desc			: Save Production Order Detail
'*  6. Comproxy List		: PD1G101.cPMngProdOrdDtl
'*  7. Modified date(First)	: 2001/06/30
'*  8. Modified date(Last) 	: 2002/07/08
'*  9. Modifier (First)		: Park, Bum-Soo
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%														'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf

Call HideStatusWnd

On Error Resume Next

Dim oPD1G101
Dim iErrorPosition																			'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim itxtSpread, itxtDtlSpread

Dim itxtDisuseReasonArray

Dim pvBtnFlag 
Dim pvTaxType
Dim pvChangeCode

Err.Clear																		'☜: Protect system from crashing


pvBtnFlag = Trim(Request("txtbtnFlag"))
pvTaxType  = "SD"


Set oPD1G101 = Server.CreateObject("PD1G101.cDSendDTax")

    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

		
Select Case pvBtnFlag	
		
	Case "Resend"
		itxtSpread = Trim(request("txtSpread"))
		itxtDtlSpread = Trim(request("txtDtlSpread"))
		Call oPD1G101.D_RE_SEND_DTAX(gStrGlobalCollection, _
								itxtSpread , _
								itxtDtlSpread, _
								iErrorPosition)
								
		If CheckSYSTEMError(Err,True) = True Then
			Set oPD1G101 = Nothing
			Response.Write "<Script Language=VBScript>" & vbCrLF
			'Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
			Response.Write "Call parent.SheetFocus2(" & iErrorPosition & ", 1)" & vbCrLF
			Response.Write "</Script>" & vbCrLF
			Response.End
		End If						
	
	Case "ChangeDocStatus"
		itxtSpread = Trim(request("txtSpread"))
		pvChangeCode = Trim(request("txtChangeStatus"))
		itxtDisuseReasonArray = Trim(request("txtDtlSpread"))
		Call oPD1G101.D_CHANGE_STATUS_DTAX(gStrGlobalCollection, _
										itxtSpread, _
										pvChangeCode, _
										itxtDisuseReasonArray, _
										iErrorPosition)
		If CheckSYSTEMError(Err,True) = True Then
			Set oPD1G101 = Nothing
			If Isempty(iErrorPosition) Then iErrorPosition = 0
			Response.Write "<Script Language=VBScript>" & vbCrLF
			'Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
			Response.Write "Call parent.SheetFocus2(" & iErrorPosition & ", 1)" & vbCrLF
			Response.Write "</Script>" & vbCrLF
			Response.End
		End If

	Case "Save"
		itxtSpread = Trim(request("txtSpread"))

		Call oPD1G101.D_SET_IVNO(gStrGlobalCollection, _
								itxtSpread, _
								"MM", _
								iErrorPosition)
		If CheckSYSTEMError(Err,True) = True Then
			Set oPD1G101 = Nothing
			If Isempty(iErrorPosition) Then iErrorPosition = 0
			Response.Write "<Script Language=VBScript>" & vbCrLF
			'Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
			Response.Write "Call parent.SheetFocus2(" & iErrorPosition & ", 1)" & vbCrLF
			Response.Write "</Script>" & vbCrLF
			Response.End
		End If
End Select
	

If Not (oPD1G101 Is nothing) Then
	Set oPD1G101 = Nothing								'☜: Unload Comproxy	
End If
        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>


