<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4111mb2.asp
'*  4. Program Name			: Manage Production Order
'*  5. Program Desc			: Create / Update / Delete
'*  6. Comproxy List		: +PP4C103.cPMngProdOrd
'*  7. Modified date(First)	: 2000/03/30
'*  8. Modified date(Last) 	: 2002/07/09
'*  9. Modifier (First)		: Kim, GyoungDon
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

On Error Resume Next

Dim oPP4C103
Dim iErrorPosition
Dim strProdOrderNo
Dim lgIntFlgMode
Dim strTxtSpread

Err.Clear																		'☜: Protect system from crashing

lgIntFlgMode = CInt(Request("txtFlgMode"))										'☆ : 입력/수정용 ComProxy Dll 사용 변수 

Set oPP4C103 = Server.CreateObject("PP4C103.cPMngProdOrd")

strTxtSpread = Request("txtSpread")

'Call ServerMesgBox(Request("txtSpread"), vbCritical, I_MKSCRIPT)
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
	
Call  oPP4C103.P_MANAGE_PRODUCTION_ORDER(gStrGlobalCollection, _
										strTxtSpread, strProdOrderNo, iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
	Set oPP4C103 = Nothing
	Response.End
End If

Set oPP4C103 = Nothing								'☜: Unload Comproxy	

    	
If lgIntFlgMode = OPMD_UMODE Then
	strProdOrderNo = Request("txtProdOrderNo1")												'☆: Production Order No.
End If

%>
<Script Language=vbscript>
    With parent	 	
		.frm1.txtProdOrderNo.Value = "<%=ConvSPChars(Cstr(strProdOrderNo))%>"										'☆: Production Order No.
		If .lgOPMDMode = "UPDATE" Then
		   	Call .DbSaveOk(False)
		Else
			Call .DbDeleteOk()
		End If
    End With
</Script>