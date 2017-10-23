<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Finance
'*  2. Function Name        : Finance Managements
'*  3. Program ID           : F5103mb1
'*  4. Program Name         : 어음수표번호자동채번 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/09/26
'*  8. Modified date(Last)  : 2002/08/08
'*  9. Modifier (First)     : hersheys
'* 10. Modifier (Last)      : Shin Myoung_Ha
'* 11. Comment              : 1. FilterVar()함수 적용 - 2002/07/31
'*							  2. FilterVar()함수 제거(Com에서 적용함) - 2002/08/08
'**********************************************************************************************




'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->


<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call LoadBasisGlobalInf()

Call HideStatusWnd

On Error Resume Next														'☜: 
ERR.CLEAR

Dim PAFG515				       				                                '☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim strMode
Dim	I1_b_bank
Dim I2_f_note

ReDim I2_f_note(5)

Const A824_I2_note_kind = 0
Const A824_I2_issue_dt = 1
Const A824_I3_note_no_head = 2
Const A824_I3_from_note_no = 3
Const A824_I3_to_note_no = 4

strMode = Trim(Request("txtMode"))

'☜ : 현재 상태를 받음 
If strMode = Trim(UID_M0002) Then	
	Call SubBizSaveMulti()
End If

Sub SubBizSaveMulti()
  
    ON ERROR RESUME NEXT
    Err.Clear                                                               '☜: Protect system from crashing        
			
    Set PAFG515 = Server.CreateObject("PAFG515.cFExecNoteNoSvr")       
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End IF    
    I1_b_bank						= Trim(Request("txtBankCd"))
    I2_f_note(A824_I2_note_kind)	= Request("cboNoteKind")
    I2_f_note(A824_I2_issue_dt)		= UNIConvDate(Request("txtIssueDt"))    
    I2_f_note(A824_I3_note_no_head)	= UCase(Trim(Request("txtNoteNo")))
    I2_f_note(A824_I3_from_note_no)	= Trim(Request("txtFromNo"))
    I2_f_note(A824_I3_to_note_no)	= Trim(Request("txtToNo"))

	'##############################################
	'DEBUG
	'Response.Write I2_f_note(A824_I2_note_kind)
	'Response.write I2_f_note(A824_I2_issue_dt)
	'Response.Write I2_f_note(A824_I3_note_no_head)
	'Response.Write I2_f_note(A824_I3_from_note_no)
	'Response.Write I2_f_note(A824_I3_to_note_no)
	'##############################################
	
		
	Call PAFG515.FN0032_EXECUTE_NOTE_NO_SVR(gStrGlobalCollection,I1_b_bank,I2_f_note)
	
	If CheckSYSTEMError(Err, True) = True Then
		Set PAFG515 = Nothing		
		Exit Sub
	ELSE
		Call DisplayMsgBox("990000", vbOKOnly, "", "", I_MKSCRIPT)	'정상처리되었습니다.
	End If
	
End Sub
%>

