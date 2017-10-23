<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Finance
'*  2. Function Name        : Finance Managements
'*  3. Program ID           : F5103mb1
'*  4. Program Name         : ������ǥ��ȣ�ڵ�ä�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/09/26
'*  8. Modified date(Last)  : 2002/08/08
'*  9. Modifier (First)     : hersheys
'* 10. Modifier (Last)      : Shin Myoung_Ha
'* 11. Comment              : 1. FilterVar()�Լ� ���� - 2002/07/31
'*							  2. FilterVar()�Լ� ����(Com���� ������) - 2002/08/08
'**********************************************************************************************




'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->


<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call LoadBasisGlobalInf()

Call HideStatusWnd

On Error Resume Next														'��: 
ERR.CLEAR

Dim PAFG515				       				                                '�� : �Է�/������ ComProxy Dll ��� ���� 
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

'�� : ���� ���¸� ���� 
If strMode = Trim(UID_M0002) Then	
	Call SubBizSaveMulti()
End If

Sub SubBizSaveMulti()
  
    ON ERROR RESUME NEXT
    Err.Clear                                                               '��: Protect system from crashing        
			
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
		Call DisplayMsgBox("990000", vbOKOnly, "", "", I_MKSCRIPT)	'����ó���Ǿ����ϴ�.
	End If
	
End Sub
%>

