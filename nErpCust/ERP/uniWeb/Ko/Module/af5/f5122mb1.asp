<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_RECEIPT
'*  3. Program ID        : f5104ma
'*  4. Program �̸�      : ��������ϰ�ó�� 
'*  5. Program ����      : ��������ϰ�ó�� 
'*  6. Comproxy ����Ʈ   : f5104ma
'*  7. ���� �ۼ������   : 2000/10/16
'*  8. ���� ���������   : 2002/02/15
'*  9. ���� �ۼ���       : ����ȯ 
'* 10. ���� �ۼ���       : ������ 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'*                         -2000/10/16 : ..........
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->


<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
On Error Resume Next                                                            '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
Dim lgOpModeCRUD
Dim lPtxtNoteNo


Call HideStatusWnd                                                               '��: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------

lgOpModeCRUD = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
 
'Response.End

'Tab 1
Const C_NOTE_NO		= 0
Const C_NOTE_AMT	= 1
Const C_DUE_DT		= 2
Const C_NOTE_STS	= 3 
Const C_BANK_CD		= 4
Const C_BANK_NM		= 5
Const C_BP_CD		= 6
Const C_BP_NM		= 7
Const C_DEPT_CD		= 8
Const C_DEPT_NM		= 9
Const C_GL_NO		= 10

'TAB2, vspddata2
Const C_CNCL_NOTE_NO		= 0
Const C_CNCL_TEMP_GL_NO		= 1
Const C_CNCL_TEMP_GL_DT		= 2
Const C_CNCL_GL_NO			= 3
Const C_CNCL_GL_DT			= 4
Const C_CNCL_NOTE_AMT		= 5
Const C_CNCL_BP_CD			= 6
Const C_CNCL_BP_NM			= 7
Const C_CNCL_DEPT_CD		= 8
Const C_CNCL_DEPT_NM		= 9
Const C_CNCL_RCPT_TYPE		= 10		'��: hidden field(10~13, ��ҽ� �ʿ�)	
Const C_CNCL_ORG_CHANGE_ID	= 11
Const C_CNCL_GL_DEPT_CD		= 12		
Const C_CNCL_INTERNAL_CD	= 13		

'------ Developer Coding part (End   ) ------------------------------------------------------------------     
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '��: Query
'         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '��: Save,Update    
         Call SubBizSaveMuliti()
End Select

'==================================================================================
'	Name : SubBizSaveMuliti()
'	Description : ��Ƽ���� ���� 
'==================================================================================
Sub SubBizSaveMuliti()

	On Error Resume Next
	Err.Clear							'��: Protect system from crashing

	Call HideStatusWnd
	Dim inDx
	Dim PAFG525CD

	Dim arrRowVal,arrVal		'��: Spread Sheet �� ���� ���� Array ���� 
				
	Const C_MOVE_CHAR	   = 0
	Const C_MOVE_SEQ       = 1
	Const C_MOVE_NOTE_NO   = 2
	Const C_MOVE_DEPT_CD   = 3
	Const C_MOVE_ITEM_DESC = 4

	Const C_MOVE_CHAR2	   = 0
	Const C_MOVE_SEQ2      = 1
	Const C_NOTE_NO2	   = 2
	Const C_TEMP_GL_NO2    = 3
	Const C_GL_NO2		   = 4

	Dim IG1_note_grp		
	Const C_NOTE_NO_GRP	   = 0
	Const C_TO_DEPT_CD_GRP = 1
	Const C_ITEM_MOVE_DESC = 2

	Dim IG1_Cnc_note_grp
	Const C_CNC_NOTE_NO_GRP = 0
	Const C_CNC_TEMP_GL_NO  = 1
	Const C_CNC_GL_NO		= 2

	Dim I1_ief_supplied
	Dim I2_fr_biz_cd
	Dim I3_mvnt_dt
	Dim I4_Org_Change_ID
	Dim I5_To_head_dept_cd
	Dim I6_To_head_Biz_cd
	Dim I7_note_move_desc

	Dim I4_ORG_CHG_ID
	Const C_CHG_ORG_ID = 0
	Const C_DEPT_CD = 1

	arrRowVal = Split(Request("txtSpread"), gRowSep)

    Set PAFG525CD = server.CreateObject ("PAFG525.cFMvntNoteSvr")    
    
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

    Dim I8_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 
    Const A667_I8_a_data_auth_data_BizAreaCd = 0
    Const A667_I8_a_data_auth_data_internal_cd = 1
    Const A667_I8_a_data_auth_data_sub_internal_cd = 2
    Const A667_I8_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I8_a_data_auth(3)
	I8_a_data_auth(A667_I8_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I8_a_data_auth(A667_I8_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I8_a_data_auth(A667_I8_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I8_a_data_auth(A667_I8_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))


	If Request("hProcFg") = "CG" Then
		Redim I4_ORG_CHG_ID(C_DEPT_CD) 
		
		I1_ief_supplied    = "CREATE"
		I2_fr_biz_cd       = Request("txtFrBizCd")
		I3_mvnt_dt         = UNIConvDate(Request("txtGLDt"))
		'I4_Org_Change_ID   = Request("hOrgChangeId")	
		
		I4_ORG_CHG_ID(C_CHG_ORG_ID) = Request("hOrgChangeId")
		I4_ORG_CHG_ID(C_DEPT_CD) = Request("txtToDeptCd")

		I5_To_head_dept_cd = Trim(Request("txtToDeptCd"))
		I6_To_head_Biz_cd  = Trim(Request("txtToBizCd"))
		
		I7_note_move_desc  = Trim(Request("txtNoteDesc"))
				
		Redim IG1_note_grp(UBound(arrRowVal)-1 ,	2)
		
	    For indx = 0 To UBound(arrRowVal)-1 
	        arrVal = Split(arrRowVal(indx), gColSep)
	        IG1_note_grp(indx, C_NOTE_NO_GRP)    = arrVal(C_MOVE_NOTE_NO)
			IG1_note_grp(indx, C_TO_DEPT_CD_GRP) = arrVal(C_MOVE_DEPT_CD)
			IG1_note_grp(indx, C_ITEM_MOVE_DESC) = arrVal(C_MOVE_ITEM_DESC)
		Next

		Call PAFG525CD.F_MOVEMENT_NOTE_SVR(gStrGlobalCollection, _
												I1_ief_supplied, _
												I2_fr_biz_cd, _
												I3_mvnt_dt, _												
												I4_ORG_CHG_ID, _ 												
												I5_To_head_dept_cd, _
												I6_To_head_Biz_cd, _
												I7_note_move_desc, _	
												IG1_note_grp,_
												IG1_Cnc_note_grp, _
												I8_a_data_auth)		
	Else
		I1_ief_supplied = "DELETE"	

		Redim IG1_Cnc_note_grp(UBound(arrRowVal)-1 ,	7)

	    For indx = 0 To UBound(arrRowVal) - 1
	        arrVal = Split(arrRowVal(indx), gColSep)	        
	        IG1_Cnc_note_grp(indx, C_CNC_NOTE_NO_GRP)	 = arrVal(C_NOTE_NO2)
			IG1_Cnc_note_grp(indx, C_CNC_TEMP_GL_NO)	 = arrVal(C_TEMP_GL_NO2)
	        IG1_Cnc_note_grp(indx, C_CNC_GL_NO)			 = arrVal(C_GL_NO2)
'			Response.Write 	UBound(arrRowVal)
'			Response.end	        
	    Next	
	    
	    Call PAFG525CD.F_MOVEMENT_NOTE_SVR(gStrGlobalCollection, _
											I1_ief_supplied, _
											, _
											, _
											, _
											, _
											, _								
											, _
											, _
											IG1_Cnc_note_grp, _
											I8_a_data_auth)

	End If
	
    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG525CD = Nothing		
		Exit Sub
    End If
    
    Set PAFG525CD = Nothing
    
	Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbSaveOk()							" & vbCr
    Response.Write "</Script>									" & vbCr    
End Sub
%>