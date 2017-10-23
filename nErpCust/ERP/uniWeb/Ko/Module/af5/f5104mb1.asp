<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_RECEIPT
'*  3. Program ID        : f5104ma
'*  4. Program 이름      : 만기어음일괄처리 
'*  5. Program 설명      : 만기어음일괄처리 
'*  6. Comproxy 리스트   : f5104ma
'*  7. 최초 작성년월일   : 2000/10/16
'*  8. 최종 수정년월일   : 2002/02/15
'*  9. 최초 작성자       : 김종환 
'* 10. 최종 작성자       : 오수민 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'*                         -2000/10/16 : ..........
'**********************************************************************************************




'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->


<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next                                                            '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
Dim lgOpModeCRUD
Dim lPtxtNoteNo


Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------

lgOpModeCRUD = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
 
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
Const C_CNCL_RCPT_TYPE		= 10		'☜: hidden field(10~13, 취소시 필요)	
Const C_CNCL_ORG_CHANGE_ID	= 11
Const C_CNCL_GL_DEPT_CD		= 12		
Const C_CNCL_INTERNAL_CD	= 13		

'------ Developer Coding part (End   ) ------------------------------------------------------------------     
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
'         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update    
         Call SubBizSaveMuliti()
End Select


'==================================================================================
'	Name : SubBizQueryMulti()
'	Description : 멀티조회 정의 
'==================================================================================
Sub SubBizQueryMulti()

On Error Resume Next                                                            '☜: Protect system from crashing
Err.Clear  

	Dim PAFG520LIST	
	Dim iLngRow,iLngCol
	Dim iIntLoopCount, iIntLoopCount1
	Dim iStrData
	Dim StrNextKey
	Dim StrNextKeyNoteNo			' NoteNO 다음 값 
	Dim StrNextKeyGlNo				' GLNO 다음 값 
	Dim lgStrPrevKey
	Dim lgStrPrevKeyNoteNo			' Note NO 이전 값 
	Dim lgStrPrevKeyGlNo
	Dim lgStrPrevKeyTempGlNo
	Dim iIntMaxRows
	Dim iIntQueryCount
	
	
	Dim I1_ief_supplied
'	Const C_COMMAND = 0 
	
	Dim I2_f_note 
	
	Const C_NOTE_FG_IMP = 0
	Const C_DUE_DT_IMP = 1
	Const C_NOTE_STS_IMP = 2	
	
	Dim I3_f_note_item
	Dim I4_f_note_item
	
	Dim strBankCd 
	
	Dim E1_b_bank
	Dim EG1_export_group
	
	Dim E2_f_note
	Const C_F_NOTE_NO_EXP = 0
	
	Dim E3_a_gl
	Const C_F_NOTE_ITEM_GL_NO = 0
	
	Dim E4_a_temp_gl
	Const C_F_NOTE_ITEM_TEMP_GL_NO = 0	
	
	'MAXROWS
	Const C_SHEETMAXROWS = 100
	
	
	'일괄처리 & 일괄취소 구분 
	I1_ief_supplied = UCase(Trim(Request("cboProcFg")))
	
	'CONDITION
	Redim I2_f_note(2) 
	I2_f_note(C_NOTE_FG_IMP) = UCase(Trim(Request("cboNoteFg")))
	I2_f_note(C_DUE_DT_IMP) = UNIConvDate(Request("txtDueDtEnd"))
	I2_f_note(C_NOTE_STS_IMP) = UCase(Trim(Request("cboNoteSts")))

	'TAB2 회계 시작일 
	I3_f_note_item = UNIConvDate(Request("txtStsDtStart"))

	
	'TAB2 회계 종료일 
	I4_f_note_item = UNIConvDate(Request("txtStsDtEnd"))

	'CONDITION
	strBankCd = UCase(Trim(Request("txtBankCd")))
	
	'TAB1 MAXKEY
	lgStrPrevKeyNoteNo = Request("lgStrPrevKeyNoteNo")
	
	'TAB2 MAXKEY
	lgStrPrevKeyGlNo = Request("lgStrPrevKeyGlNo")
	lgStrPrevKeyTempGlNo = Request("lgStrPrevKeyTempGlNo")
	
	
	iIntQueryCount	= Request("lgPageNo")

	Set PAFG520LIST = server.CreateObject ("PAFG520.cFListNoteForBtchSvr")
	
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If

	If Trim(I1_ief_supplied) = "CG" Then
		Call PAFG520LIST.FN0048_LIST_NOTE_FOR_BATCH_SVR(gStrGlobalCollection, _
														C_SHEETMAXROWS, _
														I1_ief_supplied, _
														I2_f_note, _
														, _
														, _
														strBankCd, _
														lgStrPrevKeyNoteNo, _
														lgStrPrevKeyGlNo, _
														lgStrPrevKeyTempGlNo, _
														E1_b_bank, _
														EG1_export_group, _
														E2_f_note, _
														E3_a_gl, _
														E4_a_temp_gl)
	Else
		Call PAFG520LIST.FN0048_LIST_NOTE_FOR_BATCH_SVR(gStrGlobalCollection, _
														C_SHEETMAXROWS, _
														I1_ief_supplied, _
														I2_f_note, _
														I3_f_note_item, _
														I4_f_note_item, _
														strBankCd, _
														lgStrPrevKeyNoteNo, _
														lgStrPrevKeyGlNo, _
														lgStrPrevKeyTempGlNo, _
														E1_b_bank, _
														EG1_export_group, _
														E2_f_note, _
														E3_a_gl, _
														E4_a_temp_gl)
	End If
	
    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG520LIST = nothing		
		Exit Sub
    End If
    
    Set PAFG520LIST = nothing
	
	iStrData = ""
	

'	ReDim E2_f_note(0)
	lgStrPrevKeyNoteNo = E2_f_note
	
'	Redim E3_a_gl(0)
	lgStrPrevKeyGlNo = E3_a_gl
	lgStrPrevKeyTempGlNo = E4_a_temp_gl
		
	If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)          
       End If   
    Else   
       iIntQueryCount = 0
    End If
	
	
	iIntLoopCount = 0
	if isarray(EG1_export_group) Then
		If Trim(I1_ief_supplied) = "CG" Then
			For iLngRow = 0 To UBound(EG1_export_group, 1) 	
				iIntLoopCount = iIntLoopCount + 1
				If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_NOTE_NO)))
						iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, C_NOTE_AMT),	ggExchRate.DecPoint		,0)
						iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, C_DUE_DT))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_NOTE_STS)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BANK_CD)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BANK_NM)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BP_CD)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BP_NM)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_DEPT_CD)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_DEPT_NM)))
'						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_GL_NO)))
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1) 
						iStrData = iStrData & Chr(11) & Chr(12)
				Else
					lgStrPrevKeyNoteNo = EG1_export_group(UBound(EG1_export_group, 1), C_NOTE_NO)
					iIntQueryCount = iIntQueryCount + 1
					Exit For
				End If
			Next

			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				lgStrPrevKeyNoteNo = ""
				lgStrPrevKeyGlNo = ""
				lgStrPrevKeyTempGlNo = ""
			    iIntQueryCount = ""
			End If
			
			Response.Write " <Script Language=vbscript>										 " & vbCr
			Response.Write " With parent													 " & vbCr
			Response.Write "	.ggoSpread.Source		= .frm1.vspdData					 " & vbCr
			Response.Write "	.ggoSpread.SSShowData	""" & iStrData					& """" & vbCr
			Response.Write "	.frm1.hProcFg.Value	  =	""" & Trim(I1_ief_supplied)		& """" & vbCr
			Response.Write "	.frm1.hNoteFg1.Value  =	""" & I2_f_note(C_NOTE_FG_IMP)	& """" & vbCr
			Response.Write "	.frm1.hNoteSts.Value  =	""" & I2_f_note(C_NOTE_STS_IMP)	& """" & vbCr
			Response.Write "	.frm1.hDueDtEnd.Value =	""" & I2_f_note(C_DUE_DT_IMP)	& """" & vbCr
			Response.Write "	.frm1.hBankCd.Value	  =	""" & strBankCd					& """" & vbCr
			Response.Write "	.lgPageNo			  = """ & iIntQueryCount			& """" & vbCr
			Response.Write "	.lgStrPrevKeyNoteNo	  = """ & lgStrPrevKeyNoteNo		& """" & vbCr
			Response.Write "	.lgStrPrevKeyGlNo	  = """ & lgStrPrevKeyGlNo			& """" & vbCr
			Response.Write "	.lgStrPrevKeyTempGlNo	  = """ & lgStrPrevKeyTempGlNo			& """" & vbCr
			Response.Write "	.DbQueryOk													 " & vbCr
			Response.Write "End With														 " & vbCr
			Response.Write "</Script>														 " & vbCr
		Else
			For iLngRow = 0 To UBound(EG1_export_group, 1) 	
				iIntLoopCount = iIntLoopCount + 1
				If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_NOTE_NO)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_TEMP_GL_NO)))												
						iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, C_CNCL_TEMP_GL_DT))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_GL_NO)))
						iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, C_CNCL_GL_DT))						
						iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, C_CNCL_NOTE_AMT),	ggExchRate.DecPoint		,0)
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_BP_CD)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_BP_NM)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_DEPT_CD)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_DEPT_NM)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_RCPT_TYPE)))
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_ORG_CHANGE_ID)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_GL_DEPT_CD)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CNCL_INTERNAL_CD)))
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1) 
						iStrData = iStrData & Chr(11) & Chr(12)
				Else
					lgStrPrevKeyNoteNo = EG1_export_group(UBound(EG1_export_group, 1), C_CNCL_NOTE_NO)
					lgStrPrevKeyGlNo = EG1_export_group(UBound(EG1_export_group, 1), C_CNCL_GL_NO)
					lgStrPrevKeyTempGlNo = EG1_export_group(UBound(EG1_export_group, 1), C_CNCL_TEMP_GL_NO)
					iIntQueryCount = iIntQueryCount + 1
					Exit For
				End If
			Next
			
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				lgStrPrevKeyNoteNo = ""
				lgStrPrevKeyGlNo = ""
				lgStrPrevKeyTempGlNo = ""
			    iIntQueryCount = ""
			End If

			Response.Write " <Script Language=vbscript>										 " & vbCr
			Response.Write " With parent													 " & vbCr
			Response.Write "	.ggoSpread.Source		= .frm1.vspdData2					 " & vbCr
			Response.Write "	.ggoSpread.SSShowData	   """ & iStrData					& """" & vbCr
			Response.Write "	.frm1.hProcFg.Value		=  """ & Trim(I1_ief_supplied)		& """" & vbCr
			Response.Write "	.frm1.hNoteFg2.Value	=  """ & I2_f_note(C_NOTE_FG_IMP)	& """" & vbCr
			Response.Write "	.frm1.hStsDtStart.Value =  """ & I3_f_note_item(C_START_DT)	& """" & vbCr
			Response.Write "	.frm1.hStsDtEnd.Value	=  """ & I4_f_note_item(C_END_DT)	& """" & vbCr
			Response.Write "	.frm1.hBankCd.Value		=  """ & strBankCd					& """" & vbCr
			Response.Write "	.lgPageNo1				=  """ & iIntQueryCount			& """" & vbCr
			Response.Write "	.lgStrPrevKeyNoteNo1	=  """ & lgStrPrevKeyNoteNo		& """" & vbCr
			Response.Write "	.lgStrPrevKeyGlNo1		=  """ & lgStrPrevKeyGlNo			& """" & vbCr
			Response.Write "	.lgStrPrevKeyTempGlNo1=  """ & lgStrPrevKeyTempGlNo	& """" & vbCr
			Response.Write "	.DbQueryOk													 " & vbCr
			Response.Write "End With														 " & vbCr
			Response.Write "</Script>														 " & vbCr
		End If
	Else
		Call DisplayMsgBox("141200", vbOKOnly, "", "", I_MKSCRIPT)
		Exit Sub
	End If
End Sub


'==================================================================================
'	Name : SubBizSaveMuliti()
'	Description : 멀티저장 정의 
'==================================================================================
Sub SubBizSaveMuliti()

	On Error Resume Next
	Err.Clear							'☜: Protect system from crashing

	Call HideStatusWnd
	Dim inDx

	Dim PAFG520CD
	DIm NoteAcctCd

	Dim I1_ief_supplied

	Dim I2_b_acct_dept
	Const C_CHG_ORG_ID = 0
	Const C_DEPT_CD = 1

	Dim I3_f_note_item
	Const C_RCPT_TYPE = 0
	Const C_NOTE_ACCT_CD = 1

	Dim I4_b_bank
	Dim I5_b_bank_acct
	Dim I6_a_gl
	Const C_GL_DT = 0
	Const C_GL_INSRT_ID = 1

	Dim arrRowVal,arrVal		'☜: Spread Sheet 의 값을 받을 Array 변수 
				
	Const C_SELECT_CHR1    = 0
	Const C_ROW_NUM1	   = 1
	Const C_NOTE_NO1	   = 2
	Const C_TEMP_GL_NO1	   = 3
	Const C_GL_NO1		   = 4
	Const C_NOTE_ITEM_DESC1= 5

	Dim IG1_import_group		
	Const C_NOTE_NO_GRP		  = 0
	Const C_TEMP_GL_NO_GRP	  = 1
	Const C_GL_NO_GRP		  = 2
	Const C_RCPT_TYPE_GRP     = 3
	Const C_ORG_CHANGE_ID_GRP = 4
	Const C_DEPT_CD_GRP		  = 5
	Const C_INTERNAL_CD_GRP	  = 6
	Const C_NOTE_ITEM_DESC_GRP= 7

	Redim I2_b_acct_dept(C_DEPT_CD) 	
	Redim I3_f_note_item(C_NOTE_ACCT_CD)
	Redim I6_a_gl(C_GL_INSRT_ID)
	
	I2_b_acct_dept(C_CHG_ORG_ID) = Request("hOrgChangeId")
	I2_b_acct_dept(C_DEPT_CD) = UCase(Trim(Request("txtDeptCd")))
	
	I3_f_note_item(C_RCPT_TYPE) = UCase(Trim(Request("txtRcptType")))		
	I3_f_note_item(C_NOTE_ACCT_CD) = UCase(Trim(Request("txtNoteAcctCd")))

	I4_b_bank = UCase(Trim(Request("txtBankCd1")))
	I5_b_bank_acct = UCase(Trim(Request("txtBankAcctNo")))
	
	I6_a_gl(C_GL_DT) = UNIConvDate(Request("txtGLDt"))
	I6_a_gl(C_GL_INSRT_ID) = ""
	
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	
    Set PAFG520CD = server.CreateObject ("PAFG520.cFBtchNoteSvr")    

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

    Dim I7_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
    Const A826_I7_a_data_auth_data_BizAreaCd = 0
    Const A826_I7_a_data_auth_data_internal_cd = 1
    Const A826_I7_a_data_auth_data_sub_internal_cd = 2
    Const A826_I7_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I7_a_data_auth(3)
	I7_a_data_auth(A826_I7_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I7_a_data_auth(A826_I7_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I7_a_data_auth(A826_I7_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I7_a_data_auth(A826_I7_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
	If Request("hProcFg") = "CG" Then
		I1_ief_supplied = "CREATE"
				
		Redim IG1_import_group(UBound(arrRowVal) - 1,	7)
		
	    For indx = 0 To UBound(arrRowVal) - 1
	        arrVal = Split(arrRowVal(indx), gColSep)
	        IG1_import_group(indx, C_NOTE_NO_GRP) = arrVal(C_NOTE_NO1)
			IG1_import_group(indx, C_TEMP_GL_NO_GRP) = ""
	        IG1_import_group(indx, C_GL_NO_GRP) = ""
	        IG1_import_group(indx, C_RCPT_TYPE_GRP) = ""
	        IG1_import_group(indx, C_ORG_CHANGE_ID_GRP) = ""
	        IG1_import_group(indx, C_DEPT_CD_GRP) = ""
			IG1_import_group(indx, C_INTERNAL_CD_GRP) = ""	        
	        IG1_import_group(indx, C_NOTE_ITEM_DESC_GRP) = arrVal(C_NOTE_ITEM_DESC1)    	        
	    Next

		Call PAFG520CD.FN0041_BATCH_NOTE_SVR(gStrGlobalCollection, _
												I1_ief_supplied, _
												I2_b_acct_dept, _
												I3_f_note_item, _
												I4_b_bank, _
												I5_b_bank_acct, _
												I6_a_gl, _																
												IG1_import_group, _
												I7_a_data_auth)		

	Else
		I1_ief_supplied = "DELETE"		
		
		Redim IG1_import_group(UBound(arrRowVal) - 1,	7)

	    For indx = 0 To UBound(arrRowVal) - 1 
	        arrVal = Split(arrRowVal(indx), gColSep)	        
	        IG1_import_group(indx, C_NOTE_NO_GRP)		= arrVal(C_NOTE_NO1)
			IG1_import_group(indx, C_TEMP_GL_NO_GRP)	= arrVal(C_TEMP_GL_NO1)
	        IG1_import_group(indx, C_GL_NO_GRP)			= arrVal(C_GL_NO1)
	    Next	
	    
	    Call PAFG520CD.FN0041_BATCH_NOTE_SVR(gStrGlobalCollection, _
											I1_ief_supplied, _
											, _
											, _
											, _
											, _
											, _								
											IG1_import_group, _
											I7_a_data_auth)

	End If
	
    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG520CD = nothing		
		Exit Sub
    End If
    
    Set PAFG520CD = nothing
    
	Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbSaveOk()							" & vbCr
    Response.Write "</Script>									" & vbCr    
End Sub
%>

................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................
................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................
................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................
................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................
................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................
............................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................................. 
