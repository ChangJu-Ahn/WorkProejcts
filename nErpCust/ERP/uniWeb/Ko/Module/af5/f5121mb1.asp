<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_RECEIPT
'*  3. Program ID        : f5121mb
'*  4. Program 이름      : 부도어음처리 
'*  5. Program 설명      : 부도어음처리 등록 수정 삭제 조회 
'*  6. Comproxy 리스트   : f5121mb
'*  7. 최초 작성년월일   : 2003/09/17
'*  8. 최종 수정년월일   : 
'*  9. 최초 작성자       : Soo Min, Oh
'* 10. 최종 작성자       : 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->


<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next                                                            '☜: Protect system from crashing
'Err.Clear                                                                        '☜: Clear Error status

Dim txtNoteNoQry
Dim lgOpModeCRUD
Dim lPtxtNoteNo
    
Call LoadBasisGlobalInf()    
Call LoadInfTB19029B("I","*","NOCOOKIE","MB") 
Call HideStatusWnd    

'---------------------------------------Common-----------------------------------------------------------

lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
txtNoteNoQry      = Request("txtNoteNoQry")

Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim GroupCount

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'------ Developer Coding part (End   ) ------------------------------------------------------------------     
Select Case strMode
    Case CStr(UID_M0001)                                                         '☜: Query
         Call SubBizQuery()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSave()
    Case CStr(UID_M0003)                                                         '☜: Delete    		
         Call SubBizDelete()
End Select

'==================================================================================
'	Name : SubBizQuery()
'	Description : 조회 정의 
'==================================================================================
Sub SubBizQuery()

	Dim PAFG560LIST	
	Dim indx
	
	Dim I1_f_note_no
	Dim E1_f_note
	Dim E2_f_note_gl
	Dim E3_f_note_item_grp	
	
	Dim iLngRow,iLngCol
	Dim iIntLoopCount
	Dim iStrData
	Dim iStrPrevKey
	Dim iIntMaxRows
	Dim iIntQueryCount
	
	'Const E1_f_note_sts_dt = 0
    Const C_NOTE_NO = 0
    Const C_NOTE_FG = 1
    Const C_NOTE_STS = 2
    Const C_ISSUE_DT = 3
    Const C_DUE_DT = 4
    Const C_BP_CD = 5
    Const C_BP_NM = 6
    Const C_BANK_CD = 7
    Const C_BANK_NM = 8
    Const C_NOTE_DH_AMT = 9
    Const C_NOTE_STTL_AMT = 10
            
    Const C_STS_DT = 0 
	Const C_DEPT_CD = 1
	Const C_DEPT_NM = 2
	Const C_ORG_CHANGE_ID = 3
	Const C_INT_REV_AMT = 4
	Const C_INT_ACCT_CD = 5
	Const C_INT_ACCT_NM = 6
	Const C_TEMP_GL_NO = 7 
	Const C_GL_NO = 8
	Const C_NOTE_ITEM_DESC = 9
    
    
    Const C_GP_NOTE_NO = 0
    Const C_GP_STTL_TYPE = 1
    Const C_GP_STTL_NM = 2
    Const C_GP_RCPT_TYPE = 3
    Const C_GP_RCPT_NM = 4
    Const C_GP_REF_NO = 5
    Const C_GP_NOTE_ACCT_CD = 6
    Const C_GP_ACCT_NM = 7
    Const C_GP_BANK_ACCT_NO = 8
    Const C_GP_BANK_CD = 9
    Const C_GP_BANK_NM = 10
    Const C_GP_NOTE_AMT = 11
    Const C_GP_NOTe_ITEM_DESC = 12    
	
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
    
  '********************************************************  
  '                        Query
  '********************************************************     
    Set PAFG560LIST = server.CreateObject ("PAFG560.cFListDhNoteItemSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If

	
	'-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_f_note_no = Trim(Request("txtNoteNoQry"))    

    Call PAFG560LIST.FN0028_LIST_DH_NOTE_ITEM_SVR(gStrGlobalCollection, _												  
												  I1_f_note_no, _												  
												  E1_f_note,_
												  E2_f_note_gl, _
												  E3_f_note_item_grp)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG560LIST = nothing		
		Exit Sub
    End If
    
    Set PAFG560LIST = nothing 

	Response.Write "<Script Language=vbscript>  " & vbCr
   	Response.Write " with parent.frm1" & vbCr   	   	
   	Response.Write " .hNoteNo.value		= """ & ConvSPChars(E1_f_note(C_NOTE_NO)) & """											" & vbCr '☆: Note No(hidden)   	
   	Response.Write " .txtNoteNoQry.value= """ & ConvSPChars(E1_f_note(C_NOTE_NO)) & """											" & vbCr '☆: Note No(Query)   	   	   	
   	
	If Trim(E1_f_note(C_NOTE_STS)) = "DS" Then		
   		Response.Write " .txtStsDt.Text		= """ & UNIDateClientFormat(E2_f_note_gl(C_STS_DT)) & """							" & vbCr '☆: Status Date(DS)   		
   		Response.Write " .txtDeptCD.value	= """ & ConvSPChars(E2_f_note_gl(C_DEPT_CD)) & """									" & vbCr '☆: Dept Code
   		Response.Write " .txtDeptNm.value	= """ & ConvSPChars(E2_f_note_gl(C_DEPT_NM)) & """									" & vbCr '☆: Dept Name   		
   		Response.Write " .horgchangeid.value= """ & ConvSPChars(E2_f_note_gl(C_ORG_CHANGE_ID)) & """							" & vbCr '☆: Org Change Id(hidden)   	
   	End If 
   	
   	Response.Write " .cboNoteFg.value	= """ & E1_f_note(C_NOTE_FG) & """														" & vbCr '☆: Note Flag
   	Response.Write " .cboNoteSts.value	= """ & E1_f_note(C_NOTE_STS) & """														" & vbCr '☆: Note Status
   	Response.Write " .txtIssueDt.Text	= """ & UNIDateClientFormat(E1_f_note(C_ISSUE_DT)) & """									" & vbCr '☆: Issue Date
   	Response.Write " .txtDueDt.Text		= """ & UNIDateClientFormat(E1_f_note(C_DUE_DT)) & """									" & vbCr '☆: Due Date
   	Response.Write " .txtBpCd.value		= """ & ConvSPChars(E1_f_note(C_BP_CD)) & """											" & vbCr '☆: Biz Partner Code
   	Response.Write " .txtBpNM.value		= """ & ConvSPChars(E1_f_note(C_BP_NM)) & """											" & vbCr '☆: Biz Partner Name
   	Response.Write " .txtBankCd.value	= """ & ConvSPChars(E1_f_note(C_BANK_CD)) & """											" & vbCr '☆: Bank Code
   	Response.Write " .txtBankNm.Value	= """ & ConvSPChars(E1_f_note(C_BANK_NM)) & """											" & vbCr '☆: Bank Name
   	Response.Write " .txtNoteAmt.Text	= """ & UNINumClientFormat(E1_f_note(C_NOTE_DH_AMT),	ggAmtOfMoney.DecPoint	,0) & """	" & vbCr '☆: Note Amount
   	Response.Write " .txtSttlAmt.Text	= """ & UNINumClientFormat(E1_f_note(C_NOTE_STTL_AMT ),	ggAmtOfMoney.DecPoint	,0) & """	" & vbCr '☆: Settlement Amount
   	
   	If Trim(E1_f_note(C_NOTE_STS)) = "DS" Then		
   		Response.Write " .txtIntRevAmt.Text	= """ & UNINumClientFormat(E2_f_note_gl(C_INT_REV_AMT),	ggAmtOfMoney.DecPoint	,0) & """	" & vbCr '☆: Interest Revenue Amount  
   		Response.Write " .txtIntAcctCd.value= """ & ConvSPChars(E2_f_note_gl(C_INT_ACCT_CD)) & """										" & vbCr '☆: Interest Revenue Account Code
   		Response.Write " .txtIntAcctNm.value= """ & ConvSPChars(E2_f_note_gl(C_INT_ACCT_NM)) & """										" & vbCr '☆: Interest Revenue Account Name
   		Response.Write " .txtNoteDesc.Value	= """ & ConvSPChars(E2_f_note_gl(C_NOTE_ITEM_DESC)) & """									" & vbCr '☆: Note Description 
   	End If 
   	
   	Response.Write " .hTempGlNo.Value	= """ & ConvSPChars(E2_f_note_gl(C_TEMP_GL_NO)) & """									" & vbCr '☆: Temp GL No.
   	Response.Write " .hGlNo.Value		= """ & ConvSPChars(E2_f_note_gl(C_GL_NO)) & """										" & vbCr '☆: Gl No.   	
   	
   	Response.Write "End with				" & vbcr
    Response.Write "Parent.DbQueryOk		" & vbcr
    Response.Write "</Script>               " & vbCr
    
    
    '======================================== Single End ==============================================================
	iStrData = ""

	If IsEmpty(E3_f_note_item_grp) = False Then
		For iLngRow = 0 To UBound(E3_f_note_item_grp, 1) 	
			iIntLoopCount = iIntLoopCount + 1
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_STTL_TYPE ))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_RCPT_TYPE ))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_RCPT_NM ))
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_REF_NO ))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_NOTE_ACCT_CD ))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_ACCT_NM  ))
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_BANK_ACCT_NO  ))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_BANK_CD ))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & Trim(E3_f_note_item_grp(iLngRow, C_GP_BANK_NM ))					
				iStrData = iStrData & Chr(11) & E3_f_note_item_grp(iLngRow, C_GP_NOTE_AMT )					
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(E3_f_note_item_grp(iLngRow, C_GP_NOTE_ITEM_DESC )))					
				iStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1) 
				iStrData = iStrData & Chr(11) & Chr(12)
			Else
				iStrPrevKey = E3_f_note_item_grp(UBound(E3_f_note_item_grp, 1), C_SEQ)
				iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next
	End If

	
	Response.Write " <Script Language=vbscript>								 " & vbCr
	Response.Write " With parent											 " & vbCr
    Response.Write "	.ggoSpread.Source		= .frm1.vspdData			 " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData	  """ & iStrData		& """" & vbCr
    Response.Write "	.lgPageNo				= """ & iIntQueryCount	& """" & vbCr
    Response.Write "	.lgStrPrevKey			= """ & iStrPrevKey		& """" & vbCr
    Response.Write "	.DbQueryOk											 " & vbCr
    Response.Write "End With												 " & vbCr
    Response.Write "</Script>												 " & vbCr 
	
End Sub

'==================================================================================
'	Name : SubBizSave()
'	Description : 수정, 신규 
'==================================================================================
Sub SubBizSave()
	Dim PAFG560CU	
	Dim lgIntFlgMode
	Dim arrRowVal, arrVal
	Dim indx
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        

	'(공통) SINGLE DATA 
	Dim I1_f_note_gl
	Const C_CUD_NOTE_NO			= 0		' 어음번호 
	Const C_CUD_STS_DT			= 1		' 처리일자 
	Const C_CUD_DEPT_CD			= 2		' 부서 
	Const C_CUD_ORG_CHANGE_ID	= 3		' 조직개편ID	
	Const C_CUD_INT_REV_AMT		= 4		' 이자수익금액 
	Const C_CUD_INT_REV_ACCT_CD	= 5		' 이자수익계정코드 
	Const C_CUD_NOTE_DESC		= 6		' 비고 
	Const C_CUD_GL_NO			= 7		' 회계전표번호 
	Const C_CUD_TEMP_GL_NO		= 8		' 결의전표번호	

	'MULTI DATA
	Dim IG1_import_group
	Const C_CUD_CUD_FG			= 0		' CUD Flag
	Const C_CUD_SEQ				= 1		' Sequence	
	Const C_CUD_STTL_TYPE		= 2		' 처리유형 
	Const C_CUD_RCPT_TYPE		= 3		' 입금유형 
	Const C_CUD_RCPT_ACCT_CD	= 4		' 입금계정코드 
	Const C_CUD_REF_NOTE_NO		= 5		' 받을어음번호	
	Const C_CUD_BANK_ACCT_NO	= 6		' 계좌번호 
	Const C_CUD_BANK_CD			= 7		' 은행코드 
	Const C_CUD_ITEM_AMT		= 8		' 처리금액 
	Const C_CUD_NOTE_ITEM_DESC	= 9		' 비고 

	Redim I1_f_note_gl(C_CUD_TEMP_GL_NO) 
	I1_f_note_gl(C_CUD_NOTE_NO)			= Trim(Request("txtNoteNoQry"))
	I1_f_note_gl(C_CUD_STS_DT)			= Trim(Request("txtStsDt"))
	I1_f_note_gl(C_CUD_DEPT_CD)			= Trim(Request("txtDeptCD"))
	I1_f_note_gl(C_CUD_ORG_CHANGE_ID)	= Trim(Request("horgchangeid"))	
	I1_f_note_gl(C_CUD_INT_REV_AMT)		= CDbl((UniConvNum(Request("txtIntRevAmt"),1)))	
	I1_f_note_gl(C_CUD_INT_REV_ACCT_CD)	= Trim(Request("txtIntAcctCd"))
	I1_f_note_gl(C_CUD_NOTE_DESC)		= Trim(Request("txtNoteDesc"))
	I1_f_note_gl(C_CUD_GL_NO)			= Trim(Request("hGlNo"))
	I1_f_note_gl(C_CUD_TEMP_GL_NO)		= Trim(Request("hTempGlNo"))

	arrRowVal = Split(Request("txtSpread"), gRowSep)	

	Redim IG1_import_group(UBound(arrRowVal) - 1,	9)	
	    For indx = 0 To UBound(arrRowVal) - 1
	       
	        arrVal = Split(arrRowVal(indx), gColSep)
	       
	        IG1_import_group(indx, C_CUD_CUD_FG) = arrVal(0)
			IG1_import_group(indx, C_CUD_SEQ) = arrVal(1)
	        IG1_import_group(indx, C_CUD_STTL_TYPE) = arrVal(3)
	        IG1_import_group(indx, C_CUD_RCPT_TYPE) = arrVal(4)
	        IG1_import_group(indx, C_CUD_RCPT_ACCT_CD) = arrVal(5)
	        IG1_import_group(indx, C_CUD_REF_NOTE_NO) = arrVal(6)
	        IG1_import_group(indx, C_CUD_BANK_ACCT_NO) = arrVal(7)
	        IG1_import_group(indx, C_CUD_BANK_CD) = arrVal(8)
	        IG1_import_group(indx, C_CUD_ITEM_AMT) = arrVal(9)	        
	        IG1_import_group(indx, C_CUD_NOTE_ITEM_DESC) = arrVal(10)
	    Next 	
		
    Set PAFG560CU = server.CreateObject ("PAFG560.cFMngDhNoteChgSvr")   
    
    If CheckSYSTEMError(Err, True) = True Then       
       Exit Sub
    End If         
     
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)                    	         	
    
    Select Case lgIntFlgMode
		Case  OPMD_CMODE                                                             '☜ : Create								
	
			Call PAFG560CU.FN0022_MANAGE_DH_NOTE_CHG_SVR(gStrGlobalCollection, _
														"CREATE", _
														I1_f_note_gl, _
														IG1_import_group)
        Case  OPMD_UMODE                  
			Call PAFG560CU.FN0022_MANAGE_DH_NOTE_CHG_SVR(gStrGlobalCollection, _
														"UPDATE", _
														I1_f_note_gl, _
														IG1_import_group)
    End Select

    If CheckSYSTEMError(Err,True) = True Then

		Set PAFG560CU = nothing
		Exit Sub	
    End If
	 
    Set PAFG560CU = nothing
    
    lPtxtNoteNo = Request("txtNoteNoQry")

	Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbSaveOk(""" & lPtxtNoteNo	& """)	" & vbCr
    Response.Write "</Script>									" & vbCr  
    
      
End Sub

'==================================================================================
'	Name : SubBizDelete()
'	Description : 삭제 
'==================================================================================
Sub SubBizDelete()
	Dim PAFG560D	
	
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
	
	'SINGLE DATA 
	Dim I1_f_note_gl
	Const C_DEL_NOTE_NO			= 0		' 어음번호 
	Const C_DEL_STS_DT			= 1		' 처리일자 
	Const C_DEL_DEPT_CD			= 2		' 부서 
	Const C_DEL_ORG_CHANGE_ID	= 3		' 조직개편ID	
	Const C_DEL_INT_REV_AMT		= 4		' 이자수익금액 
	Const C_DEL_INT_REV_ACCT_CD	= 5		' 이자수익계정코드 
	Const C_DEL_NOTE_DESC		= 6		' 비고 
	Const C_DEL_GL_NO			= 7		' 회계전표번호 
	Const C_DEL_TEMP_GL_NO		= 8		' 결의전표번호	  
	
	
   	Redim I1_f_note_gl(C_DEL_TEMP_GL_NO) 
	I1_f_note_gl(C_DEL_NOTE_NO)			= Trim(Request("txtNoteNo"))
	I1_f_note_gl(C_DEL_STS_DT)			= ""
	I1_f_note_gl(C_DEL_DEPT_CD)			= ""
	I1_f_note_gl(C_DEL_ORG_CHANGE_ID)	= ""
	I1_f_note_gl(C_DEL_INT_REV_AMT)		= ""
	I1_f_note_gl(C_DEL_INT_REV_ACCT_CD)	= ""
	I1_f_note_gl(C_DEL_NOTE_DESC)		= ""
	I1_f_note_gl(C_DEL_GL_NO)			= Trim(Request("hGlNo"))
	I1_f_note_gl(C_DEL_TEMP_GL_NO)		= Trim(Request("hTempGlNo"))

	
    Set PAFG560D = server.CreateObject ("PAFG560.cFMngDhNoteChgSvr")    
    
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    	
	
    Call PAFG560D.FN0022_MANAGE_DH_NOTE_CHG_SVR(gStrGlobalCollection,"DELETE",I1_f_note_gl)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG560D = nothing
		Exit Sub
    End If
	 
    Set PAFG560D = nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbDeleteOk          " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub
%>	
