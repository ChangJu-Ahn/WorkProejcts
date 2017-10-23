<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_RECEIPT
'*  3. Program ID        : f5101mb
'*  4. Program 이름      : 어음정보 등록 
'*  5. Program 설명      : 어음정보 등록 수정 삭제 조회 
'*  6. Comproxy 리스트   : f5101mb
'*  7. 최초 작성년월일   : 2000/10/12
'*  8. 최종 수정년월일   : 2002/03/25
'*  9. 최초 작성자       : 김종환 
'* 10. 최종 작성자       : Soo Min, Oh
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
Err.Clear                                                                        '☜: Clear Error status

Dim txtNoteNoQry
Dim lgOpModeCRUD
Dim lPtxtNoteNo
    
Call LoadBasisGlobalInf()    
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")    

Call HideStatusWnd    

Dim sChangeOrgId                                                           '☜: Hide Processing message

sChangeOrgId = Trim(request("horgchangeid"))

'---------------------------------------Common-----------------------------------------------------------

lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
txtNoteNoQry      = Request("txtNoteNoQry")

Const C_NOTE_FG		= 0		'어음구분 
Const C_NOTE_NO		= 1		'어음번호 
Const C_DEPT_CD		= 2		'부서 
Const C_DEPT_NM		= 3		'부서명 
Const C_NOTE_STS	= 4		'어음상태 
Const C_BP_CD		= 5		'거래처 
Const C_BP_NM		= 6		'거래처명 
Const C_BANK_CD		= 7		'은행 
Const C_BANK_NM		= 8		'은행명 
Const C_ISSUE_DT	= 9		'발행일 
Const C_DUE_DT		= 10	'만기일 
Const C_CASH_RATE	= 11	'현금율 
Const C_NOTE_AMT	= 12	'어음금액 
Const C_STTL_AMT	= 13	'결제금액 
Const C_PLACE		= 14	'보관장소 
Const C_RCPT_FG		= 15	'자수타수구분 
Const C_PUBLISHER	= 16	'발행인 
Const C_NOTE_DESC	= 17	'비고 

Const C_GL_DT		= 0		'일자 
Const C_SEQ			= 1		'순번 
Const C_DR_CR_FG	= 2		'차대구분 
Const C_ITEM_AMT	= 3		'금액 
Const C_ACCT_CD		= 4		'계정코드 
Const C_ACCT_NM		= 5		'계정명 
Const C_ITEM_DESC	= 6		'적요 
Const C_GL_NO		= 7		'전표번호 
Const C_TEMP_GL_NO	= 8		'결의전표번호 

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


	Const A744_I1_a_data_auth_data_BizAreaCd = 0
	Const A744_I1_a_data_auth_data_internal_cd = 1
	Const A744_I1_a_data_auth_data_sub_internal_cd = 2
	Const A744_I1_a_data_auth_data_auth_usr_id = 3


'==================================================================================
'	Name : SubBizQuery()
'	Description : 조회 정의 
'==================================================================================
Sub SubBizQuery()

	Dim PAFG505LIST	
	Dim indx
	Dim E1_f_note, EG1_export_group	
	Dim iLngRow,iLngCol
	Dim iIntLoopCount
	Dim iStrData
	Dim iStrPrevKey
	Dim iIntMaxRows
	Dim iIntQueryCount

	Const C_SHEETMAXROWS_D = 100

	' -- 권한관리추가 
	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A744_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_internal_cd)     = Trim(Request("lghInternalCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
    
    iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    
    If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)          
       End If   
    Else   
       iIntQueryCount = ""
    End If
    
    Set PAFG505LIST = server.CreateObject ("PAFG505.cFLkUpNoteSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If

    Call PAFG505LIST.FN0019_LOOKUP_NOTE_SVR(gStrGlobalCollection, iStrPrevKey, txtNoteNoQry, C_SHEETMAXROWS_D,E1_f_note,EG1_export_group, I1_a_data_auth)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG505LIST = nothing		
		Exit Sub
    End If
    
    Set PAFG505LIST = nothing 

	Response.Write "<Script Language=vbscript>  " & vbCr
   	Response.Write " with parent.frm1" & vbCr
   	Response.Write " .cboNoteFg.value	= """ & E1_f_note(C_NOTE_FG) & """														" & vbCr
   	Response.Write " .txtNoteNo.value	= """ & ConvSPChars(E1_f_note(C_NOTE_NO)) & """											" & vbCr											'☆: Company Name
   	Response.Write " .txtIssueDt.Text	= """ & UNIDateClientFormat(E1_f_note(C_ISSUE_DT)) & """									" & vbCr					   		
   	Response.Write " .txtDueDt.Text		= """ & UNIDateClientFormat(E1_f_note(C_DUE_DT)) & """									" & vbCr								
   	Response.Write " .txtDeptCD.value	= """ & ConvSPChars(E1_f_note(C_DEPT_CD)) & """											" & vbCr											'☆: Company Name
   	Response.Write " .txtDeptNm.value	= """ & ConvSPChars(E1_f_note(C_DEPT_NM)) & """											" & vbCr										'☆: Company FullName   														
   	Response.Write " .cboNoteSts.value	= """ & E1_f_note(C_NOTE_STS) & """														" & vbCr									'☆: Currency Code
   	Response.Write " .txtBankCd.value	= """ & ConvSPChars(E1_f_note(C_BANK_CD)) & """											" & vbCr										
   	Response.Write " .txtBankNm.Value	= """ & ConvSPChars(E1_f_note(C_BANK_NM)) & """											" & vbCr						   	
   	Response.Write " .txtBpCd.value		= """ & ConvSPChars(E1_f_note(C_BP_CD)) & """											" & vbCr										'☆: Currency Name
   	Response.Write " .txtBpNM.value		= """ & ConvSPChars(E1_f_note(C_BP_NM)) & """											" & vbCr										
   	Response.Write " .txtCashRate.Text	= """ & UNINumClientFormat(E1_f_note(C_CASH_RATE),	ggExchRate.DecPoint		,0) & """	" & vbCr									
   	Response.Write " .txtNoteAmt.Text	= """ & UNINumClientFormat(E1_f_note(C_NOTE_AMT),	ggAmtOfMoney.DecPoint	,0) & """	" & vbCr									
   	Response.Write " .txtSttlAmt.Text	= """ & UNINumClientFormat(E1_f_note(C_STTL_AMT),	ggAmtOfMoney.DecPoint	,0) & """	" & vbCr											
   	Response.Write " .cboPlace.value	= """ & ConvSPChars(E1_f_note(C_PLACE)) & """											" & vbCr								
   	Response.Write " .cboRcptFg.value	= """ & ConvSPChars(E1_f_note(C_RCPT_FG)) & """											" & vbCr
   	Response.Write " .txtPublisher.Value= """ & ConvSPChars(E1_f_note(C_PUBLISHER)) & """										" & vbCr
   	Response.Write " .txtNoteDesc.Value	= """ & ConvSPChars(E1_f_note(C_NOTE_DESC)) & """										" & vbCr   		
    Response.Write "End with				" & vbcr
    Response.Write "Parent.DbQueryOk		" & vbcr
    Response.Write "</Script>               " & vbCr
    
	iStrData = ""

	If IsEmpty(EG1_export_group) = False Then
		For iLngRow = 0 To UBound(EG1_export_group, 1) 	
			iIntLoopCount = iIntLoopCount + 1
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
					iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(EG1_export_group(iLngRow, C_GL_DT)))
					iStrData = iStrData & Chr(11) & Trim(EG1_export_group(iLngRow, C_SEQ))
					iStrData = iStrData & Chr(11) & Trim(EG1_export_group(iLngRow, C_DR_CR_FG))
					iStrData = iStrData & Chr(11) & ""
					iStrData = iStrData & Chr(11) & EG1_export_group(iLngRow, C_ITEM_AMT)
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_ACCT_CD)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_ACCT_NM)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_ITEM_DESC)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_GL_NO)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_TEMP_GL_NO)))
					iStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1) 
					iStrData = iStrData & Chr(11) & Chr(12)
			Else
				iStrPrevKey = EG1_export_group(UBound(EG1_export_group, 1), C_SEQ)
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

	Dim PAFG505CU
	Dim iarrData
	Dim lgIntFlgMode
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        

	' -- 권한관리추가 
	' -- 권한관리추가 
	Const A744_I1_a_data_auth_data_BizAreaCd = 0
	Const A744_I1_a_data_auth_data_internal_cd = 1
	Const A744_I1_a_data_auth_data_sub_internal_cd = 2
	Const A744_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A744_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthhInternalCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))


	Redim iarrData(C_NOTE_DESC)


	iarrData(C_NOTE_FG)		= Trim(Request("cboNoteFg"))
	iarrData(C_NOTE_NO)		= Trim(Request("txtNoteNo"))
	iarrData(C_DEPT_CD)		= Trim(Request("txtDeptCD"))
	iarrData(C_DEPT_NM)		= Trim(Request("txtDeptNm"))
	iarrData(C_NOTE_STS)	= Trim(Request("cboNoteSts"))
	iarrData(C_BP_CD)		= Trim(Request("txtBpCd"))
	iarrData(C_BP_NM)		= Trim(Request("txtBpNM"))
	iarrData(C_BANK_CD)		= Trim(Request("txtBankCd"))
	iarrData(C_BANK_NM)		= Trim(Request("txtBankNM"))
	iarrData(C_ISSUE_DT)	= UniConvDate(Request("txtIssueDt"))
	iarrData(C_DUE_DT)		= UniConvDate(Request("txtDueDt"))
	iarrData(C_CASH_RATE)	= UniConvNum(Request("txtCashRate"),0)
	iarrData(C_NOTE_AMT)	= UniConvNum(Request("txtNoteAmt"),0)
	iarrData(C_STTL_AMT)	= UniConvNum(Request("txtSttlAmt"), 0)
	iarrData(C_PLACE)		= Trim(Request("cboPlace"))
	iarrData(C_RCPT_FG)		= Trim(Request("cboRcptFg"))
	iarrData(C_PUBLISHER)	= Trim(Request("txtPublisher"))
	iarrData(C_NOTE_DESC)	= Trim(Request("txtNoteDesc"))
	
	
    Set PAFG505CU = server.CreateObject ("PAFG505.cFMngNoteSvr")   

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
     
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
    
   
    
    Select Case lgIntFlgMode
		Case  OPMD_CMODE                                                             '☜ : Create
			  Call PAFG505CU.FN0011_MANAGE_NOTE_SVR(gStrGlobalCollection,"CREATE",sChangeOrgId,iarrData, I1_a_data_auth)
        Case  OPMD_UMODE          
			  Call PAFG505CU.FN0011_MANAGE_NOTE_SVR(gStrGlobalCollection,"UPDATE",sChangeOrgId,iarrData, I1_a_data_auth)
    End Select

    If CheckSYSTEMError(Err,True) = True Then

		Set PAFG505CU = nothing
		Exit Sub	
    End If
	 
    Set PAFG505CU = nothing
    
    lPtxtNoteNo = Request("txtNoteNo")

	Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbSaveOk(""" & lPtxtNoteNo	& """)	" & vbCr
    Response.Write "</Script>									" & vbCr  
    
      
End Sub

'==================================================================================
'	Name : SubBizDelete()
'	Description : 삭제 
'==================================================================================
Sub SubBizDelete()
	Dim PAFG505D
	Dim iarrData
	
	
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        

	' -- 권한관리추가 
	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A744_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthhInternalCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A744_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

   	Redim iarrData(C_NOTE_NO)
        	
	iarrData(C_NOTE_FG)		= Request("cboNoteFg")			
	iarrData(C_NOTE_NO)		= Request("txtNoteNo")

     Set PAFG505D = server.CreateObject ("PAFG505.cFMngNoteSvr")     
    
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

    Call PAFG505D.FN0011_MANAGE_NOTE_SVR(gStrGlobalCollection,"DELETE",sChangeOrgId,iarrData, I1_a_data_auth)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG505D = nothing
		Exit Sub
    End If
	 
    Set PAFG505D = nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbDeleteOk          " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub
%>	
