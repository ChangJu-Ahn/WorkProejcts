<%'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : Prepayment
'*  3. Program ID        : f6102mb1
'*  4. Program 이름      : 채권 청산 
'*  5. Program 설명      : 채권 청산 List, Create, Delete, Update
'*  6. Comproxy 리스트   : Ar0081, Ar0081
'*  7. 최초 작성년월일   : 2000/10/07
'*  8. 최종 수정년월일   : 2003/01/07
'*  9. 최초 작성자       : 송봉훈 
'* 10. 최종 작성자       : Jeong Yong Kyun
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************
								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜: 
Err.Clear 

Call HideStatusWnd()
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey										' 다음 값 
Dim lgStrPrevKey									' 이전 값 
Dim LngMaxRow										' 현재 그리드의 최대Row
Dim LngMaxRow3										' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          
Dim MajorFlag										' check box 를 위해 변환 값 (0,1 -> N, Y )
Dim lgIntFlgMode
Dim lgOpModeCRUD

lgIntFlgMode = Request("txtFlgMode")

lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
      '  Call SubBizQuery()
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
      '  Call SubBizSave()
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)                                                         '☜: Delete
         Call SubBizDelete()
End Select

Sub SubBizQueryMulti()

	On Error Resume Next
	
	Dim E1_b_biz_partner 
	Dim E2_b_acct_dept
	Dim E3_a_gl 
	Dim E4_a_acct
	Dim E5_a_open_ar 
	Dim E6_a_ar_adjust
	Dim EG1_export_group 
	Dim txtArNo
	Dim pPARG090 														' 조회용 ComProxy Dll 사용 변수			... 일반 
	Dim lgStrPrevKeyOne_Seq
	Dim iIntQueryCount
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	Dim lgCurrency
	
'// - Single Data 
'	Const AOpenArDocCur = 3
	Const BAcctDeptDeptCd = 1
	Const BAcctDeptDeptNm = 2
	Const AOpenArArDt     = 1
	Const BBizPartnerBpCd = 0
	Const BBizPartnerBpNm = 1
	Const AOpenArRefNo    = 2
	Const AOpenArDocCur   = 3
	
	Const AOpenArArAmt    = 6
	Const AOpenArArLocAmt = 7
	Const AOpenArArDesc   = 12
	Const AOpenArBalAmt   = 13
	Const AOpenArBalLocAmt= 14	
	Const AGlGlNo         = 15

'// - Mulity Data (SpreadSheet 1 Data)
	Const AArAdjustAdjustNo		= 0
	Const AArAdjustAdjustDt		= 1
	Const AArAdjustDocDur		= 3
	Const AArAdjustAdjustAmt	= 5
	Const AArAdjustAdjustLocAmt = 6
	Const AArAdjustTempGlNo		= 7
	Const AArAdjustAdjustDesc	= 8
	Const AArAdjustAcctCd		= 9
	Const AArAdjustAcctNm		= 10
	Const AdjustAGlGlNo			= 11

	Const C_SHEETMAXROWS_D  = 100 


	' -- 권한관리추가 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

	lgStrPrevKey = Request("lgStrPrevKey")
	txtArNo = Request("txtArNo")
	iIntQueryCount	= Request("lgPageNo")

	Set pPARG090  = Server.CreateObject("PARG090.cALkUpArAdjSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

    Call pPARG090.A_LOOKUP_AR_ADJUST_SVR( gStrGlobalCollection, Request("txtArNo"),	Request("lgStrPrevKey"), _
										  E1_b_biz_partner ,	E2_b_acct_dept ,	E3_a_gl ,		E4_a_acct, _
										  E5_a_open_ar ,		E6_a_ar_adjust,		EG1_export_group ,		I1_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set pPARG090  = Nothing
       Exit Sub
    End If    
		
    Set pPARG090  = Nothing

	lgCurrency = E5_a_open_ar(3) 'ConvSPChars(opAr0081.ExportAOpenArDocCur)
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	
	Response.Write " With parent.frm1   " & vbCr
	Response.Write " 	.txtArNo.value			=	""" & Request("txtArNo")    & """" & vbCr '"
	Response.Write " 	.txtDeptCd.value		=	""" & E2_b_acct_dept(BAcctDeptDeptCd)  & """" & vbCr 'BAcctDeptDeptCd)"
	Response.Write " 	.txtDeptNm.value		=	""" & E2_b_acct_dept(BAcctDeptDeptNm)  & """" & vbCr 'BAcctDeptDeptNm)"
	Response.Write " 	.txtArDt.text			=	""" & UNIDateClientFormat(E5_a_open_ar(AOpenArArDt))      & """" & vbCr 'AOpenArArDt)"
	Response.Write " 	.txtBpCd.value			=	""" & E1_b_biz_partner(BBizPartnerBpCd)  & """" & vbCr 'BBizPartnerBpCd)"
	Response.Write " 	.txtBpNm.value			=	""" & E1_b_biz_partner(BBizPartnerBpNm)  & """" & vbCr 'BBizPartnerBpNm)>"
	Response.Write " 	.txtRefNo.value			=	""" & E5_a_open_ar(AOpenArRefNo)      & """" & vbCr 'AOpenArRefNo)"
	Response.Write " 	.txtDocCur.value		=	""" & E5_a_open_ar(AOpenArDocCur)      & """" & vbCr '"AOpenArDocCur)"	
	Response.Write " 	.txtArAmt.text			=	""" & UNINumClientFormat(E5_a_open_ar(AOpenArArAmt),	ggAmtOfMoney.DecPoint	,0)  & """" & vbCr 'AOpenArArAmt
	Response.Write " 	.txtArLocAmt.text		=	""" & UNINumClientFormat(E5_a_open_ar(AOpenArArLocAmt),	ggAmtOfMoney.DecPoint	,0)  & """" & vbCr 'AOpenArArLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")" 
	Response.Write " 	.txtBalAmt.text		=	""" & UNINumClientFormat(E5_a_open_ar(AOpenArBalAmt),ggAmtOfMoney.DecPoint	,0)  & """" & vbCr 'AOpenArBalAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")"  
	Response.Write " 	.txtBalLocAmt.text		=	""" & UNINumClientFormat(E5_a_open_ar(AOpenArBalLocAmt),ggAmtOfMoney.DecPoint	,0)  & """" & vbCr 'AOpenArBalLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")"		
	Response.Write " 	.txtGlNo.value			=	""" & E5_a_open_ar(AGlGlNo)		   & """" & vbCr '"AGlGlNo)"
	Response.Write " 	.txtArDesc.value		=	""" & Trim(E5_a_open_ar(AOpenArArDesc))     & """" & vbCr 'AOpenArArDesc)"

	Response.Write " 	parent.lgNextNo = """ & """"	  & vbCr	' 다음 키 값 넘겨줌 
	Response.Write " 	parent.lgPrevNo = """ & """"	  & vbCr	' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음		
	
	Response.Write " End With   " & vbCr
	Response.Write " </Script>  " & vbCr       

	strData = ""
	iIntLoopCount = 0	

	If Not IsEmpty(EG1_export_group) Then
		For iLngRow = 0 To UBound(EG1_export_group, 1) 		
			iIntLoopCount = iIntLoopCount + 1
		    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then 
  	
  				strData = strData & Chr(11) & iIntLoopCount															'1  C_AdjustNo
				strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, AArAdjustAdjustDt))		'2  AArAdjustAdjustDt
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, AArAdjustAcctCd))				'3  AArAdjustAcctCd
				strData = strData & Chr(11) & ""																	'4  C_AcctCdPopUp
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, AArAdjustAcctNm))  				'5  AArAdjustAcctNm  	
				strData = strData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, AArAdjustAdjustAmt),	ggAmtOfMoney.DecPoint	,0)		' AArAdjustAdjustAmt 
				strData = strData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, AArAdjustAdjustLocAmt),	ggAmtOfMoney.DecPoint	,0) ' AArAdjustAdjustLocAmt 
				strData = strData & Chr(11) & ConvSPChars( EG1_export_group(iLngRow, AArAdjustDocDur))				' AArAdjustDocDur 	'8  C_DocCur
				strData = strData & Chr(11) & ""																	'9  C_DocCurPopUp
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, AArAdjustAdjustDesc))			' AArAdjustAdjustDesc 	'8  AdjustDesc
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, AArAdjustTempGlNo))				' AArAdjustTempGlNo 'TempGlNo       
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, AdjustAGlGlNo))					' AdjustAGlGlNo 	'10 GlNo
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, AArAdjustAdjustNo))				' AArAdjustAdjustNo 	'11 AdjustNo

				strData = strData & Chr(11) & Cstr(iLngRow + 1) & Chr(11) & Chr(12)
		    Else
				iStrPrevKey = EG1_export_group(UBound(EG1_export_group, 1), 0)
				iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next	    

	End IF

	Response.Write " <Script Language=vbscript>									" & vbCr
	Response.Write " With parent												" & vbCr
	Response.Write "	.ggoSpread.Source		=      .frm1.vspdData   " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData      """ & strData    & """" & vbCr		
	Response.Write "	.frm1.hArNo.value		= """ & txtArNo			   & """" & vbCr
	Response.Write "	.lgStrPrevKey = """ & iStrPrevKey				   & """" & vbCr
	Response.Write "	.DbQueryOk												" & vbCr		
	Response.Write " End With													" & vbCr
	Response.Write " </Script>													" & vbCr       
		
 End Sub
'--------------------------------------------------------------------------------------------------------
'                                   SAVE
'--------------------------------------------------------------------------------------------------------
Sub SubBizSaveMulti() 																'☜: 저장 요청을 받음 

	On Error Resume Next
    Err.Clear																		'☜: Protect system from 

	Dim pPARG090 																	' 저장용 ComProxy Dll 사용 변수			... 일반 

	Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt																		'☜: Group Count
	Dim strCode																		'Lookup 용 리턴 변수 
	Dim AAcctTransTypeTransType		
	Dim iCommandSent
	Dim AOpenArArNo
	Dim temptxtSpread

	' -- 권한관리추가 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
		
	iCommandSent = "UPDATE"	
	AOpenArArNo   = Trim(Request("txtArNo"))							
	AAcctTransTypeTransType	= "AR007"	
									
    LngMaxRow  = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
    LngMaxRow3 = CInt(Request("txtMaxRows3"))
	
	arrTemp = Trim(Request("txtSpread"))

    Set pPARG090 = Server.CreateObject("PARG090.cAMngArAdjSvr") 
       
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

    Call pPARG090.A_MANAGE_AR_ADJUST_SVR(gStrGlobalCollection, iCommandSent,AAcctTransTypeTransType ,AOpenArArNo, arrTemp, Request("txtSpread3"), I1_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set pPARG090 = Nothing
       Exit Sub
    End If    
    
    Set pPARG090 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

 End Sub
 
%>
