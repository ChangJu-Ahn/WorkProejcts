<%'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : Prepayment
'*  3. Program ID        : f6102mb1
'*  4. Program �̸�      : ä�� û�� 
'*  5. Program ����      : ä�� û�� List, Create, Delete, Update
'*  6. Comproxy ����Ʈ   : Ar0081, Ar0081
'*  7. ���� �ۼ������   : 2000/10/07
'*  8. ���� ���������   : 2003/01/07
'*  9. ���� �ۼ���       : �ۺ��� 
'* 10. ���� �ۼ���       : Jeong Yong Kyun
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'**********************************************************************************************
								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 
Err.Clear 

Call HideStatusWnd()
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey										' ���� �� 
Dim lgStrPrevKey									' ���� �� 
Dim LngMaxRow										' ���� �׸����� �ִ�Row
Dim LngMaxRow3										' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          
Dim MajorFlag										' check box �� ���� ��ȯ �� (0,1 -> N, Y )
Dim lgIntFlgMode
Dim lgOpModeCRUD

lgIntFlgMode = Request("txtFlgMode")

lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '��: Query
      '  Call SubBizQuery()
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '��: Save,Update
      '  Call SubBizSave()
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)                                                         '��: Delete
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
	Dim pPARG090 														' ��ȸ�� ComProxy Dll ��� ����			... �Ϲ� 
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


	' -- ���Ѱ����߰� 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

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

	Response.Write " 	parent.lgNextNo = """ & """"	  & vbCr	' ���� Ű �� �Ѱ��� 
	Response.Write " 	parent.lgPrevNo = """ & """"	  & vbCr	' ���� Ű �� �Ѱ��� , ���� ComProxy�� ����� �ȵ� ����		
	
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
Sub SubBizSaveMulti() 																'��: ���� ��û�� ���� 

	On Error Resume Next
    Err.Clear																		'��: Protect system from 

	Dim pPARG090 																	' ����� ComProxy Dll ��� ����			... �Ϲ� 

	Dim arrVal, arrTemp																'��: Spread Sheet �� ���� ���� Array ���� 
	Dim strStatus																	'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
	Dim	lGrpCnt																		'��: Group Count
	Dim strCode																		'Lookup �� ���� ���� 
	Dim AAcctTransTypeTransType		
	Dim iCommandSent
	Dim AOpenArArNo
	Dim temptxtSpread

	' -- ���Ѱ����߰� 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
		
	iCommandSent = "UPDATE"	
	AOpenArArNo   = Trim(Request("txtArNo"))							
	AAcctTransTypeTransType	= "AR007"	
									
    LngMaxRow  = CInt(Request("txtMaxRows"))											'��: �ִ� ������Ʈ�� ���� 
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
