<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%'**********************************************************************************************
'*  1. Module명          : Account
'*  2. Function명        : 
'*  3. Program ID        : f4117mb1
'*  4. Program 이름      : 채무잔액정리 
'*  5. Program 설명      : 채무잔액정리 List, Create, Delete, Update
'*  6. Complus 리스트    : 
'*  7. 최초 작성년월일   : 2000/10/07
'*  8. 최종 수정년월일   : 2002/07/02
'*  9. 최초 작성자       : 송봉훈 
'* 10. 최종 작성자       : hersheys / Park, Joon-Won
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************
												'☜ : ASP가 캐쉬되지 않도록 한다.
												'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

On Error Resume Next														'☜ : Protect system from crashing
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngMaxRow3		' 현재 그리드의 최대Row
Dim LngRow
Dim lgIntFlgMode
Dim lgOpModeCRUD

lgOpModeCRUD      = Request("txtMode")                                      '☜: Read Operation Mode (CRUD)
   
Select Case lgOpModeCRUD
    Case CStr(UID_M0001) 
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                    '☜: Save,Update
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)                                                    '☜: Delete
         Call SubBizDelete()
End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    Dim I1_a_open_ap 
	Dim I2_a_ap_adjust 
	Dim E1_a_open_ap 
	Dim E2_b_biz_partner 
	Dim E3_b_acct_dept 
	Dim EG1_export_group 
	Dim E4_a_gl 
	Dim E5_a_acct 
	Dim E6_a_ap_adjust 
    Dim txtApNo
     
	Dim iPAPG085
	Dim iIntQueryCount
	
	
	Const C_SHEETMAXROWS_D  = 100
	const C_MaxQueryReCord = 0
	
	Dim LngLastRow      
    Dim LngMaxRow       
    Dim iLngRow          
    Dim strTemp
    Dim strData
	Dim lgCurrency
	Dim iStrPrevKey
	Dim iIntLoopCount
	
'//- Single Data
    Const A305_E3_org_change_id = 0   
    Const A305_E3_dept_cd = 1
    Const A305_E3_dept_nm = 2
    
    Const A305_E2_bp_cd = 0
    Const A305_E2_bp_nm = 1
    
 	Const A305_E1_ap_no = 0 
    Const A305_E1_ap_dt = 1
    Const A305_E1_ref_no = 2
    Const A305_E1_doc_cur = 3
    Const A305_E1_ap_amt = 4
    Const A305_E1_ap_loc_amt = 5
    Const A305_E1_cls_amt = 6
    Const A305_E1_cls_loc_amt = 7
    Const A305_E1_adjust_amt = 8
    Const A305_E1_adjust_loc_amt = 9
    Const A305_E1_ap_desc = 10
    Const A305_E1_xch_rate = 11
    Const A305_E1_bal_amt = 12
    Const A305_E1_bal_loc_amt = 13
    Const A306_E1_gl_no = 14
    
    '// multi Data (Spread Sheet1 Data)
	Const A305_EG1_E1_acct_cd = 0    
    Const A305_EG1_E1_acct_nm = 1
    Const A305_EG1_E2_gl_no = 2    
    Const A305_EG1_E3_adjust_no = 3   
    Const A305_EG1_E3_adjust_dt = 4
    Const A305_EG1_E3_ref_no = 5
    Const A305_EG1_E3_doc_dur = 6
    Const A305_EG1_E3_xch_rate = 7
    Const A305_EG1_E3_adjust_amt = 8
    Const A305_EG1_E3_adjust_loc_amt = 9
    Const A305_EG1_E3_temp_gl_no = 10
    Const A305_EG1_E3_adjust_desc = 11 

	' -- 조회용 
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
		
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
	
	lgStrPrevKey  = Request("lgStrPrevKey")
	txtApNo       = Request("txtApNo")
	
	Set iPAPG085 = server.CreateObject ("PAPG085.cALkUpApAdjSvr") 
	
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
 
	Call iPAPG085.A_LOOKUP_AP_ADJUST_SVR(gStrGlobalCollection, txtApNo, lgStrPrevKey, _
									      E1_a_open_ap, E2_b_biz_partner, E3_b_acct_dept, EG1_export_group, _
									      E4_a_gl,  E5_a_acct, E6_a_ap_adjust, I1_a_data_auth)
	                                        
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG085 = Nothing		
		Exit Sub
    End If

	Set iPAPG085 = Nothing  
	
	
    lgCurrency = Trim(E1_a_open_ap(A305_E1_doc_cur))
    
 	Response.Write "<Script Language=vbscript>                                                         " & vbCr
   	Response.Write " with parent.frm1                                                                  " & vbCr
	Response.Write " .txtApNo.value      = """ & Request("txtApNo")                               & """" & vbCr
	Response.Write " .txtDeptCd.value    = """ & ConvSPChars(E3_b_acct_dept(A305_E3_dept_cd))     & """" & vbCr
	Response.Write " .txtDeptNm.value    = """ & ConvSPChars(E3_b_acct_dept(A305_E3_dept_nm))     & """" & vbCr
    Response.Write " .txtApDt.text       = """ & UNIDateClientFormat(E1_a_open_ap(A305_E1_ap_dt)) & """" & vbCr		 		
	Response.Write " .txtBpCd.value      = """ & ConvSPChars(E2_b_biz_partner(A305_E2_bp_cd))     & """" & vbCr		 			 		
    Response.Write " .txtBpNm.value      = """ & ConvSPChars(E2_b_biz_partner(A305_E2_bp_nm))     & """" & vbCr		 			 				 	
    Response.Write " .txtRefNo.value     = """ & ConvSPChars(E1_a_open_ap(A305_E1_ref_no))        & """" & vbCr		 			 				 			 
    Response.Write " .txtDocCur.value    = """ & ConvSPChars(E1_a_open_ap(A305_E1_doc_cur))       & """" & vbCr		 			 				 			 		 		
    Response.Write " .txtApAmt.value     = """ & UNIConvNumDBToCompanyByCurrency(E1_a_open_ap(A305_E1_ap_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")	            & """" & vbCr		 			 				 			 		 				
    Response.Write " .txtApLocAmt.value  = """ & UNIConvNumDBToCompanyByCurrency(E1_a_open_ap(A305_E1_ap_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr		 			 				 			 		 						 			
    Response.Write " .txtBalAmt.value    = """ & UNIConvNumDBToCompanyByCurrency(E1_a_open_ap(A305_E1_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")	            & """" & vbCr		 			 				 			 		 						 				
    Response.Write " .txtBalLocAmt.value = """ & UNIConvNumDBToCompanyByCurrency(E1_a_open_ap(A305_E1_bal_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr		 			 				 			 		 						 				
    Response.Write " .txtGlNo.value      = """ & E1_a_open_ap(A306_E1_gl_no)                      & """" & vbCr		 			 				 			 		 						 				    
    Response.Write " .txtApDesc.value    = """ & ConvSPChars(E1_a_open_ap(A305_E1_ap_desc))       & """" & vbCr		 			 				 			 		 						 				        		 		
    Response.Write "  End with					                                                       " & vbcr
    Response.write " Parent.lgNextNo     = """"                                                        " & vbCr		          ' 다음 키 값 넘겨줌 
    Response.write " Parent.lgPrevNo     = """"                                                        " & vbCr		          ' 이전 키 값 넘겨줌 
    Response.Write "Parent.DbQueryOk			                                                       " & vbcr
    Response.Write "</Script>                                                                          " & vbCr  

	strData = ""
	iIntLoopCount = 0	

	If Not IsEmpty(EG1_export_group) Then
		For iLngRow = 0 To UBound(EG1_export_group, 1) 		
			iIntLoopCount = iIntLoopCount + 1
		    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then 
  		        strData = strData & Chr(11) & iIntLoopCount															'1  C_AdjustNo
				strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow,A305_EG1_E3_adjust_dt))  '2 C_AdjustDt
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,A305_EG1_E1_acct_cd))			'3  C_AcctCd
				strData = strData & Chr(11) & ""																	'4  C_AcctCdPopUp
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,A305_EG1_E1_acct_nm))  			'5  C_AcctNm 

				strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,A305_EG1_E3_adjust_amt),	lgCurrency,ggAmtOfMoneyNo, "X" , "X")	
				strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,A305_EG1_E3_adjust_loc_amt),	gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") 
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,6))								'8  C_DocCur
				strData = strData & Chr(11) & ""																	'9  C_DocCurPopUp

				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,A305_EG1_E3_adjust_desc))		'8  AdjustDesc
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,A305_EG1_E3_temp_gl_no))	        'TempGlNo       
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,A305_EG1_E2_gl_no))	  		    '10 GlNo
				strData = strData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,A305_EG1_E3_adjust_no))	   	    '11 RefNo

				strData = strData & Chr(11) & Cstr(iLngRow + 1) & Chr(11) & Chr(12)
		    Else
				iStrPrevKey = EG1_export_group(UBound(EG1_export_group, 1), A305_EG1_E3_adjust_no)
				iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next
	End if		    
	 
	Response.Write " <Script Language=vbscript>	                              " & vbCr
	Response.Write " With parent                                              " & vbCr
	Response.Write "	.ggoSpread.Source	 =   .frm1.vspdData       " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData   """ & strData     & """" & vbCr		
	Response.Write "	.frm1.hRcptNo.value		     = """ & txtApNo     & """" & vbCr
	Response.Write "	.lgStrPrevKey                = """ & iStrPrevKey & """" & vbCr
	Response.Write "	.DbQueryOk				                              " & vbCr
	Response.Write " End With                                                 " & vbCr
	Response.Write " </Script>                                                " & vbCr 
	
End Sub    	 

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim iPAPG085 																	' 저장용 ComProxy Dll 사용 변수			... 일반 
	Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt																		'☜: Group Count
	Dim strCode																		'Lookup 용 리턴 변수 
	Dim AAcctTransTypeTransType		
	Dim AOpenApApNo
	Dim temptxtSpread
	
    Const A372_IG1_I1_select_char = 0
    Const A372_IG1_I1_count = 2    
    Const A372_IG1_I3_adjust_dt = 3
    Const A372_IG1_I2_acct_cd = 4    
    Const A372_IG1_I3_adjust_amt = 5
    Const A372_IG1_I3_adjust_loc_amt = 6
    Const A372_IG1_I3_doc_dur = 7
    Const A372_IG1_I3_adjust_no = 8    
    Const A372_IG1_I3_adjust_desc = 9

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
	    								
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    
    AOpenApApNo   = Trim(Request("txtApNo"))							
	AAcctTransTypeTransType	= "AP006" 

    LngMaxRow  = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
    LngMaxRow3 = CInt(Request("txtMaxRows3"))
    
    Set iPAPG085 = Server.CreateObject("PAPG085.cAMngApAdjSvr") 

    If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
    End If    
 
    Call iPAPG085.A_MANAGE_AP_ADJUST_SVR(gStrGlobalCollection, AAcctTransTypeTransType ,AOpenApApNo, Request("txtSpread"), Request("txtSpread3"), I1_a_data_auth)

	If CheckSYSTEMError(Err, True) = True Then					
		Set PAPG085Data = Nothing
		Exit Sub
    End If    
   
    Set PAPG085Data = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
    
 End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
                                                                     '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
 '---------- Developer Coding part (Start) ---------------------------------------------------------------    
 '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	
End Sub


'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub


'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>


