<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Reference Popup Business Part												*
'*  3. Program ID           : 																			*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Reference Popup															*
'*  7. Modified date(First) : 2006/09/19																*
'*  8. Modified date(Last)  : 																			*
'*  9. Modifier (First)     : Jeng Yong Kyun															*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :																			*
'*                            																			*
'********************************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE>{{미결통합참조}}</TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->


<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>
<Script Language="VBScript">
Option Explicit                                            '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "a5150rb2.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수
Const C_MaxKey          = 31					                      '☆: SpreadSheet의 키의 갯수

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          
Dim lgPopUpR                                              
Dim IsOpenPop  

Dim Strflag

Dim arrReturn
Dim arrParent
Dim arrParam

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' {{사업장}}
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부{{부서}}
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부{{부서}}(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인

 '------ Set Parameters from Parent ASP ------ 
	arrParent = window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam = arrParent(1)
		
	top.document.title = "{{미결통합참조}}"

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()

	Redim arrReturn(0,0)

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1

	Self.Returnvalue = arrReturn
	'OpenCondition9.style.display = "none"
	' 권한관리 추가
	If UBound(arrParam) > 10 Then
		lgAuthBizAreaCd		= arrParam(11)
		lgInternalCd		= arrParam(12)
		lgSubInternalCd		= arrParam(13)
		lgAuthUsrID			= arrParam(14)
	End If	

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim lsMode 
	Dim frDt
	Dim strYear, strMonth, strDay

	Call	ExtractDateFrom("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)
	frDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01") 
	OpenCondition9.style.display = "none"
	
	txtBpCd.value		= arrParam(0)
	txtBpNm.value		= arrParam(1)
	If arrParam(8) <> "" Then
		txtDocCur.value		= arrParam(8)
	Else
		txtDocCur.value		= arrParam(2)	
	End If	

	lsMode				= arrParam(3)	
	txtBizCd.value		= arrParam(4)
	txtBizNm.value		= arrParam(5)
	txtFrOpenDt.Text    = frDt
	txtToOpenDt.Text    = arrParam(6)
	htxtAllcDt.value	= arrParam(6)
    htxtAllcAlt.value	= arrParam(7)
    htxtParentGlNo.value= arrParam(9)
    hOrgChangeId.value  = PopupParent.gChangeOrgId
    
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
'	If txtBpCd.value <> "" Then				
'		Call ggoOper.SetReqAttr(txtBpCd,   "Q")		
'	Else		
'		Call ggoOper.SetReqAttr(txtBpCd,   "N")		
'	End If
	
	If  arrParam(8) <> "" Then				
		Call ggoOper.SetReqAttr(txtDocCur,   "Q")		
	Else		
		Call ggoOper.SetReqAttr(txtDocCur,   "N")		
	End If	
	
	If  txtBizCd.value <> "" Then				
		Call ggoOper.SetReqAttr(txtBizCd,   "Q")		
	Else	
		IF lsMode = "Q" Then
			Call ggoOper.SetReqAttr(txtBizCd,   "N")		
		Else	
			Call ggoOper.SetReqAttr(txtBizCd,   "D")		
		END IF	
	End If	
	
End Sub

 '******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
'======================================================================================================
'   Event Name : OpenCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function  OpenCurrencyInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If txtDocCur.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "{{거래통화팝업}}"					' 팝업 명칭
	arrParam(1) = "b_currency"							' TABLE 명칭
	arrParam(2) = Trim(txtDocCur.value)							 	    ' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "{{거래통화}}" 			
	
    arrField(0) = "CURRENCY"							' Field명(0)
    arrField(1) = "CURRENCY_DESC"						' Field명(1)
    
    
    arrHeader(0) = "{{거래통화}}"						' Header명(0)
    arrHeader(1) = "{{거래통화명}}"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    txtDocCur.focus
		Exit Function
	Else
		Call SetCurrencyInfo(arrRet)
	End If	

End Function

'======================================================================================================
'   Event Name : SetCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function SetCurrencyInfo(Byval arrRet)
	
		txtDocCur.value = arrRet(0)
		txtDocCur.focus
	
End Function

 '------------------------------------------  OpenDeptCd()  -------------------------------------------------
'	Name : OpenDeptCd()
'	Description : Cost PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptCd()
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = txtFrOpenDt.text							'  Code Condition
   	arrParam(1) = txtToOpenDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = txtDeptCd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  

	' 권한관리 추가
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID	
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(PopUpParent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'------------------------------------------  SetDeptCd()  --------------------------------------------------
'	Name : SetDeptCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetDept(Byval arrRet)
	hOrgChangeId.value = arrRet(2)

	txtDeptCd.value = arrRet(0)
	txtDeptNm.value = arrRet(1)		
	txtFrOpenDt.text = arrRet(4)
	txtToOpenDt.text = arrRet(5)
	txtDeptCd.focus
	lgBlnFlgChgValue = True

End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(11)
	Dim arrField(6)
	Dim arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If iWhere = 1 Then
		If UCase(txtBpCd.className) = "PROTECTED" Then Exit Function
	End If
	
	IsOpenPop = True

	Select Case Trim(UCase(cboOpenType.value))
		Case "AR"
			arrParam(0) = strCode										'Code Condition
   			arrParam(1) = "A_OPEN_AR"									'채권과 연계({{거래처}} 유무)
			arrParam(2) = txtFrOpenDt.Text								'FrDt
			arrParam(3) = txtToOpenDt.Text								'ToDt
			arrParam(4) = "B"											'B :매출 S: 매입 T: 전체
			Select Case iWhere
				Case 1
					arrParam(5) = "PAYER"								'SUP :{{공급처}} PAYTO: {{지급처}} SOL:주문처 PAYER :수금처 INV:세금계산 	
				Case 2
					arrParam(5) = "SOL"									'SUP :{{공급처}} PAYTO: {{지급처}} SOL:주문처 PAYER :수금처 INV:세금계산 	
			End Select
			
			arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
				"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")				
		Case "AP"
			arrParam(0) = strCode										'Code Condition
   			arrParam(1) = "A_OPEN_AP"									'채무와 연계({{거래처}} 유무)
			arrParam(2) = txtFrOpenDt.Text								'FrDt
			arrParam(3) = txtToOpenDt.Text								'ToDt
			arrParam(4) = "B"											'B :매출 S: 매입 T: 전체
			Select Case iWhere
				Case 1
					arrParam(5) = "PAYTO"								'SUP :{{공급처}} PAYTO: {{지급처}} SOL:주문처 PAYER :수금처 INV:세금계산 	
				Case 2
					arrParam(5) = "SUP"									'SUP :{{공급처}} PAYTO: {{지급처}} SOL:주문처 PAYER :수금처 INV:세금계산 	
			End Select			
			
			arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
				"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")			
		Case "PP","PR","SS"
			arrParam(0) = "{{거래처팝업}}"
			arrParam(1) = "B_BIZ_PARTNER " 
			If iWhere = 1 then
				arrParam(2) = Trim(txtBpCd.Value)
			Else
				arrParam(2) = Trim(txtBpCd2.Value)
			End If	
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "{{거래처}}"			
	
			arrField(0) = "BP_CD"	
			arrField(1) = "BP_NM"	
	   
			arrHeader(0) = "{{거래처}}"		
			arrHeader(1) = "{{거래처명}}"						

			arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")		
	End Select	
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscBpCdPopup(iWhere)
		Exit Function
	Else
		Call SetBpCd(arrRet,iWhere)
	End If	
End Function

'------------------------------------------  EscBpCdPopup()  --------------------------------------------------
'	Name : EscBpCdPopup()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function EscBpCdPopup(ByVal BpPos)'
	
	If BpPos = 1 Then
		txtBpCd.focus
	Else
		txtBpCd2.focus
	End If
				
	lgBlnFlgChgValue = True
	
End Function
'------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetBpCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBpCd(Byval arrRet,ByVal BpPos)
	
	If BpPos = 1 Then
		txtBpCd.value = arrRet(0)		
		txtBpNm.value = arrRet(1)
		txtBpCd.focus
	Else
		txtBpCd2.value = arrRet(0)
		txtBpNm2.value = arrRet(1)
		txtBpCd2.focus
	End If
				
	lgBlnFlgChgValue = True
		
End Function

'------------------------------------------  OpenBizCd()  -------------------------------------------------
'	Name : OpenBizCd()
'	Description : Cost PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If txtBizCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "{{사업장팝업}}"					' 팝업 명칭
	arrParam(1) = "B_BIZ_AREA"						' TABLE 명칭
	arrParam(2) = Trim(txtBizCd.Value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
		' 권한관리 추가
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	arrParam(5) = "{{사업장}}"			
	
    arrField(0) = "BIZ_AREA_CD"						' Field명(0)
    arrField(1) = "BIZ_AREA_NM"						' Field명(1)
    
    arrHeader(0) = "{{사업장}}"						' Header명(0)
    arrHeader(1) = "{{사업장명}}"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	IF 	arrRet(0) <> "" then		
		Call SetBizCd(arrRet)
	Else
		txtBizCd.focus
		Exit Function
	end if
	
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetBizCd()  --------------------------------------------------
'	Name : SetBizCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizCd(Byval arrRet)
	
	txtBizCd.value = arrRet(0)		
	txtBizNm.value = arrRet(1)
	txtBizCd.focus
	lgBlnFlgChgValue = True				
	
End Function

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet				'권한관리 추가  
	Dim IntRetCD, IntRetCD1
	Dim strFrom, strWhere, strFrom1, strWhere1
	Dim arrVal, arrVal1, arrVal2, arrVal3, arrVal4, arrVal5, arrVal6, arrVal7
	DIm stbl_id, scol_id, sdata_id, stbl_id2, scol_id2, sdata_id2	 							  
	Dim strgChangeOrgId

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere			
		Case 0
			arrParam(0) = "{{계정코드팝업}}"											' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 											' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A_ACCT.MGNT_FG = " & FilterVar("Y", "''", "S") & "  and A_ACCT.mgnt_type = " & FilterVar("9", "''", "S") & " "    ' Where Condition
			arrParam(5) = "{{계정코드}}"												' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"												' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"												' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"												' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"												' Field명(3)
			
			arrHeader(0) = "{{계정코드}}"												' Header명(0)
			arrHeader(1) = "{{계정코드명}}"												' Header명(1)
			arrHeader(2) = "{{그룹코드}}"												' Header명(2)
			arrHeader(3) = "{{그룹명}}"													' Header명(3)
		Case 1
			If txtMgntCd1.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If

		    Call QueryCtrlVal()

			stbl_id = hTblId.value
			scol_id = hDataColmID.value
			arrVal3 = hDataColmNm.value

			If stbl_id = "" Then
				IsOpenPop = False
				Exit Function
			End If

			strFrom = " A_OPEN_ACCT A, " & stbl_id & " B "

			If Trim(txtAcctCd.value) <> ""  Then
				strWhere = " ACCT_CD =  " & FilterVar(txtAcctCd.value, "''", "S") & ""
				strWhere = strWhere  & " AND A.MGNT_VAL1 = B."&scol_id & " AND STATUS <> " & FilterVar("C", "''", "S") & " "
			Else
				strWhere = " A.MGNT_VAL1 = B."&scol_id
			End If

			arrParam(0) = "{{미결코드1팝업}}"											' 팝업 명칭 
			arrParam(1) = strFrom		    											' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = strWhere														' Where Condition
			arrParam(5) = "{{미결코드}}"												' 조건필드의 라벨 명칭 

			arrField(0) = "A.MGNT_VAL1"	    											' Field명(0)
			arrField(1) = "B."&arrVal3 	    											' Field명(1)

			arrHeader(0) = "{{미결관리1}}"												' Header명(0)
			arrHeader(1) = "{{미결코드}}"												' Header명(1)
		Case 2			
			If txtMgntCd2.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			Call QueryCtrlVal2()
			
			stbl_id2 = hTblId2.value
			scol_id2 = hDataColmID2.value
			arrVal7  = hDataColmNm2.value

			If stbl_id2 = "" Then
				IsOpenPop = False
				Exit Function
			End If

			strFrom1 = " A_OPEN_ACCT A, " & stbl_id2 & " B "
			strWhere1 = " ACCT_CD =  " & FilterVar(txtAcctCd.value, "''", "S") & ""
			strWhere1 = strWhere1  & " AND  A.MGNT_VAL2 = B."&scol_id2 & " AND STATUS <> " & FilterVar("C", "''", "S") & " "

			arrParam(0) = "{{미결코드2팝업}}"											' 팝업 명칭 
			arrParam(1) = strFrom1	    												' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = strWhere1														' Where Condition
			arrParam(5) = "{{미결코드}}"												' 조건필드의 라벨 명칭 

			arrField(0) = "MGNT_VAL2"	    											' Field명(0)
			arrField(1) = "B."&arrVal7 	    											' Field명(1)
   
			arrHeader(0) = "{{미결관리1}}"												' Header명(0)
			arrHeader(1) = "{{미결코드}}"
		Case 3
			arrParam(0) = "{{카드사팝업}}"									' 팝업 명칭 
			arrParam(1) = "b_card_co A"						' TABLE 명칭 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = ""														' Where Condition			
			arrParam(5) = txtCardCoCd.Alt										' 조건필드의 라벨 명칭 

			arrField(0) = "A.CARD_CO_CD"						' Field명(0)
			arrField(1) = "A.CARD_CO_NM"						' Field명(1)
   
			arrHeader(0) = txtCardCoCd.Alt					' Header명(0)
			arrHeader(1) = txtCardCoNm.Alt					' Header명(1)
		Case 4
			arrParam(0) = "{{카드번호팝업}}"								' 팝업 명칭 
			arrParam(1) = "B_CREDIT_CARD"	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = txtCardNo.Alt								' 조건필드의 라벨 명칭 

			arrField(0) = "CREDIT_NO"										' Field명(0)
			arrField(1) = "USE_USER_ID"
			arrField(2) = "CREDIT_NM"										' Field명(0)

			arrHeader(0) = txtCardNo.Alt									' Header명(0)
			arrHeader(1) = "{{사용자}}"
			arrHeader(2) = "{{카드명}}"									' Header명(1)
		Case 5
			arrParam(0) = "{{카드관리자팝업}}"								' 팝업 명칭 
			arrParam(1) = "(select distinct(use_user_id) usr_id from B_CREDIT_CARD ) a left join Haa010t b on a.usr_id=b.emp_no "	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = txtFrCardUserId.Alt								' 조건필드의 라벨 명칭 

			arrField(0) = " a.usr_id  "										' Field명(0)
			arrField(1) = " b.name "

			arrHeader(0) = "{{카드관리자}}"
			arrHeader(1) = "{{카드관리자명}}"									' Header명(1)			
		Case 6
			arrParam(0) = "{{카드관리자팝업}}"								' 팝업 명칭 
			arrParam(1) = "(select distinct(use_user_id) usr_id from B_CREDIT_CARD ) a left join Haa010t b on a.usr_id=b.emp_no "	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = txtToCardUserId.Alt								' 조건필드의 라벨 명칭 

			arrField(0) = " a.usr_id  "										' Field명(0)
			arrField(1) = " b.name "

			arrHeader(0) = "{{카드관리자}}"
			arrHeader(1) = "{{카드관리자명}}"									' Header명(1)				
			
	End Select
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			txtAcctCd.focus
		ElseIf iWhere = 1 Then
			txtMgntCd1.focus
		ElseIf iWhere = 2 Then
			txtMgntCd2.focus     
		Elseif 	iWhere = 3 Then
			txtCardCoCd.focus
		Elseif 	iWhere = 4 Then
			txtCardNo.focus
		Elseif 	iWhere = 5 Then
			txtFrCardUserId.focus			
		Elseif 	iWhere = 6 Then
			txtToCardUserId.focus			
		End If  
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	Select Case iWhere				
		Case 0
			txtAcctCd.focus
			txtAcctCd.Value = arrRet(0)
			txtAcctNm.value = arrRet(1)
			Call txtAcctcd_Onchange()
		Case 1
			txtMgntCd1.focus
			txtMgntCd1.value =  arrRet(0)	
			txtMgntCd1Nm.value =  arrRet(1)	
		Case 2
			txtMgntCd2.focus
			txtMgntCd2.value =  arrRet(0)		
			txtMgntCd2Nm.value =  arrRet(1)
		Case 3
			txtCardCoCd.Value = arrRet(0)
			txtCardCoNm.value = arrRet(1)
		Case 4
			txtCardNo.value = arrRet(0)
		Case 5
			txtFrCardUserId.value = arrRet(0)
			txtFrCardUserNm.value = arrRet(1)
		Case 6
			txtToCardUserId.value = arrRet(0)
			txtToCardUserNm.value = arrRet(1)
	End Select

	lgBlnFlgChgValue = True
End Function


'========================================  2.3 LoadInfTB19029()  =========================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "A","NOCOOKIE","RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("I", "A", "NOCOOKIE", "RA") %>
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
Function OKClick()
	Dim ii ,jj ,kk

	
	if vspdData.SelModeSelCount > 0 Then 			
		Redim arrReturn(vspdData.SelModeSelCount - 1,C_MaxKey)
		kk = 0
		For ii = 0 To vspdData.MaxRows - 1
			vspdData.Row = ii + 1			
			If vspdData.SelModeSelected Then
				For jj = 0 To C_MaxKey - 1
					vspdData.Col	 = GetKeyPos("A",jj + 1)		
					arrReturn(kk,jj) = vspdData.Text
					
				Next			
'				arrReturn(kk,C_MaxKey) = txtDocCur.value
				kk = kk + 1
			End If
		Next	
	End If			
	
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)
		
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
	Err.clear
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1050", "''", "S") & "  order by minor_cd", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(cboOpenType ,lgF0  ,lgF1  ,Chr(11))			
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    vspddata.OperationMode = 5
    Call SetZAdoSpreadSheet("a5150RA2","S","A","V20080515",PopupParent.C_SORT_DBAGENT,vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock() 
End Sub

Sub InitSpreadSheetOPEN()
    vspddata.OperationMode = 5
    Call SetZAdoSpreadSheet("a5150RA2A","S","A","V20080525",PopupParent.C_SORT_DBAGENT,vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock() 
End Sub


'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	vspdData.ReDraw = False
	ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
	vspdData.ReDraw = True
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029()														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)        
    
    Call ggoOper.LockField(Document, "N")                                   
    
	Call SetDefaultVal()						'1
	Call InitVariables()						'2		//logic은 1->2순으로 처리되어야 함.				
	Call InitSpreadSheet()
	Call InitComboBox()	
	cboOpenType.value ="AP"
    call  cboOpenType_OnChange()
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    FncQuery = False                                            
    
    Err.Clear                                                   

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	If Not ChkQueryDate Then
		Exit Function
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables() 											
	ggoSpread.Source = vspdData
	ggoSpread.ClearSpreadData
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)    
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal
	Dim iChkLocalCur

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)

	If chkLocalCur.checked = True Then
		iChkLocalCur = "Y"
	Else	
		iChkLocalCur = "N"
	End If     

	If txtAllcAmt = "" Then
		txtAllcAmt.text = 0
	End If     	

    strVal = BIZ_PGM_ID

    If lgIntFlgMode  <> PopupParent.OPMD_UMODE Then   ' This means that it is first search
		strVal = strVal & "?txtOpenType="		& Trim(cboOpenType.value)
		strVal = strVal & "&txtFrOpenDt="		& UniConvDateAToB(txtFrOpenDt.Text,popupparent.gDateFormat,popupparent.gServerDateFormat)		
		strVal = strVal & "&txtToOpenDt="		& UniConvDateAToB(txtToOpenDt.Text,popupparent.gDateFormat,popupparent.gServerDateFormat)	
		strVal = strVal & "&txtDocCur="			& Trim(txtDocCur.value)
		strVal = strVal & "&txtOrgChangeId="	& Trim(hOrgChangeId.value)			
		strVal = strVal & "&txtDeptCd="			& Trim(txtDeptCd.value)					
		strVal = strVal & "&txtBpCd="			& Trim(txtBpCd.value)
		strVal = strVal & "&txtBizCd="			& Trim(txtBizCd.value)					
		strVal = strVal & "&txtBpCd2="			& Trim(txtBpCd2.value)
		strVal = strVal & "&txtProject="		& Trim(txtProject.value)
		strVal = strVal & "&txtAllcAmt="		& Trim(txtAllcAmt.text)
		strVal = strVal & "&txtRefNo="			& Trim(txtRefNo.value)
		strVal = strVal & "&txtAcctCd="			& Trim(txtAcctCd.value)
		strVal = strVal & "&txtGlNo="			& Trim(txtGlNo.value)
		strVal = strVal & "&txtMgntCd1="		& Trim(txtMgntCd1.value)
		strVal = strVal & "&txtMgntCd2="		& Trim(txtMgntCd2.value)									
		strVal = strVal & "&txtCardCoCd="		& Trim(txtCardCoCd.value)
		strVal = strVal & "&txtCardNo="			& Trim(txtCardNo.value)
		strVal = strVal & "&txtFrCardUserId="	& Trim(txtFrCardUserId.value)
		strVal = strVal & "&txtToCardUserId="	& Trim(txtToCardUserId.value)				
		strVal = strVal & "&chkLocalCur="		& Trim(iChkLocalCur)
    Else
		strVal = strVal & "?txtOpenType="		& Trim(htxtOpenType.value)
		strVal = strVal & "&txtFrOpenDt="		& Trim(htxtFrOpenDt.value)				
		strVal = strVal & "&txtToOpenDt="		& Trim(htxtToOpenDt.value)				
		strVal = strVal & "&txtDocCur="			& Trim(htxtDocCur.value)
		strVal = strVal & "&txtOrgChangeId="	& Trim(hOrgChangeId.value)			
		strVal = strVal & "&txtDeptCd="			& Trim(htxtDeptCd.value)					
		strVal = strVal & "&txtBpCd="			& Trim(htxtBpCd.value)
		strVal = strVal & "&txtBizCd="			& Trim(htxtBizCd.value)					
		strVal = strVal & "&txtBpCd2="			& Trim(htxtBpCd2.value)
		strVal = strVal & "&txtProject="		& Trim(htxtProject.value)
		strVal = strVal & "&txtAllcAmt="		& Trim(htxtAllcAmt.value)
		strVal = strVal & "&txtRefNo="			& Trim(htxtRefNo.value)
		strVal = strVal & "&txtAcctCd="			& Trim(htxtAcctCd.value)
		strVal = strVal & "&txtGlNo="			& Trim(htxtGlNo.value)
		strVal = strVal & "&txtMgntCd1="		& Trim(htxtMgntCd1.value)
		strVal = strVal & "&txtMgntCd2="		& Trim(htxtMgntCd2.value)									
		strVal = strVal & "&txtCardCoCd="		& Trim(htxtCardCoCd.value)
		strVal = strVal & "&txtCardNo="			& Trim(htxtCardNo.value)
		strVal = strVal & "&txtFrCardUserId="	& Trim(htxtFrCardUserId.value)
		strVal = strVal & "&txtToCardUserId="	& Trim(htxtToCardUserId.value)
		strVal = strVal & "&chkLocalCur="		& Trim(hchkLocalCur.value)
	End If   



	strVal = strVal & "&txtParentGLNo="			& Trim(htxtParentGlNo.value)
	strVal = strVal & "&txtDeptCd_alt="			& Trim(txtDeptCd.alt)					
	strVal = strVal & "&txtBpCd_alt="			& Trim(txtBpCd.alt)
	strVal = strVal & "&txtBizCd_alt="			& Trim(txtBizCd.alt)					
	strVal = strVal & "&txtBpCd2_alt="			& Trim(txtBpCd2.alt)
	strVal = strVal & "&txtAcctCd_alt="			& Trim(txtAcctCd.alt)
	strVal = strVal & "&txtMgntCd1_alt="		& Trim(txtMgntCd1.alt)
	strVal = strVal & "&txtMgntCd2_alt="		& Trim(txtMgntCd2.alt)									
	strVal = strVal & "&txtCardCoCd_alt="		& Trim(txtCardCoCd.alt)
	strVal = strVal & "&txtFrCardUserId_alt="	& Trim(txtFrCardUserId.alt)
	strVal = strVal & "&txtToCardUserId_alt="	& Trim(txtToCardUserId.alt)	
	strVal = strVal & "&txtFrDueDt="	& UniConvDateAToB(Trim(txtFrDueDt.text),popupparent.gDateFormat,popupparent.gServerDateFormat)
	strVal = strVal & "&txtToDueDt="	& UniConvDateAToB(Trim(txtToDueDt.text),popupparent.gDateFormat,popupparent.gServerDateFormat)
	
	 'UniConvDateAToB(Trim(txtFrDueDt.text),parent.gDateFormat,parent.gServerDateFormat)
	

	strVal = strVal & "&lgPageNo="       & lgPageNo         
	strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
	strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
	strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
 
	Call RunMyBizASP(MyBizASP, strVal)							

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode

	If vspdData.MaxRows > 0 Then
		vspdData.Focus
	End If

End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenSortPopup()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Popupparent.SORTW_WIDTH & "px; dialogHeight=" & Popupparent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables()
       Call InitSpreadSheet()       
   End If
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OkClick()
		End If
	End If
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SPC"   
    
    If Row = 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
	Call SetSpreadColumnValue("A",vspdData,Col,Row)	        
	
    If vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyPress
'   Event Desc : 
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function
	
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
End Sub



'=======================================================================================================
'   Function Name : ChkQueryDate
'   Function Desc : 
'=======================================================================================================
Function ChkQueryDate()
	chkQueryDate= True
	
	If CompareDateByFormat(txtFrOpenDt.text, txtFrOpenDt.text, txtFrOpenDt.Alt, txtToOpenDt.Alt, _
   	           "970025", txtFrOpenDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
		chkQueryDate= False
		txtFrOpenDt.focus
		Exit Function
	End If
	
'	If CompareDateByFormat(txtArDt.text,htxtAllcDt.Value,txtArDt.Alt,htxtAllcAlt.value, _
 '  	           "970025",txtArDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
'	   chkQueryDate= False
'	   txtArDt.focus
'	   Exit Function
'	End If
	
'	If CompareDateByFormat(txtToArDt.text,htxtAllcDt.Value,txtToArDt.Alt, htxtAllcAlt.value,_
 '  	           "970025",txtToArDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
'	   chkQueryDate= False
'	   txtToArDt.focus
'	   Exit Function
'	End If

End Function

Sub cboOpenType_OnChange()
	Dim	i
	Dim IntRetCD
	OpenCondition9.style.display = "none"
	
	Select Case Trim(UCase(cboOpenType.value))
		Case "AR"
			If OpenCondition2.style.display = "none" Then
				OpenCondition2.style.display = ""
			End If			
			If OpenCondition3.style.display = "none" Then
				OpenCondition3.style.display = ""
			End If			
			If OpenCondition4.style.display = "none" Then
				OpenCondition4.style.display = ""
			End If			
			If OpenCondition5.style.display = "none" Then
			Else
				OpenCondition5.style.display = "none"
			End If			
			If OpenCondition6.style.display = "none" Then
			Else
				OpenCondition6.style.display = "none"
			End If			
			If OpenCondition7.style.display = "none" Then
			Else
				OpenCondition7.style.display = "none"
			End If
			If OpenCondition8.style.display = "none" Then
			Else
				OpenCondition8.style.display = "none"
			End If		
			OpenCondition9.style.display = ""
			spnNm1.innerHTML = "{{수금처}}"
			spnNm2.innerHTML = "{{주문처}}"
			spnNm3.innerHTML = "{{채권금액}}"
			txtBpCd.Alt = "{{수금처}}"
			txtBpCd2.Alt = "{{주문처}}"			
			txtAllcAmt.Alt = "{{채권금액}}"
			Call ggoOper.SetReqAttr(txtProject,   "D")			
			call InitSpreadSheet
			Call ClearSpnField
		Case "AP"
			If OpenCondition2.style.display = "none" Then
				OpenCondition2.style.display = ""
			End If			
			If OpenCondition3.style.display = "none" Then
				OpenCondition3.style.display = ""
			End If			
			If OpenCondition4.style.display = "none" Then
				OpenCondition4.style.display = ""
			End If			
			If OpenCondition5.style.display = "none" Then
			Else
				OpenCondition5.style.display = "none"
			End If			
			If OpenCondition6.style.display = "none" Then
			Else
				OpenCondition6.style.display = "none"
			End If			
			If OpenCondition7.style.display = "none" Then
			Else
				OpenCondition7.style.display = "none"
			End If
			If OpenCondition8.style.display = "none" Then
			Else
				OpenCondition8.style.display = "none"
			End If			
			OpenCondition9.style.display = ""
			spnNm1.innerHTML = "{{지급처}}"
			spnNm2.innerHTML = "{{공급처}}"
			spnNm3.innerHTML = "{{채무금액}}"
			txtBpCd.Alt = "{{지급처}}"
			txtBpCd2.Alt = "{{공급처}}"			
			txtAllcAmt.Alt = "{{채무금액}}"		
			Call ggoOper.SetReqAttr(txtProject,   "Q")
			call InitSpreadSheet
			Call ClearSpnField
		Case "PP"
			If OpenCondition2.style.display = "none" Then
				OpenCondition2.style.display = ""
			End If			
			If OpenCondition3.style.display = "none" Then
			Else
				OpenCondition3.style.display = "none"
			End If			
			If OpenCondition4.style.display = "none" Then
				OpenCondition4.style.display = ""
			End If			
			If OpenCondition5.style.display = "none" Then
			Else
				OpenCondition5.style.display = "none"
			End If			
			If OpenCondition6.style.display = "none" Then
			Else
				OpenCondition6.style.display = "none"
			End If			
			If OpenCondition7.style.display = "none" Then
			Else
				OpenCondition7.style.display = "none"
			End If
			If OpenCondition8.style.display = "none" Then
			Else
				OpenCondition8.style.display = "none"
			End If			
			spnNm1.innerHTML = "{{거래처}}"
			spnNm2.innerHTML = "{{거래처}}"
			spnNm3.innerHTML = "{{선급금액}}"
			txtBpCd.Alt = "{{거래처}}"
			txtBpCd2.Alt = "{{거래처}}"			
			txtAllcAmt.Alt = "{{선급금액}}"
			Call ggoOper.SetReqAttr(txtProject,   "Q")
			call InitSpreadSheet			
			Call ClearSpnField
		Case "PR"
			If OpenCondition2.style.display = "none" Then
				OpenCondition2.style.display = ""
			End If			
			If OpenCondition3.style.display = "none" Then
				OpenCondition3.style.display = ""
			End If			
			If OpenCondition4.style.display = "none" Then
				OpenCondition4.style.display = ""
			End If			
			If OpenCondition5.style.display = "none" Then
			Else
				OpenCondition5.style.display = "none"
			End If			
			If OpenCondition6.style.display = "none" Then
			Else
				OpenCondition6.style.display = "none"
			End If			
			If OpenCondition7.style.display = "none" Then
			Else
				OpenCondition7.style.display = "none"
			End If
			If OpenCondition8.style.display = "none" Then
			Else
				OpenCondition8.style.display = "none"
			End If			
			spnNm1.innerHTML = "{{거래처}}"
			spnNm2.innerHTML = "{{거래처}}"
			spnNm3.innerHTML = "{{선수금액}}"
			txtBpCd.Alt = "{{거래처}}"
			txtBpCd2.Alt = "{{거래처}}"			
			txtAllcAmt.Alt = "{{선수금액}}"
			Call ggoOper.SetReqAttr(txtProject, "D")
			Call ggoOper.SetReqAttr(txtBpCd2,   "Q")									
			call InitSpreadSheet						
			Call ClearSpnField
		Case "SS"
			If OpenCondition2.style.display = "none" Then
				OpenCondition2.style.display = ""
			End If			
			If OpenCondition3.style.display = "none" Then
				OpenCondition3.style.display = ""
			End If			
			If OpenCondition4.style.display = "none" Then
				OpenCondition4.style.display = ""
			End If			
			If OpenCondition5.style.display = "none" Then
			Else
				OpenCondition5.style.display = "none"
			End If			
			If OpenCondition6.style.display = "none" Then
			Else
				OpenCondition6.style.display = "none"
			End If			
			If OpenCondition7.style.display = "none" Then
			Else
				OpenCondition7.style.display = "none"
			End If
			If OpenCondition8.style.display = "none" Then
			Else
				OpenCondition8.style.display = "none"
			End If								
			spnNm1.innerHTML = "{{거래처}}"
			spnNm2.innerHTML = "{{거래처}}"
			spnNm3.innerHTML = "{{가수금액}}"
			txtBpCd.Alt = "{{거래처}}"
			txtBpCd2.Alt = "{{거래처}}"			
			txtAllcAmt.Alt =  "{{가수금액}}"
			Call ggoOper.SetReqAttr(txtProject, "D")			
			Call ggoOper.SetReqAttr(txtBpCd2,   "Q")
			call InitSpreadSheet
			Call ClearSpnField
		Case "U9"
			If OpenCondition2.style.display = "none" Then
			Else
				OpenCondition2.style.display = "none"
			End If			
			If OpenCondition3.style.display = "none" Then
			Else
				OpenCondition3.style.display = "none"
			End If			
			If OpenCondition4.style.display = "none" Then
			Else
				OpenCondition4.style.display = "none"
			End If			
			If OpenCondition5.style.display = "none" Then
				OpenCondition5.style.display = ""
			End If			
			If OpenCondition6.style.display = "none" Then
				OpenCondition6.style.display = ""
			End If			
			If OpenCondition7.style.display = "none" Then
			Else
				OpenCondition7.style.display = "none"
			End If
			If OpenCondition8.style.display = "none" Then
			Else
				OpenCondition8.style.display = "none"
			End If			
			OpenCondition9.style.display = ""
			call InitSpreadSheetOPEN()
			Call ClearSpnField
		Case "U6"
			If OpenCondition2.style.display = "none" Then
			Else
				OpenCondition2.style.display = "none"
			End If			
			If OpenCondition3.style.display = "none" Then
			Else
				OpenCondition3.style.display = "none"
			End If			
			If OpenCondition4.style.display = "none" Then
			Else
				OpenCondition4.style.display = "none"
			End If			
			If OpenCondition5.style.display = "none" Then
			Else
				OpenCondition5.style.display = "none"
			End If			
			If OpenCondition6.style.display = "none" Then
			Else
				OpenCondition6.style.display = "none"
			End If			
			If OpenCondition7.style.display = "none" Then
				OpenCondition7.style.display = ""
			End If
			If OpenCondition8.style.display = "none" Then
				OpenCondition8.style.display = ""
			End If
			OpenCondition9.style.display = ""
			call InitSpreadSheetOPEN()			
			Call ClearSpnField			
		Case Else	
		OpenCondition9.style.display = "none"
			If OpenCondition2.style.display = "none" Then
			Else
				OpenCondition2.style.display = "none"
			End If			
			If OpenCondition3.style.display = "none" Then
			Else
				OpenCondition3.style.display = "none"
			End If			
			If OpenCondition4.style.display = "none" Then
				OpenCondition4.style.display = "none"
			End If			
			If OpenCondition5.style.display = "none" Then
			Else
				OpenCondition5.style.display = "none"
			End If			
			If OpenCondition6.style.display = "none" Then
			Else
				OpenCondition6.style.display = "none"
			End If			
			If OpenCondition7.style.display = "none" Then
			Else
				OpenCondition7.style.display = "none"
			End If
			If OpenCondition8.style.display = "none" Then
			Else
				OpenCondition8.style.display = "none"
			End If			
			spnNm1.innerHTML = "{{거래처}}"
			spnNm2.innerHTML = "{{거래처}}"
			spnNm3.innerHTML = "{{금액}}"
			txtBpCd.Alt = "{{거래처}}"
			txtBpCd2.Alt = "{{거래처}}"			
			txtAllcAmt.Alt = "{{금액}}"
			Call ClearSpnField
	End Select	

	ggoSpread.Source = vspddata
	ggoSpread.ClearSpreadData()

	lgBlnFlgChgValue = True
End Sub

Sub ClearSpnField()
	txtBpCd2.value = ""
	txtBpNm2.value = ""	
	txtAllcAmt.Text = 0
End Sub

Function QueryCtrlVal()
    Dim ArrRet

    Call CommonQueryRs("TBL_ID, KEY_COLM_ID1, DATA_COLM_NM,COLM_DATA_TYPE", _
                       "A_ACCT A, A_CTRL_ITEM B", _
                       "A.mgnt_cd1 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(txtAcctCd.value, "''", "S"),_
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ArrRet 	= Split(lgF0,Chr(11))

	If Trim(ArrRet(0)) <> "" then
		Strflag = "1"
		hTblId.value  = ArrRet(0)

		ArrRet 	= Split(lgF1,Chr(11))
		hDataColmID.value  = ArrRet(0)
		ArrRet 	= Split(lgF2,Chr(11))
		hDataColmNm.value = ArrRet(0)
	Else
		Strflag = "2"
		If replace(lgF3,Chr(11),"") = "D" Then
			txtMgntCd1Nm.value = "YYYY-MM-DD"
		Elseif replace(lgF3,Chr(11),"") = "N" Then
			txtMgntCd1Nm.value = "{{숫자는 구분자없이}}"
		End If	 
	End If

End Function

Function QueryCtrlVal2()
    Dim ArrRet

    Call CommonQueryRs("TBL_ID, KEY_COLM_ID1, DATA_COLM_NM,COLM_DATA_TYPE", _
                       "A_ACCT A, A_CTRL_ITEM B", _
                       "A.mgnt_cd2 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(txtAcctCd.value, "''", "S"),_
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ArrRet 	= Split(lgF0,Chr(11))

	If Trim(ArrRet(0)) <> "" Then
		Strflag = "1"
		hTblId2.value  = ArrRet(0)
		
		ArrRet 	= Split(lgF1,Chr(11))
		hDataColmID2.value  = ArrRet(0)
		ArrRet 	= Split(lgF2,Chr(11))
		hDataColmNm2.value = ArrRet(0)
	Else
		Strflag = "2"		
		If replace(lgF3,Chr(11),"") = "D" Then
			txtMgntCd2Nm.value = "YYYY-MM-DD"
		Elseif replace(lgF3,Chr(11),"") = "N" Then
			txtMgntCd2Nm.value = "{{숫자는 구분자없이}}"
		End If	 
	End If
End Function

Function txtAcctCd_Onchange()
    txtAcctCd_OnChange = False

	Call CommonQueryRs("distinct A_ACCT.ACCT_CD, ACCT_NM ","A_ACCT, A_ACCT_CTRL_ASSN","A_ACCT.ACCT_CD = '" & txtAcctCd.value & "' AND A_ACCT.acct_cd = a_acct_ctrl_assn.acct_cd" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If (lgF0 <> "X") And (Trim(lgF0) <> "") Then 
		txtAcctNm.value = Left(lgF1, Len(lgF1)-1)    
		txtMgntCd1.value = ""
		txtMgntCd1Nm.value = ""
		txtMgntCd2.value = ""
		txtMgntCd2Nm.value = ""

		Call CommonQueryRs("CTRL_NM", _
                   "A_ACCT A, A_CTRL_ITEM B", _
                   "A.mgnt_cd1 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(txtAcctCd.value, "''", "S"),_
                   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If lgF0 <> "" then
			CtrlCd.innerHTML = REPLACE(lgF0,Chr(11),"") 
		Else
			CtrlCd.innerHTML = "{{미결코드1}}" 
		End if

		Call CommonQueryRs("CTRL_NM", _
                   "A_ACCT A, A_CTRL_ITEM B", _
                   "A.mgnt_cd2 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(txtAcctCd.value, "''", "S"),_
                   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If lgF0 <> "" then
			CtrlCd2.innerHTML = REPLACE(lgF0,Chr(11),"")
		Else
			CtrlCd2.innerHTML = "{{미결코드2}}"
		End if

		Call ggoOper.SetReqAttr(txtMgntCd1,		"D")
		Call ggoOper.SetReqAttr(txtMgntCd2,		"D")
		Call ggoOper.SetReqAttr(txtMgntCd1Nm,	"Q")
		Call ggoOper.SetReqAttr(txtMgntCd2Nm,	"Q")
		txtAcctCd.focus
	Else       
		txtAcctCd.value = ""
		txtAcctNm.value = ""
		txtMgntCd1.value = ""
		txtMgntCd1Nm.value = ""
		txtMgntCd2.value = ""
		txtMgntCd2Nm.value = ""
		CtrlCd.innerHTML = "{{미결코드1}}"
		CtrlCd2.innerHTML = "{{미결코드2}}"
		Call ggoOper.SetReqAttr(txtMgntCd1,		"Q")
		Call ggoOper.SetReqAttr(txtMgntCd2,		"Q")
		Call ggoOper.SetReqAttr(txtMgntCd1Nm,	"Q")
		Call ggoOper.SetReqAttr(txtMgntCd2Nm,	"Q")      
		'txtCtrlVal.value = ""
		'txtCtrlValNm.value = ""       
		txtAcctCd.focus       
	End If   

    txtAcctCd_OnChange = True
End Function

'=======================================================================================================
'   Event Name : txtFrOpenDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtFrOpenDt_DblClick(Button)
    If Button = 1 Then
        txtFrOpenDt.Action = 7                        
        Call SetFocusToDocument("P")
		txtFrOpenDt.Focus 
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToOpenDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtToOpenDt_DblClick(Button)
    If Button = 1 Then
        txtToOpenDt.Action = 7                        
        Call SetFocusToDocument("P")
		txtToOpenDt.Focus 
    End If
End Sub


Sub txtFrOpenDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub txtToOpenDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub  txtFrDueDt_DblClick(Button)
    If Button = 1 Then
        txtFrDueDt.Action = 7                        
        Call SetFocusToDocument("P")
		txtFrDueDt.Focus 
    End If
End Sub


Sub txtFrDueDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

  
Sub  txtToDueDt_DblClick(Button)
    If Button = 1 Then
        txtToDueDt.Action = 7                        
        Call SetFocusToDocument("P")
		txtToDueDt.Focus 
    End If
End Sub


Sub txtToDueDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub





Sub txtBpCd_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR ID="StdCondition" style="display: yes">
						<TD CLASS=TD5 NOWRAP>{{미결구분}}</TD>
						<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOpenType" tag="12" STYLE="WIDTH:82px:" ALT="{{미결구분}}"><OPTION VALUE="" selected></OPTION></SELECT></TD>								
						<TD CLASS=TD5 NOWRAP>{{미결일자}}</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrOpenDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="{{미결시작일자}}" id=OBJECT3></OBJECT>');</SCRIPT>								
						&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToOpenDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="{{미결종료일자}}" id=OBJECT4></OBJECT>');</SCRIPT></TD>												
					</TR>
					<TR ID="StdCondition1" style="display: yes">
						<TD CLASS=TD5 NOWRAP>{{거래통화}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="{{거래통화}}" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="12NXXU"><IMG align=top name=btnCalType onclick="vbscript:OpenCurrencyInfo()" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
											 <INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkLocalCur ID=chkLocalCur tag="">{{자국통화}}</TD>
						<TD CLASS=TD5 NOWRAP>{{부서}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="{{부서}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenDeptCd()"> <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14" ALT="{{부서명}}"></TD>					
					</TR>
					<TR ID="OpenCondition2" style="display: none">						
						<TD CLASS=TD5 NOWRAP><span id="spnNm1">{{거래처}}</span></TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11NXXU" Alt="{{거래처}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(txtBpCd.Value, 1)"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="{{거래처명}}"></TD>				
						<TD CLASS=TD5 NOWRAP>{{사업장}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="{{사업장}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizCd()"> <INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14" ALT="{{사업장명}}"></TD>					
					</TR>
					<TR ID="OpenCondition3" style="display: none">
						<TD CLASS=TD5 NOWRAP><span id="spnNm2">{{거래처2}}</span></TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd2" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11NXXU" ALT="{{거래처2}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(txtBpCd2.Value, 2)"> <INPUT TYPE=TEXT NAME="txtBpNm2" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="{{거래처명2}}"></TD>
						<TD CLASS=TD5 NOWRAP>{{프로젝트번호}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="{{프로젝트번호}}" MAXLENGTH=25 SIZE=25 tag="1X"></TD>
					</TR> 
					<TR ID="OpenCondition4" style="display: yes">
						<TD CLASS=TD5 NOWRAP><span id="spnNm3">{{금액}}</span></TD>
						<TD CLASS=TD6 NOWRAP>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAllcAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="{{금액}}" tag="11X2" id=OBJECT1></OBJECT>');</SCRIPT>											
						</TD>
						<TD CLASS=TD5 NOWRAP>{{참조번호}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME=txtRefNo ALT="{{참조번호}}" MAXLENGTH=25 SIZE=25 tag="1X"></TD>						
					
					</TR> 
					
		
					
					<TR ID="OpenCondition5" style="display: none">				
						<TD CLASS=TD5 NOWRAP>{{계정코드}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="{{계정코드}}" MAXLENGTH="10" SIZE=11 tag ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtAcctCd.Value,0)">
											 <INPUT NAME="txtAcctNm" ALT="{{계정명}}"   MAXLENGTH="20" SIZE=18 tag ="14XXXU"></TD>
						<TD CLASS=TD5 ID="CtrlCd" NOWRAP>{{미결코드1}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMgntCd1" ALT="{{미결코드1}}" MAXLENGTH="30" SIZE=20 tag ="14XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtMgntCd1.Value,1)">
											 <INPUT NAME="txtMgntCd1Nm" ALT="{{미결코드명1}}"   MAXLENGTH="30" SIZE=18 tag ="14XXXU"></TD>
					</TR>
					<TR ID="OpenCondition6" style="display: none">
						<TD CLASS=TD5 NOWRAP>{{전표번호}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtGlNo" SIZE=18 MAXLENGTH=18 tag="1XXXXU" ALT="{{전표번호}}"></TD>
						<TD CLASS=TD5 ID="CtrlCd2" NOWRAP>{{미결코드2}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMgntCd2" ALT="{{미결코드2}}" MAXLENGTH="30" SIZE=20 tag ="14XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtMgntCd2.Value,2)">
											 <INPUT NAME="txtMgntCd2Nm" ALT="{{미결코드명2}}"   MAXLENGTH="30" SIZE=18 tag ="14XXXU"></TD>
					</TR>
					<TR ID="OpenCondition7" style="display: none">				
                        <TD CLASS=TD5 NOWRAP>{{카드사}}</TD>
                        <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCardCoCd"  SIZE="10" MAXLENGTH="10" TAG="11xxxU" ALT="{{카드사}}"><IMG SRC="../../image/btnPopup.gif" NAME="bntCardCoCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtCardCoCd.value,3)">
											 <INPUT TYPE=TEXT NAME="txtCardCoNm"  SIZE=20   TAG="14xxxU" ALT="{{카드사명}}"></TD>
						<TD CLASS=TD5 NOWRAP>{{카드번호}}</TD>
                        <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCardNo"  SIZE=20 MAXLENGTH=20 TAG="11XXXU" ALT="{{카드번호}}"><IMG SRC="../../image/btnPopup.gif" NAME="btnCardNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtCardNo.value,4)"></TD>											
					</TR>
					<TR ID="OpenCondition8" style="display: none">				
                        <TD CLASS=TD5 NOWRAP>{{카드관리자}}</TD>
                        <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFrCardUserId" SIZE="8" MAXLENGTH="12" TAG="11xxxU" ALT="{{카드관리자1}}"><IMG SRC="../../image/btnPopup.gif" NAME="bntFrCardUserId" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtFrCardUserId.value,5)"> <INPUT TYPE=TEXT NAME="txtFrCardUserNm"  SIZE=10   TAG="14xxxU" ALT="{{카드관리자명1}}">&nbsp;~
											 <INPUT TYPE=TEXT NAME="txtToCardUserId" SIZE="8" MAXLENGTH="12" TAG="11xxxU" ALT="{{카드관리자2}}"><IMG SRC="../../image/btnPopup.gif" NAME="bntToCardUserId" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtToCardUserId.value,6)"> <INPUT TYPE=TEXT NAME="txtToCardUserNm"  SIZE=10   TAG="14xxxU" ALT="{{카드관리자명2}}">
                        </TD>
						<TD CLASS=TD5 NOWRAP></TD>
                        <TD CLASS=TD6 NOWRAP></TD>											
					</TR>
					
					<TR ID="OpenCondition9" style="display: ">				
						<TD CLASS=TD5 NOWRAP>{{만기일자}}</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrDueDt" CLASS=FPDTYYYYMMDD tag="11" Title="FPDATETIME" ALT="{{만기시작일자}}" id=OBJECT3></OBJECT>');</SCRIPT>								
						&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToDueDt" CLASS=FPDTYYYYMMDD tag="11" Title="FPDATETIME" ALT="{{만기종료일자}}" id=OBJECT4></OBJECT>');</SCRIPT></TD>												

						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>						
					</TR> 
					
					
										 
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
	<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vspdData tag="2"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Call FncQuery()"></IMG>
						&nbsp;<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					
					</TD>
					<TD ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>

		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtOpenType"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtFrOpenDt"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtToOpenDt"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtDocCur"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"		tag="14" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtDeptCd"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtBpCd"			tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizCd"			tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtBpCd2"			tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtProject"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtAllcAmt"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtRefNo"			tag="14" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtAcctCd"		tag="14" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtGlNo"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtMgntCd1"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtMgntCd2"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtCardCoCd"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtCardNo"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtFrCardUserId"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtToCardUserId"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hchkLocalCur"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtAllcDt"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtAllcAlt"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtParentGlNo"	tag="14" Tabindex="-1">
<INPUT TYPE=hidden NAME="hTblId"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hDataColmID"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hDataColmNm"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hTblId2"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hDataColmID2"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hDataColmNm2"		tag="24" Tabindex="-1">
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

