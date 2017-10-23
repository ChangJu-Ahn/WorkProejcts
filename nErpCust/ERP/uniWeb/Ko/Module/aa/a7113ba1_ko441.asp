<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Asset Management
'*  3. Program ID           : a7113ba1_ko441.asp
'*  4. Program Name         : 감가상각전표처리(site)
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             AS0052
'                             
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2000/03/05
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   ************************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->

<!--========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit  
 '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_PGM_ID = "a7113bb1_ko441.asp"  
 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
Dim lgAnswer
Dim srtToday

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 



Function OpenDeptCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg   	


	If IsOpenPop = True Or UCase(frm1.txtDeptCd.className) = "PROTECTED" Then Exit Function
	IsOpenPop = True

	arrParam(0) = "전표생성부서 팝업"	
	arrParam(1) = "  ( SELECT h.DEPT_CD, h.DEPT_NM FROM B_ACCT_DEPT h(nolock) join B_COST_CENTER i(nolock) on h.cost_cd = i.cost_cd" & vbcr
	arrParam(1) = 	arrParam(1) & " WHERE h.ORG_CHANGE_ID =(select distinct org_change_id " & vbcr
	arrParam(1) = 	arrParam(1) & " from b_acct_dept where org_change_dt = ( select max(org_change_dt) " & vbcr
	arrParam(1) = 	arrParam(1) & " from b_acct_dept where org_change_dt <=  " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & ")) " & vbcr
	If Trim(frm1.txtBizAreaCd.value) <> "" then
		arrParam(1) = 	arrParam(1) & " AND i.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd.value, "''", "S") & vbcr
	end if	

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(1) = 	arrParam(1) & " AND i.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	End If

	If lgInternalCd <> "" Then
		arrParam(1) = 	arrParam(1) & " AND h.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
	End If

	If lgSubInternalCd <> "" Then
		arrParam(1) = 	arrParam(1) & " AND h.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
	End If

	arrParam(1) = 	arrParam(1) & " ) A " & vbcr


	'arrParam(1) = "B_ACCT_DEPT A"				
	arrParam(2) = Trim(frm1.txtDeptCd.value)
	arrParam(3) = "" 
	'arrParam(4) = "A.ORG_CHANGE_ID = '" & parent.gChangeOrgId & "'"
	arrParam(5) = "부서코드"			
	
    arrField(0) = "A.DEPT_CD"	
    arrField(1) = "A.DEPT_Nm"
    
    arrHeader(0) = "부서코드"		
    arrHeader(1) = "부서코드명"		
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		field_fg = 8
		Call SetReturnVal(arrRet,field_fg)
	End If	
End Function

'===========================================================================
' Function Name : OpenBizAreaCd
' Function Desc : OpenDeptCode Reference Popup
'===========================================================================
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg  

	If IsOpenPop = True Or UCase(frm1.txtBizAreaCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"						' TABLE 명칭 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name COndition

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

	arrParam(5) = "사업장"			
	
    arrField(0) = "BIZ_AREA_CD"						' Field명(0)
    arrField(1) = "BIZ_AREA_NM"						' Field명(1)
    
    arrHeader(0) = "사업장코드"					' Header명(0)
    arrHeader(1) = "사업장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
	    Exit Function
	Else
		field_fg = 7
		Call SetReturnVal(arrRet,field_fg)
	End If	

End Function

 '------------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg
			case 8  'OpenDept
			.txtDeptCd.focus
			.txtDeptCd.value  = Trim(arrRet(0))
			.txtDeptNm.value  = arrRet(1)
			Call txtDeptCd_OnChange
			case 7   'OpenBizAreaCd
			.txtBizAreaCd.focus
			.txtBizAreaCd.value  = Trim(arrRet(0))
			.txtBizAreaNm.value  = arrRet(1)
			Call txtBizAreaCd_onchange
		End select	
	End With
	
End Function



 '------------------------------------------  fnButtonExec()  --------------------------------------------------
'	Name : fnButtonExec()
'	Description : 전표처리 수행부 
'--------------------------------------------------------------------------------------------------------- 

Function fnButtonExec()
    Dim strVal       		
	Dim strFrdt,strTodt
	Dim strWkDt
	Dim strDeptCd
	Dim RetFlag
	Dim strYear
	Dim strMonth
	Dim strDay
	Dim strBizAreaCd

    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then        '⊙: Check contents area
       Exit Function
    End If

	RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
	'RetFlag = Msgbox("작업을 수행 하시겠습니까?", vbOKOnly + vbInformation, "정보")

	If RetFlag = VBNO Then
		Exit Function
	End IF  
	
    Err.Clear    	
    
    Call LayerShowHide(1) 

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002							'☜: 비지니스 처리 ASP의 상태 
    
    if frm1.Rb_WK1.checked = true then
		strVal = strVal & "&txtRadio=" & "1"								'☆: 조회 조건 데이타 
		strVal  = strVal & "&txtOrgChangeId=" & Trim(frm1.hOrgChangeId.value) 
    else
		strVal = strVal & "&txtRadio=" & "2"								'☆: 조회 조건 데이타 
		strVal  = strVal & "&txtOrgChangeId=" & parent.gChangeOrgId
	end if

    if frm1.Rb_WK3.checked = true then
		strVal = strVal & "&txtRadio2=" & "1"								'☆: 조회 조건 데이타 
    else
		strVal = strVal & "&txtRadio2=" & "2"								'☆: 조회 조건 데이타 
	end if      
	
	
	Call ExtractDateFrom(frm1.txtFrdt.Text,frm1.txtFrdt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

    strFrDt	  = strYear & strMonth
	strWkDt	  =  UniConvDateToYYYYMMDD(frm1.txtGLdt.text, gDateFormat, parent.gServerDateType)
	strDeptCd = frm1.txtDeptCd.value 
	strBizAreaCd = frm1.txtBizAreaCd.value
				
    strVal  = strVal & "&txtFrdt=" & strFrdt
    strVal  = strVal & "&txtGLdt=" & strWkDt       
    strVal  = strVal & "&txtDeptCd=" & strDeptCd
    strVal  = strVal & "&txtBizAreaCd=" & strBizAreaCd

    'strVal = strVal & "&txtStatus=" & "confirm" 
	'gAnswer = "confirm"    	       

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

end function	

Function fnButtonExecOk()
    Dim IntRetCD 

    IntRetCD = DisplayMsgBox("990000","X","X","X")   '☜ 바뀐부분    
	   '''Msgbox "정상처리되었습니다."
End Function
 '#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################

'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	Dim lgF2By2


	If Trim(frm1.txtGldt.Text = "") Then    
		Exit sub
    End If
    IsOpenPop = True

		'----------------------------------------------------------------------------------------


		strSelect	=			 " dept_cd, org_change_id, internal_cd "
		strFrom		=			 " b_acct_dept(NOLOCK) "
		strWhere	=  " COST_CD IN ( "
		strWhere	= strWhere & " 	SELECT COST_CD "
		strWhere	= strWhere & " 	FROM B_COST_CENTER "
		strWhere	= strWhere & " 	WHERE BIZ_AREA_CD=" & FilterVar(frm1.txtBizAreaCd.value, "''", "S")
		strWhere	= strWhere & " 		) "
		strWhere	= strWhere & " 		AND ORG_CHANGE_ID =(select distinct org_change_id "
		strWhere	= strWhere & " 			 from b_acct_dept where org_change_dt = ( select max(org_change_dt) "
		strWhere	= strWhere & " 				 from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & ")) "
		strWhere	= strWhere  & " And dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

'msgbox "Select  " & strSelect  & " From " & strFrom & " Where " & strWhere
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  

			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
			jj = Ubound(arrVal1,1)
			For ii = 0 to jj - 1

				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
		End If
	    IsOpenPop = False
		frm1.txtDeptCd.focus

		'----------------------------------------------------------------------------------------

End Sub


'========================================================================================
' Function Name : FncSave
' Function Desc : 
'========================================================================================


Function FncSave()
End Function

Function FncQuery()
End Function

Function FncPrint() 
    Call parent.FncPrint()                                              
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
    
End Function

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "BA") %>
End Sub

 '******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)       
    Call ggoOper.FormatDate(frm1.txtFrDT, gDateFormat, 2)
    Call ggoOper.LockField(Document, "N")									'⊙: Lock  Suitable  Field
    srtToday = "<% = GetSvrDate %>"
    frm1.txtFrdt.focus 
	frm1.fpDateTime1.Text = UNIMonthClientFormat(srtToday)
	frm1.fpDateTime3.Text = UNIDateClientFormat(srtToday)
    Call SetToolbar("10000000000011")
	Call radio3_onchange()
	frm1.hOrgChangeId.value = parent.gChangeOrgId

	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing
	
End Sub


'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'======================================================================================================
'   Event Name : DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtfrdt_DblClick(Button)
    If Button = 1 Then
       frm1.txtfrdt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtfrdt.Focus       
    End If
End Sub
'======================================================================================================
'   Event Name : DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================

Sub txtgldt_DblClick(Button)
    If Button = 1 Then
       frm1.txtgldt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtgldt.Focus
    End If
End Sub


Function radio1_onchange()
	IF frm1.Rb_WK4.checked =  True Then
	ggoOper.SetReqAttr frm1.txtGldt,		 "N"    '전표생성일자 
	ggoOper.SetReqAttr frm1.txtDeptCd,		 "N"    '부서 
	End If
End Function
Function radio2_onchange()
	IF frm1.Rb_WK4.checked =  True Then
		ggoOper.SetReqAttr frm1.txtGldt,		 "Q"    '전표생성일자 
		ggoOper.SetReqAttr frm1.txtDeptCd,		 "Q"    '부서 
		ggoOper.SetReqAttr frm1.txtBizAreaCd,		 "N"    '사업장 
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
	End If
	'ggoOper.SetReqAttr frm1.txtGldt,		 "Q"    '전표생성일자 
	'ggoOper.SetReqAttr frm1.txtDeptCd,		 "Q"    '부서 
	'frm1.txtDeptCd.value = ""
	'frm1.txtDeptNm.value = ""
	'ggoOper.SetReqAttr frm1.chkShowDt1,		 "Q"    '전표생성일자 
	'ggoOper.SetReqAttr frm1.chkShowDt2,		 "Q"    '전표생성일자 
	'frm1.chkShowDt1.Checked = False
	'frm1.chkShowDt2.Checked = False

End Function
Function radio3_onchange()
	ggoOper.SetReqAttr frm1.txtGldt,		 "Q"    '전표생성일자 
	ggoOper.SetReqAttr frm1.txtDeptCd,		 "Q"    '부서 
	ggoOper.SetReqAttr frm1.txtBizAreaCd,		 "Q"    '사업장 
	frm1.txtDeptCd.value = ""
	frm1.txtDeptNm.value = ""
	frm1.txtBizAreaCd.value = ""
	frm1.txtBizAreaNm.value = ""
End Function
Function radio4_onchange()
	IF frm1.Rb_WK1.checked =  True Then
		ggoOper.SetReqAttr frm1.txtGldt,		 "N"    '전표생성일자 
		ggoOper.SetReqAttr frm1.txtDeptCd,		 "N"    '부서 
		ggoOper.SetReqAttr frm1.txtBizAreaCd,		 "N"    '사업장 
	Else
		ggoOper.SetReqAttr frm1.txtGldt,		 "Q"    '전표생성일자 
		ggoOper.SetReqAttr frm1.txtDeptCd,		 "Q"    '부서 
		ggoOper.SetReqAttr frm1.txtBizAreaCd,		 "N"    '사업장 
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.txtBizAreaCd.value = ""
		frm1.txtBizAreaNm.value = ""
	End If


End Function
'========================================================================================================
' Name : chkShowBp_onchange
' Desc : 
'========================================================================================================
Sub chkShowDt_onchange()
	If frm1.chkShowDt.checked = True Then
		frm1.txtShowDt.value = "Y"
	Else
		frm1.txtShowDt.value = "N"	
	End If
End Sub
Sub txtBizAreaCd_onchange()
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	Dim lgF2By2


	If Trim(frm1.txtGldt.Text = "") Or Trim(frm1.txtDeptCd.value) = "" Then    
		Exit sub
	End If

    IsOpenPop = True


		'----------------------------------------------------------------------------------------
		strSelect	=			 " A.DEPT_CD ,A.DEPT_NM  "
		strFrom		=			 " ( SELECT DEPT_CD ,DEPT_NM FROM B_ACCT_DEPT    "
		strFrom		= strFrom & " WHERE COST_CD IN (  SELECT COST_CD  FROM B_COST_CENTER    "
		strFrom		= strFrom & " WHERE BIZ_AREA_CD= " & FilterVar(frm1.txtBizAreaCd.value, "''", "S") & " AND ORG_CHANGE_ID =(select distinct org_change_id  from b_acct_dept where org_change_dt = ( select max(org_change_dt)   "
		strFrom		= strFrom & " from b_acct_dept where org_change_dt <=  " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & ")))) A "

		If UCase(frm1.txtDeptCd.className) <> "PROTECTED" Then 
			strFrom		= strFrom & " Where  A.DEPT_CD = " & FilterVar(frm1.txtDeptCd.value, "''", "S") 
		End If

		'msgbox "Select  " & strSelect  & " From " & strFrom & " Where " & strWhere
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
'			frm1.txtBizAreaCd.value = ""
'			frm1.txtBizAreaNm.value = ""
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
'			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
'			jj = Ubound(arrVal1,1)
'			For ii = 0 to jj - 1
'
'				arrVal2 = Split(arrVal1(ii), chr(11))
'				frm1.hOrgChangeId.value = Trim(arrVal2(2))
'			Next	
		End If
	    IsOpenPop = False

	'Call txtDeptCd_OnChange
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio2 ID=Rb_WK3 Checked onclick=radio3_onchange()><LABEL FOR=Rb_WK3>부서배분</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio2 ID=Rb_WK4 onclick=radio4_onchange()><LABEL FOR=Rb_WK4>전표처리</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업구분</TD>
								<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 Checked onclick=radio1_onchange()><LABEL FOR=Rb_WK1>생성</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK2 onclick=radio2_onchange()><LABEL FOR=Rb_WK2>취소</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrdt" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT=작업년월 id=fpDateTime1> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=13 MAXLENGTH=10 tag="12XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo2" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd.value,0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>전표일자</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtGldt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT=전표일자 id=fpDateTime3> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>생성부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=13 MAXLENGTH=10 tag="12XXXU" ALT="생성부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript: OpenDeptCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=25 tag="14"></TD>
							</TR>
						</TABLE>
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
					<TD><BUTTON NAME="btn배치" CLASS="CLSMBTN" OnClick="VBScript:Call fnButtonExec()" Flag=1>실 행</BUTTON> &nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>

		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
            <IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

