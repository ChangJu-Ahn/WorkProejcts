<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5120ma1
'*  4. Program Name         : 결의전표조회 
'*  5. Program Desc         : 결의전표조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/01/23
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Lim YOung Woon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		


<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">         </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "AcctCtrl.vbs">							</SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit
	
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID 		= "a5120mb1.asp"                              '☆: Biz Logic ASP Name

'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 3					                          '☆: SpreadSheet의 키의 갯수 
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          
Dim lgMaxFieldCount
Dim lgCookValue 
Dim lgSaveRow 
Dim IsOpenPop 

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


<%                                               
   BaseDate     = GetSvrDate                                                                  'Get DB Server Date
%>

'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0 
End Sub

'========================================================================================================
Sub SetDefaultVal()
	
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
	
	frm1.txtFromGlDt.text	= UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	frm1.txtToglDt.Text		= UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","QA") %>
End Sub

'========================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal

	Const CookieSplit = 4877						

	If Kubun = 0 Then                                              ' Called Area
       strTemp = ReadCookie(CookieSplit)

       If strTemp = "" then Exit Function

       arrVal = Split(strTemp, parent.gRowSep)


       WriteCookie CookieSplit , ""
	
	ElseIf Kubun = 1 then                                         ' If you want to call
		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue		
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End IF

	
End Function


'=========================================================================================================
Sub InitComboBox()
	
	Err.clear
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & "   order by minor_nm ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlInputType ,lgF0  ,lgF1  ,Chr(11))
	 
End Sub

'========================================================================================================
Sub InitSpreadSheet()
    
    Call SetZAdoSpreadSheet("A5120QA1", "S", "A", "V20021108", parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
    Call SetSpreadLock()      
         
End Sub

'========================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029														
	
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables														
	Call SetDefaultVal

	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")										
    Call InitComboBox()
    Call CookiePage(0)
	frm1.txtAmtFr.Text	= ""
	frm1.txtAmtTo.Text	= ""
    frm1.txtFromGlDt.focus


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

'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub


'==========================================================================================
Sub txtDeptCD_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromGlDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtToGlDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
			
		End If
	End IF

End Sub
'==========================================================================================

Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromGlDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromGlDt.Focus
    End If
End Sub
'==========================================================================================

Sub txtFromGlDt_Change() 
    lgBlnFlgChgValue = True
End Sub
'==========================================================================================

Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToGlDt.Focus
    End If
End Sub
'==========================================================================================

Sub txtToGlDt_Change() 
    lgBlnFlgChgValue = True
End Sub
'==========================================================================================

Sub txtFromGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
			frm1.txtToGlDt.focus
			Call FncQuery
	End If
End Sub
'==========================================================================================

Sub txtToGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromGlDt.focus
		Call FncQuery
	End If
End Sub

'========================================================================================================
Function FncQuery() 
	Dim IntRetCD

    FncQuery = False                                            
    
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    
    Call InitVariables
    
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtFromGlDt.text,frm1.txtToGlDt.text,frm1.txtFromGlDt.Alt,frm1.txtToGlDt.Alt, _
                        "970025",frm1.txtFromGlDt.UserDefinedFormat,parent.gComDateType,True) = False Then			
		Exit Function
    End If

	If frm1.txtAmtTo.text <> "" Then
		If UNICDbl(frm1.txtAmtTo.text) < UNICDbl(frm1.txtAmtFr.text) Then
			Call DisplayMsgBox("970023","X",frm1.txtAmtTo.Alt,frm1.txtAmtFr.Alt)                         '☜ : 숫자영 
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")


    If frm1.txtBizArea.value = "" Then
		frm1.txtBizAreaNm.value = ""
    End If
    
    If frm1.txtCOST_CENTER_CD.value = "" Then
		frm1.txtCOST_CENTER_NM.value = ""
    End If
    
    If frm1.txtdeptcd.value = "" Then
		frm1.txtdeptnm.value = ""
    End If
    
    If frm1.txtUsr_Id.value = "" Then
		frm1.txtUsr_Nm.value = ""
    End If

	IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("124600","X","X","X")           '⊙: Display Message(There is no changed data.)
		Exit Function
	End if
    
    '-----------------------
    'Query function call area
    '-----------------------
    IF DbQuery	 = False Then															'☜: Query db data
       Exit Function
    End IF
       
    FncQuery = True												

End Function


'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function DbQuery() 
	Dim strVal

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode = parent.OPMD_CMODE Then   ' This means that it is first search
        
			strVal = strVal & "?txtMode=" & parent.UID_M0001	
			strVal = strVal & "&txtFromGlDt=" & Trim(.txtFromGlDt.text)
			strVal = strVal & "&txtToGlDt=" & Trim(.txtToGlDt.text)						'☜: 
			strVal = strVal & "&txtBizArea=" & UCase(Trim(.txtBizArea.value))
			strVal = strVal & "&txtBizArea1=" & UCase(Trim(.txtBizArea1.value))			
			strVal = strVal & "&txtCOST_CENTER_CD=" & Trim(.txtCOST_CENTER_CD.value)
			strVal = strVal & "&txtdeptcd=" & UCase(Trim(.txtdeptcd.value))				'☆: 조회 조건 데이타 
		    strVal = strVal & "&cboGlInputType=" & Trim(.cboGlInputType.value)
		'	strVal = strVal & "&lgStrPreglno=" & lgStrPreglno		
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&OrgChangeId=" & Trim(.hOrgChangeId.Value)
			strVal = strVal & "&txtDesc=" & Trim(.txtDesc.Value)
			strVal = strVal & "&txtRefNo=" & .txtRefNo.value
			strVal = strVal & "&txtAmtFr=" & .txtAmtFr.text
			strVal = strVal & "&txtAmtTo=" & .txtAmtTo.text
			strVal = strVal & "&txtUsr_ID=" & .txtUsr_ID.VALUE
        Else
            strVal = strVal & "?txtMode=" & parent.UID_M0001	
			strVal = strVal & "&txtFromGlDt=" & Trim(.htxtFromGlDt.value)
			strVal = strVal & "&txtToGlDt=" & Trim(.htxtToGlDt.value)						'☜: 
			strVal = strVal & "&txtBizArea=" & UCase(Trim(.htxtBizArea.value))
			strVal = strVal & "&txtBizArea1=" & UCase(Trim(.htxtBizArea1.value))			
			strVal = strVal & "&txtCOST_CENTER_CD=" & Trim(.htxtCOST_CENTER_CD.value)
			strVal = strVal & "&txtdeptcd=" & UCase(Trim(.htxtdeptcd.value))				'☆: 조회 조건 데이타 
		    strVal = strVal & "&cboGlInputType=" & Trim(.hcboGlInputType.value)
		'	strVal = strVal & "&lgStrPreglno=" & lgStrPreglno		
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&OrgChangeId=" & Trim(.hOrgChangeId.Value)
			strVal = strVal & "&txtDesc=" & Trim(.htxtDesc.Value)
			strVal = strVal & "&txtRefNo=" & .htxtRefNo.value
			strVal = strVal & "&txtAmtFr=" & .htxtAmtFr.value
			strVal = strVal & "&txtAmtTo=" & .htxtAmtTo.value
			strVal = strVal & "&txtUsr_ID=" & .HtxtUsr_ID.VALUE
        End If   

    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&lgAuthorityFlag="   & EnCoding(lgAuthorityFlag)            '권한관리 추가 
		strVal = strVal & "&txtGlNo=" & .txtGlNo.value				


		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

        Call RunMyBizASP(MyBizASP, strVal)							
    End With


    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1
	CALL vspdData_Click(1, 1)
End Function

'==========================================================================================

Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
Dim arrStrRet				'권한관리 추가   							  

dim strgChangeOrgId

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 0
			arrStrRet =  AutorityMakeSql("DEPT",parent.gChangeOrgId, "","","","")	'권한관리 추가   							  
			
			arrParam(0) = "부서코드 팝업"					' 팝업 명칭 
			arrParam(1) = arrstrRet(0)										'권한관리 추가   							  				
			arrParam(2) = UCase(Trim(frm1.txtDeptCd.Value))	' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = arrstrRet(1)										'권한관리 추가   							  

			' 권한관리 추가 
'			If lgInternalCd <>  "" Then
'				arrParam(4) = " INTERNAL_CD =" & FilterVar(lgInternalCd, "''", "S")			' Where Condition
'			Else
'				arrParam(4) = ""
'			End If
'
'			If lgSubInternalCd <>  "" Then
'				arrParam(4) = " INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
'			Else
'				arrParam(4) = ""
'			End If


			arrParam(5) = "부서 코드"			
	
   			arrField(0) = "DEPT_CD"	     				' Field명(0)
			arrField(1) = "DEPT_NM"			    		' Field명(1)
		
			arrHeader(0) = "부서코드"					' Header명(0)
			arrHeader(1) = "부서코드명"				' Header명(1)
    
		Case 1,3
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			' 권한관리 추가 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If
			
			arrParam(5) = "사업장코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"						' Field명(0)
			arrField(1) = "BIZ_AREA_NM"						' Field명(1)
    
			arrHeader(0) = "사업장코드"			' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)
			
		Case 2
			arrParam(0) = "코스트센타 팝업"					' 팝업 명칭 
			arrParam(1) = "B_COST_CENTER"							' TABLE 명칭 
			arrParam(2) = strCode			       				    ' Code Condition
			arrParam(3) = ""										' Name Cindition

			' 권한관리 추가 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "코스트센타"			
	
		    arrField(0) = "COST_CD"									' Field명(0)
			arrField(1) = "COST_NM"									' Field명(1)
    
			arrHeader(0) = "코스트센타코드"					' Header명(0)
			arrHeader(1) = "코스트센타명"						' Header명(1)
			
		Case 4
			arrParam(0) = "전표번호 팝업"					' 팝업 명칭 
			arrParam(1) = "a_gl"							' TABLE 명칭 
			arrParam(2) = strCode			       				    ' Code Condition
			arrParam(3) = ""										' Name Cindition

			arrParam(4) = ""
			arrParam(5) = "전표번호"			
	
		    arrField(0) = "gl_no"									' Field명(0)
			
    
			arrHeader(0) = "전표번호"					' Header명(0)
			
		Case 5
			arrParam(0) = "작성자 팝업"						' 팝업 명칭 
			arrParam(1) = "A_GL A, Z_USR_MAST_REC B"							' TABLE 명칭 
			arrParam(2) = strCode			       				    ' Code Condition
			arrParam(3) = ""										' Name Cindition

			' 권한관리 추가 
			If lgAuthUsrID <>  "" Then
'hanc				arrParam(4) = " A.INSRT_USER_ID*=B.USR_ID AND A.INSRT_USER_ID=" & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
				arrParam(4) = " A.INSRT_USER_ID=B.USR_ID AND A.INSRT_USER_ID=" & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			Else
'hanc				arrParam(4) = " A.INSRT_USER_ID*=B.USR_ID "
				arrParam(4) = " A.INSRT_USER_ID=B.USR_ID "
			End If

			arrParam(5) = "작성자"			
	
		    arrField(0) = "A.INSRT_USER_ID"									' Field명(0)
			arrField(1) = "B.USR_NM"									' Field명(1)
    
			arrHeader(0) = "작성자"					' Header명(0)
			arrHeader(1) = "작성자명"						' Header명(1)				
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If	

	Call FocusAfterPopup (iWhere)

End Function

'==========================================================================================

Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function	
	
	With frm1.vspdData
	   if .maxrows > 0 Then	
		.Row = .ActiveRow
		.Col = 2

	
		arrParam(0) = Trim(.Text)	'회계전표번호 
		arrParam(1) = ""			'Reference번호 
	   End if	
	End With
	
	
'	arrParam(0) =  Trim(GetKeyPosVal("A", 1))	'전표번호 
'	arrParam(1) = ""			      
	IsOpenPop = True 
	
	iCalledAspName = AskPRAspName("a5120ra1")     
	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'==========================================================================================
Function SetPopUp(ByRef arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0	     'DEPT
				.txtdeptcd.value		= UCase(Trim(arrRet(0)))
				.txtdeptNm.value		= arrRet(1)
				
			Case 1		' Biz area
				.txtBizArea.value		= UCase(Trim(arrRet(0)))
				.txtBizAreaNm.value		= arrRet(1)
				
			Case 2
				.txtCOST_CENTER_CD.value = arrRet(0)
				.txtCOST_CENTER_NM.value = arrRet(1)
			Case 3		' Biz area
				.txtBizArea1.value		= UCase(Trim(arrRet(0)))
				.txtBizAreaNm1.value		= arrRet(1)
				
			Case 4		' Biz area
				.txtGlNo.value		= UCase(Trim(arrRet(0)))
			Case 5
				.txtUsr_ID.value = arrRet(0)
				.txtUsr_NM.value = arrRet(1)		
		End Select

		'lgBlnFlgChgValue = True
	End With
End Function
'=======================================================================================================
Function FocusAfterPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtdeptcd.focus
			Case 1  
				.txtBizArea.focus
			Case 2 
				.txtCOST_CENTER_CD.focus
		End Select    
	End With

End Function

'==========================================================================================
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtFromGlDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToGlDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  
	
	' 권한관리 추가 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID


	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus		
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'==========================================================================================

Function SetDept(Byval arrRet)
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtFromGlDt.text = arrRet(4)
		frm1.txtToGlDt.text = arrRet(5)
		frm1.txtDeptCd.focus
End Function

'========================================================================================================
Function OpenOrderPopup(ByVal pSpdNo)

	Dim arrRet	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function

'========================================================================================================
Sub PopZAdoConfigGrid()
	
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then	
		Exit Sub
	End If		
	Call OpenOrderPopup("A")
	
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function
	
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
		
    Call SetPopUpMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If  
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
	
	If Col < 1 Then Exit Sub
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)	 	
End Sub
	
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then
	'	If lgStrPreglno <> "" Then
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub
'========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtFromEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtFromEnterDt.Focus
	End If
End Sub
'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtToEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtToEnterDt.Focus
	End If
End Sub

'========================================================================================================
Sub fpdtFromEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
Sub fpdtToEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'==========================================================================================
Sub txtAmtFr_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call fncQuery()
    End if
End Sub
'==========================================================================================
Sub txtAmtTo_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call fncQuery()
    End if
End Sub
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtFromGlDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtToGlDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
					'msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></td>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>회계일자</TD>	                                                  
						            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name="txtFromGlDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일자" tag="12" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
						                                 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name="txtToGlDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일자" tag="12" id=fpDateTime2></OBJECT>');</SCRIPT></TD>								
						            <TD CLASS=TD5 NOWRAP>부서코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtdeptcd" ALT="부서코드" Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
														 <INPUT NAME="txtdeptnm" ALT="부서명"   Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag="14N"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>사업장</TD>										
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizArea"   ALT="사업장"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizArea.Value, 1)">
														 <INPUT NAME="txtBizAreaNm" ALT="사업장명" Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N">&nbsp;~</TD>
									<TD CLASS=TD5 NOWRAP>금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAmtFr" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="11XXXX" ALT="시작금액" id=OBJECT1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAmtTo" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="11XXXX" ALT="종료금액" id=OBJECT2></OBJECT>');</SCRIPT>
										 </TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>										
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizArea1"   ALT="사업장"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizArea1.Value, 3)">
														 <INPUT NAME="txtBizAreaNm1" ALT="사업장명" Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
									<TD CLASS=TD5 NOWRAP>전표번호</TD>										
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGlNo"   ALT="전표번호"   Size="12" MAXLENGTH="18" STYLE="TEXT-ALIGN: left" tag   ="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtGlNo.Value, 4)"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>코스트센타</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCOST_CENTER_CD" MAXLENGTH="10" SIZE=12 ALT ="코스트센타 코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="Call OpenPopup(frm1.txtCOST_CENTER_CD.value, 2)">
														 <INPUT NAME="txtCOST_CENTER_NM" MAXLENGTH="20" SIZE=24 STYLE="TEXT-ALIGN:left" ALT ="코스트센타명" tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="1N"STYLE="WIDTH:82px:"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								 <TR>
									<TD CLASS=TD5 NOWRAP>작성자</TD>
									<TD CLASS=TD6 NOWRAP> <INPUT NAME="txtUsr_ID" MAXLENGTH="12" SIZE=12 ALT ="작성자" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="Call OpenPopup(frm1.txtUsr_ID.value, 5)">
														  <INPUT NAME="txtUsr_NM" MAXLENGTH="20" SIZE=24 STYLE="TEXT-ALIGN:left" ALT ="작성자명" tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>참조번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="참조번호" MAXLENGTH="30" SIZE="20" tag="11XXXU" ></TD>				
								 </TR>
								 <TR>
									<TD CLASS=TD5 NOWRAP>비고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="35" tag="11" ></TD>
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
				<TR>
					<TD WIDTH="100%" HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" COLSPAN ="2">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> Title="Spread" height="100%" id=vaSpread1 name=vspdData width="100%" tag="23"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD18 HEIGHT=20 NOWRAP>차대합계</TD>
								<TD CLASS=TD19>
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrlocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계(거래)" id=fpDoubleSingle1></OBJECT>');</SCRIPT>													
									&nbsp;
									&nbsp;						
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="대변합계(자국)" tag="24X2" id=fpDoubleSingle4></OBJECT>');</SCRIPT>
								</TD>
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
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1"></IFRAME></TD>		
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtFromGlDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtToGlDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizArea" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizArea1" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtCOST_CENTER_CD" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtdeptcd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hcboGlInputType" tag="24"TABINDEX="-1"> 
<INPUT TYPE=HIDDEN NAME="horgchangeid" tag="14"TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtglno" tag="24"TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtDesc" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtRefNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtAmtFr" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtAmtTo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtUSr_Id" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
