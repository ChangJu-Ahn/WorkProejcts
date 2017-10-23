
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5118ma1
'*  4. Program Name         : 결의전표출력 
'*  5. Program Desc         : 결의전표출력 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/12/08
'*  8. Modified date(Last)  : 2004/01/12
'*  9. Modifier (First)     : 송문길 
'* 10. Modifier (Last)      : Kim Chang Jin
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

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->


<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE = "VBScript"		SRC = "../../inc/incCliMAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript"		SRC = "../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript"		SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript"		SRC = "../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript"		SRC = "../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript"		SRC = "../../inc/incEB.vbs">				</SCRIPT>


<SCRIPT LANGUAGE="VBScript">

Option Explicit	


'========================================================================================================= 

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag 
Dim lgIntFlgMode               ' Variable is for Operation Status 
Dim lgF2By2

Dim IsOpenPop

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 



'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

End Sub

'========================================================================================================= 

Sub SetDefaultVal()	

	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
	frm1.txtDateFr.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	frm1.txtDateTo.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
End Sub
'=======================================================================================

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","OA") %>
<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","OA") %>
End Sub

'=======================================================================================
Function OpenPopUp(strCode, iWhere)
	Dim arrRet
	Dim arrParam(8), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 0, 1
			arrParam(0) = frm1.txtDateFr.Text
			arrParam(1) = frm1.txtDateTo.Text

			' 권한관리 추가 
			arrParam(5)	= lgAuthBizAreaCd
			arrParam(6)	= lgInternalCd
			arrParam(7)	= lgSubInternalCd
			arrParam(8)	= lgAuthUsrID
	
'		Case 2
'			arrParam(0) = "부서코드 팝업"								' 팝업 명칭 
'			arrParam(1) = "B_ACCT_DEPT"    									' TABLE 명칭 
'			arrParam(2) = strCode											' Code Condition
'			arrParam(3) = ""												' Name Cindition
'			arrParam(4) = "ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value, "''", "S") & ""		' Where Condition
'			arrParam(5) = "부서코드"									' 조건필드의 라벨 명칭 
'
'			arrField(0) = "DEPT_CD"	     									' Field명(0)
'			arrField(1) = "DEPT_NM"			    							' Field명(1)
'
'			arrHeader(0) = "부서코드"									' Header명(0)
'			arrHeader(1) = "부서명"										' Header명(1)
			
		Case 3, 5
			arrParam(0) = "사업장코드 팝업"								' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 										' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition

			' 권한관리 추가 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "사업장코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"										' Field명(0)
			arrField(1) = "BIZ_AREA_NM"										' Field명(1)
    
			arrHeader(0) = "사업장코드"									' Header명(0)
			arrHeader(1) = "사업장명"									' Header명(1)

		Case 4
			arrParam(0) = "전표입력경로팝업"								' 팝업 명칭 
			arrParam(1) = "B_MINOR" 										' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("A1001", "''", "S") & " "												' Where Condition
			arrParam(5) = "전표입력경로코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "MINOR_CD"										' Field명(0)
			arrField(1) = "MINOR_NM"										' Field명(1)
    
			arrHeader(0) = "전표입력경로코드"									' Header명(0)
			arrHeader(1) = "전표입력경로명"									' Header명(1)			
			
			
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

    Select Case iWhere
		Case 0, 1
			arrRet = window.showModalDialog("a5101ra1.asp", Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		Case Else
			arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select
		
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetReturnPopUp(arrRet, iWhere)
	End If	
	
	Call EscPopup(iWhere)	

End Function
'=======================================================================================

Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtTempGlNoFr.focus
			Case 1
				.txtTempGlNoTo.focus
'			Case 2
'				.txtDeptCd.focus
			Case 3
				.txtBizAreaCd.focus
			Case 4
				.txtInputType.focus
			Case 5
				.txtBizAreaCd.focus
		End Select
	End With
	
End Function

'=======================================================================================

Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	arrParam(0) = frm1.txtDateFr.text								'  Code Condition
   	arrParam(1) = frm1.txtDateTo.Text
'	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
'	arrParam(4) = "F"									' 결의일자 상태 Condition  
	

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

'=======================================================================================

Function SetDept(Byval arrRet)
		
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtDateFr.text = arrRet(4)
		frm1.txtDateTo.text = arrRet(5)
		frm1.txtDeptCd.focus

End Function


'=======================================================================================

Function SetReturnPopUp(ByVal arrRet, ByVal iWhere)
	
	Select Case iWhere
		Case 0		'회계전표번호 
			frm1.txtTempGlNoFr.value = UCase(Trim(arrRet(0)))
		Case 1		'회계전표번호 
			frm1.txtTempGlNoTo.value = UCase(Trim(arrRet(0)))
		Case 2		'부서코드 
			frm1.txtDeptCd.value = UCase(Trim(arrRet(0)))
			frm1.txtDeptNm.value = arrRet(1)
		Case 3		'사업장코드 
			frm1.txtBizAreaCd.value = UCase(Trim(arrRet(0)))
			frm1.txtBizAreaNm.value = arrRet(1)
		Case 4		'입력경로 
			frm1.txtInputType.value = UCase(Trim(arrRet(0)))
			frm1.txtInputTypeNm.value = arrRet(1)
		Case 5		'사업장코드 
			frm1.txtBizAreaCd1.value = UCase(Trim(arrRet(0)))
			frm1.txtBizAreaNm1.value = arrRet(1)							
		Case Else
	End select	

End Function

'=======================================================================================

Sub Form_Load()

    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format

    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call InitVariables                            '⊙: Initializes local global Variables
    Call SetDefaultVal
    Call SetToolbar("10000000000011")				'⊙: 버튼 툴바 제어 

	frm1.txtDeptCd.focus 

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

'=======================================================================================
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
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtDateFr.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtDateTo.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
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
		'----------------------------------------------------------------------------------------

End Sub

'=======================================================================================================
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
        Call SetFocusToDocument("M")
        frm1.fpDateTime1.focus

    End If
End Sub
'=======================================================================================

Sub txtDateTo_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime2.Action = 7
        Call SetFocusToDocument("M")
        frm1.fpDateTime2.focus
    End If
End Sub


'=======================================================================================================
Sub SetPrintCond(StrEbrFile, StrUrl)

    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarBizAreaCd1, VarTempGlNoFr, VarTempGlNoTo, varGlPutType

	Dim	strAuthCond

	StrEbrFile = "a5118ma1_1"
	
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, parent.gDateFormat, parent.gServerDateType)	
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, parent.gDateFormat, parent.gServerDateType)	
	VarDeptCd    = "%"
	
	If frm1.txtBizAreaCd.value = "" then
		frm1.txtBizAreaNm.value = ""
		VarBizAreaCd = " "
	else 
		VarBizAreaCd = FilterVar(frm1.txtBizAreaCD.value,"","SNM")
	end if
	
	If frm1.txtBizAreaCd1.value = "" then
		frm1.txtBizAreaNm1.value = ""
		VarBizAreaCd1 = "ZZZZZZZZZZ"
	else 
		VarBizAreaCd1 = FilterVar(frm1.txtBizAreaCD1.value,"","SNM")
	end if
	
	VarTempGlNoFr = " "
	VarTempGlNoTo = "ZZZZZZZZZZZZZZZZZZ"

	If Trim(frm1.txtDeptCd.value)		<> "" Then VarDeptCd		= UCase(Trim(frm1.txtDeptCd.value))
	If Trim(frm1.txtTempGlNoFr.value)	<> "" Then VarTempGlNoFr	= UCase(Trim(frm1.txtTempGlNoFr.value))
	If Trim(frm1.txtTempGlNoTo.value)	<> "" Then VarTempGlNoTo	= UCase(Trim(frm1.txtTempGlNoTo.value))
	If Trim(frm1.txtInputType.value)	<> "" Then 
		varGlPutType	= UCase(Trim(frm1.txtInputType.value))	
	Else
		varGlPutType	= "%"
		frm1.txtInputTypeNm.value	= ""
	End IF


	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_TEMP_GL.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_TEMP_GL.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_TEMP_GL.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_TEMP_GL.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	
	StrUrl = StrUrl & "DateFr|"			& VarDateFr
	StrUrl = StrUrl & "|DateTo|"		& VarDateTo
	StrUrl = StrUrl & "|DeptCd|"		& VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|"		& VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|"	& VarBizAreaCd1
	StrUrl = StrUrl & "|TempGlNoFr|"	& VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|"	& VarTempGlNoTo
	StrUrl = StrUrl & "|GlPutType|"		& varGlPutType	
	StrUrl = StrUrl & "|OrgChangeId|"	& parent.gChangeOrgId
	
	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond
	'StrUrl = StrUrl & "|gUsrId|" & parent.gUsrNM    
	StrUrl = StrUrl & "|gUsrId|" & parent.gUsrId      	'>>air
	StrUrl = StrUrl & "|LoginDeptNm|" & parent.gDepart	'>>air	
	
End Sub


'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
    Dim StrEbrFile
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	'----------------------------------------------
	'결의일자 Check
	'----------------------------------------------
	If CompareDateByFormat(frm1.txtDateFr.text,frm1.txtDateTo.text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
                        "970025",frm1.txtDateFr.UserDefinedFormat,parent.gComDateType,True) = False Then		
		frm1.txtDateFr.focus                        	
		Exit Function
	End If
	
	'----------------------------------------------
	'결의번호 Check
	'----------------------------------------------
	frm1.txtTempGlNoFr.value = UCase(Trim(frm1.txtTempGlNoFr.value))
	frm1.txtTempGlNoTo.value = UCase(Trim(frm1.txtTempGlNoTo.value))
	
	If frm1.txtTempGlNoFr.value <> "" And frm1.txtTempGlNoTo.value <> "" Then
		If frm1.txtTempGlNoFr.value > frm1.txtTempGlNoTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtTempGlNoFr.Alt, frm1.txtTempGlNoTo.Alt)
			frm1.txtTempGlNoFr.focus 
			Exit Function
		End If
	End If

	Call SetPrintCond(StrEbrFile, StrUrl)

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
    
	Call FncEBRPrint(EBAction,ObjName,StrUrl)
	
End Function

'========================================================================================
Function FncBtnPreview() 
	Dim strUrl
    Dim StrEbrFile
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	'----------------------------------------------
	'결의일자 Check
	'----------------------------------------------
	If CompareDateByFormat(frm1.txtDateFr.text,frm1.txtDateTo.text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
                        "970025",frm1.txtDateFr.UserDefinedFormat,parent.gComDateType,True) = False Then		
		frm1.txtDateFr.focus                        	
		Exit Function
	End If
	
	'----------------------------------------------
	'결의번호 Check
	'----------------------------------------------
	frm1.txtTempGlNoFr.value = UCase(Trim(frm1.txtTempGlNoFr.value))
	frm1.txtTempGlNoTo.value = UCase(Trim(frm1.txtTempGlNoTo.value))
	
	If frm1.txtTempGlNoFr.value <> "" And frm1.txtTempGlNoTo.value <> "" Then
		If frm1.txtTempGlNoFr.value > frm1.txtTempGlNoTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtTempGlNoFr.Alt, frm1.txtTempGlNoTo.Alt)
			frm1.txtTempGlNoFr.focus 
			Exit Function
		End If
	End If


	Call SetPrintCond(StrEbrFile, StrUrl)

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
    
	Call FncEBRPreview(ObjName,StrUrl)
		
End Function

'=======================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
End Function

'========================================================================================

Function FncExit()
    FncExit = True
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
								<TD CLASS="TD5" NOWRAP>결의일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateFr" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT="시작결의일자" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
													   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateTo" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT="종료결의일자" id=fpDateTime2></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>입력부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="입력부서코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()"> 
													   <INPUT TYPE="Text" NAME="txtDeptNm" SIZE=25 tag="14X" ALT="입력부서명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,3)"> 
													   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명">&nbsp;~</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,5)"> 
													   <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>전표입력경로</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtInputType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="전표입력경로코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtInputType.Value,4)">
								                       <INPUT TYPE="Text" NAME="txtInputTypeNm" SIZE=25 tag="14X" ALT="전표입력경로명"></TD>
							</TR>														
							<TR>
								<TD CLASS="TD5" NOWRAP>결의번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtTempGlNoFr" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="시작결의번호" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNoFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtTempGlNoFr.Value,0)">&nbsp;~&nbsp;
													   <INPUT TYPE="Text" NAME="txtTempGlNoTo" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="종료결의번호" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNoTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtTempGlNoTo.Value,1)">
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">	
</FORM>
</BODY>
</HTML>	
