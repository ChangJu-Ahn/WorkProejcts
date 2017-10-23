
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5122oa1_ko441
'*  4. Program Name         : 회계전표집계표출력 
'*  5. Program Desc         : 회계전표집계표출력 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/12/12
'*  8. Modified date(Last)  : 2004/01/12
'*  9. Modifier (First)     : 안혜진 
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

Option Explicit																'☜: indicates that All variables must be declared in advance


'******************************************  1.2 Global 변수/상수 선언  ***********************************

Dim lgBlnFlgChgValue 
Dim lgIntFlgMode     

Dim IsOpenPop

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE  
    lgBlnFlgChgValue = False         


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

'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","OA") %>
<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","OA") %>
End Sub


'==========================================================================================
Function OpenRefGl(iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(8)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5104ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	IsOpenPop = True

	arrParam(0) = frm1.txtDateFr.Text
	arrParam(1) = frm1.txtDateTo.Text
	
	' 권한관리 추가 
	arrParam(5)	= lgAuthBizAreaCd
	arrParam(6)	= lgInternalCd
	arrParam(7)	= lgSubInternalCd
	arrParam(8)	= lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		     
	IsOpenPop = False
	
	If arrRet(0) <> ""  Then			
		Select Case iWhere
		Case 0		'회계전표번호 
			frm1.txtGlNoFr.value = UCase(Trim(arrRet(0)))
		Case 1		'회계전표번호 
			frm1.txtGlNoTo.value = UCase(Trim(arrRet(0)))
		End Select
	End If

	Call EscPopUp( iWhere)
	
End Function


'==========================================================================================
'   Event Name : OpenPopUp
'   Event Desc :
'==========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
'		Case 0, 1
'			arrParam(0) = "회계전표 팝업"				' 팝업 명칭 
'			arrParam(1) = "A_Gl a, B_ACCT_DEPT b"		' TABLE 명칭 
'			arrParam(2) = strCode							' Code Condition
'			arrParam(3) = ""								' Name Cindition
'			arrParam(4) = "a.DEPT_CD=b.DEPT_CD and b.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value, "''", "S") & ""	' Where Condition
'			arrParam(5) = "회계전표번호"				' 조건필드의 라벨 명칭 
'
'			arrField(0) = "a.Gl_No"									' Field명(0)
'			arrField(1) = "DD" & parent.gColSep & "a.gl_dt"									' Field명(1)
'			arrField(2) = "b.DEPT_NM"								' Field명(2)
'			arrField(3) = "F3" & parent.gColSep & "a.cr_Amt"						   		' Field명(3)
'
'			arrHeader(0) = "회계전표번호"								' Header명(0)
'			arrHeader(1) = "회계일자"									' Header명(1)
'			arrHeader(2) = "부서명"										' Header명(0)
'			arrHeader(3) = "발생금액"									' Header명(1)
			
		Case 3, 4
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
		
			
		Case 5
			arrParam(0) = "승인자 팝업"						' 팝업 명칭 
			arrParam(1) = "A_GL A JOIN Z_USR_MAST_REC B ON  A.UPDT_USER_ID= B.USR_ID"							' TABLE 명칭 
			arrParam(2) = strCode			       				    ' Code Condition
			arrParam(3) = ""										' Name Cindition

            		arrParam(4) ="1=1"
			arrParam(5) = "승인자"			
	
		    arrField(0) = "A.UPDT_USER_ID"									' Field명(0)
			arrField(1) = "B.USR_NM"									' Field명(1)
    
			arrHeader(0) = "승인자"					' Header명(0)
			arrHeader(1) = "승인자명"						' Header명(1)			
			
						
		Case Else
			Exit Function
	End Select
    
    Select Case iWhere
	Case 0, 1
'		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
'		 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetReturnPopUp(arrRet, iWhere)
	End If	
	Call EscPopUp( iWhere)
End Function

'=======================================================================================

Function SetReturnPopUp(ByRef arrRet, ByVal iWhere)
	
	Select Case iWhere
'		Case 0		'결의전표번호 
'			frm1.txtGlNoFr.value = UCase(Trim(arrRet(0)))
'		Case 1		'결의전표번호 
'			frm1.txtGlNoTo.value = UCase(Trim(arrRet(0)))
'		Case 2		'부서코드 
'			frm1.txtDeptCd.value = UCase(Trim(arrRet(0)))
'			frm1.txtDeptNm.value = arrRet(1)
		Case 3		'사업장코드 
			frm1.txtBizAreaCd.value = UCase(Trim(arrRet(0)))
			frm1.txtBizAreaNm.value = arrRet(1)
		Case 4		'사업장코드 
			frm1.txtBizAreaCd1.value = UCase(Trim(arrRet(0)))
			frm1.txtBizAreaNm1.value = arrRet(1)
			
		Case 5		'승인자 
			frm1.txtusrid.value = UCase(Trim(arrRet(0)))
			frm1.txtusridnm.value = arrRet(1)
			
		Case Else
	End select	

End Function

'=======================================================================================

Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtGlNoFr.focus
			Case 1
				.txtGlNoTo.focus
'			Case 2
'				.txtDeptCd.focus
'			Case 3
'				.txtBizAreaCd.focus
'			Case 4
'				.txtBizAreaCd1.focus
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
	'arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
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
	
	If arrRet(0) <> "" Then
		Call SetDept(arrRet)
	End If	
	frm1.txtDeptCd.focus
End Function

'=======================================================================================

Function SetDept(Byval arrRet)
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtDateFr.text = arrRet(4)
		frm1.txtDateTo.text = arrRet(5)
End Function

'========================================================================================================= 
Sub Form_Load()
	
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)    
    Call InitVariables                            '⊙: Initializes local global Variables
    Call SetDefaultVal
    
    '----------  Coding part  -------------------------------------------------------------	
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

'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'=======================================================================================================
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
        Call SetFocusToDocument("M")
        frm1.fpDateTime1.focus
    End If
End Sub

Sub txtDateTo_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime2.Action = 7
        Call SetFocusToDocument("M")
        frm1.fpDateTime2.focus
    End If
End Sub


'=======================================================================================================
Sub SetPrintCond(StrEbrFile, StrUrl)

	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarBizAreaCd1, VarGlNoFr, VarGlNoTo
    
	Dim strAuthCond
	Dim txtDateProveFr,txtDateProveTo,txtusrid
	
	
	StrEbrFile = "a5122ma1_ko441"


	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, parent.gDateFormat, parent.gServerDateType) 
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, parent.gDateFormat, parent.gServerDateType)  
	
	txtDateProveFr = UniConvDateToYYYYMMDD(frm1.txtDateProveFr.Text, parent.gDateFormat, parent.gServerDateType) 
	txtDateProveTo = UniConvDateToYYYYMMDD(frm1.txtDateProveTo.Text, parent.gDateFormat, parent.gServerDateType)  
	txtusrid = trim(frm1.txtusrid.value)
	
	if txtDateProveTo="" then txtDateProveTo ="2999-12-31"
	
	
	VarDeptCd    = "" & FilterVar("%", "''", "S") & ""
	
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
	
	varGlNoFr    = "" & FilterVar(" ", "''", "S") & " "
	varGlNoTo    = "" & FilterVar("zzzzzzzzzzzzzzzzzz", "''", "S") & " "
	
	If Len(frm1.txtDeptCd.value)    > 0 Then VarDeptCd    = " " & FilterVar(UCase(frm1.txtDeptCd.value), "''", "S") & ""
	If Len(frm1.txtGlNoFr.value)    > 0 Then varGlNoFr    = " " & FilterVar(UCase(frm1.txtGlNoFr.value), "''", "S") & ""
	If Len(frm1.txtGlNoTo.value)    > 0 Then varGlNoTo    = " " & FilterVar(UCase(frm1.txtGlNoTo.value), "''", "S") & ""

	
	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_GL.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_GL.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_GL.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_GL.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	StrUrl = StrUrl & "glfrdt|"			& VarDateFr
	StrUrl = StrUrl & "|gltodt|"		& VarDateTo
	StrUrl = StrUrl & "|gldeptcd|"		& VarDeptCd
	StrUrl = StrUrl & "|BizAreacd|"		& VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|"	& VarBizAreaCd1
	StrUrl = StrUrl & "|glfrno|"		& varGlNoFr
	StrUrl = StrUrl & "|gltono|"		& varGlNoTo
	StrUrl = StrUrl & "|txtDateProveFr|"		& txtDateProveFr
	StrUrl = StrUrl & "|txtDateProveTo|"		& txtDateProveTo
	StrUrl = StrUrl & "|txtusrid|"		& txtusrid
	
'	msgbox StrUrl
	
	
	
	

	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond

	
End Sub


'=======================================================================================================
Function FncBtnPrint() 

    Dim StrUrl

    Dim StrEbrFile
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

'    On Error Resume Next                                                    '☜: Protect system from crashing
    
    If CompareDateByFormat(frm1.txtDateFr.text,frm1.txtDateTo.text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
                        "970025",frm1.txtDateFr.UserDefinedFormat,parent.gComDateType,True) = False Then		
		frm1.txtDateFr.focus                        	
		Exit Function
	End If 
    
   	'----------------------------------------------
	'전표번호 Check
	'----------------------------------------------
	frm1.txtGlNoFr.value = UCase(Trim(frm1.txtGlNoFr.value))
	frm1.txtGlNoTo.value = UCase(Trim(frm1.txtGlNoTo.value))
	
	If frm1.txtGlNoFr.value <> "" And frm1.txtGlNoTo.value <> "" Then
		If frm1.txtGlNoFr.value > frm1.txtGlNoTo.value Then
			Call DisplayMsgBox("970025","X", frm1.txtGlNoFr.Alt, frm1.txtGlNoTo.Alt)
			frm1.txtGlNoFr.focus 
			Exit Function
		End If
	End If

	Call SetPrintCond(StrEbrFile, StrUrl)

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")

	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function

'========================================================================================

Function BtnPreview() 
	'On Error Resume Next                                                    '☜: Protect system from crashing
    
    Dim StrUrl

    Dim StrEbrFile
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtDateFr.text,frm1.txtDateTo.text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
                        "970025",frm1.txtDateFr.UserDefinedFormat,parent.gComDateType,True) = False Then		
		frm1.txtDateFr.focus                        	
		Exit Function
	End If 
	
   If CompareDateByFormat(frm1.txtDateProveFr.text,frm1.txtDateProveTo.text,frm1.txtDateProveFr.Alt,frm1.txtDateProveTo.Alt, _
                        "970025",frm1.txtDateFr.UserDefinedFormat,parent.gComDateType,True) = False Then		
		frm1.txtDateProveFr.focus                        	
		Exit Function
	End If 
 
 
 
 
   	'----------------------------------------------
	'전표번호 Check
	'----------------------------------------------
	frm1.txtGlNoFr.value = UCase(Trim(frm1.txtGlNoFr.value))
	frm1.txtGlNoTo.value = UCase(Trim(frm1.txtGlNoTo.value))
	
	If frm1.txtGlNoFr.value <> "" And frm1.txtGlNoTo.value <> "" Then
		If frm1.txtGlNoFr.value > frm1.txtGlNoTo.value Then
			Call DisplayMsgBox("970025","X", frm1.txtGlNoFr.Alt, frm1.txtGlNoTo.Alt)
			frm1.txtGlNoFr.focus 
			Exit Function
		End If
	End If

	Call SetPrintCond(StrEbrFile, StrUrl)

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	
	Call FncEBRPreview(ObjName,StrUrl)
		
End Function


'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function


'========================================================================================

Function FncExcel() 
End Function


'=======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function

'=======================================================================================================

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
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
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
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>회계일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateFr" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="시작회계일자" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
													   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateTo" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="종료회계일자" id=fpDateTime2></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>부서코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="입력부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()"> <INPUT TYPE="Text" NAME="txtDeptNm" SIZE=25 tag="14" ALT="입력부서명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,3)"> 
													   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14" ALT="사업장명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,4)"> 
													   <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtGlNoFr" SIZE=25 MAXLENGTH=18 tag="11XXXU" ALT="시작회계번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNoFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefGL(0)">&nbsp;~&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtGlNoTo" SIZE=25 MAXLENGTH=18 tag="11XXXU" ALT="종료회계번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNoTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefGL(1)"></TD>
							</TR>
							
							
							
							<TR>
								<TD CLASS="TD5" NOWRAP>승인일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateProveFr" CLASS=FPDTYYYYMMDD tag="11" Title="FPDATETIME" ALT="시작승인일자" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
													   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateProveTo" CLASS=FPDTYYYYMMDD tag="11" Title="FPDATETIME" ALT="종료승인일자" id=fpDateTime2></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5" NOWRAP>승인자</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtusrId" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="승인자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtusrId.Value,5)"> 
													   <INPUT TYPE="Text" NAME="txtusrIdnm" SIZE=25 tag="14" ALT="승인자명"></TD>
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
					<TD><BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()"  Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=<%=BizSize%>>
		<TD WIDTH=1% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=1% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
		<TD WIDTH=99% HEIGHT=<%=BizSize%>>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>

