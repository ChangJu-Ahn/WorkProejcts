<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name       : 인사/급여관리 
'*  2. Function Name     : 연말정산관리 
'*  3. Program ID        : h9204oa1
'*  4. Program Name      : 연월차대장출력 
'*  5. Program Desc      : 연월차대장출력 
'*  6. Comproxy 리스트   :
'*  7. 최초 작성년월일   : 2001/06/01
'*  8. 최종 수정년월일   : 2003/06/16
'*  9. 최초 작성자       : TGS 최용철 
'* 10. 최종 작성자       : Lee SiNa
'* 11. 전체 comment      :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBsCRIPT"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow
<%StartDate	= GetSvrDate%>
'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear,strMonth,strDay
	frm1.txtpay_yymm.focus
	Call ggoOper.FormatDate(frm1.txtpay_yymm, parent.gDateFormat, 2)
	Call ExtractDateFrom("<%=StartDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)
	frm1.txtpay_yymm.Year	= strYear
	frm1.txtpay_yymm.Month	= strMonth
	frm1.txtpay_yymm.Day	= strDay
End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("P", "H", "NOCOOKIE", "OA") %>
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitVariables

    Call FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)						' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
    Call SetToolbar("1000000000000111")

End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If
End Function

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt            
    Dim LastDate

	If Not chkField(Document, "1") Then													'⊙: This function check indispensable field
       Exit Function
    End If
    
    strBasDt	= UniConvYYYYMMDDToDate(parent.gDateFormat,frm1.txtpay_yymm.Year,Right("0" & frm1.txtpay_yymm.Month,2),frm1.txtpay_yymm.Day)
	LastDate    = UNIGetLastDay (strBasDt,parent.gDateFormat)									'Last  day of this month

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value											' Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value											' Code Condition	
	End If

    arrParam(1) = LastDate
	arrParam(2) = lgUsrIntcd

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
               frm1.txtFr_dept_cd.focus
             Case "1"
               frm1.txtTo_dept_cd.focus
        End Select
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If

End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
		     Case "0"
               .txtFr_dept_cd.value = arrRet(0)
               .txtFr_dept_nm.value = arrRet(1)
               .txtFr_internal_cd.value = arrRet(2)
               .txtFr_dept_cd.focus
             Case "1"
               .txtTo_dept_cd.value = arrRet(0)
               .txtTo_dept_nm.value = arrRet(1)
               .txtTo_internal_cd.value = arrRet(2)
               .txtTo_dept_cd.focus
        End Select
	End With
End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================%>

Function FncBtnPrint()
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
	Dim ObjName
    If Not chkField(Document, "1") Then											'⊙: This function check indispensable field
       Exit Function
    End If

	dim bas_yymm, fr_dept_cd, to_dept_cd , rFrDept ,rToDept , IntRetCd

	StrEbrFile = "h9204oa1"

    Dim strYear,strMonth,strDay

    Call ExtractDateFrom(frm1.txtpay_yymm.Text,frm1.txtpay_yymm.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

	bas_yymm = strYear&strMonth

	fr_dept_cd = frm1.txtFr_dept_cd.value
	to_dept_cd = frm1.txtTo_dept_cd.value

    If  txtFr_Dept_cd_Onchange() then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange()  then
        Exit Function
    End If

    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value

    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If

	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If

    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value

    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF
    END IF

	condvar = "bas_yymm|" & bas_yymm
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd
	condvar = condvar & "|to_dept_cd|" & to_dept_cd

 	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 	call FncEBRPrint(EBAction , ObjName , condvar)

End Function

'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()

    If Not chkField(Document, "1") Then											'⊙: This function check indispensable field
       Exit Function
    End If

	dim condvar
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile

	dim bas_yymm, fr_dept_cd, to_dept_cd , rFrDept ,rToDept , IntRetCd
	Dim ObjName
	StrEbrFile = "h9204oa1"

    Dim strYear,strMonth,strDay

    Call ExtractDateFrom(frm1.txtpay_yymm.Text,frm1.txtpay_yymm.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

	bas_yymm = strYear&strMonth

	fr_dept_cd = frm1.txtFr_dept_cd.value
	to_dept_cd = frm1.txtTo_dept_cd.value

    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If

    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value

    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If

	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If

    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value

    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF
    END IF

	condvar = "bas_yymm|" & bas_yymm
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd
	condvar = condvar & "|to_dept_cd|" & to_dept_cd
 
 	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , condvar)

End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
	Call Parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim strBasDt   
    Dim LastDate

    strBasDt	= UniConvYYYYMMDDToDate(parent.gServerDateFormat,frm1.txtpay_yymm.Year,Right("0" & frm1.txtpay_yymm.Month,2),frm1.txtpay_yymm.Day)
	LastDate    = UNIGetLastDay (strBasDt,parent.gServerDateFormat)									'Last  day of this month	

    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , LastDate , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")										'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")										'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtFr_dept_cd_Onchange = true
            Exit Function
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if
    End if

End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim strBasDt    
    Dim LastDate

	strBasDt	= UniConvYYYYMMDDToDate(parent.gServerDateFormat,frm1.txtpay_yymm.Year,Right("0" & frm1.txtpay_yymm.Month,2),frm1.txtpay_yymm.Day)
	LastDate    = UNIGetLastDay (strBasDt,parent.gServerDateFormat)									'Last  day of this month	

    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , LastDate , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtTo_dept_cd_Onchange = true
            Exit Function
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if
    End if

End Function

'=======================================================================================================
'   Event Name : txt________Keypress
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================

Sub txtFr_dept_cd_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

Sub txtto_dept_cd_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub
'========================================================================================================
' Name : txtPay_yymm_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtPay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 	
		frm1.txtPay_yymm.Action = 7
		frm1.txtPay_yymm.focus
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>연월차대장출력</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
								<TD CLASS=TD5  NOWRAP>해당년월</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h9204oa1_txtPay_yymm_txtPay_yymm.js'></script></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU">&nbsp;~
		                                             <INPUT NAME="txtFr_internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU">
		                    </TR>
		                    <TR>
		                        <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" ALT="부서코드" TYPE="Text" SIZE=10 MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                         <INPUT NAME="txtto_dept_nm" ALT="부서코드명" TYPE="Text" SIZE=20 MAXLENGTH="40"  tag="14XXXU">
    			                                     <INPUT NAME="txtTo_internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
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
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()">미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()">인쇄</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>

</BODY>
</HTML>

