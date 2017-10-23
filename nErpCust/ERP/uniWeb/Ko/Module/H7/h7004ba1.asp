<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 상여지급율생성 
*  3. Program ID           : H7004ba1
*  4. Program Name         : H7004ba1
*  5. Program Desc         : 상여관리/상여지급율생성 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/04
*  8. Modified date(Last)  : 2001/06/04
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H7004bb1.asp"
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()

	frm1.txtbas_dt.text  = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) 
	frm1.txtprov_dt.text = frm1.txtbas_dt.text

	frm1.txtbonus_yymm_dt.text = UniConvDateAToB(frm1.txtbas_dt.text, Parent.gDateFormat, Parent.gDateFormatYYYYMM)
	frm1.txtbonus_yymm_dt.focus()

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "BA") %>
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtPay_cd, lgF0, lgF1, Chr(11))

    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0040", "''", "S") & "  AND MINOR_CD BETWEEN " & FilterVar("2", "''", "S") & " AND " & FilterVar("9", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtProv_type, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
	End If
    arrParam(2) = ""
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmpName()
'	Description : Item Popup에서 Return되는 값 setting(grid외에서 사용)
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'======================================================================================================
'   Event Name : txtEmp_no_OnChange
'   Event Desc : 사번(성명)이 변경될 경우 
'=======================================================================================================
Function txtEmp_no_OnChange()

    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

	frm1.txtName.value = ""
    If  frm1.txtEmp_no.value = "" Then
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_OnChange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if

End Function

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD
	Dim strProvDt, strToDt, LastDate, StartDate, EndDate, rDate
	Dim strYear,strMonth,strDay 
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Exit Function
	End If
	if txtEmp_no_OnChange() then
		Exit Function
	End If
	
	strDay = frm1.txtbonus_yymm_dt.Year & Right("0" & frm1.txtbonus_yymm_dt.month, 2)    
	
	rDate = UNIGetLastDay(frm1.txtbonus_yymm_dt.Text, Parent.gDateFormatYYYYMM)
	rDate = UniConvDateAToB(rDate, Parent.gDateFormat, Parent.gServerDateFormat)

    IF  FuncAuthority(frm1.txtProv_type.value, rDate , Parent.gUsrID) = "N" THEN
        '"상여 마감처리된 일 입니다."
        Call DisplayMsgBox("800313","X","X","X")
        Call BtnDisabled(0)
        exit function
    END IF

	 If ValidDateCheck(frm1.txtDilig_strt_dt, frm1.txtDilig_end_dt) = False then 	
        frm1.txtDilig_strt_dt.value = ""
        frm1.txtDilig_end_dt.value = ""
        frm1.txtDilig_strt_dt.focus
        exit function
    END IF

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	if LayerShowHide(1) = False then
		Exit Function
	end if	

	ExeReflect = False                                                          '⊙: Processing is NG
    
    strProvDt = frm1.txtBonus_yymm_dt.Year & Right("0" & frm1.txtBonus_yymm_dt.Month, 2)
    strToDt   = frm1.txtBas_dt.Year & Right("0" & frm1.txtBas_dt.Month, 2) & Right("0" & frm1.txtBas_dt.Day, 2)    
    LastDate  = frm1.txtProv_dt.Year & Right("0" & frm1.txtProv_dt.month, 2) & Right("0" & frm1.txtProv_dt.Day, 2)     

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	strVal = strVal & "&txtBonus_yymm_dt=" & Trim(strProvDt)
	strVal = strVal & "&txtProv_type=" & frm1.txtProv_type.value
	strVal = strVal & "&txtBas_dt=" & Trim(strToDt)
	strVal = strVal & "&txtProv_dt=" & Trim(LastDate)

	strVal = strVal & "&txtPay_cd=" & frm1.txtPay_cd.value

    if  frm1.txtDilig_strt_dt.text = "" then
        strVal = strVal & "&txtDilig_strt_dt=25001231"
    else
		StartDate = frm1.txtDilig_strt_dt.Year & right("0" & frm1.txtDilig_strt_dt.Month,2) & right("0" & frm1.txtDilig_strt_dt.Day,2)
		strVal = strVal & "&txtDilig_strt_dt=" & Trim(StartDate)
    end if

    if  frm1.txtDilig_end_dt.text = "" then
        strVal = strVal & "&txtDilig_end_dt=25001231"
    else
		Call ExtractDateFrom(frm1.txtDilig_end_dt.Text,frm1.txtDilig_end_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
		'EndDate = frm1.txtDilig_end_dt.Year & frm1.txtDilig_end_dt.Month & frm1.txtDilig_end_dt.Day
		EndDate = strYear & strMonth & strDay
		strVal = strVal & "&txtDilig_end_dt=" & Trim(EndDate)
    end if

    ' Business Logic에서 emp_no check('%')
    strVal = strVal & "&txtEmp_no=" & frm1.txtEmp_no.value
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG

End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	window.status = "작업 완료"

End Function

Function ExeReflectNo()				            '☆: 처리할 데이타가 없습니다.
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("800161","X","X","X")
	window.status = "작업 완료"

End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = ""
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)

		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

Function FuncAuthority(Pay_type, Pay_yymmdd, Emp_no)

    Dim strRet
    Dim IntRetCD

    strRet = "N"
    IntRetCD = CommonQueryRs("close_type, Convert(char(10),close_dt,20), emp_no","hda270t","org_cd=" & FilterVar("1", "''", "S") & "  and pay_gubun=" & FilterVar("Z", "''", "S") & "  and pay_type= " & FilterVar(Pay_type, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  IntRetCD = false then
        strRet = "Y"
    else

        SELECT CASE Replace(lgF0, Chr(11), "")
        	CASE "1" '마감형태 : 정상 
        	    IF UNIGetLastDay(Replace(lgF1, Chr(11), ""),Parent.gServerDateFormat) <= Pay_yymmdd THEN 
        	        strRet = "Y"
        		ELSE
        	        strRet = "N" 
        		END IF
           CASE "2" '마감형태 : 마감 
        	    
        	    IF  UNIGetLastDay(Replace(lgF1, Chr(11), ""),Parent.gServerDateFormat) < Pay_yymmdd THEN 
        	        strRet = "Y" 
        		ELSE
        	        strRet = "N" 
        	    END IF
        END SELECT
        
    end if

    FuncAuthority = strRet

End Function

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtBonus_yymm_dt, Parent.gDateFormat, 2)

	Call InitVariables                                                     '⊙: Setup the Spread sheet
	
	Call InitComboBox()
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
'   Event Name : txtbonus_yymm_dt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtbonus_yymm_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtbonus_yymm_dt.Action = 7
		frm1.txtbonus_yymm_dt.focus
	End If
End Sub

'======================================================================================================
'   Event Name : txtBas_dt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtBas_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtBas_dt.Action = 7
		frm1.txtBas_dt.focus
	End If
End Sub

'======================================================================================================
'   Event Name : txtProv_dt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtProv_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtProv_dt.Action = 7
		frm1.txtProv_dt.focus
	End If
End Sub

'======================================================================================================
'   Event Name : txtDilig_strt_dt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtDilig_strt_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtDilig_strt_dt.Action = 7
		frm1.txtDilig_strt_dt.focus
	End If
End Sub

'======================================================================================================
'   Event Name : txtDilig_end_dt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtDilig_end_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtDilig_end_dt.Action = 7
		frm1.txtDilig_end_dt.focus
	End If
End Sub

'==========================================================================================
' Function Name : FncQuery
' Function Desc : 
'============================================================================================
Function FncQuery()

End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

'======================================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'=======================================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>상여지급율생성</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_40%>WIDTH=100%>   
						   	<TR>
								<TD CLASS=TD5 NOWRAP>상여년월</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h7004ba1_fpDateTime2_txtBonus_yymm_dt.js'></script></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>상여구분</TD>
							    <TD CLASS="TD6" NOWRAP><SELECT Name="txtProv_type" ALT="상여구분" STYLE="WIDTH: 133px" tag="12"></SELECT></TD>
							</TR>                       
							<TR>
								<TD CLASS=TD5 NOWRAP>기준일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h7004ba1_fpDateTime2_txtBas_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>지급일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h7004ba1_fpDateTime2_txtProv_dt.js'></script></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>급여구분</TD>
							    <TD CLASS=TD6 NOWRAP><SELECT Name="txtPay_cd" ALT="급여구분" STYLE="WIDTH: 133px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>근태기간</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h7004ba1_fpDateTime2_txtDilig_strt_dt.js'></script>
											  &nbsp; ~ &nbsp;
											  <script language =javascript src='./js/h7004ba1_fpDateTime2_txtDilig_end_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대상자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT ID="txtEmp_no" NAME="txtEmp_no" SIZE=13 MAXLENGTH=13 STYLE="TEXT-ALIGN: left" tag="11X" ALT="대상자코드명"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">
								                     <INPUT TYPE=TEXT ID="txtName" NAME="txtName" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ALT="대상자코드">
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
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD Width = 10> &nbsp </TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag="1">실행</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
