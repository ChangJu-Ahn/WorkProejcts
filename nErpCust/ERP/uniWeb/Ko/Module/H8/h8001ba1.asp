<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 급여인상분계산 
*  3. Program ID           : H8001ba1
*  4. Program Name         : H8001ba1
*  5. Program Desc         : 급.상여소급분관리/급여인상분계산 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/07
*  8. Modified date(Last)  : 2003/06/13
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
Const BIZ_PGM_ID = "H8001bb1.asp"
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

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtfr_pay_yymm.Year = strYear 		 '년월일 default value setting
	frm1.txtfr_pay_yymm.Month = strMonth
	
	frm1.txtto_pay_yymm.Year = strYear 		 '년월일 default value setting
	frm1.txtto_pay_yymm.Month = strMonth

    frm1.txtfr_pay_yymm.focus()

	frm1.txtprov_cd.value = "P" ' 급여인상분계산 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "BA") %>
End Sub
'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   If pOpt = "Q" Then
      lgKeyStream = Frm1.txtWarrentNo.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
   Else
      lgKeyStream = Frm1.txtMajorCd.Value & Parent.gColSep         'You Must append one character(Parent.gColSep)
   End If   
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtPay_cd, lgF0, lgF1, Chr(11))

    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0040", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtProv_cd, lgF0, lgF1, Chr(11))


    iCodeArr = "1" & Chr(11) & "2" & Chr(11)
    iNameArr = "당월" & Chr(11) & "익월" & Chr(11)
    Call SetCombo2(frm1.txtStd_mm, iCodeArr, iNameArr, Chr(11))
    Call SetCombo2(frm1.txtProv_mm, iCodeArr, iNameArr, Chr(11))

End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    
    Call ggoOper.FormatDate(frm1.txtfr_pay_yymm, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtto_pay_yymm, Parent.gDateFormat, 2)

    Call AppendNumberPlace("6","2","0")
    Call ggoOper.FormatNumber(frm1.txtStd_dd, 31, 0)
    Call ggoOper.FormatNumber(frm1.txtProv_dd, 31, 0)

	Call InitVariables                                                     '⊙: Setup the Spread sheet
	
	Call InitComboBox()   

	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
			
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()


End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call FncQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncNew()	
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
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
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


'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD
    Dim strChk_pay_cd
    Dim strNum
    Dim intCnt
    Dim strRetro_yn

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If
	
	
    Dim strFrPayYYYY
    Dim strFrPayMM
    Dim strFrPayDt
    
    Dim strToPayYYYY
    Dim strToPayMM
    Dim strToPayDt
  	if txtEmp_no_OnChange() then
		Exit Function
	end if  
	strFrPayYYYY = frm1.txtfr_pay_yymm.Year
	strFrPayMM = frm1.txtfr_pay_yymm.Month
	If len(strFrPayMM) = 1 Then
		strFrPayMM = "0" & strFrPayMM
	End if
	strFrPayDt = strFrPayYYYY & strFrPayMM
    
	strToPayYYYY = frm1.txtto_pay_yymm.Year
	strToPayMM = frm1.txtto_pay_yymm.Month
	If len(strToPayMM) = 1 Then
		strToPayMM = "0" & strToPayMM
	End if
	strToPayDt = strToPayYYYY & strToPayMM
    	

    IF  strFrPayDt > strToPayDt THEN
        Call DisplayMsgBox("970027","X","급여년월","X")		'급여년월 오류입니다.
        frm1.txtfr_pay_yymm.focus
        Call BtnDisabled(0)
        Exit function
    END IF
    
	IntRetCD = DisplayMsgBox("800459",Parent.VB_YES_NO,"X","X")
	If IntRetCD = vbYes Then
        strRetro_yn = "Y"
    else
        strRetro_yn = "N"
	End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If LayerShowHide(1) = False Then
	     Call BtnDisabled(0)
	     Exit Function
	End If
	Call BtnDisabled(1)

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0006
	strVal = strVal & "&txtfr_pay_yymm=" & strFrPayDt
	strVal = strVal & "&txtto_pay_yymm=" & strToPayDt
	strVal = strVal & "&txtstd_mm=" & frm1.txtStd_mm.value
	if  Cint(frm1.txtstd_dd.text) < 10 then
	    strVal = strVal & "&txtstd_dd=0" & frm1.txtstd_dd.text
	else
	    strVal = strVal & "&txtstd_dd=" & frm1.txtstd_dd.text
	end if
	strVal = strVal & "&txtprov_mm=" & frm1.txtprov_mm.value

	if  Cint(frm1.txtprov_dd.text) < 10 then
	    strVal = strVal & "&txtprov_dd=0" & frm1.txtprov_dd.text
	else
	    strVal = strVal & "&txtprov_dd=" & frm1.txtprov_dd.text
	end if
	strVal = strVal & "&txtprov_cd=" & frm1.txtProv_cd.value
    strVal = strVal & "&txtpay_cd=" & frm1.txtpay_cd.value
    strVal = strVal & "&txtRetro_yn=" & strRetro_yn

    ' Business Logic에서 emp_no check('%')
    strVal = strVal & "&txtEmp_no=" & frm1.txtemp_no.value

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG

	Call BtnDisabled(0)
End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")

End Function
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
	Dim IntRetCD 

    Call DisplayMsgBox("800161","X","X","X")

End Function

Sub txtfr_Pay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtfr_Pay_yymm.Action = 7
		frm1.txtfr_Pay_yymm.focus
	End If
End Sub

Sub txtto_Pay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")
		frm1.txtto_Pay_yymm.Action = 7
		frm1.txtto_Pay_yymm.focus
	End If
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

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급여인상분계산</font></td>
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
								<TD CLASS="TD5" NOWRAP>급여년월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h8001ba1_txtfr_Pay_yymm_txtfr_Pay_yymm.js'></script> ~
                                                       <script language =javascript src='./js/h8001ba1_txtto_Pay_yymm_txtto_Pay_yymm.js'></script>
                                </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급구분</TD>
								<TD CLASS="TD6" NOWRAP><SELECT Name="txtProv_cd" ALT="지급구분" CLASS ="cbonormal" tag="14"></SELECT></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5  NOWRAP>기준일</TD>
								<TD CLASS=TD6  NOWRAP><table><tr><td><SELECT Name="txtStd_mm" ALT="기준일" CLASS ="cbonormal" tag="12"></SELECT></td>
								                      <td><script language =javascript src='./js/h8001ba1_txtStd_dd_txtStd_dd.js'></script></td>
								                      <td>일(0:말일)</td></tr></table>
								</TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5  NOWRAP>지급일</TD>
								<TD CLASS=TD6  NOWRAP><table><tr><td><SELECT Name="txtProv_mm" ALT="지급일" CLASS ="cbonormal" tag="12"></SELECT></td>
								                     <td> <script language =javascript src='./js/h8001ba1_txtProv_dd_txtProv_dd.js'></script></td>
								                     <td>일(0:말일)</td></tr></table>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>급여구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtPay_cd" ALT="급여구분" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대상자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="대상자" TYPE="Text" MAXLENGTH="13" SIZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmp()">&nbsp;<INPUT NAME="txtName" TYPE="Text" MAXLENGTH="30" SIZE=20 tag=14XXXU></TD>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


