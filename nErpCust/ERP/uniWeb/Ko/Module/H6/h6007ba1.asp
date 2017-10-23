<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : ���ڱݱ޿��ݿ� 
*  3. Program ID           : H6007ba1
*  4. Program Name         : H6007ba1
*  5. Program Desc         : �޿�����/���ڱݱ޿��ݿ� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/04
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
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Const BIZ_PGM_ID = "H6007bb1.asp"
Dim IsOpenPop          
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
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtPay_cd, iCodeArr, iNameArr, Chr(11))
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	frm1.txtdt_from.text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
    frm1.txtdt_from.focus()

	frm1.txtdt_to.text = frm1.txtdt_from.text

	Call ggoOper.FormatDate(frm1.txtPay_yymm, Parent.gDateFormat, 2)

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtPay_yymm.Year = strYear 		 '����� default value setting
	frm1.txtPay_yymm.Month = strMonth 

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
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables                                        

    Call InitComboBox

    Call SetDefaultVal
    Call SetToolbar("1000000000000111")
    
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
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    FncDelete = True                                                             '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '��: Protect system from crashing
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
'	Name : OpenCode()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function OpenCode()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	        arrParam(0) = "���ޱ��� �˾�"			' �˾� ��Ī 
	        arrParam(1) = " b_minor "			 		' TABLE ��Ī 
	        arrParam(2) = frm1.txtprov_type.value       ' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = " major_cd = " & FilterVar("H0040", "''", "S") & " "		' Where Condition
	        arrParam(5) = "���ޱ����ڵ�"		    ' TextBox ��Ī 
	
            arrField(0) = " minor_cd "					' Field��(0)
            arrField(1) = " minor_nm "				    ' Field��(1)
    
            arrHeader(0) = "���ޱ����ڵ�"			' Header��(0)
            arrHeader(1) = "���ޱ��и�"			    ' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtprov_type.focus
		Exit Function
	Else
		Call SetCode(arrRet)
	End If	

End Function


'========================================================================================================
'	Name : SetCode()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================

Function SetCode(Byval arrRet)

        frm1.txtprov_type.value = arrRet(0)
        frm1.txtprov_type_nm.value = arrRet(1)
		frm1.txtprov_type.focus
End Function


'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line(grid�ܿ��� ���) 
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
'	Description : Item Popup���� Return�Ǵ� �� setting(grid�ܿ��� ���)
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		Call ggoOper.ClearField(Document, "2")					 '��: Clear Contents  Field
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'======================================================================================================
'   Event Name : txtEmp_no_OnChange
'   Event Desc : ���(����)�� ����� ��� 
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
    			Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            end if
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

    Dim IntRetCD    
	dim strVal
	Dim strProvDt, strToDt, LastDate
	Dim strYear,strMonth,strDay 

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	If txtEmp_no_OnChange() Then
		Exit Function
	End If

	If ValidDateCheck(frm1.txtDt_from, frm1.txtDt_to) = False then 
        exit function
     end if

	If IntRetCD = vbNo Then
		Exit Function
	End If
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")

	if LayerShowHide(1) = False then
		Exit Function
	end if	

	Call ExtractDateFrom(frm1.txtDt_from.Text,frm1.txtDt_from.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strProvDt = strYear & strMonth & strDay
	Call ExtractDateFrom(frm1.txtDt_to.Text,frm1.txtDt_to.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strToDt = strYear & strMonth & strDay
	Call ExtractDateFrom(frm1.txtPay_yymm.Text,frm1.txtPay_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	LastDate = strYear & strMonth

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0006
	strVal = strVal & "&txtDt_from=" & Trim(strProvDt)
	strVal = strVal & "&txtDt_to=" & Trim(strToDt)
	strVal = strVal & "&txtPay_yymm=" & Trim(LastDate)	
	strVal = strVal & "&txtPay_cd=" & frm1.txtPay_cd.value

    strVal = strVal & "&txtEmp_no=" & frm1.txtEmp_no.value
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

End Function

Sub FuncBtnRunOK()
    Call DisplayMsgBox("800154","X","X","X")
End Sub

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function ExeReflectOk()				            '��: ���� ������ ���� ���� 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	window.status = "�۾� �Ϸ�"

End Function

Function ExeReflectNo()				            '��: ó���� ����Ÿ�� �����ϴ�.
	Dim IntRetCD 

'	IntRetCD =DisplayMsgBox("800161","X","X","X")
	window.status = "�۾� �Ϸ�"

End Function

'========================================================================================================
' Name : txtYyyymm_DblClick
' Desc : �޷� Popup�� ȣ�� 
'========================================================================================================
Sub txtdt_from_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtdt_from.Action = 7
		frm1.txtdt_from.focus
	End If
End Sub

Sub txtdt_to_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")
		frm1.txtdt_to.Action = 7
		frm1.txtdt_to.focus
	End If
End Sub

Sub txtPay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtPay_yymm.Action = 7
		frm1.txtPay_yymm.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtProv_type_Onchange             
'   Event Desc :
'========================================================================================================
Function txtProv_type_Onchange()
    Dim IntRetCd

    If  frm1.txtProv_type.value = "" Then
		frm1.txtProv_type_nm.value = ""
    Else
        IntRetCd = Parent.FuncCodeName(1, "H0040", frm1.txtProv_type.value)
        If  IntRetCd = frm1.txtProv_type.value then
			Call DisplayMsgBox("970027","X","���ޱ���","X")
			frm1.txtProv_type_nm.value = ""
            frm1.txtProv_type.focus
            Set gActiveElement = document.ActiveElement   
            txtProv_type_Onchange = true    
        Else
			frm1.txtProv_type_nm.value = IntRetCd
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ڱݱ޿��ݿ�</font></td>
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
								<TD CLASS=TD5  NOWRAP>���Ⱓ</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h6007ba1_txtdt_from_txtdt_from.js'></script>
								                        ~<script language =javascript src='./js/h6007ba1_txtdt_to_txtdt_to.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ڱݹݿ���</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h6007ba1_txtPay_yymm_txtPay_yymm.js'></script>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�޿�����</TD>
	                        	<TD CLASS="TD6" NOWRAP><SELECT Name="txtPay_cd" ALT="�޿�����" STYLE="WIDTH: 133px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>			
								
							<TR>
								<TD CLASS=TD5 NOWRAP>�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="�����" TYPE="Text" MAXLENGTH=13 SIZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmpName (1)">&nbsp;<INPUT NAME="txtName" TYPE="Text" MAXLENGTH=30 SIZE=20 tag=14XXXU></TD>	
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
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:ExeReflect()" Flag="1">����</BUTTON>&nbsp;
		            </TD>
					<TD WIDTH=* ALIGN="right"></TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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


