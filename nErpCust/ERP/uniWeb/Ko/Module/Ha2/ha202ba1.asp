<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : �����߰�װ�� 
*  3. Program ID           : Ha202ba1
*  4. Program Name         : Ha202ba1
*  5. Program Desc         : �����������/�����߰�װ�� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/05
*  8. Modified date(Last)  : 2003/06/16
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
'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'=======================================================================================================
Const BIZ_PGM_ID = "Ha202bb1.asp"
Dim StartDate
Dim IsOpenPop          

StartDate	= "<%=GetSvrDate%>"

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()

	frm1.txtBas_dt.text = UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtRetro_bas_dt.text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)
	
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "BA") %>
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
    			Call DisplayMsgbox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call DisplayMsgbox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
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
	Dim strYear,strMonth,strDay

	On Error Resume Next
	
	ExeReflect = False

	If Not chkField(Document, "1") Then
		Exit Function
	End If
	if  txtEmp_no_OnChange() then
		Exit Function
	end if

	IntRetCD = DisplayMsgbox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	IF LayerShowHide(1) = False Then
		Exit Function
	End If
	
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001

	strVal = strVal & "&txtBas_dt=" & frm1.txtBas_dt.year &	right("00" &frm1.txtBas_dt.month,2) & right("00"& frm1.txtBas_dt.day,2)
	strVal = strVal & "&txtRetro_bas_dt=" & frm1.txtRetro_bas_dt.year & right("00" & frm1.txtRetro_bas_dt.month,2) & right("00" & frm1.txtRetro_bas_dt.day,2)

	if  frm1.txtcalcu_logic1.checked = true then
	    strVal = strVal & "&txtcalcu_logic=1"
	elseif  frm1.txtcalcu_logic2.checked = true then
	    strVal = strVal & "&txtcalcu_logic=2"
	elseif  frm1.txtcalcu_logic3.checked = true then
	    strVal = strVal & "&txtcalcu_logic=3"
	elseif  frm1.txtcalcu_logic4.checked = true then
	    strVal = strVal & "&txtcalcu_logic=4"
	end if

	if  frm1.txtpay_logic1.checked = true then
	    strVal = strVal & "&txtpay_logic=M"
	else
	    strVal = strVal & "&txtpay_logic=D"
	end if


    ' Business Logic���� emp_no check('%')
    strVal = strVal & "&txtEmp_no=" & frm1.txtemp_no.value

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

	ExeReflect = True                                                           '��: Processing is NG

End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function ExeReflectOk()				            '��: ���� ������ ���� ���� 
	call DisplayMsgbox("990000","X","X","X")
	window.status = "�۾� �Ϸ�"
End Function
Function ExeReflectNo()				            '��: ����� �ڷᰡ �����ϴ� 
    Call DisplayMsgbox("800161","X","X","X")
	window.status = "�۾� �Ϸ�"
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
	arrParam(2) = lgUsrIntCd
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
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()

	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

	Call InitVariables                                                     '��: Setup the Spread sheet
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'��: ��ư ���� ���� 
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
'   Event Name : txtyear_yymm_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtBas_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtBas_dt.Action = 7
		frm1.txtBas_dt.focus
	End If
End Sub

Sub txtRetro_bas_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtRetro_bas_dt.Action = 7
		frm1.txtRetro_bas_dt.focus
	End If
End Sub

Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE,False)
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
<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����߰�װ��</font></td>
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
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>						
						    <TR>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/ha202ba1_fpDateTime1_txtBas_dt.js'></script></TD>
							</TR>
							<TR>	
								<TD CLASS=TD5 NOWRAP>�ұޱ�����</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/ha202ba1_present_dt_txtRetro_bas_dt.js'></script></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="�����" TYPE="Text" MAXLENGTH=13 SIZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmp()">&nbsp;<INPUT NAME="txtName" TYPE="Text" MAXLENGTH=30 SIZE=20 tag=14XXXU></TD>
						    </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP>������</TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCalcu_logic" TAG="1X" VALUE="������ӱ�*�ټӰ���/12" CHECKED ID="txtCalcu_logic1"><LABEL FOR="txtCalcu_logic1">������ӱ�*�ټӰ���/12</LABEL>
                            </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP></TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCalcu_logic" TAG="1X" VALUE="������ӱ�*�ټ��ϼ�/365"  ID="txtCalcu_logic2"><LABEL FOR="txtCalcu_logic2">������ӱ�*�ټ��ϼ�/365</LABEL>
                            </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP></TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCalcu_logic" TAG="1X" VALUE="��/��/��������"  ID="txtCalcu_logic3"><LABEL FOR="txtCalcu_logic3">��/��/��������</LABEL>
                            </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP></TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCalcu_logic" TAG="1X" VALUE="������ӱ�*30*�ټ��ϼ�/365"  ID="txtCalcu_logic4"><LABEL FOR="txtCalcu_logic4">������ӱ�*30*�ټ��ϼ�/365</LABEL>
                            </TR>
                            <TR>
		        				<TD CLASS="TD5" NOWRAP>��ձ޿��������</TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtPay_logic" TAG="1X" VALUE="����������ӱݻ���" CHECKED ID="txtPay_logic1"><LABEL FOR="txtPay_logic1">����������ӱݻ���</LABEL>
                            </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP></TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtPay_logic" TAG="1X" VALUE="�ϴ�������ӱݻ���"  ID="txtPay_logic2"><LABEL FOR="txtPay_logic2">�ϴ�������ӱݻ���</LABEL>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>����</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=100><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=100 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
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


