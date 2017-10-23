<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : �ٹ�Į���ٻ��� 
*  3. Program ID           : H4002ma1
*  4. Program Name         : H4002ma1
*  5. Program Desc         : ���°���/�ٹ�Į���ٻ��� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/04
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : Hwang Jeong-won
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBSCRIPT">
Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
Const BIZ_PGM_ID = "H4001mb2.asp"

Dim IsOpenPop          

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
	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtStrYear.focus
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtStrYear.Year = strYear 		 '����� default value setting
	frm1.txtStrYear.Month = strMonth 
	
	frm1.txtEndYear.Year = strYear 		 '����� default value setting
	frm1.txtEndYear.Month = strMonth 
	
	frm1.txtDays.disabled = True	
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf  '�ɼ� money������ decimal point�� ���� ������ �ֱ����ؼ�..�ʿ�?
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A( "Q", "H","NOCOOKIE","MA")%>
End Sub

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 
	Call BtnDisabled(1)
	Dim strVal
	Dim txtStrYear,txtEndYear
	Dim IntRetCD

	ExeReflect = False                                                          '��: Processing is NG
    
	On Error Resume Next                                                   '��: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If txtBA_CD_OnChange() = true Then
		Call BtnDisabled(0)
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

    If Not( ValidDateCheck(frm1.txtStrYear, frm1.txtEndYear)) Then
		Call BtnDisabled(0)    
        Exit Function
    End If

	txtStrYear = frm1.txtStrYear.Year & right("0" & frm1.txtStrYear.Month,2)
	txtEndYear = frm1.txtEndYear.Year & right("0" & frm1.txtEndYear.Month,2)
	
	If LayerShowHide(1) = False then
    		Exit Function 
    End if

	strVal = BIZ_PGM_ID & "?txtMode="		& parent.UID_M0006
	strVal = strVal     & "&txtStrYear="	& txtStrYear
	strVal = strVal     & "&txtEndYear="	& txtEndYear	
	strVal = strVal     & "&txtBA_cd="		& frm1.txtBA_cd.value
    strVal = strVal     & "&txtWork="		& frm1.cboWork.value

    If frm1.txtDay5YN.checked = True Then	'��5�� ���� ����(N:��5�� �������,1:��~6:��) - 2006.04.24
		strVal = strVal     & "&txtDays="		& frm1.txtDays.value
	Else
		strVal = strVal     & "&txtDays=N"
	End If

	Call RunMyBizASP(MyBizASP, strVal)	                                        '��: �����Ͻ� ASP �� ���� 

	ExeReflect = True                                                           '��: Processing is NG
	Call BtnDisabled(0)
End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function ExeReflectOk()				            '��: ���� ������ ���� ���� 
	call DisplayMsgBox("990000","X","X","X")
	frm1.txtStrYear.focus
End Function

Function ExeReflectNo()				            '��: ����� �ڷᰡ �����ϴ� 
    Call DisplayMsgBox("800161","X","X","X")
	frm1.txtStrYear.focus
End Function

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()
	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtStrYear, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtEndYear, parent.gDateFormat, 2)

	Call InitVariables                                                     '��: Setup the Spread sheet
	Call InitComboBox
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
'   Event Name : txtStrYear_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtStrYear_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 	
		frm1.txtStrYear.Action = 7
		frm1.txtStrYear.focus
	End If
End Sub

Sub txtEndYear_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 	
		frm1.txtEndYear.Action = 7
		frm1.txtEndYear.focus
	End If
End Sub
'==========================================================================================
' Function Name : FncQuery
' Function Desc : 
'============================================================================================
Function FncQuery()

End Function

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

'========================================================================================================
' Name : OpenbizareaInfo()
' Desc : developer describe this line
'========================================================================================================
Function OpenbizareaInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����POPUP"					' �˾� ��Ī %>
	arrParam(1) = "B_BIZ_AREA"						' TABLE ��Ī %>
	arrParam(2) = strCode							' Code Condition%>
	arrParam(3) = ""								' Name COndition%>
	arrParam(4) = ""								' Where Condition%>
	arrParam(5) = "������ڵ�"			
	
    arrField(0) = "BIZ_AREA_CD"						' Field��(0)%>
    arrField(1) = "BIZ_AREA_NM"						' Field��(1)%>    
    arrHeader(0) = "������ڵ�"					' Header��(0)%>
    arrHeader(1) = "������"					' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBA_CD.focus
	    Exit Function
	Else
		Call SetbizareaInfo(arrRet)
	End If	

End Function

'========================================================================================================
' Name : SetBizAreaInfo()
' Desc : developer describe this line
'========================================================================================================
Function SetBizAreaInfo(ByVal arrRet)

	With frm1
		.txtBA_CD.value = arrRet(0)
		.txtBA_NM.value = arrRet(1)		
		.txtBA_CD.focus
	End With
	
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()

    Dim iCodeArr
    Dim iNameArr
        
    Err.Clear                                                               '��: Clear error no
	On Error Resume Next

    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0047", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 'unicode
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboWork, iCodeArr, iNameArr, Chr(11))
    
    iCodeArr = "1" & Chr(11) & "2" & Chr(11) & "3" & Chr(11) & "4" & Chr(11) & "5" & Chr(11) & "6"& Chr(11)
    iNameArr = "��" & Chr(11) & "ȭ" & Chr(11) & "��" & Chr(11) & "��" & Chr(11) & "��" & Chr(11) & "��"& Chr(11)
    Call SetCombo2(frm1.txtDays, iCodeArr, iNameArr, Chr(11))
	frm1.txtDays.selectedIndex = "5"
	
End Sub
'========================================================================================================
'   Event Name : txtBA_CD_OnChange
'   Event Desc : 
'========================================================================================================
Function txtBA_CD_OnChange()    
    Dim IntRetCd
 
    If frm1.txtBA_CD.value = "" Then
        frm1.txtBA_NM.value = ""
    ELSE    
        IntRetCd = CommonQueryRs(" biz_area_nm "," b_biz_area "," biz_area_cd =  " & FilterVar(frm1.txtBA_CD.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 'unicode
        If IntRetCd = false Then
            Call DisplayMsgBox("800272","X","X","X")                         '�� : ������ȣ�� �ٸ��� �Է� �Ͻʽÿ�.
            frm1.txtBA_NM.value = ""
            frm1.txtBA_CD.focus
            Set gActiveElement = document.activeElement  
            txtBA_CD_OnChange = True
            Exit Function
        Else
            frm1.txtBA_NM.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

Sub txtDay5YN_OnClick()
	If frm1.txtDay5YN.checked = True Then		
		frm1.txtDays.disabled = False
	Else
		frm1.txtDays.disabled = True	
	End If
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ٹ�ī��������</font></td>
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
						<TABLE <%=LR_SPACE_TYPE_60%>WIDTH=100%>   
						    	<TR>
									<TD CLASS="TD5" NOWRAP>�����Ⱓ</TD>

									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtStrYear NAME="txtStrYear" CLASS=FPDTYYYYMM tag="12X1" ALT="�������۳��" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtEndYear NAME="txtEndYear" CLASS=FPDTYYYYMM tag="12X1" ALT="����������" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>

								<TR>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBA_CD" MAXLENGTH="10" SIZE=10  ALT ="������ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenbizareaInfo(frm1.txtba_cd.value)">&nbsp;
									<INPUT NAME="txtBA_NM" MAXLENGTH="50" SIZE=30 ALT ="������" tag="14X"></TD>									
								</TR>			

							    <TR>
									<TD CLASS="TD5" NOWRAP>�ٹ�������</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboWork" tag="12X" CLASS ="cbonormal" ALT="�ٹ�������"></SELECT></TD>
							    </TR>
							    <TR>
									<TD CLASS="TD5" NOWRAP>��5����������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="txtDay5YN" ID="txtDay5YN">
									<SELECT NAME="txtDays" tag="12X" CLASS ="cbonormal" ALT="��5����������"></SELECT></TD>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>����</BUTTON></TD>
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
