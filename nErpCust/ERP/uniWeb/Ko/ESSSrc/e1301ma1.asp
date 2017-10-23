<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../ESSinc/IncServer.asp"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->

<Script Language="VBScript">
Option Explicit 

Const BIZ_PGM_ID      = "e1301mb1.asp"						           '��: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim Grid1
Dim Emp_no

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################
Sub LoadInfTB19029()
	<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "I", "H") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.txtYear.Value,"-","")) & gColSep
    else

        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.txtYear.Value,"-","")) & gColSep
    end if

End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    dim strSQL, IntRetCD
    
	iCodeArr = ""
	iNameArr = ""

    strSQL = " org_cd = " & FilterVar("1", "''", "S") & " AND pay_gubun = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE = " & FilterVar("*", "''", "S") & " "
    IntRetCD = CommonQueryRs(" year(close_dt) close_year "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
		iDx = Replace(lgF0, Chr(11), "") +1
	end if
	iCodeArr = cdbl(idx) & Chr(11) & iCodeArr
	iNameArr = cdbl(idx) & Chr(11) & iNameArr
	   
    Call SetCombo2(frm1.txtYear, iCodeArr, iNameArr, Chr(11))

End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status

	parent.document.All("nextprev").style.VISIBILITY = "hidden"

	call LoadInfTB19029()
    Call InitComboBox()

    Call LayerShowHide(0)

    Call SetToolBar("10010")

    frm1.txtEmp_no.value = parent.txtEmp_no.Value
    
    Call LockField(Document)

    Call DbQuery(1)

End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : ������ ��ȯ�̳� ȭ���� ���� ��� �����ؾ� �� ���� ó�� 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing
    Set Grid1 = Nothing
End Sub

Function DbQuery(ppage)
    Dim strVal
    Err.Clear                                                                    '��: Clear err status

    DbQuery = False                                                              '��: Processing is NG
    
    If frm1.txtYear.value = "" then
		Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
    
    If len(frm1.txtYear.value)<>4 then
		Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
    
    Call ClearField(document,2)
    Call LayerShowHide(1)
    if ppage = 1 then
        Call MakeKeyStream("Q")
    else
        Call MakeKeyStream("S")        
    end if

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic
    
    DbQuery = True                                                               '��: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '��: Clear err status
	lgIntFlgMode      = OPMD_UMODE                                              '��: Indicates that current mode is Create mode
    'Call Grid1.ShowData(frm1,1)
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)
    Call SetToolBar("10010")

End Function

Function DbQueryFail()
    Err.Clear
	lgIntFlgMode = ""

    Call ClearField(Document,2)
    Call SetToolBar("10000")

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
   
    If frm1.txtYear.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
    
    If len(frm1.txtYear.value)<>4 then
        Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
   
	if ChkField(Document, "2") then
		exit function
	end if
	
	Call LayerShowHide(1)
	Call MakeKeyStream("S")
	
	With Frm1
		.txtMode.value        = "UID_M0002"                                        '��: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '��: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call DbQuery(2)
End Function

Function DbSaveFail()
	Call DisplayMsgBox("990024","X","X","X")
End Function

Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Sub Print_onClick()
    Call SubPrint(MyBizASP)
End Sub


Sub GRID_PAGE_OnChange()
End Sub

Sub DELETE_OnClick()
    Call Grid1.DeleteClick()
End Sub

Sub CANCEL_OnClick()
    Call Grid1.CancelClick()
End Sub

</SCRIPT>

<!-- #Include file="../ESSinc/uniSimsClassID.inc" -->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
                            <tr> 
								<td width="80" height="27" bgcolor="D4E5E8" class="base1">���</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">����</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">����</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">�μ�</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
                            <tr> 
								<td width="80" height="30" bgcolor="D4E5E8" class=base1 valign=middle>����⵵
								</td>
								<td width="85" bgcolor="FFFFF" align=center>
								    <SELECT Name="txtYear" tabindex=-1 class=base2>
								    </SELECT>
								</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
                            </tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>
					<TR>
					    <TD class="ftgray">&nbsp;
							<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">�ҵ����</font></strong></td>                               
						<TD>
					</TR>
					<tr> 
					    <td height="3"></td>
					</tr>
					<TR><TD>
						<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		            		<TR>
                        	    <TD CLASS="ctrow01" >�����</TD>
	                           	<TD CLASS="ctrow02">
		            		        <INPUT CLASS="ftgray" TYPE="CHECKBOX" NAME="rdoSpouse_t" ID="rdoPhantomType1" disabled></INPUT>
		            		    </TD>
                        	    <TD CLASS="ctrow01" >�γ���</TD>
		            		    <TD CLASS="ctrow02">
		            		        <INPUT CLASS="ftgray" TYPE="CHECKBOX" NAME="rdoLady_t" ID="rdoPhantomType2" disabled></INPUT>
		            		    </TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >�ξ���(��)</TD>
                            	<TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtSupp_old_cnt_t" ALT="�ξ���(��)" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT>
                            	</TD>
		            		    <TD CLASS="ctrow01" >�ξ���(��)</TD>
                            	<TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtSupp_young_cnt_t" ALT="�ξ���(��)" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT>
                            	</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >�����(65�̻�)</TD>
                            	<TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtOld_cnt_t1" ALT="�����" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT>
                            	</TD>
		            		    <TD CLASS="ctrow01" >�����(70�̻�)</TD>
                            	<TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtOld_cnt_t2" ALT="�����" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT>
                            	</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >�ڳ������</TD>
                            	<TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtChl_rear_inwon_t" ALT="�ڳ������" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT>
                            	</TD>
		            		    <TD CLASS="ctrow01" >�����</TD>
                            	<TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtParia_cnt_t" ALT="�����" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT>
                            	</TD>
		            		</TR>
		            	</TABLE>
                    </TD></TR>
					<tr> 
					    <td height="3"></td>
					</tr>
					<TR><TD>
						<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		            		<TR>
		            		    <TD CLASS="ctrow01" >�� Ÿ �� ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtOther_insur_amt" ALT="��Ÿ����" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�� �� �� ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtMed_insur_amt" ALT="�ǰ�����" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >�� �� �� ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtEmp_insur_amt" ALT="��뺸��" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�� �� �� ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtNational_pension_amt" ALT="���ο���" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >��������뺸��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtDisabled_insur_amt" ALT="��������뺸��" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >&nbsp;</TD>
		            		    <TD CLASS="ctrow02">&nbsp;</TD>
		            		</TR>
		            	</TABLE>
                    </TD></TR>
					<tr> 
					    <td height="3"></td>
					</tr>
					<TR><TD>
						<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		            		<TR>
		            		    <TD CLASS="ctrow01" >���α�����</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtPer_edu_amt" ALT="���α�����" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�����Ư��������</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtDisabled_edu_amt" ALT="�����Ư��������" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            		    </TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >���߰�����</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtFam_edu_amt" ALT="���߰�����" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�ڳ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtFam_edu_cnt" ALT="�ڳ��" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT> ��
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >��ġ��������</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtKind_edu_amt" ALT="��ġ��������" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�ڳ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtKind_edu_cnt" ALT="�ڳ��" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT> ��
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >���б�����</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtUniv_edu_amt" ALT="���б�����" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�ڳ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtUniv_edu_cnt" ALT="�ڳ��" TYPE="Text" MAXLENGTH=5 SiZE=10 style='TEXT-ALIGN: right;' readonly></INPUT> ��
		            			</TD>
		            		</TR>
		            	</TABLE>
                    </TD></TR>
					<tr> 
					    <td height="3"></td>
					</tr>
					<TR><TD>
						<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		            		<TR>
		            		    <TD CLASS="ctrow01" >�Ϲ��Ƿ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtTot_med_amt" ALT="�Ϲ��Ƿ��" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >����/�����/������Ƿ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtSpeci_med_amt" ALT="�������Ƿ��" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		</TR>
		            	</TABLE>
                    </TD></TR>
					<tr> 
					    <td height="3"></td>
					</tr>
					<TR><TD>
						<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		            		<TR>
		            		    <TD CLASS="ctrow01" >������α�</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtLegal_contr_amt" ALT="������α�" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >��ġ�ڱݱ�α�</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtPoli_contr_amt1" ALT="��ġ�ڱݱ�α�" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >������(75%)</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtTaxLaw_contr_amt2" ALT="Ư�ʱ�α�(100%)" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >Ư�ʱ�α�(50%)</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtTaxLaw_contr_amt" ALT="Ư�ʱ�α�(50%)" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		</TR>		                        		
		            		<TR>
		            		    <TD CLASS="ctrow01" >�츮�������ձ�α�</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtOurstock_contr_amt" ALT="�츮�������ձ�α�" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >������α�</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtApp_contr_amt" ALT="������α�" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		</TR>	
		            		<TR>
		            		    <TD CLASS="ctrow01" >�뵿���պ�</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtPriv_contr_amt" ALT="�뵿���պ�" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >&nbsp;</TD>
		            		    <TD CLASS="ctrow02">&nbsp;</TD>
		            		</TR>			                        		
		            	</TABLE>
                    </TD></TR>    
					<tr> 
					    <td height="3"></td>
					</tr>
					<TR><TD>
						<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		            		<TR>
		            		    <TD CLASS="ctrow01" >��������/���Աݻ�ȯ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtHouse_fund_amt" ALT="��������/���Աݻ�ȯ��" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >��������������Ա����ڻ�ȯ��(15��̸�)</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtLong_house_loan_amt" ALT="��������������Ա����ڻ�ȯ��1" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >��������������Ա����ڻ�ȯ��(15���̻�)</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtLong_house_loan_amt1" ALT="��������������Ա����ڻ�ȯ��2" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >��ȥ/���/�̻��</TD>
		            		    <TD CLASS="ctrow02">
		            		        <INPUT CLASS="form01" NAME="txtCeremony_cnt" ALT="Ƚ��" TYPE="Text" MAXLENGTH=3 SiZE=3 style='TEXT-ALIGN: right;'></INPUT>ȸ
		            				<INPUT CLASS="form02" NAME="txtCeremony_amt" ALT="��ȥ��ʺ�" TYPE="Text" MAXLENGTH=14 SiZE=20 style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		</TR>		                        		
		            		<TR>
		            		    <TD CLASS="ctrow01" >�ܱ��α�����/������</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtFore_edu_amt" ALT="�ܱ��α�����/������" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�츮�����⿬��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtOur_stock_amt" ALT="�츮����" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            		    </TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >���ο���(2000������)</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtIndiv_anu_amt" ALT="���ο���(2000������)" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >��������(2001������)</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtIndiv_anu2_amt" ALT="��������(2001������)" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >�����������ھ�</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtinvest2_sub_amt" ALT="�����������ھ�" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�������ݼҵ����</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtRetire_pension" ALT="�������ݼҵ����" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>

		            			</TD>
		            		</TR>		                        		
		            	</TABLE>
                    </TD></TR>
					<tr> 
					    <td height="3"></td>
					</tr>
					<TR><TD>
						<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		            		<TR>
		            			<TD CLASS="ctrow01" >�ſ�/����/����ī��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtCard_use_amt" ALT="�ſ�/����/����ī��" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >���ݿ�����</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtCard2_use_amt" ALT="���ݿ�����" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >�п������γ���</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form02" NAME="txtInstitution_giro" ALT="�п������γ���" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;' readonly></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >�ܱ��ҵ�</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtFore_income_amt" ALT="�ܱ��ҵ�" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >������</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtAfter_bonus_amt" ALT="������" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >��Ÿ�ҵ�</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtOther_income_amt" ALT="��Ÿ�ҵ�" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		</TR>		                        		
		            	</TABLE>
                    </TD></TR>
					<tr> 
					    <td height="3"></td>
					</tr>
					<TR><TD>
						<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		            		<TR>
		            		    <TD CLASS="ctrow01" >�������Ա����ڻ�ȯ��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtHouse_repay_amt" ALT="�������Ա����ڻ�ȯ��" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
			            		    <TD CLASS="ctrow01" >���ٳ������հ���</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtTax_Union_Ded" ALT="���ٳ������հ���" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>

		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01">�ܱ����μ���</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtFore_pay_amt" ALT="�ܱ����μ���" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >���ٹ����������</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtSave_tax_sub_amt" ALT="���ٹ����������" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		</TR>
		            		<TR>
		            		    <TD CLASS="ctrow01" >�ҵ漼��</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtIncome_redu_amt" ALT="�ҵ漼��" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		    <TD CLASS="ctrow01" >������</TD>
		            		    <TD CLASS="ctrow02">
		            				<INPUT CLASS="form01" NAME="txtTaxes_redu_amt" ALT="������" TYPE="Text" MAXLENGTH=14 SiZE=20  style='TEXT-ALIGN: right;'></INPUT>
		            			</TD>
		            		</TR>
		            	</TABLE>
                    </TD></TR>
                </TABLE>
            </TD>
        </TR>
		<TR>
			<TD height=5></TD>
		</TR>
    </TABLE>

    <TABLE cellSpacing=2 cellPadding=0 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">
</FORM>	

</BODY>
</HTML>
