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
<%

    Dim RYear
    Dim Emp_no
    Dim txtYear
    Emp_no = Trim(Request("txtEmp_no"))
    RYear = Trim(Request("txtYear"))
%>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1303mb2.asp"						           '��: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../ESSinc/lgvariables.inc" --> 

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim RYear
Dim Emp_no

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & RYear & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep                    ' Internal_cd
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & RYear & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
    end if

End Sub        

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status

    RYear = "<%=RYear%>"

    Call LayerShowHide(0)

    Call SetToolBar("00000")

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
    
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic
    
    DbQuery = True                                                               '��: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '��: Clear err status

End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '��: Clear err status

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
		
	Call LayerShowHide(1)

	With Frm1
		.txtMode.value        = "UID_M0002"                                        '��: Save
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
    Call DbQuery()
End Function


'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '��: Processing is OK
    Err.Clear                                                                    '��: Clear err status
    

    Call MakeKeyStream("N")


    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & "UID_M0001"                     '��: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"	                         '��: Direction

    StrSQL = " emp_no, name, dept_nm, entr_dt, group_entr_dt, resent_promote_dt, " 
    StrSQL = StrSQL & " (select b_minor.minor_nm from b_minor where b_minor.minor_cd = roll_pstn and b_minor.major_cd=" & FilterVar("H0002", "''", "S") & ") as roll_pstn "

        Frm1.txtEmp_no.Value  = lgF0
        Frm1.txtName.Value  = lgF1
        frm1.txtDept_nm.value = lgF2
        frm1.txtroll_pstn.value = lgF6
        frm1.txtresent_promote_dt.value = lgF5

        frm1.txtentr_dt.value = lgF3
        frm1.txtgroup_entr_dt.value = lgF4

	Call RunMyBizASP(MyBizASP, strVal)										     '��: Run Biz 

    FncNext = True                                                               '��: Processing is OK
	
End Function

'========================================================================================================
' Name : goBackForm
' Desc : 
'========================================================================================================
Function goBackForm1() 
    goBackForm1 = False                                                              '��: Processing is OK
    Err.Clear                                                                    '��: Clear err status
	history.back 
    goBackForm1 = True                                                               '��: Processing is OK
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
    Call DbQuery()
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
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>

                    <TR>
                        <TD>
                            <TABLE cellSpacing=0 cellPadding=0 width=100% border=0 bgcolor=#DDDDDD>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		    <TD CLASS=ctrow03 width=75%>�������޿�</TD>
		                        		    <TD CLASS=ctrow04 width=25%><INPUT CLASS=form02 NAME="txtincome_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>�ٷμҵ����</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS=form02 NAME="txtincome_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>�������ٷμҵ�ݾ�</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS=form02 NAME="txtIncome_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		     <TD CLASS=ctrow03 rowspan="4" width=25%>�⺻����</TD>    
		                        		     <TD CLASS=ctrow03 width=30%>����</TD>       
		                        		     <TD CLASS="ctrow04"  width=20%></TD>
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtper_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>       
		                        		<TR>
		                        		     <TD CLASS=ctrow03  width=30%>�����</TD>       
		                        		     <TD CLASS="ctrow04"  width=20%></TD>
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtspouse_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=30%>�ξ簡��</TD>     
		                        		     <TD CLASS="ctrow04"  width=20%>
		                        				<INPUT CLASS="form02" NAME="txtsupp_old_cnt" TYPE="Text" MAXLENGTH=30 SiZE=6 style='TEXT-ALIGN: right;' readonly></INPUT> ��
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtsupp_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>   
		                        		     <TD CLASS=ctrow03 rowspan="5" width=25%>�߰�����</TD>    
		                        		     <TD CLASS=ctrow03  width=30%>��ο��(65���̻�)</TD>     
		                        		     <TD CLASS="ctrow04"  width=20%>
		                        				<INPUT CLASS="form02" NAME="txtold_cnt1" TYPE="Text" MAXLENGTH=30 SiZE=6 style='TEXT-ALIGN: right;' readonly></INPUT> ��
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=50% rowspan="2">
		                        				<INPUT CLASS="form02" NAME="txtold_sub_amt1" TYPE="Text" MAXLENGTH=30 SiZE=15  style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=30%>��ο��(70���̻�)</TD>     
		                        		     <TD CLASS="ctrow04"  width=20%>
		                        				<INPUT CLASS="form02" NAME="txtold_cnt2" TYPE="Text" MAXLENGTH=30 SiZE=6 style='TEXT-ALIGN: right;' readonly></INPUT> ��
		                        		     </TD>
		                        		</TR> 		                        		 
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=30%>�����</TD>     
		                        		     <TD CLASS="ctrow04"  width=20%>
		                        				<INPUT CLASS="form02" NAME="txtparia_cnt" TYPE="Text" MAXLENGTH=30 SiZE=6 style='TEXT-ALIGN: right;' readonly></INPUT> ��
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtparia_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>
		                        		     <TD CLASS=ctrow03  width=30%>�γ���</TD>       
		                        		     <TD CLASS="ctrow04"  width=20%></TD>
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtlady_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=30%>�ڳ������</TD>     
		                        		     <TD CLASS="ctrow04"  width=20%>
		                        				<INPUT CLASS="form02" NAME="txtchl_rear" TYPE="Text" MAXLENGTH=30 SiZE=6 style='TEXT-ALIGN: right;' readonly></INPUT> ��
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtchl_rear_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		    <TD CLASS=ctrow03 width=75%>�Ҽ��������߰�����</TD>
		                        		    <TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtsmall_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
                                
                                    <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		    <TD CLASS=ctrow03 width=75%>���ݺ�������</TD>
		                        		    <TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="national_pension_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
                                
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>   
		                        		     <TD CLASS=ctrow03 rowspan="7" width=25%>Ư������</TD>    
		                        		     <TD CLASS=ctrow03  width=50%>�����</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtInsur_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>�Ƿ��</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtMed_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>������</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtEdu_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>   

		                        		<TR>
		                        		     <TD CLASS=ctrow03  width=50%>�����ڱ�</TD>       
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txth_house_fund_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>��α�</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtcontr_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>��ȥ��ʺ�</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtCeremony_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		
		                        		</TR>		                        		 
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>��(�Ǵ�ǥ�ذ���)</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtstd_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		
		                        		</TR>   
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		    <TD CLASS=ctrow03 width=75%>�����ҵ�ݾ�</TD>
		                        		    <TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtSub_income_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>���ο�������ҵ����</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtIndiv_anu_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>��������ҵ����</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtIndiv_anu2_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS=ctrow03 width=75%>���ڼ������ڵ�ҵ����</TD>
		                        		    <TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtInvest_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>�츮�����⿬�� �ҵ����</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtOur_stock_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>�ſ�ī�����</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtcard_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>�������ݼҵ����</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtRetire_pension" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>		                        		
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>�ܱ��αٷ��ڱ�����</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtFore_edu_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>		                        		
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>���ռҵ����ǥ��</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtTax_std_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=75%>���⼼��</TD>
		                        			<TD CLASS=ctrow04 width=25%><INPUT CLASS="form02" NAME="txtCalu_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>   
		                        		     <TD CLASS=ctrow03 rowspan="7" width=25%>���װ����׼��װ���</TD>    
		                        		     <TD CLASS=ctrow03  width=50%>�ٷμҵ�</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtincome_tax_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>�������Ա�</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txthouse_repay_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>�ܱ�����</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtFore_pay_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>��ġ�ڱݱ�αݼ��װ���</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtPolicontr_tax_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 		                        		
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>���װ�����</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txttax_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>���鼼��</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtRedu_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=ctrow03  width=50%>���ٳ������հ���</TD>     
		                        		     <TD CLASS="ctrow04"  width=25%>
		                        				<INPUT CLASS="form02" NAME="txtTax_Union_Ded" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>		                        		 
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		     <TD CLASS=ctrow03 width=25%></TD>    
		                        		     <TD CLASS=ctrow03 width=18% align=center>�ҵ漼</TD>       
		                        		     <TD CLASS=ctrow03 width=18% align=center>�ֹμ�</TD>       
		                        		     <TD CLASS=ctrow03 width=18% align=center>��Ư��</TD>       
		                        		     <TD CLASS=ctrow03 width=18% align=center>��</TD>       
		                        		</TR>       
		                        		<TR>
		                        		     <TD CLASS=ctrow03  width=25%>��������</TD>       
		                        		     <TD CLASS="ctrow04"  width=16%>
		                        				<INPUT CLASS="form02" NAME="txtDec_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=17%>
		                        				<INPUT CLASS="form02" NAME="txtDec_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     
		                        		     <TD CLASS="ctrow04"  width=17%>
		                        				<INPUT CLASS="form02" NAME="txtDec_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=18%>
		                        				<INPUT CLASS="form02" NAME="txtdec_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>
		                        		     <TD CLASS=ctrow03  width=25%>�ⳳ�μ���</TD>       
		                        		     <TD CLASS="ctrow04"  width=16%>
		                        				<INPUT CLASS="form02" NAME="txtold_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=17%>
		                        				<INPUT CLASS="form02" NAME="txtBefore_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     
		                        		     <TD CLASS="ctrow04"  width=17%>
		                        				<INPUT CLASS="form02" NAME="txtold_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=18%>
		                        				<INPUT CLASS="form02" NAME="txtold_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>
		                        		     <TD CLASS=ctrow03  width=25%>����¡������</TD>       
		                        		     <TD CLASS="ctrow04"  width=16%>
		                        				<INPUT CLASS="form02" NAME="txtincome_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=17%>
		                        				<INPUT CLASS="form02" NAME="txtRes_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     
		                        		     <TD CLASS="ctrow04"  width=17%>
		                        				<INPUT CLASS="form02" NAME="txtfarm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04"  width=18%>
		                        				<INPUT CLASS="form02" NAME="txtf_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        	</TABLE>
                                </TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
					<TR>
					    <TD height=10></TD>
					</TR>
					<TR>
					    <TD CLASS="ctrow06" align=center height=30>
							<img SRC="../ESSimage/button_15.gif" border="0" OnClick="vbscript: call goBackForm1()" name="printprev" alt='���ư���' onMouseOver="javascript:this.src='../ESSimage/button_r_15.gif';" onMouseOut="javascript:this.src='../ESSimage/button_15.gif';">
					    </TD>
					</TR>
					<TR>
					    <TD CLASS="ctrow06" align=center height=30>
							���� ����� ������ �ٸ� �� �ֽ��ϴ�.
					    </TD>
					</TR>
                </TABLE>
            </TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=2 cellPadding=0 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
</FORM>	

</BODY>
</HTML>
