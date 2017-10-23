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
<!-- #Include file="../../inc/IncServer.asp"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/CommStyleSheet.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
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
<!-- #Include file="../../inc/lgvariables.inc" --> 

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

<!-- #Include file="../../inc/uniSimsClassID.inc" -->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=722 border=0 bgcolor=#ffffff>
                    <TR height=26 valign=middle>
                        <TD class=base1>���:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=15 ></TD>
                        <TD class=base1>����:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10 ></TD>
                        <TD class=base1>����:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10 ></TD>
                        <TD class=base1>�μ�:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15 ></TD>
                    </TR>   
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=0 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		    <TD CLASS=TDFAMILY_TITLE6 width=75%>�������޿�</TD>
		                        		    <TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtincome_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>�ٷμҵ����</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtincome_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>�������ٷμҵ�ݾ�</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtIncome_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE6 rowspan="4" width=25%>�⺻����</TD>    
		                        		     <TD CLASS=TDFAMILY_TITLE6 width=30%>����</TD>       
		                        		     <TD CLASS="TDFAMILY6"  width=20%></TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtper_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>       
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=30%>�����</TD>       
		                        		     <TD CLASS="TDFAMILY6"  width=20%></TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtspouse_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=30%>�ξ簡��</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=20%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtsupp_old_cnt" TYPE="Text" MAXLENGTH=30 SiZE=6 tag="24" style='TEXT-ALIGN: right;'></INPUT> ��
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtsupp_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6 rowspan="5" width=25%>�߰�����</TD>    
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=30%>��ο��(65���̻�)</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=20%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_cnt1" TYPE="Text" MAXLENGTH=30 SiZE=6 tag="24" style='TEXT-ALIGN: right;'></INPUT> ��
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=50% rowspan="2">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_sub_amt1" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24"  style='TEXT-ALIGN: right;height:40'></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=30%>��ο��(70���̻�)</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=20%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_cnt2" TYPE="Text" MAXLENGTH=30 SiZE=6 tag="24" style='TEXT-ALIGN: right;'></INPUT> ��
		                        		     </TD>
		                        		</TR> 		                        		 
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=30%>�����</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=20%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtparia_cnt" TYPE="Text" MAXLENGTH=30 SiZE=6 tag="24" style='TEXT-ALIGN: right;'></INPUT> ��
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtparia_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=30%>�γ���</TD>       
		                        		     <TD CLASS="TDFAMILY6"  width=20%></TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtlady_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=30%>�ڳ������</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=20%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtchl_rear" TYPE="Text" MAXLENGTH=30 SiZE=6 tag="24" style='TEXT-ALIGN: right;'></INPUT> ��
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtchl_rear_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		    <TD CLASS=TDFAMILY_TITLE6 width=75%>�Ҽ��������߰�����</TD>
		                        		    <TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtsmall_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6 rowspan="7" width=25%>Ư������</TD>    
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>�����</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtInsur_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>�Ƿ��</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtMed_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>������</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtEdu_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   

		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>�����ڱ�</TD>       
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txth_house_fund_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>��α�</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtcontr_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>��ȥ��ʺ�</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtCeremony_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		
		                        		</TR>		                        		 
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>��(�Ǵ�ǥ�ذ���)</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtstd_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		
		                        		</TR>   
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		    <TD CLASS=TDFAMILY_TITLE6 width=75%>�����ҵ�ݾ�</TD>
		                        		    <TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtSub_income_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>���ο�������ҵ����</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtIndiv_anu_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>��������ҵ����</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtIndiv_anu2_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS=TDFAMILY_TITLE6 width=75%>���ڼ������ڵ�ҵ����</TD>
		                        		    <TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtInvest_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>�츮�����⿬�� �ҵ����</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtOur_stock_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>�ſ�ī�����</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtcard_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>�ܱ��αٷ��ڱ�����</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtFore_edu_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>		                        		
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>���ռҵ����ǥ��</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtTax_std_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE6 width=75%>���⼼��</TD>
		                        			<TD CLASS=TDFAMILY6 width=25%><INPUT NAME="txtCalu_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6 rowspan="6" width=25%>���װ����׼��װ���</TD>    
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>�ٷμҵ�</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtincome_tax_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>�������Ա�</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthouse_repay_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>�ܱ�����</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtFore_pay_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>  
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>��ġ�ڱݱ�αݼ��װ���</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtPolicontr_tax_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR> 		                        		
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>���װ�����</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txttax_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=50%>���鼼��</TD>     
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtRedu_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        	</TABLE>
                                </TD></TR>
                                <TR><TD>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE6 width=25%></TD>    
		                        		     <TD CLASS=TDFAMILY_TITLE6 width=16%>�ҵ漼</TD>       
		                        		     <TD CLASS=TDFAMILY_TITLE6 width=17%>�ֹμ�</TD>       
		                        		     <TD CLASS=TDFAMILY_TITLE6 width=17%>��Ư��</TD>       
		                        		     <TD CLASS=TDFAMILY_TITLE6 width=25%>��</TD>       
		                        		</TR>       
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=25%>��������</TD>       
		                        		     <TD CLASS="TDFAMILY6"  width=16%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtDec_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=17%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtDec_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     
		                        		     <TD CLASS="TDFAMILY6"  width=17%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtDec_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtdec_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=25%>�ⳳ�μ���</TD>       
		                        		     <TD CLASS="TDFAMILY6"  width=16%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=17%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtBefore_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     
		                        		     <TD CLASS="TDFAMILY6"  width=17%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE6  width=25%>����¡������</TD>       
		                        		     <TD CLASS="TDFAMILY6"  width=16%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtincome_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=17%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtRes_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     
		                        		     <TD CLASS="TDFAMILY6"  width=17%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtfarm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=13 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY6"  width=25%>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtf_amt" TYPE="Text" MAXLENGTH=30 SiZE=15 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR> 
										<TR>
										    <TD colspan=5 align=center height=50>
												<INPUT style="WIDTH: 150px; HEIGHT: 20px" TYPE=button NAME=gobackForm VALUE="���ư���" OnClick="vbscript:call goBackForm1()">
										    </TD>
										</TR>
										<TR>
										    <TD colspan=5 align=center height=50>
												���� ����� ������ �ٸ� �� �ֽ��ϴ�.
										    </TD>
										</TR>
		                        	
		                        	</TABLE>
                                </TD></TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
           
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=2 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
</FORM>	

</BODY>
</HTML>
