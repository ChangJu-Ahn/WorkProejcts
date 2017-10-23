<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>


<!-- #Include file="../../inc/IncServer.asp"  -->

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
    Emp_no = Trim(Request("txtEmp_no"))
    RYear = Trim(Request("txtYear"))
%>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "e1302mb1.asp"						           '��: Biz Logic ASP Name

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
'                        5.1 Common Method-1
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "Q", "H") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.cboYear.Value,"-","")) & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.cboYear.Value,"-","")) & gColSep
    end if



End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
Dim lgYear,i
dim strSQL, IntRetCd
    strSQL = " org_cd = " & FilterVar("1", "''", "S") & " AND pay_gubun = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE = " & FilterVar("*", "''", "S") & " "
    IntRetCD = CommonQueryRs(" year(close_dt) close_year "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
		lgYear = Replace(lgF0, Chr(11), "") 
	end if
	
	For i=lgYear To lgYear-10 step -1
		Call SetCombo(frm1.cboYear, i, i)
	next

	frm1.cboYear.value = CStr(lgYear)
    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status

    RYear = "<%=RYear%>"
    Emp_no = "<%=Emp_no%>"
    
    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call LayerShowHide(0)

    Call SetToolBar("10000")
    Call InitComboBox()

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
    Dim strSQL

    FncNext = False                                                              '��: Processing is OK
    Err.Clear                                                                    '��: Clear err status

    Call MakeKeyStream("N")
    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & "UID_M0001"                     '��: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"	                         '��: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '��: Run Biz 

    FncNext = True                                                               '��: Processing is OK
	
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
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=749 border=0 bgcolor=#ffffff>
                    <TR height=26 valign=middle>
                        <TD class=base1>���:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 tag=14></TD>
                        <TD class=base1>����:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>����:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>�μ�:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15  tag=14></TD>
                    </TR>   
                    <TR height=26 valign=middle>
                        <TD class=base1>���꿬��:
						    <SELECT Name="cboYear" tabindex=-1 STYLE="WIDTH: 100px">
						    </SELECT>
                        </TD>
                        <TD></TD>
		            	<TD class=base1></TD>
		            	<TD></TD>
                    </TR>

                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=0 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
                                <TR><TD>
		                        	<FIELDSET><LEGEND ALIGN="LEFT">��������/����¡������</LEGEND>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=40% valign=middle>����</TD>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=15% valign=middle>�޿�</TD>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=15% valign=middle>��</TD>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=15% valign=middle>������</TD>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=15% valign=middle>�հ�</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS=TDFAMILY_TITLE4 width=40%>1.���ٹ����ٷμҵ���Աݾ�</TD>
		                        		    <TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtNew_pay_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		    <TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtNew_bonus_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		    <TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtafter_bonus_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		    <TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txta_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=40%>2.���ٹ����ٷμҵ���Աݾ�</TD>
		                        			<TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtpay_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtbonus_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtold_after_bonus_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtb_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=40%>3.�ٷμҵ���Աݾ�</TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtincome_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=40%>4.�ٷμҵ����</TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txtincome_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=TDFAMILY_TITLE4 width=40%>5.�ٷμҵ�ݾ�</TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%></TD>
		                        			<TD CLASS=TDFAMILY4 width=15%><INPUT NAME="txthfa050t_income_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
        
                                <TR><TD>
		                        	<FIELDSET><LEGEND ALIGN="LEFT">��������</LEGEND>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="34%" colspan="2" align="middle">�����׸�</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="08%" align="middle">��������</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="08%" align="middle">������</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="34%" colspan="2" align="middle">�����׸�</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="08%" align="middle">��������</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="08%" align="middle">������</TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4 rowspan="4">6.�⺻����</TD>    
		                        		     <TD CLASS=TDFAMILY_TITLE4 >���ΰ���</TD>       
		                        		     <TD CLASS="TDFAMILY4" colspan="2">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtper_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 rowspan="5">7.�߰�����</TD>       
		                        		     <TD CLASS=TDFAMILY_TITLE4>����μ�</TD>       
		                        		     <TD CLASS="TDFAMILY4">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtparia_cnt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtparia_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>       
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4>�����(Y/N)</TD>       
		                        		     <TD CLASS="TDFAMILY4">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtspouse" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtspouse_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4>��ο���(65���̻�)</TD>      
		                        		     <TD CLASS="TDFAMILY4">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_cnt1" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" rowspan="2">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE4 >�ξ���(��55,��60���̻�)</TD>     
		                        		     <TD CLASS="TDFAMILY4" width="11%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtsupp_old_cnt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" rowspan="2">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtsupp_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4>��ο���(70���̻�)</TD>      
		                        		     <TD CLASS="TDFAMILY4">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_cnt2" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE4 >�ξ���(20������/�ʰ������)</TD>      
		                        		     <TD CLASS="TDFAMILY4" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtsupp_young_cnt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 >�γ��ڼ����ֿ���(Y/N)</TD>      
		                        		     <TD CLASS="TDFAMILY4" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtlady" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtlady_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE4 colspan="2">8.�Ҽ��������߰�����</TD>
		                        		     <TD CLASS="TDFAMILY4" colspan="2">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtsmall_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 >�ڳ������(6������)</TD>
		                        		     <TD CLASS="TDFAMILY4" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtchl_rear" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtchl_rear_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     
		                        		</TR>
		                        		<TR>   
		                        		     <TD CLASS=TDFAMILY_TITLE4 colspan="2">9.����������</TD>
		                        		     <TD CLASS="TDFAMILY4" colspan="2">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtd_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 colspan="2">&nbsp;</TD>
		                        		     <TD CLASS="TDFAMILY4" colspan="2">&nbsp;</TD>
		                        		     
		                        		</TR>
		                        		
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
        
                                <TR><TD>
		                        	<FIELDSET><LEGEND ALIGN="LEFT">Ư�����װ���</LEGEND>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="30%" colspan="3" align="middle">�����׸�</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="08%" align="middle">��������</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="08%" align="middle">������</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="38%" colspan="2" align="middle">�����׸�</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="08%" align="middle">��������</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="08%" align="middle">������</TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="6%" rowspan="10">10.Ư������</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="4%" rowspan="4">�����</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">�ǷẸ���</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtinsur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtmed_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="6%" rowspan="3">�����ڱ�</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="20%" >��������/���Աݻ�ȯ��</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa030t_house_fund_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%" rowspan="3">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa050t_house_fund_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">��뺸���</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa030t_emp_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa050t_emp_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="20%" >��������������Ա����ڻ�ȯ��(15��̸�)</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtlong_house_loan_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>		                        		     
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">��Ÿ���强�����</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa030t_other_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa050t_other_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="20%" >��������������Ա����ڻ�ȯ��(15���̻�)</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtlong_house_loan_amt1" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>	
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">��������뺸���</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa030t_disabled_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa050t_disabled_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="6%" rowspan="7">��α�</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="20%">������α�</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtlegal_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%" rowspan="7">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtcontr_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>			                        		     
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="4%" rowspan="2">�Ƿ��</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">�Ϲ��Ƿ��</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txttot_med_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%" rowspan="2">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtmed_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="20%">��ġ�ڱݱ�α�(04/3/11 ����)</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtPoli_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">����/�����/������Ƿ��</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtspeci_med_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="20%">��ġ�ڱݱ�α�(04/3/12 ����)</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtPoli_contr_amt1" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
			                        		     <TD CLASS=TDFAMILY_TITLE4  width="4%" rowspan="3">������</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">���α�����</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtper_edu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%" rowspan="3">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtedu_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" >Ư�ʱ�α�</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%" >		                        		     
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtTaxLaw_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>		                        		     
		                        		     </TD>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">����������</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtedu_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" >�츮�������ձ�α�</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%" >		                        		     
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtOurstock_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>		                        		     
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="10%">�����Ư��������</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtDisabled_edu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="20%">������α�</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtapp_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="2">��ȥ��ʺ�</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="hfa030t_Ceremony_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%" > 
		                        				<INPUT CLASS="NUM_FIELD" NAME="hfa050t_Ceremony_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>		                        		     
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="20%">�뵿���պ�</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtpriv_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="3"  >11.�� �Ǵ� ǥ�ذ���</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%" colspan="5"></TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtstd_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>		                        				                        		
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="3">12.���ο��������</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa030t_indiv_anu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa050t_indiv_anu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="2">13.�츮�����⿬��</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%"></TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtOur_stock_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  colspan="3" width="26%">14.���ڼҵ����</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%"></TD>		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtinvest_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>

		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="2">15.ī��ҵ����</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%"></TD>		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtcard_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
											 <TD CLASS=TDFAMILY_TITLE4  colspan="3" width="26%">���ݺ�������</TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa030t_National_pension_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="12%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthfa050t_National_pension_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>		                        		
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="50%" colspan="2">�ܱ��α�����/������</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%"></TD>		                        		     
		                        		     <TD CLASS="TDFAMILY4" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtFore_edu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
                      					<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="3"></TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%">		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%">
		                        		     </TD>
		                        		
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="50%" colspan="2">16.�ҵ������</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%"></TD>		                        		     
		                        		     <TD CLASS="TDFAMILY4" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtsum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>		                        		
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="3">17.�ҵ����ǥ��</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%">		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txttax_std_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="2">18.���⼼��</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%"></TD>		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtcalu_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="6%" colspan="2" rowspan="4">19.���װ���</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="20%" >�ٷμҵ�</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%" ></TD>		                        		     		                        		     		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%" align=left>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtincome_tax_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width="26%" colspan="2">�����ڱ����Ա����ڻ�ȯ��</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%"></TD>		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%" >
		                        				<INPUT CLASS="NUM_FIELD" NAME="txthouse_repay_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="20%" >�ܱ����μ��װ���</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%" ></TD>		                        		     		                        		     		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%" align=left>
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtFore_pay_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>		                        		
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="26%" colspan="2">20.���װ�����</TD>
		                        		     <TD CLASS="TDFAMILY4" width="24%"></TD>		                        		     
		                        		     <TD CLASS="TDFAMILY4" width="24%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txttax_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>

                                <TR><TD>
		                        	<FIELDSET><LEGEND ALIGN="LEFT">��漼��/����¡������</LEGEND>
		                        	<TABLE  border=0 width="100%" cellSpacing=1 cellPadding=0>
		                        		 <TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width=40% valign="middle">����</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width=15% valign="middle">�ҵ漼</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width=15% valign="middle">�ֹμ�</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width=15% valign="middle">��Ư��</TD>
		                        		     <TD CLASS=TDFAMILY_TITLE4 width=15% valign="middle">��</TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="28%">21.���꼼��</TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtdec_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtdec_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtdec_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtdec_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="28%">22.���ٹ���¡������</TD>  
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtnew_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtnew_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtnew_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtincome_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="28%">23.�����ٹ�������</TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtold_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=TDFAMILY_TITLE4  width="28%">24.¡���ؾ��Ҽ���</TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtincome_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtres_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtfarm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="TDFAMILY4" width="18%">
		                        				<INPUT CLASS="NUM_FIELD" NAME="txtf_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 tag="24" style='TEXT-ALIGN: right;'></INPUT>
		                        		     </TD>
		                        		</TR>

		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=2 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
</FORM>	

</BODY>
</HTML>
