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
Option Explicit                                                        '��: indicates that All variables must be declared in advance
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const BIZ_PGM_ID      = "e1303mb1.asp"						           '��: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1303ma2.asp"						           '��: Biz Logic ASP Name
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
<!-- #Include file="../ESSinc/lgvariables.inc" --> 
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "I", "H") %>
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    dim strSQL, IntRetCd

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
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
        lgKeyStream = lgKeyStream & "Q" & gColSep
    Elseif pOpt = "P" Then
        lgKeyStream = Trim(parent.parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
        lgKeyStream = lgKeyStream & "P" & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
    end if
End Sub        

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status

	Call LoadInfTB19029()
    Call InitComboBox()
    'parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call LayerShowHide(0)

    Call SetToolBar("00000")

    Call LockField(Document)

    Call DbQueryEmp(1)

End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : ������ ��ȯ�̳� ȭ���� ���� ��� �����ؾ� �� ���� ó�� 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	'Set gActiveElement = Nothing
    'Set Grid1 = Nothing
End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : ������ ��ȯ�̳� ȭ���� ���� ��� �����ؾ� �� ���� ó�� 
'========================================================================================
Function dodata(strdo)
	dim IntRetCD
	
'    IntRetCD = CommonQueryRs(" * "," HFA031T "," EMP_NO = '" & parent.txtEmp_no.Value & "' AND YEAR_YY = '" &  frm1.txtYear.Value &"'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

'	if  IntRetCD = true then
		if strdo="1" then
			IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
			If IntRetCD = vbNo Then
				Exit Function
			End If

			call dbquery("2")
		else 
			call dbquery("1")		
		end if
'	else
'	    Call DisplayMsgBox("800186","X","X","X")		
'	end if

End Function

Function DbQueryEmp(ppage)
    Dim strVal

    Err.Clear                                                                    '��: Clear err status

    DbQueryEmp = False                                                              '��: Processing is NG
    'If Grid1.ChkChange() Then Exit Function
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="      & "UID_M0001"                   '��: Query
    strVal = strVal     & "&txtKeyStream=" & lgKeyStream                   '��: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '��: Run biz logic

    DbQueryEmp = True                                                               '��: Processing is NG
End Function


Function DbQueryOk()

    Dim strVal
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
Function DbQuery(ppage)
	Dim strVal
	Dim strDate
    Err.Clear                                                                    '��: Clear err status
		
	DbQuery = False														         '��: Processing is NG
		
	if ChkField(Document, "2") then
		exit function
	end if

	if ppage = "2" Then

		Call MakeKeyStream("P")

		strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '��: Query
		strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key

		Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic
	Else
		
		Call MakeKeyStream("Q")
		
		With Frm1

	        if  Num_chk(.PAY_TAX_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .PAY_TAX_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.BONUS_TAX_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .BONUS_TAX_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.AFTER_BONUS_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .AFTER_BONUS_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.BEFORE_INCOME_TAX_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .BEFORE_INCOME_TAX_AMT.focus()
	            exit function
	        end if


	        if  Num_chk(.BEFORE_RES_TAX_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .BEFORE_RES_TAX_AMT.focus()
	            exit function
	        end if
	    
	    
	        if  Num_chk(.OLD_SUPP_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .OLD_SUPP_CNT.focus()
	            exit function
	        end if

	        if  Num_chk(.YOUNG_SUPP_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .YOUNG_SUPP_CNT.focus()
	            exit function
	        end if

	        if  Num_chk(.OLD_CNT1.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .OLD_CNT1.focus()
	            exit function
	        end if

	        if  Num_chk(.OLD_CNT2.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .OLD_CNT2.focus()
	            exit function
	        end if
	        
	        if  Num_chk(.PARIA_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .PARIA_CNT.focus()
	            exit function
	        end if


	        if  Num_chk(.CHL_REAR_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .CHL_REAR_CNT.focus()
	            exit function
	        end if

	        if  Num_chk(.MED_INSUR.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .MED_INSUR.focus()
	            exit function
	        end if

	        if  Num_chk(.EMP_INSUR.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .EMP_INSUR.focus()
	            exit function
	        end if

	        if  Num_chk(.OTHER_INSUR.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .OTHER_INSUR.focus()
	            exit function
	        end if

	        if  Num_chk(.MED_SPPORT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .MED_SPPORT.focus()
	            exit function
	        end if

	        if  Num_chk(.SPECI_MED.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .SPECI_MED.focus()
	            exit function
	        end if

	        if  Num_chk(.PER_EDU.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .PER_EDU.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY1_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY1_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY2_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY2_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY3_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY3_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY4_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY4_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY1_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY1_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY2_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY2_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY3_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY3_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY4_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY4_AMT.focus()
	            exit function
	        end if
	        
	        if  Num_chk(.HOUSE_FUND.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .HOUSE_FUND.focus()
	            exit function
	        end if

	        if  Num_chk(.LONG_HOUSE_LOAN_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .LONG_HOUSE_LOAN_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.txtLegal_contr_amt.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .txtLegal_contr_amt.focus()
	            exit function
	        end if

	        if  Num_chk(.txtApp_contr_amt.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .txtApp_contr_amt.focus()
	            exit function
	        end if

	        if  Num_chk(.INDIV_ANU2.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .INDIV_ANU2.focus()
	            exit function
	        end if

	        if  Num_chk(.NATIONAL_PENSION_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .NATIONAL_PENSION_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.CARD_USE_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .CARD_USE_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.txtInstitution_giro.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .txtInstitution_giro.focus()
	            exit function
	        end if

	        if  Num_chk(.txtRetire_pension.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .txtRetire_pension.focus()
	            exit function
	        end if
	        	        
	        if  Num_chk(.HOUSE_REPAY.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .HOUSE_REPAY.focus()
	            exit function
	        end if

	        if  Num_chk(.FORE_INCOME.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FORE_INCOME.focus()
	            exit function
	        end if

	        if  Num_chk(.FORE_PAY.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FORE_PAY.focus()
	            exit function
	        end if

	        if  Num_chk(.INCOME_REDU.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .INCOME_REDU.focus()
	            exit function
	        end if

	        if  Num_chk(.TAXES_REDU.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .TAXES_REDU.focus()
	            exit function
	        end if

	        if  Num_chk(.txtTax_Union_Ded.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .txtTax_Union_Ded.focus()
	            exit function
	        end if
	        
			Call LayerShowHide(1)
		'------ Developer Coding part (End )   -------------------------------------------------------------- 

			.txtMode.value        = "UID_M0002"                                        '��: Save
	'		.txtFlgMode.value     = lgIntFlgMode
	        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
		End With
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    End if 

    DbQuery  = True                                                               '��: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Dim strVal

    strVal = BIZ_PGM_ID1 & "?txtEmp_no=" & frm1.txtEmp_no.value
    strVal = strVal & "&txtYear=" & frm1.txtYear.value

    document.location = strVal
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

Sub txtYear_OnChange()
    Call DbQueryEmp(1)
End Sub

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

Sub FncPrintPrev()
	Dim strDate
    Dim strVal
    Err.Clear                                                                    '��: Clear err status

	With Frm1

        if  Date_chk(.year_yy.value & "0101", strDate) = True then
            '.year_yy.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .year_yy.focus()
            exit sub
        end if

        if  Date_chk(.Bas_dt.value, strDate) = True then
            .Bas_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .Bas_dt.focus()
            exit sub
        end if
    End With

    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

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
                        <TD>
                        <TABLE cellSpacing=1 cellPadding=0 width=100% border=0>
                            <TR>
							    <TD class="ftgray">&nbsp;
									<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">�⺻����</font></strong></td>                               
								<TD>
                            </TR>
							<tr> 
							    <td height="2"></td>
							</tr>
                            <TR><TD>
		                        <TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
							    <TR>
	            			    	<TD CLASS=ctrow01 >�������޿�</TD>
	            			    	<TD CLASS=ctrow02 align=left>
	            			    	    <INPUT CLASS="form01" NAME="PAY_TAX_AMT" ALT="�������޿�"  TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            			    	</TD>
	            			    	<TD CLASS=ctrow01 >��������</TD>
	            			    	<TD CLASS=ctrow02 align=left>
	            			    	    <INPUT CLASS="form01" NAME="BONUS_TAX_AMT" ALT="�������޿�" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            			    	</TD>
							    </TR>
							    <TR>
	            			    	<TD CLASS=ctrow01 >������</TD>
	            			    	<TD CLASS=ctrow02 align=left>
	            			    	    <INPUT CLASS="form01" NAME="AFTER_BONUS_AMT" ALT="������" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            			    	</TD>
	            			    	<TD CLASS=ctrow01>���ܱٷμҵ�</TD>
	            			    	<TD CLASS=ctrow02 align=left>
	            			    	    <INPUT CLASS="form01" NAME="FORE_INCOME"  ALT="���ܱٷμҵ�" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            			    	</TD>	            	        	
							    </TR>
							    <TR>
	            			    	<TD CLASS=ctrow01 >��Ÿ�ҵ�</TD>
	            			    	<TD CLASS=ctrow02 align=left  >
	            			    	    <INPUT CLASS="form01" NAME="other_income" ALT="��Ÿ�ҵ�" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            			    	</TD>
	            			    	<TD CLASS=ctrow01 >�ܱ��αٷ��ںи��������뿩��</TD>
	            			    	<TD CLASS=ctrow02 align=left >
	            			    	    <INPUT CLASS="ftgray" NAME="txtForeign_separate_tax_yn"  TYPE=CHECKBOX>
	            			    	</TD>	            	        	
							    </TR>                            
							    <TR>
	            			    	<TD CLASS=ctrow01 >�ⳳ�μҵ漼��</TD>
	            			    	<TD CLASS=ctrow02 align=left>
	            			    	    <INPUT CLASS="form01" NAME="BEFORE_INCOME_TAX_AMT" ALT="�ⳳ�μҵ漼��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            			    	</TD>                            
	            			    	<TD CLASS=ctrow01 >�ⳳ���ֹμ���</TD>
	            			    	<TD CLASS=ctrow02 align=left>
	            			    	    <INPUT CLASS="form01" NAME="BEFORE_RES_TAX_AMT" ALT="�ⳳ���ֹμ���" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            			    	</TD>
							    </TR>
	                			</TABLE>
		                    </TD></TR>
							<tr> 
							    <td height="5"></td>
							</tr>
                            <TR>
							    <TD class="ftgray">&nbsp;
									<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">�⺻����</font></strong></td>                               
								<TD>
                            </TR>
							<tr> 
							    <td height="2"></td>
							</tr>
                            <TR><TD>
		                        <TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
							    <TR>
	            			    	<TD CLASS=ctrow01 >����ڰ���</TD>
	            			    	<TD CLASS=ctrow02 align=left colspan=3>
	            			    	    <INPUT CLASS="ftgray" TYPE=CHECKBOX NAME="SPOUSE">
	            			    	</TD>
							    </TR>
							    <TR>
	            			    	<TD CLASS=ctrow01 >�ξ���(��)</TD>
	            			    	<TD CLASS=ctrow02 align=left>
	            			    	    <INPUT CLASS="form01" NAME="YOUNG_SUPP_CNT" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right' ID="Text1">
	            			    	</TD>
	            			    	<TD CLASS=ctrow01 >�ξ���(��)</TD>
	            			    	<TD CLASS=ctrow02 align=left>
	            			    	    <INPUT CLASS="form01" NAME="OLD_SUPP_CNT" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            			    	</TD>
							    </TR>
								</TABLE>
							</TD></TR>
							<tr> 
							    <td height="5"></td>
							</tr>
                            <TR>
							    <TD class="ftgray">&nbsp;
									<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">�߰�����</font></strong></td>                               
								<TD>
                            </TR>
							<tr> 
							    <td height="2"></td>
							</tr>
                            <TR><TD>
		                        <TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
						        <TR>
	            		        	<TD CLASS=ctrow01>��ο�����(65���̻�)</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="OLD_CNT1" ALT="��ο�����" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>����ڰ���</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="PARIA_CNT" ALT="����ڰ���" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>��ο�����(70���̻�)</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="OLD_CNT2" ALT="��ο�����" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>�γ��ڰ���</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="ftgray" TYPE=CHECKBOX NAME="LADY">
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>�ڳ����</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="CHL_REAR_CNT" ALT="�ڳ����" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>&nbsp;</TD>
	            		        	<TD CLASS=ctrow02 align=left>&nbsp;</TD>
						        </TR>                            
								</TABLE>
							</TD></TR>

							<tr> 
							    <td height="5"></td>
							</tr>
                            <TR>
							    <TD class="ftgray">&nbsp;
									<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">Ư������</font></strong></td>                               
								<TD>
                            </TR>
							<tr> 
							    <td height="2"></td>
							</tr>
                            <TR><TD>
		                        <TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
						        <TR>
	            		        	<TD CLASS=ctrow01>���ΰǰ������</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="MED_INSUR" ALT="���ΰǰ������" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>��뺸���</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="EMP_INSUR" ALT="��뺸���" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>���强�����</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="OTHER_INSUR" ALT="���强�����" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>��������뺸���</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="Disabled_INSUR" ALT="��������뺸���" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>�Ϲ��Ƿ��</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="MED_SPPORT" ALT="�Ϲ��Ƿ��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>����/�����/������Ƿ��</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="SPECI_MED" ALT="��ְ���Ƿ��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>���α�����</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="PER_EDU" ALT="���α�����" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>��ġ��������</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        		<INPUT CLASS="form01" NAME="FAMILY1_CNT" ALT="��" TYPE="Text" MAXLENGTH=3 SiZE=3 style='TEXT-ALIGN: right;'></INPUT>�� 
	            		        	    <INPUT CLASS="form01" NAME="FAMILY1_AMT" ALT="���߰�����"  TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>���߰�����</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        		<INPUT CLASS="form01" NAME="FAMILY2_CNT" ALT="��" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>�� 
	            		        	    <INPUT CLASS="form01" NAME="FAMILY2_AMT" ALT="����2������" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>���б�����</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        		<INPUT CLASS="form01" NAME="FAMILY3_CNT" ALT="��" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>�� 
	            		        	    <INPUT CLASS="form01" NAME="FAMILY3_AMT" ALT="����3������" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'></TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>�����Ư��������</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        		<INPUT CLASS="form01" NAME="FAMILY4_CNT" ALT="��" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>�� 
	            		        	    <INPUT CLASS="form01" NAME='FAMILY4_AMT' ALT="����4������" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>��ȥ��ʺ�</TD>
	            		        	<TD CLASS=ctrow02 align=left>
						            	<INPUT CLASS="form01" NAME="txtCeremony_cnt" ALT="Ƚ��" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>ȸ 
						            	<INPUT CLASS="form01" NAME="txtCeremony_amt" ALT="��ȥ��ʺ�" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" style='TEXT-ALIGN: right;'></INPUT>
	            		        	</TD>	            	        	
						        </TR>
						        <TR>
						            <TD CLASS="ctrow01" >������α�</TD>
						            <TD CLASS="ctrow02">
						        		<INPUT CLASS="form01" NAME="txtLegal_contr_amt" ALT="������α�" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22" style='TEXT-ALIGN: right;'></INPUT>
						        	</TD>
						            <TD CLASS="ctrow01" >��ġ�ڱݱ�α�</TD>
						            <TD CLASS="ctrow02">
						        		<INPUT CLASS="form01" NAME="txtPoli_contr_amt1" ALT="��ġ�ڱݱ�α�1" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22" style='TEXT-ALIGN: right;'></INPUT>
						        	</TD>
						        </TR>
						        <TR>
						            <TD CLASS="ctrow01" >Ư�ʱ�α�(100%)</TD>
						            <TD CLASS="ctrow02">
						        		<INPUT CLASS="form01" NAME="txtTaxLaw_contr_amt2" ALT="Ư�ʱ�α�(100%)" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22" style='TEXT-ALIGN: right;'></INPUT>
						        	</TD>
						            <TD CLASS="ctrow01" >Ư�ʱ�α�(50%)</TD>
						            <TD CLASS="ctrow02">
						        		<INPUT CLASS="form01" NAME="txtTaxLaw_contr_amt" ALT="Ư�ʱ�α�(50%)" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22" style='TEXT-ALIGN: right;'></INPUT>
						        	</TD>
						        </TR>		                        		
						        <TR>
						            <TD CLASS="ctrow01" >�츮�������ձ�α�</TD>
						            <TD CLASS="ctrow02">
						        		<INPUT CLASS="form01" NAME="txtOurstock_contr_amt" ALT="�츮�������ձ�α�" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22" style='TEXT-ALIGN: right;'></INPUT>
						        	</TD>
						            <TD CLASS="ctrow01" >������α�</TD>
						            <TD CLASS="ctrow02">
						        		<INPUT CLASS="form01" NAME="txtApp_contr_amt" ALT="������α�" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22" style='TEXT-ALIGN: right;'></INPUT>
						        	</TD>
						        </TR>	
						        <TR>
						            <TD CLASS="ctrow01" >�뵿���պ�</TD>
						            <TD CLASS="ctrow02">
						        		<INPUT CLASS="form01" NAME="txtPriv_contr_amt" ALT="�뵿���պ�" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22" style='TEXT-ALIGN: right;'></INPUT>
						        	</TD>
	            		        	<TD CLASS=ctrow01>��������/���Աݻ�ȯ��</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="HOUSE_FUND" ALT="��������/���Աݻ�ȯ��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>	
						        <TR>
	            		        	<TD CLASS=ctrow01>��������������Ա����ڻ�ȯ��(15��̸�)</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="LONG_HOUSE_LOAN_AMT" ALT="��������������Ա����ڻ�ȯ��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>��������������Ա����ڻ�ȯ��(15���̻�)</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="LONG_HOUSE_LOAN_AMT1" ALT="��������������Ա����ڻ�ȯ��1" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>		                    
								</TABLE>
							</TD></TR>
							<tr> 
							    <td height="5"></td>
							</tr>
                            <TR>
							    <TD class="ftgray">&nbsp;
									<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">��Ÿ�ҵ����</font></strong></td>                               
								<TD>
                            </TR>
							<tr> 
							    <td height="2"></td>
							</tr>
                            <TR><TD>
		                        <TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
						        <TR>
	            		        	<TD CLASS=ctrow01>���ο��ݺ��Ծ�(2000������)</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="INDIV_ANU" ALT="���ο��ݺ��Ծ�(2000������)" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>���ο��ݺ��Ծ�(2001������)</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="INDIV_ANU2" ALT="���ο��ݺ��Ծ�(2001������)" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>���ο��ݺ��Ծ�</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="NATIONAL_PENSION_AMT" ALT="���ο��ݺ��Ծ�" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>�츮�����⿬��</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="txtOur_stock_amt" ALT="�츮�����⿬��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>�����������ھ�</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="txtinvest2_sub_amt"  ALT="�����������ھ�" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>                            
	            		        	<TD CLASS=ctrow01>�ܱ��αٷ��ڱ�����</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="FORE_EDU_AMT" ALT="�ܱ��αٷ��ڱ�����" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>	            	        	
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>�ſ�/����/����ī��</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="CARD_USE_AMT"  ALT="�ſ�/����/����ī��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>���ݿ�����</TD>
	            		        	<TD CLASS=ctrow02 align=left >
	            		        	    <INPUT CLASS="form01" NAME="CARD2_USE_AMT"  ALT="���ݿ�����" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>�п������γ���</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="txtInstitution_giro"  ALT="�п������γ���" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>�������ݼҵ����</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="txtRetire_pension"  ALT="�������ݼҵ����" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR> 						                                    
								</TABLE>
							</TD></TR>
							<tr> 
							    <td height="5"></td>
							</tr>
                            <TR>
							    <TD class="ftgray">&nbsp;
									<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">���װ��� �� ���װ���</font></strong></td>                               
								<TD>
                            </TR>
							<tr> 
							    <td height="2"></td>
							</tr>
                            <TR><TD>
		                        <TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
						         <TR>
	            		        	<TD CLASS=ctrow01>�̺о��������Ա����ڻ�ȯ��</TD>
	            		        	<TD CLASS=ctrow02 align=left >
	            		        	    <INPUT CLASS="form01" NAME="HOUSE_REPAY"  ALT="�̺о��������Ա����ڻ�ȯ��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>                            
	            		        	<TD CLASS=ctrow01>�ܱ����μ���</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="FORE_PAY"  ALT="�ܱ����μ���" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>                            
						        <TR>
	            		        	<TD CLASS=ctrow01>���鼼�� �ҵ漼��</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="INCOME_REDU"  ALT="���鼼�� �ҵ漼��" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>���鼼�� ������</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="TAXES_REDU"  ALT="���鼼�� ������" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
						        </TR>
						        <TR>
	            		        	<TD CLASS=ctrow01>���ٳ������հ���</TD>
	            		        	<TD CLASS=ctrow02 align=left>
	            		        	    <INPUT CLASS="form01" NAME="txtTax_Union_Ded"  ALT="���ٳ������հ���" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            		        	</TD>
	            		        	<TD CLASS=ctrow01>&nbsp;</TD>
	            		        	<TD CLASS=ctrow02 align=left>&nbsp;</TD>
						        </TR>						        
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			 </TD>
		</TR>
		<TR valign=middle height=50>
			<TD align=center>
				<img src="../ESSimage/button_16.gif" onclick = "javascript:dodata('1')" alt="�����ڷ����"  onMouseOver="javascript:this.src='../ESSimage/button_r_16.gif';this.style.cursor='hand';" onMouseOut="javascript:this.src='../ESSimage/button_16.gif';"></a>
				<img src="../ESSimage/button_14.gif" onclick = "javascript:dodata('2')" alt="���" onMouseOver="javascript:this.src='../ESSimage/button_r_14.gif';this.style.cursor='hand';" onMouseOut="javascript:this.src='../ESSimage/button_14.gif';"></a>
			</TD>
		</TR>
    </TABLE>

    <TABLE cellSpacing=0 cellPadding=0 border=0 width=100%>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">
    <INPUT TYPE=HIDDEN NAME="txtres_no">
    <INPUT TYPE=HIDDEN NAME="txtdomi">
    <INPUT TYPE=HIDDEN NAME="txtaddr">
    <INPUT TYPE=HIDDEN NAME="txtentr_dt">
    <INPUT TYPE=HIDDEN NAME="txtretire_dt">

</FORM>	
</BODY>
</HTML>

