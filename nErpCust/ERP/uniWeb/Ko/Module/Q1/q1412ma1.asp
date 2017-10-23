<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1412MA1
'*  4. Program Name         : �������ý��� ���ȭ�� 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>T�˻� � ������ �ý��� ���</TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                                             

Const PGM_JUMP_ID1 = "Q1413MA1.asp"	'��������� 
Const PGM_JUMP_ID2 = "Q1413MA2.asp"	'��������� 
Const PGM_JUMP_ID3 = "Q1413MA3.asp"	'��������� 
Const PGM_JUMP_ID4 = "Q1413MA4.asp"	'�������2ȸ 
Const PGM_JUMP_ID5 = "Q1411MA1"		'����ȭ�� 
Const PGM_JUMP_ID6 = "Q1413MA5.asp"	'�跮������ 
Const PGM_JUMP_ID7 = "Q1413MA6.asp"	'�跮������ 
	
Dim lgNextNo					'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo					' ""

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgMpsFirmDate
Dim lgLlcGivenDt				
Dim IsOpenPop          
Dim gSelframeFlg

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                                               	'��: Indicates that current mode is Create mode
    lgIntGrpCount = 0                                                     	  	'��: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    	
    IsOpenPop = False						'��: ����� ���� �ʱ�ȭ 
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

	Dim ICase
	Dim IType1
	Dim IType2
	Dim IType3
	Dim IType4
	Dim IType5
	Dim IAssr1
	Dim IAssr2
	Dim IAssr3
	Dim IAssr4
	Dim IAssr5

	If ReadCookie("txtInsCase") <> "" Then					
		ICase= ReadCookie("txtInsCase")
			Select Case ICase
				Case 01:
					 frm1.txtBestInsCase.Value = "���"
				Case 02:
					 frm1.txtBestInsCase.Value = "�跮"
			 End Select
	End If
	
	If ReadCookie("txtInsType1") <> "" Then
		IType1= ReadCookie("txtInsType1")		
			Select Case IType1
				Case 0100:
					 frm1.txtBestType1.Value = "������"
				Case 02:
					 frm1.txtBestType1.Value = "������"
				Case 0300:
					 frm1.txtBestType1.Value = "������"
				Case 0400:
					 frm1.txtBestType1.Value = "���ӻ�����"
				Case 0500:
					 frm1.txtBestType1.Value = "�ŷڼ��˻�"
				Case 0600:
					 frm1.txtBestType1.Value = "�˻��� ��Ȯ�� ����"
				Case 0700:
					 frm1.txtBestType1.Value = "üũ�˻�"
				Case 0800:
					 frm1.txtBestType1.Value = "���� ������ ����˻�"
				Case 0900:
					 frm1.txtBestType1.Value = "�Һ��� �����˻�"
				Case 1000:
					 frm1.txtBestType1.Value = "���˻�"
				Case 1100:
					 frm1.txtBestType1.Value = "�����˻�"			
			 End Select
	End If	
	If ReadCookie("txtInsType2") <> "" Then
		IType2= ReadCookie("txtInsType2")
			Select Case IType2
				Case 0100:
					 frm1.txtBestType2.Value = "������"
				Case 02:
					 frm1.txtBestType2.Value = "������"				
				Case 0300:
					 frm1.txtBestType2.Value = "������"
				Case 0400:
					 frm1.txtBestType2.Value = "���ӻ�����"
				Case 0500:
					 frm1.txtBestType2.Value = "�ŷڼ��˻�"
				Case 0600:
					 frm1.txtBestType2.Value = "�˻��� ��Ȯ�� ����"
				Case 0700:
					 frm1.txtBestType2.Value = "üũ�˻�"
				Case 0800:
					 frm1.txtBestType2.Value = "���� ������ ����˻�"
				Case 0900:
					 frm1.txtBestType2.Value = "�Һ��� �����˻�"
				Case 1000:
					 frm1.txtBestType2.Value = "���˻�"
				Case 1100:
					 frm1.txtBestType2.Value = "�����˻�"			
			 End Select
	End If
	
	
	If ReadCookie("txtInsType3") <> "" Then
		IType3= ReadCookie("txtInsType3")
			Select Case IType3
				Case 0100:
					 frm1.txtBestType3.Value = "������"
				Case 02:
					 frm1.txtBestType3.Value = "������"				
				Case 0300:
					 frm1.txtBestType3.Value = "������"
				Case 0400:
					 frm1.txtBestType3.Value = "���ӻ�����"
				Case 0500:
					 frm1.txtBestType3.Value = "�ŷڼ��˻�"
				Case 0600:
					 frm1.txtBestType3.Value = "�˻��� ��Ȯ�� ����"
				Case 0700:
					 frm1.txtBestType3.Value = "üũ�˻�"
				Case 0800:
					 frm1.txtBestType3.Value = "���� ������ ����˻�"
				Case 0900:
					 frm1.txtBestType3.Value = "�Һ��� �����˻�"
				Case 1000:
					 frm1.txtBestType3.Value = "���˻�"
				Case 1100:
					 frm1.txtBestType3.Value = "�����˻�"			
			 End Select
	End If	
	If ReadCookie("txtInsType4") <> "" Then
		IType4= ReadCookie("txtInsType4")
			Select Case IType4
				Case 0100:
					 frm1.txtBestType4.Value = "������"
				Case 02:
					 frm1.txtBestType4.Value = "������"				
				Case 0300:
					 frm1.txtBestType4.Value = "������"
				Case 0400:
					 frm1.txtBestType4.Value = "���ӻ�����"
				Case 0500:
					 frm1.txtBestType4.Value = "�ŷڼ��˻�"
				Case 0600:
					 frm1.txtBestType4.Value = "�˻��� ��Ȯ�� ����"
				Case 0700:
					 frm1.txtBestType4.Value = "üũ�˻�"
				Case 0800:
					 frm1.txtBestType4.Value = "���� ������ ����˻�"
				Case 0900:
					 frm1.txtBestType4.Value = "�Һ��� �����˻�"
				Case 1000:
					 frm1.txtBestType4.Value = "���˻�"
				Case 1100:
					 frm1.txtBestType4.Value = "�����˻�"				 
			 End Select
	End If
	If ReadCookie("txtInsType5") <> "" Then
		IType5= ReadCookie("txtInsType5")
			Select Case IType5
				Case 0100:
					 frm1.txtBestType5.Value = "������"
				Case 02:
					 frm1.txtBestType5.Value = "������"
				Case 0300:
					 frm1.txtBestType5.Value = "������"
				Case 0400:
					 frm1.txtBestType5.Value = "���ӻ�����"
				Case 0500:
					 frm1.txtBestType5.Value = "�ŷڼ��˻�"
				Case 0600:
					 frm1.txtBestType5.Value = "�˻��� ��Ȯ�� ����"
				Case 0700:
					 frm1.txtBestType5.Value = "üũ�˻�"
				Case 0800:
					 frm1.txtBestType5.Value = "���� ������ ����˻�"
				Case 0900:
					 frm1.txtBestType5.Value = "�Һ��� �����˻�"
				Case 1000:
					 frm1.txtBestType5.Value = "���˻�"
				Case 1100:
					 frm1.txtBestType5.Value = "�����˻�"			
			 End Select
	End If
	
	If ReadCookie("txtInsAssureance1") <> "" Then
		IAssr1= ReadCookie("txtInsAssureance1")
			Select Case IAssr1
				Case 01:
					 frm1.txtAssureance1.Value = "AQOL����"
				Case 02:
					 frm1.txtAssureance1.Value = "LTPD����"
			 End Select			 
	End If

	If ReadCookie("txtInsAssureance2") <> "" Then
		IAssr2= ReadCookie("txtInsAssureance2")
			Select Case IAssr2
				Case 01:
					 frm1.txtAssureance2.Value = "AQOL����"
				Case 02:
					 frm1.txtAssureance2.Value = "LTPD����"
			 End Select		
	End If
	If ReadCookie("txtInsAssureance3") <> "" Then
		IAssr3= ReadCookie("txtInsAssureance3")
			Select Case IAssr3
				Case 01:
					 frm1.txtAssureance3.Value = "AQOL����"
				Case 02:
					 frm1.txtAssureance3.Value = "LTPD����"
			 End Select		
	End If
	If ReadCookie("txtInsAssureance4") <> "" Then
		IAssr4= ReadCookie("txtInsAssureance4")
			Select Case IAssr4
				Case 01:
					 frm1.txtAssureance4.Value = "AQOL����"
				Case 02:
					 frm1.txtAssureance4.Value = "LTPD����"
			 End Select		
	End If
	If ReadCookie("txtInsAssureance5") <> "" Then
		IAssr5= ReadCookie("txtInsAssureance5")
			Select Case IAssr5
				Case 01:
					 frm1.txtAssureance5.Value = "AQOL����"
				Case 02:
					 frm1.txtAssureance5.Value = "LTPD����"
			 End Select		
	End If	
	
	If ReadCookie("txtFitnessDegree1") <> "" Then
		frm1.txtBestDegree1.Value = ReadCookie("txtFitnessDegree1")
	End If
	If ReadCookie("txtFitnessDegree2") <> "" Then
		frm1.txtBestDegree2.Value = ReadCookie("txtFitnessDegree2")
	End If
	If ReadCookie("txtFitnessDegree3") <> "" Then
		frm1.txtBestDegree3.Value = ReadCookie("txtFitnessDegree3")
	End If
	If ReadCookie("txtFitnessDegree4") <> "" Then
		frm1.txtBestDegree4.Value = ReadCookie("txtFitnessDegree4")
	End If
	If ReadCookie("txtFitnessDegree5") <> "" Then
		frm1.txtBestDegree5.Value = ReadCookie("txtFitnessDegree5")
	End If
	
		
	frm1.txtBestInsCaseType.Value = frm1.txtBestInsCase.Value & " " & frm1.txtBestType1.Value & " " & frm1.txtAssureance1.Value
	frm1.txtJumpCd.Value = ICase & IType1 & IAssr1				'Jump�� ���� �ڵ� 
	
	WriteCookie "txtInsCase", ""
	WriteCookie "txtInsType1", ""
	WriteCookie "txtInsType2", ""
	WriteCookie "txtInsType3", ""
	WriteCookie "txtInsType4", ""
	WriteCookie "txtInsType5", ""
	WriteCookie "txtFitnessDegree1", ""
	WriteCookie "txtFitnessDegree2", ""
	WriteCookie "txtFitnessDegree3", ""
	WriteCookie "txtFitnessDegree4", ""
	WriteCookie "txtFitnessDegree5", ""
	
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function InsTypeResult()
	Dim sPGM
	Dim arrRet
	Dim TypeNumber

	Select Case frm1.txtJumpCd.Value    	
		Case "010100":
			sPGM = PGM_JUMP_ID1				'��� �������˻� 
		
		Case "010201":
   				sPGM = PGM_JUMP_ID2		'��� �������˻�AOQL���� 
   		
		Case "010202":
   			sPGM = PGM_JUMP_ID2		'��� �������˻�LTPD���� 

		Case "010300":    		
    		sPGM = PGM_JUMP_ID3		'��� �������˻� 

		'Case "010100_":
    		'sPGM = PGM_JUMP_ID4		'��� ������2ȸ�˻� 
    		
    		
    	Case "020100":
    		sPGM = PGM_JUMP_ID6		'�跮 �������˻� 
		
		Case "020300":
    			sPGM = PGM_JUMP_ID7		'�跮 �������˻� 

		Case "010400":				'��� ���ӻ������˻� 
			Call DisplayMsgBox("229922", "X", "X", "X") 		'����� �������� �ʽ��ϴ� 
			Exit Function	
		Case Else
			Call DisplayMsgBox("229922", "X", "X", "X") 		'����� �������� �ʽ��ϴ� 
			Exit Function					
	End Select		
	
	Navigate sPGM
			
	WriteCookie "txtInsTypeNumber", (TypeNumber)

End Function

'=============================================  2.3.3()  ======================================
'=	Event Name : ReturnClick
'=	Event Desc :
'========================================================================================================
Function ReturnClick()
	PgmJump(PGM_JUMP_ID5)
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'��: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'��: Lock  Suitable  Field
	Call InitVariables																'��: Initializes local global variables
   	Call SetDefaultVal
   	'----------  Coding part  -------------------------------------------------------------
    	
   	Call SetToolbar("10000000000111")
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    	
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	FncQuery = False
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	FncNew = False
End Function

'========================================================================================
' Function Name : Fnc
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	FncDelete = False
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	FncSave = False
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = False
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = False
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	FncInsertRow = False
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = False
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call Parent.FncPrint()
	FncPrint = True
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
	FncPrev = False
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	FncNext = False
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
	Call parent.FncExport(Parent.C_SINGLE)					'��: ȭ�� ���� 
	FncExcel = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	FncFind = False
	Call parent.FncFind(Parent.C_SINGLE, False)     
	FncFind = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> BORDER=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>������ �˻��� ���</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
						    	</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD  VALIGN="TOP" WIDTH=100% CLASS="Tab11">
			<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%> </TD>
				</TR>
				<TR>
					<TD VALIGN="top" WIDTH="100%" HEIGHT=*>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS="TD5" HEIGHT=10 NOWRAP>������ �˻���</TD>
									<TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%">
										<INPUT NAME="txtBestInsCaseType" SIZE="32" MAXLENGTH="25" ALT="������ �˻���" TAG="24" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" HEIGHT=10 NOWRAP></TD>
									<TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%"></TR>
								<TR>
									<TD CLASS="TD5" HEIGHT=10 NOWRAP>������ ��ǰ����</TD>
									<TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%">
										<INPUT NAME="txtBestInsCase" SIZE="20" MAXLENGTH="20" ALT="������ ��ǰ����" TAG="24" ></TD>	
								</TR>
								<TR>
									<TD CLASS="TD5" HEIGHT=20 NOWRAP></TD>
									<TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%"></TD>
								
								</TR>
								<!-- /* Issue: ȭ�� ���� - START %/-->
								<TR>
									<TD CLASS="TD5" HEIGHT=20 NOWRAP></TD>
									<TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%">
										<INPUT NAME="lblBestType" value="�˻���" SIZE="20" STYLE="border-color: SteelBlue; BACKGROUND-COLOR: SteelBlue; Text-Align: center; Color: White">
										<INPUT NAME="lblAssureance" value="�������" SIZE="20" STYLE="border-color: SteelBlue; BACKGROUND-COLOR: SteelBlue; Text-Align: center; Color: White">
								 		<INPUT NAME="lblBestDegree" value="���յ�" SIZE="10" STYLE="border-color: SteelBlue; BACKGROUND-COLOR: SteelBlue; Text-Align: center; Color: White">
								 	</TD>
								</TR>
								
								<TR>
									<TD CLASS="TD5" HEIGHT=10 NOWRAP>1.</TD>
									<TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%">
										<INPUT NAME="txtBestType1" SIZE="20" MAXLENGTH="20" TAG="24">
										<INPUT NAME="txtAssureance1" SIZE="20" MAXLENGTH="10" TAG="24">
								 		<INPUT NAME="txtBestDegree1" SIZE="10" MAXLENGTH="10" TAG="24" STYLE="Text-Align: Right">
								 	</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" HEIGHT=10 NOWRAP>2.</TD>
									<TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%">
										<INPUT NAME="txtBestType2" SIZE="20" MAXLENGTH="20" TAG="24">
									 	<INPUT NAME="txtAssureance2" SIZE="20" MAXLENGTH="10" TAG="24">
									 	<INPUT NAME="txtBestDegree2" SIZE="10" MAXLENGTH="10" TAG="24" STYLE="Text-Align: Right"></TD>
								</TR>								
								<TR>
									<TD CLASS="TD5" HEIGHT=10 NOWRAP>3.</TD>
									<TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%">
										<INPUT NAME="txtBestType3" SIZE="20" MAXLENGTH="20" TAG="24">
									 	<INPUT NAME="txtAssureance3" SIZE="20" MAXLENGTH="10" TAG="24">
									 	<INPUT NAME="txtBestDegree3" SIZE="10" MAXLENGTH="10" TAG="24" STYLE="Text-Align: Right"></TD>
								</TR>
								<TR>
									 <TD CLASS="TD5" HEIGHT=10 NOWRAP>4.</TD>
									 <TD CLASS="TD656" HEIGHT=10 NOWRAP style="width:100%">
										<INPUT NAME="txtBestType4" SIZE="20" MAXLENGTH="20" TAG="24">
										<INPUT NAME="txtAssureance4" SIZE="20" MAXLENGTH="10" TAG="24">
										<INPUT NAME="txtBestDegree4" SIZE="10" MAXLENGTH="10" TAG="24" STYLE="Text-Align: Right"></TD>
								</TR>
								<TR>
									 <TD CLASS="TD5" HEIGHT=10 NOWRAP>5.</TD>
									 <TD CLASS="TD656" HEIGHT=10 NOWRAP colspan=3 style="width:100%">
										<INPUT NAME="txtBestType5" SIZE="20" MAXLENGTH="20" TAG="24">
									 	<INPUT NAME="txtAssureance5" SIZE="20" MAXLENGTH="10" TAG="24">
									 	<INPUT NAME="txtBestDegree5" SIZE="10" MAXLENGTH="10" TAG="24" STYLE="Text-Align: Right"></TD>
								</TR>
								<!-- /* Issue: ȭ�� ���� - END %/-->
							</TABLE>
						</FIELDSET>
					</TD>	
				</TR>
			</TABLE>
		</DIV>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
    	<TD WIDTH="100%">
    		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
	   			<TR>
	   				<TD WIDTH=10>&nbsp;</TD>
    				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:InsTypeResult">�˻���	����</A>&nbsp;|&nbsp;<A href="vbscript:ReturnClick()">������ �ý��� ����</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
    			</TR>
    		</TABLE>
    	</TD>
    </TR>
    	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtJumpCd" TAG="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

