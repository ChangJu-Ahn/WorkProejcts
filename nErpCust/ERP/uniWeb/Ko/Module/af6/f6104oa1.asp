<%@ LANGUAGE="VBSCRIPT" %>


<!--'**********************************************************************************************
'*  1. Module Name          : Finance
'*  2. Function Name        : Finance Management
'*  3. Program ID           : f6104ma1.asp
'*  4. Program Name         : ���ޱ�CheckList��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/18
'*  8. Modified date(Last)  : 2003/01/08
'*  9. Modifier (First)     : Hersheys
'* 10. Modifier (Last)      : Kim Chang Jin
'* 11. Comment              :
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################

'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<!--
'=============================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                              '��: indicates that All variables must be declared in advance 

'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################


'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* 


'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

'Const BIZ_PGM_ID = "f6104mb1.asp"  

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt											 '��: �����Ͻ� ���� ASP���� �����ϹǷ� Dim 

Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim IsOpenPop          

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 



'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 


'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 

Sub SetDefaultVal()
<%
Dim dtToday
dtToday = GetSvrDate
%>

	frm1.fpDateTime1.Text = UniConvDateAToB("<%=dtToday%>", Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.fpDateTime2.Text = UniConvDateAToB("<%=dtToday%>", Parent.gServerDateFormat,Parent.gDateFormat)
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'*****************************************  2.1 Pop-Up �Լ�   ********************************************
'	���: Pop-Up 
'********************************************************************************************************* 
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0, 4
			arrParam(0) = "����� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA"	 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition

			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "������ڵ�"				' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"					' Field��(0)
			arrField(1) = "BIZ_AREA_NM"					' Field��(0)
    
			arrHeader(0) = "������ڵ�"				' Header��(0)
			arrHeader(1) = "������"				' Header��(0)
		Case 3	'������� 
			If frm1.txtPaymType.className = Parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = frm1.txtPaymType.Alt
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD AND B.SEQ_NO = 2 AND B.REFERENCE = " & FilterVar("PP", "''", "S") & "  "
			arrParam(5) = frm1.txtPaymType.Alt
	
			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
				    
			arrHeader(0) = frm1.txtPaymType.Alt
			arrHeader(1) = frm1.txtPaymTypeNm.Alt			
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


'===========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================== 

'++++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------

Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' ����� 
				.txtBizAreaCD.focus
			Case 1		' ���۰ŷ�ó 
				.txtFromBpCd.focus
			Case 2		' ����ŷ�ó 
				.txtToBpCd.focus
			Case 3		' ������� 
				.txtPaymType.focus
			Case 4		' �����1
				.txtBizAreaCD1.focus
		End Select
	End With
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' ����� 
				.txtBizAreaCd.value = arrRet(0)
				.txtBizAreaNm.value = arrRet(1)
				.txtBizAreaCD.focus
			Case 1		' ���۰ŷ�ó 
				.txtFromBpCd.value = arrRet(0)
				.txtFromBpNm.value = arrRet(1)
				.txtFromBpCd.focus
			Case 2		' ����ŷ�ó 
				.txtToBpCd.value = arrRet(0)
				.txtToBpNm.value = arrRet(1)
				.txtToBpCd.focus
			Case 3		' ������� 
				.txtPaymType.value = arrRet(0)
				.txtPaymTypeNm.value = arrRet(1)
				.txtPaymType.focus
			Case 4		' �����1
				.txtBizAreaCd1.value = arrRet(0)
				.txtBizAreaNm1.value = arrRet(1)
				.txtBizAreaCD1.focus
		End Select
	End With
End Function

Function SetPrintCond(strEbrFile, strCond)

	Dim strFromBizAreaCd, strToBizAreaCd, strFromDt, strToDt, strFromBpCd, strToBpCd, strPaymType, strOrgChangeId

	Dim	strAuthCond

	SetPrintCond = False

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	If CompareDateByFormat(frm1.fpDateTime1.text,frm1.fpDateTime2.text,frm1.fpDateTime1.Alt,frm1.fpDateTime2.Alt, _
        	               "970025",frm1.fpDateTime1.UserDefinedFormat,Parent.gComDateType, true) = False Then
	   frm1.fpDateTime1.focus
	   Exit Function
	End If

	
	strFromBizAreaCd	= FilterVar(frm1.txtBizAreaCD.value,"","SNM")
	strToBizAreaCd		= FilterVar(frm1.txtBizAreaCD1.value,"","SNM")

	strFromDt	= ggoOper.RetFormat(frm1.fpDateTime1.text, "yyyyMMDD")
	strToDt		= ggoOper.RetFormat(frm1.fpDateTime2.text, "yyyyMMDD")

	strFromBpCd		= FilterVar(frm1.txtFromBpCD.value,"","SNM")
	strToBpCd		= FilterVar(frm1.txtToBpCd.value,"","SNM")

	strPaymType		= FilterVar(frm1.txtPaymType.value ,"","SNM")
	strOrgChangeId	= FilterVar(Parent.gChangeOrgId,"","SNM")
	
	If strFromBizAreaCd = "" Then
		strFromBizAreaCd = ""
		frm1.txtBizAreaNM.value = ""
	End If

	If strToBizAreaCd = "" Then
		strToBizAreaCd = "ZZZZZZZZZZZ"
		frm1.txtBizAreaNM1.value = ""
	End If

	If strFromBpCd = "" Then
		strFromBpCd = ""
		frm1.txtFromBpNm.value = ""
	End If	

	If strToBpCd = "" Then
		strToBpCd = "ZZZZZZZZZZ"
		frm1.txtToBpNm.value = ""
	End If
	
	If strPaymType = "" Then
		strPaymType = "%"
	End If

	StrEbrFile	= "f6104ma1.ebr"

	strCond = strCond & "FromBizAreaCd|"		& strFromBizAreaCd
	strCond = strCond & "|ToBizAreaCd|"			& strToBizAreaCd
	strCond = strCond & "|FromPrPaymDt|"		& strFromDt
	strCond = strCond & "|ToPrPaymDt|"			& strToDt
	strCond = strCond & "|FromBpCd|"			& strFromBpCd
	strCond = strCond & "|ToBpCd|"				& strToBpCd
	strCond = strCond & "|PaymType|"			& strPaymType
	strCond = strCond & "|ChangeOrgID|"			& strOrgChangeId


	' ���Ѱ��� �߰� 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_PRPAYM.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_PRPAYM.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_PRPAYM.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_PRPAYM.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	strCond = strCond & "|strAuthCond|"	& strAuthCond

	SetPrintCond = True

End Function

Function FncBtnPrint() 
	On Error Resume Next

	Dim strEbrFile, strCond, objName

	If SetPrintCond(strEbrFile, strCond) = False Then 
		Exit Function
	End If

	objName = AskEBDocumentName(strEbrFile,"ebr")	
	
	Call FncEBRPrint(EBAction,StrEbrFile,strCond)	
	
End Function

Function FncBtnPreview()
	On Error Resume Next
	
	Dim strEbrFile, strCond, objName

	If SetPrintCond(strEbrFile, strCond) = False Then 
		Exit Function
	End If

	objName = AskEBDocumentName(strEbrFile,"ebr")	

	Call FncEBRPreview(StrEbrFile,strCond)	
	
End Function


'###########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################

'*****************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'===========================================  3.1.1 Form_Load()  =========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029																'��: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
    Call SetDefaultVal

	frm1.txtBizAreaCD.focus

	' ���Ѱ��� �߰� 
	Dim xmlDoc

	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc)

	' �����		
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ�		
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text

	' ���κμ�(��������)		
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text

	' ����						
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text

	Set xmlDoc = Nothing
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssueDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtFromIssueDt.Focus        
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssueDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToIssueDt.Focus
    End If
End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

'Function FncQuery() 
'End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call Parent.fncPrint()    
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>

<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100% HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
							 	<TD CLASS="TD5" NOWRAP>�������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtFromIssueDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�����������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
											           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtToIssueDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�����������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=12 MAXLENGTH=10 ALT="�����" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">
											           <INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=25 MAXLENGTH=50 STYLE="TEXT-ALIGN: Left" ALT="������" tag="14X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtBizAreaCD1" NAME="txtBizAreaCD1" SIZE=12 MAXLENGTH=10 ALT="�����" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD1.Value, 4)">
											           <INPUT TYPE=TEXT ID="txtBizAreaNM1" NAME="txtBizAreaNM1" SIZE=25 MAXLENGTH=50 STYLE="TEXT-ALIGN: Left" ALT="������" tag="14X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtFromBPCd" NAME="txtFromBPCd" SIZE=12 MAXLENGTH=10  ALT="�ŷ�ó" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBP(frm1.txtFromBpCd.Value, 1)">
											           <INPUT TYPE=TEXT ID="txtFromBPNm" NAME="txtFromBPNm" SIZE=25 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" ALT="�ŷ�ó��" tag="14X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;~&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtToBPCd" NAME="txtToBPCd" SIZE=12 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" ALT="�ŷ�ó" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtToBpCd.Value, 2)">&nbsp;
											           <INPUT TYPE=TEXT ID="txtToBPNm" NAME="txtToBPNm" SIZE=25 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" ALT="�ŷ�ó" tag="14X" ></TD>
							</TR>
							
							<TR>
							<TD CLASS="TD5" NOWRAP>�������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPaymType" SIZE=10 MAXLENGTH=2 tag="11NXXU" ALT="�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPaymType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtPaymType.Value,3)">
													   <INPUT TYPE=TEXT NAME="txtPaymTypeNm" SIZE=25 tag="24" ALT="���������"></TD>	
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" OnClick="VBScript:Call FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" OnClick="VBScript:Call FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TabIndex="-1">	
</FORM>
</BODY>
</HTML>

