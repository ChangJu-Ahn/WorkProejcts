<%@ LANGUAGE="VBSCRIPT" %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : PRERECEIPT
'*  3. Program ID		    : F7104ma1
'*  4. Program Name         : �����ݹ߻�Check LIst
'*  5. Program Desc         : �����ݹ߻�Check LIst
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/12/15
'*  8. Modified date(Last)  : 2003/01/08
'*  9. Modifier (First)     : Hee Jung, Kim
'* 10. Modifier (Last)      : Kim Chang Jin
'* 11. Comment              :
'======================================================================================================= -->
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
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance


'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'��: �����Ͻ� ���� ASP�� 
'Const BIZ_PGM_ID = "a7120mb1.asp"			'��: �����Ͻ� ���� ASP�� 


'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag 
Dim lgIntFlgMode               ' Variable is for Operation Status 


'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

    '---- Coding part--------------------------------------------------------------------    
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 

Sub SetDefaultVal()
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	dtToday = "<%=GetSvrDate%>"
	Call parent.ExtractDateFrom(dtToday, Parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")                                                     


	frm1.fpDateTime1.Text = StartDate
	frm1.fpDateTime2.Text = EndDate
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>
End Sub


'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
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
		Call EscReturnVal(iWhere)	
		Exit Function
	Else		
		Call SetReturnVal(arrRet, iWhere)
	End If	

End Function
'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function OpenPopup(Byval strCode, Byval Cond)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	select case Cond
		case "biz", "biz1"
			arrParam(0) = "������ڵ��˾�"			' �˾� ��Ī 
			arrParam(1) = "b_biz_area"						' TABLE ��Ī 
			arrParam(2) = strCode      						' Code Condition
			arrParam(3) = ""							' Name Cindition

			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "������ڵ�"				' �����ʵ��� �� ��Ī 
	
			arrField(0) = "BIZ_AREA_CD"						' Field��(0)
			arrField(1) = "BIZ_AREA_NM"						' Field��(1)
    
			arrHeader(0) = "������ڵ�"				' Header��(0)
			arrHeader(1) = "������"					' Header��(1)

	end select    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscReturnVal(Cond)	
		Exit Function
	Else		
		Call SetReturnVal(arrRet, Cond)
	End If	
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 


'------------------------------------------  SetReturnVal()  ---------------------------------------------
'	Name : SetReturnVal()
'	Description : Account Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetReturnVal(ByVal arrRet, ByVal field_fg)	
	Select case field_fg
		case "biz"	'biz area Popup
			frm1.txtBizAreaCd.Value		= arrRet(0)
			frm1.txtBizAreaNm.Value		= arrRet(1)
			frm1.txtBizAreaCD.focus
		case "FrBp"	'biz partner Popup
			frm1.txtFrbpcd.value 		= arrRet(0)
			frm1.txtFrbpnm.value 		= arrRet(1)
			frm1.txtFrbpcd.focus
		case "ToBp"	'biz partner Popup
			frm1.txtTobpcd.value		= arrRet(0)
			frm1.txtTobpnm.value		= arrRet(1)			
			frm1.txtTobpcd.focus
		case "biz1"	'biz area Popup
			frm1.txtBizAreaCd1.Value		= arrRet(0)
			frm1.txtBizAreaNm1.Value		= arrRet(1)
			frm1.txtBizAreaCD1.focus
	End select	
End Function

'------------------------------------------  SetReturnVal()  ---------------------------------------------
'	Name : SetReturnVal()
'	Description : Account Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function EscReturnVal(ByVal field_fg)	
	Select case field_fg
		case "biz"	'biz area Popup
			frm1.txtBizAreaCD.focus
		case "FrBp"	'biz partner Popup
			frm1.txtFrbpcd.focus
		case "ToBp"	'biz partner Popup
			frm1.txtTobpcd.focus
		case "biz1"	'biz area Popup
			frm1.txtBizAreaCD1.focus
	End select	
End Function
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 


'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################

'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

	' GetGlobalVar �Լ��� �ݵ�� ClassLoad �Լ����� ���� ȣ��Ǿ�� �մϴ�.
    Call LoadInfTB19029                           '��: Load table , B_numeric_format
	' ���� Page�� Form Element���� Clear�Ѵ�. 
    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking    
    Call InitVariables                            '��: Initializes local global Variables
    Call SetDefaultVal
    
    Call SetToolbar("10000000000000")				'��: ��ư ���� ���� 
	frm1.txtBizAreacd.focus	

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


'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 


'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Event ó��	
'********************************************************************************************************* 

'======================================================================================================
'   Event Name : txtFrYymm_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtFrYymm_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
        Call SetFocusToDocument("M")
		Frm1.fpDateTime1.Focus
		
    End If
End Sub
Sub txtToYymm_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime2.Action = 7
        Call SetFocusToDocument("M")
		Frm1.fpDateTime2.Focus
		
    End If
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Function SetPrintCond(strEbrFile, StrUrl)
	Dim varBizArea, varBizArea1, varFrBp, varToBp, varFromdt, varTodt
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

	StrEbrFile = "f7104ma1"
		
	If Len(frm1.txtBizAreaCd.value) < 1 Then
		varBizArea    = ""
		frm1.txtBizAreaNM.value = ""
	Else		
		varBizArea = FilterVar(Trim(frm1.txtBizAreaCd.value),"","SNM")
	End If
		
	If Len(frm1.txtBizAreaCd1.value) < 1 Then
		varBizArea1    = "ZZZZZZZZZ"
		frm1.txtBizAreaNM1.value = ""
	Else		
		varBizArea1 = FilterVar(Trim(frm1.txtBizAreaCd1.value),"","SNM")
	End If
		
	If Len(frm1.txtFrBpCd.value) < 1 Then
		varFrBp = ""
		frm1.txtFrBpNm.value = ""
	Else
		varFrBp = FilterVar(Trim(frm1.txtFrBpCd.value),"","SNM")
	End If
		
	If Len(frm1.txtToBpCd.value) < 1 Then
		varToBp = "ZZZZZZZZZZ"
		frm1.txtToBpNm.value = ""
	Else
		varToBp = FilterVar(Trim(frm1.txtToBpCd.value),"","SNM")
	End If	
		
	varFromDt = UNIConvDateToYYYYMMDD(frm1.fpDateTime1.Text, gDateFormat, "")
	varToDt   = UNIConvDateToYYYYMMDD(frm1.fpDateTime2.Text, gDateFormat, "")
		
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	StrUrl = StrUrl & "varFrBP|"		& varFrBp
	StrUrl = StrUrl & "|varToBP|"		& varToBp
	StrUrl = StrUrl & "|varFromDt|"		& varFromDt
	StrUrl = StrUrl & "|varToDt|"		& varToDt
	StrUrl = StrUrl & "|varBizArea|"	& varBizArea
	StrUrl = StrUrl & "|varBizArea1|"	& varBizArea1

	' ���Ѱ��� �߰� 
	strAuthCond		= "	"
		
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_PRRCPT.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_PRRCPT.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_PRRCPT.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_PRRCPT.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond

	SetPrintCond = True

	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Function


'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrEbrFile, StrUrl, ObjName

	On Error Resume Next                                                    '��: Protect system from crashing
	
	If SetPrintCond(strEbrFile, StrUrl) = False Then 
		Exit Function
	End If
    
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")

	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
			
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

function FncBtnPreview()
	Dim StrEbrFile, StrUrl, ObjName

	On Error Resume Next                                                    '��: Protect system from crashing

	If SetPrintCond(strEbrFile, StrUrl) = False Then 
		Exit Function
	End If
    
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")

	Call FncEBRPreview(ObjName,StrUrl)	
	
End Function



'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 


'********************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

'Function FncQuery() 
'    FncQuery = True
'End Function



'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function


'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************** 

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	

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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=12 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" ALT="�Ű�����" tag="1XN" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 'biz')">
											    <INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=25 MAXLENGTH=50 STYLE="TEXT-ALIGN: Left" ALT="�Ű�����" tag="14X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtBizAreaCD1" NAME="txtBizAreaCD1" SIZE=12 MAXLENGTH=10 ALT="�����" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD1.Value, 'biz1')">
											           <INPUT TYPE=TEXT ID="txtBizAreaNM1" NAME="txtBizAreaNM1" SIZE=25 MAXLENGTH=50 STYLE="TEXT-ALIGN: Left" ALT="������" tag="14X" ></TD>
							</TR>						
							<TR>
								<TD CLASS="TD5" NOWRAP>�ŷ�ó�ڵ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFrBpCd" NAME="txtFrBpCd" SIZE=12 MAXLENGTH=10 tag="11X" ALT="�ŷ�ó�ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtFrBpCd.value, 'FrBp')"> <INPUT TYPE="Text" NAME="txtFrBpNm" SIZE=25 MAXLENGTH=30 tag="14" ALT="�ŷ�ó��">&nbsp;~&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txttoBpCd" NAME="txtToBpCd" SIZE=12 MAXLENGTH=10 tag="11X" ALT="�ŷ�ó�ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtToBpCd.value, 'ToBp')"> <INPUT TYPE="Text" NAME="txtToBpNm" SIZE=25 MAXLENGTH=30 tag="14" ALT="�ŷ�ó��"></TD>
							</TR>																			
							<TR>
								<TD CLASS="TD5" NOWRAP>�߻�����</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrYymm" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=���۹߻����� id=fpDateTime1> </OBJECT>');</SCRIPT> ~
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToYymm" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=����߻����� id=fpDateTime2> </OBJECT>');</SCRIPT>
								</TD>								
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" OnClick="VBScript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname"    TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname"   TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar"  TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date"     TABINDEX="-1"> 	
</FORM>
</BODY>
</HTML>
