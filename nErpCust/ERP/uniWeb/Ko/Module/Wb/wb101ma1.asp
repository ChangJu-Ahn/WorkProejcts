<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : VB101MA1
'*  4. Program Name         : Company Register(�����������)
'*  5. Program Desc         : ����������� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/12/27
'*  8. Modified date(Last)  : 2004/12/27
'*  9. Modifier (First)     : LSHSAT
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'***********************************************************************k*********************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '��: indicates that All variables must be declared in advance 


'********************************************  1.2 Global ����/��� ����  *********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global ��� ����  ====================================
'==========================================================================================================

Const BIZ_MNU_ID = "WB101MA1"
Const BIZ_PGM_ID = "wb101mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID = "Wb101mb2.asp"
Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt											 '��: �����Ͻ� ���� ASP���� �����ϹǷ� Dim 

Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        
Dim lgLoadOk, gSelframeFlg

Dim C_SEQ_NO
Dim C_W_TYPE
Dim C_W_NAME
Dim C_W_RGST_NO1
Dim C_W_MGT_NO
Dim C_W_RGST_NO
Dim C_W_RGST_NO2
Dim C_W_CO_ADDR
Dim C_W_HOME_ADDR

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
    
'	frm1.txtCO_CD.value = parent.wgCO_CD
'	frm1.txtco_cd.focus  
End Sub

Sub InitSpreadPosVariables()
    C_SEQ_NO			= 1
    C_W_TYPE			= 2
    C_W_NAME			= 3
    C_W_RGST_NO1		= 4
    C_W_MGT_NO			= 5
    C_W_RGST_NO			= 6
    C_W_RGST_NO2		= 7
    C_W_CO_ADDR			= 8
    C_W_HOME_ADDR		= 9
End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub



'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name :InitComboBox_Five()
'	Description : 
'------------------------------------------------------------------------------------------------------------
Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1018', '" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox_One()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 

Sub InitComboBox_One()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," dbo.ufn_TB_MINOR('W1009', '" & C_REVISION_YM & "')  "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCOMP_TYPE1 ,lgF0  ,lgF1  ,Chr(11))
End Sub

'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox_Two()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox_Two()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1010', '" & C_REVISION_YM & "')  "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDEBT_MULTIPLE ,lgF0  ,lgF1  ,Chr(11))
End Sub

 
'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name :InitComboBox_Three()
'	Description : 
'------------------------------------------------------------------------------------------------------------
Sub InitComboBox_Three()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1013', '" & C_REVISION_YM & "')  "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCOMP_TYPE2 ,lgF0  ,lgF1  ,Chr(11))
End Sub

'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name :InitComboBox_Four()
'	Description : 
'------------------------------------------------------------------------------------------------------------
Sub InitComboBox_Four()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1003', '" & C_REVISION_YM & "')  "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboHOLDING_COMP_FLG ,lgF0  ,lgF1  ,Chr(11))
End Sub


'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name :InitComboBox_Five()
'	Description : 
'------------------------------------------------------------------------------------------------------------
Sub InitComboBox_Five()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1018', '" & C_REVISION_YM & "')  "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE_Body ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20050701",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_W_HOME_ADDR + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    Call AppendNumberPlace("6","3","1")

    ggoSpread.SSSetEdit		C_SEQ_NO,		"����", 10,,,100,1
	ggoSpread.SSSetEdit		C_W_TYPE,		"����", 10,,,10,1
	ggoSpread.SSSetEdit		C_W_NAME,		"����", 8,,,30,1
    ggoSpread.SSSetEdit		C_W_RGST_NO1,	"��Ϲ�ȣ", 15,,,13
    ggoSpread.SSSetMask		C_W_MGT_NO,		"������ȣ", 8, 2,"U-9999-9"
    ggoSpread.SSSetMask		C_W_RGST_NO,	"����ڹ�ȣ", 10, 2,"999-99-99999"
    ggoSpread.SSSetMask		C_W_RGST_NO2,	"�ֹε�Ϲ�ȣ", 14, 2,"999999-9999999"
    ggoSpread.SSSetEdit		C_W_CO_ADDR,	"����������", 30,,,140
    ggoSpread.SSSetEdit		C_W_HOME_ADDR,	"�ּ�", 30,,,140
    
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
	'Call InitSpreadComboBox()
	
	.ReDraw = true
	
    'Call SetSpreadLock 
    
    End With
End Sub


'============================================  �׸��� �Լ�  ====================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
	ggoSpread.SpreadLock C_SEQ_NO, -1, C_W_TYPE

	If frm1.cboEX_RECON_FLG.value = "Y" Then
 		ggoSpread.SSSetRequired C_W_NAME, -1, -1
 		ggoSpread.SSSetRequired C_W_RGST_NO1, -1, -1
		ggoSpread.SSSetRequired C_W_MGT_NO, -1, -1
		ggoSpread.SSSetRequired C_W_RGST_NO, -1, -1
		ggoSpread.SSSetRequired C_W_RGST_NO2, -1, -1
		ggoSpread.SSSetRequired C_W_CO_ADDR, -1, -1
		ggoSpread.SSSetRequired C_W_HOME_ADDR, -1, -1
	Else
		ggoSpread.SSSetUndoColor C_W_NAME,-1,C_W_HOME_ADDR,-1
		ggoSpread.SpreadUnLock C_W_NAME,-1,C_W_HOME_ADDR,-1
	End If
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData

		.ReDraw = False
 
		ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_W_TYPE, pvStartRow, pvEndRow
 			
		If frm1.cboEX_RECON_FLG.value = "Y" Then
 			ggoSpread.SSSetRequired C_W_NAME, pvStartRow, pvEndRow
 			ggoSpread.SSSetRequired C_W_RGST_NO1, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W_MGT_NO, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W_RGST_NO, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W_RGST_NO2, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W_CO_ADDR, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W_HOME_ADDR, pvStartRow, pvEndRow
		End If	    
		.ReDraw = True
    
    End With
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO			= iCurColumnPos(1)
			C_W_TYPE			= iCurColumnPos(2)
			C_W_NAME			= iCurColumnPos(3)
			C_W_RGST_NO1		= iCurColumnPos(4)
			C_W_MGT_NO			= iCurColumnPos(5)
			C_W_RGST_NO			= iCurColumnPos(6)
			C_W_RGST_NO2		= iCurColumnPos(7)
			C_W_CO_ADDR			= iCurColumnPos(8)
			C_W_HOME_ADDR		= iCurColumnPos(9)
    End Select    
End Sub 
'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenCompanyInfo()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

Function OpenCompanyInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"						' �˾� ��Ī 
	arrParam(1) = "TB_COMPANY_HISTORY"						' TABLE ��Ī 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"

    arrField(0) = "CO_CD "							' Field��(0)
    arrField(1) = "CO_NM"							' Field��(1)
    arrField(2) = "FISC_YEAR"						' Field��(2)
    arrField(3) = "REP_TYPE"						' Field��(3)

    arrHeader(0) = "�����ڵ�"						' Header��(0)
    arrHeader(1) = "���θ�"							' Header��(1)
    arrHeader(2) = "�������"						' Header��(2)
    arrHeader(3) = "�Ű���"						' Header��(3)

	arrRet = window.showModalDialog("wb101ra1.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCO_CD.focus
	    Exit Function
	Else
		Call SetCompanyInfo(arrRet,iWhere)
	End If	

End Function



'------------------------------------------  SetItemInfo()  -------------------------------------------------
'	Name : SetCostInfo()
'	Description : Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------------
Function SetCompanyInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtCO_CD.focus
			.txtCO_CD.value     = arrRet(0)
			.txtCO_FULLNM.value = arrRet(1)
			.txtFISC_YEAR.text = arrRet(2)
			.cboREP_TYPE.value = arrRet(3)
		End If
'		lgBlnFlgChgValue = False
	End With

End Function


'========================================================================================================= 
Function OpenTaxOfficeInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���Ҽ����� �˾�"							' �˾� ��Ī 
	arrParam(1) = "dbo.ufn_TB_MINOR('W1079', '" & C_REVISION_YM & "') "					' TABLE ��Ī 
	arrParam(2) =  Trim(strCode)						' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = " "									' Where Condition
	arrParam(5) = "���Ҽ�����"

    arrField(0) = "MINOR_CD"							' Field��(0)
    arrField(1) = "MINOR_NM"							' Field��(1)

    arrHeader(0) = "�������ڵ�"							' Header��(0)
    arrHeader(1) = "��������"								' Header��(1)

	arrRet = window.showModalDialog("../../comasp/adoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		frm1.txtTAX_OFFICE.focus
	    Exit Function
	Else
		Call SetTaxOfficeInfo(arrRet,iWhere)
	End If
End Function

'========================================================================================================= 
Function SetTaxOfficeInfo(Byval arrRet,byval iWhere)'
	With frm1
		If iWhere = 0 Then

			.txtTAX_OFFICE.focus
			.txtTAX_OFFICE.value = arrRet(0)
			.txtTAX_OFFICE_Nm.value = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================= 
Function OpenIndclassInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"							' �˾� ��Ī 
	arrParam(1) = "tb_std_income_rate"					' TABLE ��Ī 
	arrParam(2) =  Trim(strCode)						' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = " "									' Where Condition
	arrParam(5) = "����"

    arrField(0) = "left(STD_INCM_RT_CD, 2)"				' Field��(0)
    arrField(1) = "BUSNSECT_NM"							' Field��(1)

    arrHeader(0) = "�����ڵ�"							' Header��(0)
    arrHeader(1) = "���¸�"								' Header��(1)

	arrRet = window.showModalDialog("../../comasp/adoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		frm1.txtInd_class.focus
	    Exit Function
	Else
		Call SetOpenIndclassInfo(arrRet,iWhere)
	End If
End Function

'========================================================================================================= 
Function SetOpenIndclassInfo(Byval arrRet,byval iWhere)'
	With frm1
		If iWhere = 0 Then

			.txtInd_class.focus
			.txtInd_class.value = arrRet(0)
			.txtInd_class_Nm.value = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================= 
Function OpenIndTypeInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"							' �˾� ��Ī 
	arrParam(1) = "tb_std_income_rate"					' TABLE ��Ī 
	arrParam(2) =  Trim(strCode)					 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "left(STD_INCM_RT_CD, 2) = '" & Frm1.txtInd_class.value & "' "					' Where Condition
	arrParam(5) = "����"

	arrField(0) = "STD_INCM_RT_CD"									' Field��(0)
	arrField(1) = "BUSNSECT_NM"									' Field��(1)
	arrField(2) = "DETAIL_NM"									' Field��(2)
	arrField(3) = "FULL_DETAIL_NM"									' Field��(3)

    arrHeader(0) = "�����ڵ�"							' Header��(0)
    arrHeader(1) = "����"							' Header��(1)
    arrHeader(2) = "������"							' Header��(2)
    arrHeader(3) = "������"							' Header��(3)

	arrRet = window.showModalDialog("../../comasp/adoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=520px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtInd_Type.focus
	    Exit Function
	Else
		Call SetOpenIndTypeInfo(arrRet,iWhere)
	End If
End Function

'========================================================================================================= 
Function SetOpenIndTypeInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtInd_Type.focus
			.txtInd_Type.value = arrRet(0)
			.txtInd_Type_Nm.value = arrRet(2)
			If .txtHOME_TAX_MAIN_IND.Value <> "" And .txtHOME_TAX_MAIN_IND.Value <> arrRet(0) Then
				.txtHOME_TAX_MAIN_IND.Value = arrRet(0)
				.txtHOME_TAX_MAIN_IND_NM.Value = arrRet(2)
			End If
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================= 
Function OpenTaxMainInd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�־��� �˾�"							' �˾� ��Ī 
	arrParam(1) = "tb_std_income_rate"								' TABLE ��Ī 
	arrParam(2) =  strCode								' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "�־���"
	
    arrField(0) = "ED09" & Parent.gColSep & "STD_INCM_RT_CD"							' Field��(0)
	arrField(1) = "ED22" & Parent.gColSep & "DETAIL_NM"									' Field��(2)
	arrField(2) = "ED08" & Parent.gColSep & "BUSNSECT_NM"								' Field��(1)
	arrField(3) = "ED15" & Parent.gColSep & "FULL_DETAIL_NM"							' Field��(3)

    arrHeader(0) = "�־����ڵ�"							' Header��(0)
    arrHeader(1) = "�־�����"							' Header��(2)
    arrHeader(2) = "����"								' Header��(1)
    arrHeader(3) = "������"							' Header��(3)

	arrRet = window.showModalDialog("../../comasp/adoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=520px; dialogHeight=550px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		frm1.txtHOME_TAX_MAIN_IND.focus
	    Exit Function
	Else
		Call SetOpenTaxMainInd(arrRet,iWhere)
	End If
End Function

'========================================================================================================= 
Function SetOpenTaxMainInd(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtHOME_TAX_MAIN_IND.focus
			.txtHOME_TAX_MAIN_IND.value = trim(arrRet(0))
			.txtHOME_TAX_MAIN_IND_NM.value = trim(arrRet(1))

		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================= 
Function OpenBankCD(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"							' �˾� ��Ī 
	arrParam(1) = "dbo.ufn_TB_MINOR('W1020', '" & C_REVISION_YM & "') "								' TABLE ��Ī 
	arrParam(2) =  strCode								' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""					' Where Condition
	arrParam(5) = "����"

    arrField(0) = "MINOR_CD"							' Field��(0)
    arrField(1) = "MINOR_NM"							' Field��(1)

    arrHeader(0) = "�����ڵ�"							' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		frm1.txtBANK_CD.focus
	    Exit Function
	Else
		Call SetOpenBankCD(arrRet,iWhere)
	End If
End Function


'========================================================================================================= 
Function SetOpenBankCD(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtBANK_CD.focus
			.txtBANK_CD.value = arrRet(0)
			.txtBANK_NM.value = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoNM, IntRetCD
	sCoNM		= frm1.txtCO_NM.value
	
    IntRetCD = DisplayMsgBox("WB0003", parent.VB_YES_NO, sCoNM, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If

	Call ggoOper.LockField(Document, "N")
    Call ggoOper.ClearField(Document, "2")	
	Call InitVariables			
	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtCO_CD="			 & Frm1.txtCO_CD.Value      '��: Query Key        
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
	' 2. ��������ǥ�� �ڻ��Ѱ�, ��ä�Ѱ�-�����޹��μ�, �ں���+�����޹��μ�+�ֽĹ����ʰ���+��������-�ֽĹ�����������-�������� �������� 
	lgBlnFlgChgValue = True
End Function

Sub document_onkeydown()
	Dim pObj
	Set pObj = window.event.srcElement 
	If pObj.TagName = "INPUT" And Left(pObj.GetAttribute("Tag"), 1) = "2" Then lgBlnFlgChgValue = True
End Sub

Sub ChangeEvents()
	lgBlnFlgChgValue = True
End Sub

'====================================== �� �Լ� =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetToolBar("1110100000000111")
	Else
		Call SetToolBar("1110100000000111")
	End If
	
	Call ElementVisible(frm1.bttnPreview,0) 
	Call ElementVisible(frm1.bttnPrint,0) 
	
	If frm1.txtCO_CD_Body.readOnly = False Then
		window.setTimeout "javascript:FocusObj('txtCO_CD_Body')", 100
	Else
		window.setTimeout "javascript:FocusObj('txtCO_NM')", 100
	End If

End Function

Function ClickTab2()	
	Dim i, blnChange

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetToolBar("1110110100000111")
	Else
		If frm1.vspdData.MaxRows > 0 Then
			Call SetToolBar("1110111100000111")
		Else
			Call SetToolBar("1110110100000111")
		End if		
	End If
	
	Call ElementVisible(frm1.bttnPreview,1) 
	Call ElementVisible(frm1.bttnPrint,1) 
	
	frm1.txtAGENT_NM.focus
End Function

'========================================================================================================= 
Sub Form_Load()
	
    lgLoadOk = False

    Call InitVariables																'��: Initializes local global variables
    Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6","4","0")
    'Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("1110100000000111")
    Call InitComboBox
    Call InitComboBox_One
    Call InitComboBox_Two
    Call InitComboBox_Three
	Call InitComboBox_Four
	Call InitComboBox_Five


	'Call ggoOper.FormatDate(frm1.txtFirstDeprYyyymm, parent.gDateFormat, 2)
    'Call ggoOper.FormatDate(frm1.txtLastDeprYyyymm, parent.gDateFormat, 2)
    'Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR_Body, parent.gDateFormat,3)
    Call ggoOper.FormatDate(frm1.txtFOUNDATION_DT, parent.gDateFormat, 1)
    Call ggoOper.FormatDate(frm1.txtFISC_START_DT, parent.gDateFormat, 1)
    Call ggoOper.FormatDate(frm1.txtFISC_END_DT, parent.gDateFormat, 1)
    Call ggoOper.FormatDate(frm1.txtHOME_ANY_START_DT, parent.gDateFormat, 1)
    Call ggoOper.FormatDate(frm1.txtHOME_ANY_END_DT, parent.gDateFormat, 1)
    
	Call InitSpreadSheet
	With frm1
		.txtOWN_RGST_NO.Mask = "999-99-99999"
		.txtOWN_RGST_NO.AlignTextH = 1
		.txtLAW_RGST_NO.Mask = "999999-9999999"
		.txtLAW_RGST_NO.AlignTextH = 1	
		.txtREPRE_RGST_NO.Mask = "999999-9999999"
		.txtREPRE_RGST_NO.AlignTextH = 1
		.txtAGENT_RGST_NO.Mask = "999-99-99999"
		.txtAGENT_RGST_NO.AlignTextH = 1
		.txtAPPO_NO.Mask = "9-9999"
		.txtAPPO_NO.AlignTextH = 1
		.txtRECON_BAN_NO.Mask = "9-9999"
		.txtRECON_BAN_NO.AlignTextH = 1
		.txtRECON_MGT_NO.Mask = "9-9999-9"
		.txtRECON_MGT_NO.AlignTextH = 1
			
	End With	
	Call InitData
	
	Call ElementVisible(frm1.bttnPreview,0) 
	Call ElementVisible(frm1.bttnPrint,0) 
	
	frm1.txtco_cd.focus 

    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgLoadOk = True

	Dim sCookey
	Dim sCoCd, sFiscYear, sRepType
	sCoCd		= ReadCookie("gCoCd")
	sFiscYear	= ReadCookie("gFiscYear")
	sRepType	= ReadCookie("gRepType")

	If sCoCd <> "" Then
		With frm1
			.txtCO_CD.value = sCoCd
			.txtFISC_YEAR.text = sFiscYear
			.cboREP_TYPE.value = sRepType
		End With
			
		Call FncQuery
	Else

		frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
		frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
		frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
		If "<%=wgREP_TYPE%>" <> "" Then
			frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
		End If
	
		If frm1.txtCO_CD.value <> "" And frm1.txtFISC_YEAR.text <> "" Then
			Call FncQuery
		End If

	End If
'	FncQuery

End Sub

'============================================  ����� �Լ�  ====================================
Sub InitData()
	frm1.cboREP_TYPE.value = "1"
	frm1.txtREVISION_YM.value = C_REVISION_YM
	frm1.cboCOMP_TYPE1.value = "1"
	frm1.cboDEBT_MULTIPLE.value = "01"
	frm1.cboCOMP_TYPE2.value = "1"
	frm1.cboHOLDING_COMP_FLG.value = "1"
	frm1.cboREP_TYPE_Body.value = "1"
	frm1.cboEX_RECON_FLG.value = "N"
	frm1.cboEX_54_FLG.value = "N"
	frm1.cboSUBMIT_FLG.value = "2"
	frm1.cboUSE_FLG.value = "Y"	
End Sub

Sub SetFieldAtt()

	If Frm1.cboEX_RECON_FLG.value = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtAGENT_NM, "N")
		Call ggoOper.SetReqAttr(frm1.txtRECON_BAN_NO, "N")
		Call ggoOper.SetReqAttr(frm1.txtRECON_MGT_NO, "N")
		Call ggoOper.SetReqAttr(frm1.txtAGENT_RGST_NO, "N")
		Call ggoOper.SetReqAttr(frm1.txtREQUEST_DT, "N")
		Call ggoOper.SetReqAttr(frm1.txtAPPO_NO, "N")
		Call ggoOper.SetReqAttr(frm1.txtAPPO_DT, "N")
		Call ggoOper.SetReqAttr(frm1.txtAPPO_DESC, "N")
		Call SetSpreadLock
	Else
		Call ggoOper.SetReqAttr(frm1.txtAGENT_NM, "D")
		Call ggoOper.SetReqAttr(frm1.txtRECON_BAN_NO, "D")
		Call ggoOper.SetReqAttr(frm1.txtRECON_MGT_NO, "D")
		Call ggoOper.SetReqAttr(frm1.txtAGENT_RGST_NO, "D")
		Call ggoOper.SetReqAttr(frm1.txtREQUEST_DT, "D")
		Call ggoOper.SetReqAttr(frm1.txtAPPO_NO, "D")
		Call ggoOper.SetReqAttr(frm1.txtAPPO_DT, "D")
		Call ggoOper.SetReqAttr(frm1.txtAPPO_DESC, "D")
		Call SetSpreadLock
	End If
End Sub

Function ChkFiscDate()
	Dim i, iDGap, iMGap
	Dim dFisc_Start_Dt, dFisc_End_Dt
	
	ChkFiscDate = True
	
	If frm1.txtFISC_START_DT.Text = "" Or frm1.txtFISC_END_DT.Text = "" Then Exit Function
	dFisc_Start_Dt = CDate(frm1.txtFISC_START_DT.Text)
	dFisc_End_Dt = CDate(frm1.txtFISC_END_DT.Text)
	
	iDGap = DateDiff("d", dFisc_Start_Dt, dFisc_End_Dt)
	iMGap = DateDiff("m", dFisc_Start_Dt, dFisc_End_Dt)
	
	If iDGap > 365 Then
		MsgBox "��� ����, ������ �Ⱓ�� Ȯ���Ͻʽÿ�", vbInformation
		ChkFiscDate = False
		Exit Function
	'ElseIf iMGap > 6 And Frm1.cboREP_TYPE_Body.Value = "2" Then
	'	msgbox "��� ����, ������ �Ⱓ�� Ȯ���Ͻʽÿ�"
	'	ChkFiscDate = False
	'	Exit Function
	End If
End Function

'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	'Call GetFISC_DATE
End Sub

Sub txtCO_CD_onChange()	' �����ڵ� ����� 
	Dim arrVal
	
	If Len(frm1.txtCO_CD.Value) > 0 Then
		If CommonQueryRs("CO_NM", " TB_COMPANY_HISTORY " , " CO_CD = '" & frm1.txtCO_CD.Value &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    	arrVal				= Split(lgF0, Chr(11))
			frm1.txtCO_FULLNM.Value	= arrVal(0)
		Else
			Call DisplayMsgBox("970000", "x",frm1.txtCO_CD.alt & " '" & UCase(Me.Value) & "' " ,"x")
			frm1.txtCO_CD.Value	= ""
			frm1.txtCO_FULLNM.Value	= ""
		End If
	Else
		frm1.txtCO_FULLNM.Value = ""
	End If

End Sub

Sub txtTAX_OFFICE_onChange()	' ���Ҽ������ڵ� ����� 
	Dim arrVal
	
	If Len(frm1.txtTAX_OFFICE.Value) > 0 Then
		If CommonQueryRs("MINOR_NM", " dbo.ufn_TB_MINOR('W1079', '" & C_REVISION_YM & "') " , "MINOR_CD = '" & frm1.txtTAX_OFFICE.Value &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    	arrVal				= Split(lgF0, Chr(11))
			frm1.txtTAX_OFFICE_NM.Value	= arrVal(0)
		Else
			Call DisplayMsgBox("970000", "x",frm1.txtTAX_OFFICE.alt & " '" & UCase(Me.Value) & "' ","x")
			frm1.txtTAX_OFFICE.Value	= ""
			frm1.txtTAX_OFFICE_NM.Value	= ""
		End If
	Else
		frm1.txtTAX_OFFICE_NM.Value = ""
	End If

End Sub

Sub txtBANK_CD_onChange()	' �����ڵ� ����� 
	Dim arrVal
	
	If Len(frm1.txtBANK_CD.Value) > 0 Then
		If CommonQueryRs("MINOR_NM", " dbo.ufn_TB_MINOR('W1020', '" & C_REVISION_YM & "')  " , "MINOR_CD = '" & frm1.txtBANK_CD.Value &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    	arrVal				= Split(lgF0, Chr(11))
			frm1.txtBANK_NM.Value	= arrVal(0)
		Else
			Call DisplayMsgBox("970000", "x",frm1.txtBANK_CD.alt & " '" & UCase(Me.Value) & "' ","x")
			frm1.txtBANK_CD.Value	= ""
			frm1.txtBANK_NM.Value	= ""
		End If
	Else
		frm1.txtBANK_NM.Value = ""
	End If

End Sub

Sub txtHOME_TAX_MAIN_IND_onChange()	' �����ڵ� ����� 
	Dim arrVal
	
	If Len(frm1.txtHOME_TAX_MAIN_IND.Value) > 0 Then
		If CommonQueryRs("DETAIL_NM", " tb_std_income_rate " , "STD_INCM_RT_CD = '" & frm1.txtHOME_TAX_MAIN_IND.Value &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    	arrVal				= Split(lgF0, Chr(11))
			frm1.txtHOME_TAX_MAIN_IND_NM.Value	= arrVal(0)
		Else
			Call DisplayMsgBox("970000", "x",frm1.txtHOME_TAX_MAIN_IND.alt & " '" & UCase(Me.Value) & "' ","x")
			frm1.txtHOME_TAX_MAIN_IND.value = ""
			frm1.txtHOME_TAX_MAIN_IND_NM.Value	= ""
		End If
	Else
		frm1.txtHOME_TAX_MAIN_IND_NM.Value = ""
	End If

End Sub

Sub txtFISC_YEAR_Body_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR_Body.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR_Body.Focus
    End If
End Sub

Sub txtFISC_YEAR_Body_Change()
	With frm1

		.txtFISC_START_DT.text = .txtFISC_YEAR_Body.text & "-01-01"
		.txtFISC_END_DT.text = .txtFISC_YEAR_Body.text & "-12-31"
	End With
End Sub
'=======================================================================================================
'   Event Name : txtFOUNDATION_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFOUNDATION_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFOUNDATION_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFOUNDATION_DT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtINCOM_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtINCOM_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtINCOM_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtINCOM_DT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtREQUEST_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtREQUEST_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtREQUEST_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtREQUEST_DT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtINCOM_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtAPPO_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtAPPO_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtAPPO_DT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFOUNDATION_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtHOME_FILE_MAKE_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtHOME_FILE_MAKE_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtHOME_FILE_MAKE_DT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFISC_START_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFISC_START_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_START_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_START_DT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFISC_END_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFISC_END_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_END_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_END_DT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFISC_START_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtHOME_ANY_START_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtHOME_ANY_START_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtHOME_ANY_START_DT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFISC_END_DT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtHOME_ANY_END_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtHOME_ANY_END_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtHOME_ANY_END_DT.Focus
    End If
End Sub

'=======================================================================================================

'=======================================================================================================
Sub txtFOUNDATION_DT_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtFISC_START_DT_Change()
    lgBlnFlgChgValue = True
'    Call ChkFiscDate()
End Sub

'=======================================================================================================
Sub txtFISC_END_DT_Change()
    lgBlnFlgChgValue = True
'    Call ChkFiscDate()
End Sub

'=======================================================================================================
Sub txtHOME_ANY_START_DT_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtHOME_ANY_END_DT_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtINCOM_DT_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtHOME_FILE_MAKE_DT_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtAPPO_DT_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtREQUEST_DT_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtOWN_RGST_NO_change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLAW_RGST_NO_change()
	lgBlnFlgChgValue = True
End Sub

Sub txtREPRE_RGST_NO_change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : txtFISC_YEAR_Body_OnChange()
' Function Desc : 
'========================================================================================

'========================================================================================================= 
Sub cboImdpalignopt_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub cboTaxPolicy_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub cboCurPolicy_OnChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================
Sub cboXCH_RATE_FG_OnChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================
Sub cboOpenAcctFg_OnChange() 
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub cboXchErrorUseFg_OnChange() 
	lgBlnFlgChgValue = True
End Sub

Sub cboEX_RECON_FLG_onChange()
	lgBlnFlgChgValue = True

	If lgLoadOk = True Then
		Call SetFieldAtt()
	End If
End Sub

'Sub cboEX_54_FLG_onChange()
'	lgBlnFlgChgValue = True
'End Sub

Sub Document_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboUSE_FLG_onKeyDown()
	If window.event.keyCode = 9 Then
		If frm1.txtCO_CD_Body.readOnly = False Then
			window.setTimeout "javascript:FocusObj('txtCO_CD_Body')", 100
		Else
			window.setTimeout "javascript:FocusObj('txtCO_NM')", 100
		End If
	End If 
End Sub

Sub txtAPPO_DESC_onKeyDown()
	If window.event.keyCode = 9 Then
		window.setTimeout "javascript:FocusObj('txtAGENT_NM')", 100
	End If 
End Sub

Sub txtAGENT_RGST_NO_onKeyDown()
	If window.event.keyCode = 9 Then
		If frm1.vspdData.MaxRows = 0 Then
			window.setTimeout "javascript:FocusObj('txtREQUEST_DT')", 100
		End If
	End If 
End Sub

Sub FocusObj(Byval pObjNm)
	Dim pObj
	On Error Resume Next	' -- Object ���� .Select�� ������ 
	Set pObj = document.all(pObjNm)
	If Not pObj is Nothing Then
		If Trim(pObj.value) <> "" Then
			pObj.Focus
			pObj.Select
		Else
			pObj.Focus
		End If
	End If
End Sub
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1111111111")    
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
        Exit Sub
    End If


End Sub

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'========================================================================================
Function FncQuery() 
    Dim IntRetCD

    FncQuery = False
    Err.Clear

  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

  '-----------------------
    'Query function call area
    '----------------------- 

    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

'    Call DbQuery
    FncQuery = True
End Function


'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
	Call InitData
    Call SetToolbar("1110100000000111")
    lgIntFlgMode = parent.OPMD_CMODE

	window.setTimeout "javascript:FocusObj('txtCO_CD')", 100	'frm1.txtCO_CD.focus

    FncNew = True

End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim strYear,strMonth,strDay
    Dim strYear1,strMonth1,strDay1

	FncSave = False
	Err.Clear

	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
	    Exit Function
	End If
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                             '��: Check contents area
	   Exit Function
	End If

	If Not isNumeric(Replace(frm1.txtOWN_RGST_NO.text, "-", "")) Then
		Call DisplayMsgBox("WC0036", parent.VB_INFORMATION,"����ڵ�Ϲ�ȣ","����")
		window.setTimeout "javascript:FocusObj('txtOWN_RGST_NO')", 100	'frm1.txtOWN_RGST_NO.focus
		Exit Function
	End If
	
	If Not isNumeric(Replace(frm1.txtLAW_RGST_NO.text, "-", "")) Then
		Call DisplayMsgBox("WC0036", parent.VB_INFORMATION,"���ε�Ϲ�ȣ","����")
		window.setTimeout "javascript:FocusObj('txtLAW_RGST_NO')", 100	'frm1.txtLAW_RGST_NO.focus
		Exit Function
	End If

	If frm1.txtREPRE_RGST_NO.text <> "" Then
		If Not isNumeric(Replace(frm1.txtREPRE_RGST_NO.text, "-", "")) Then
			Call DisplayMsgBox("WC0036", parent.VB_INFORMATION,"��ǥ���ֹι�ȣ","����")
			window.setTimeout "javascript:FocusObj('txtREPRE_RGST_NO')", 100	'frm1.txtLAW_RGST_NO.focus
			Exit Function
		End If
	End If
		
	If CompareDateByFormat(frm1.txtFISC_Start_DT.text,frm1.txtFISC_End_DT.text,frm1.txtFISC_Start_DT.Alt,frm1.txtFISC_End_DT.Alt, _
        	               "970024",frm1.txtFISC_Start_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
	   window.setTimeout "javascript:FocusObj('txtFISC_Start_DT')", 100	'frm1.txtFISC_Start_DT.focus
	   Exit Function
	End If
	

	If CompareDateByFormat(frm1.txtFOUNDATION_DT.text,frm1.txtFISC_START_DT.text,frm1.txtFOUNDATION_DT.Alt,frm1.txtFISC_START_DT.Alt, _
        	               "970025",frm1.txtFOUNDATION_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
	'	frm1.txtFISC_Start_DT.focus

		window.setTimeout "javascript:FocusObj('txtFISC_Start_DT')", 100	
		Exit Function
	End If
	
  
	If CompareDateByFormat(frm1.txtHOME_ANY_START_DT.text,frm1.txtHOME_ANY_END_DT.text,frm1.txtHOME_ANY_START_DT.Alt,frm1.txtHOME_ANY_END_DT.Alt, _
        	               "970024",frm1.txtHOME_ANY_START_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
	   window.setTimeout "javascript:FocusObj('txtHOME_ANY_START_DT')", 100	'frm1.txtHOME_ANY_START_DT.focus
	   Exit Function
	End If

	If CompareDateByFormat(frm1.txtFISC_Start_DT.text,frm1.txtHOME_ANY_START_DT.text,frm1.txtFISC_Start_DT.Alt,frm1.txtHOME_ANY_START_DT.Alt, _
        	               "970024",frm1.txtHOME_ANY_START_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
	   window.setTimeout "javascript:FocusObj('txtHOME_ANY_START_DT')", 100	'frm1.txtHOME_ANY_START_DT.focus
	   Exit Function
	End If
	
	IF  ChkFiscDate	= False then
		Exit Function
	End If
	
	If frm1.cboEX_RECON_FLG.value = "Y" Then
		Dim iRow, iMaxRows, iCol
		With frm1.vspdData
			iMaxRows = .MaxRows
			For iRow = 1 To iMaxRows
				.Row = iRow 
				For iCol = C_W_NAME To C_W_HOME_ADDR
					.Col = iCol
					If Trim(.Value) = "" Then
						.Row = 0
						Call DisplayMsgBox("X", parent.VB_INFORMATION, .Text & "��(��) �Է��Ͻʽÿ�","X")
						.focus
						.Col = iCol : .Row = iRow
						.Action = 0
						Exit Function
					End If
				Next
			Next
		End With
	End If

	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False then
		Exit Function
	End If

	FncSave = True
End Function


'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    lgIntFlgMode = parent.OPMD_CMODE											'Indicates that current mode is Crate mode

     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'This function lock the suitable field
    
	lgBlnFlgChgValue = True

    frm1.txtCO_CD_Body.value = ""

    frm1.txtCO_CD_Body.focus
    
End Function


'========================================================================================
Function FncCancel()
     On Error Resume Next
    
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
    
    lgBlnFlgChgValue = True
End Function


'========================================================================================
Function FncInsertRow(pvRowCnt)

	Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo, i

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
   
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		iRow = .ActiveRow+1

		.ReDraw = False
			
		' SEQ_NO �� �׸��忡 �ִ� ���� 
		iSeqNo = GetMaxSpreadVal(frm1.vspdData , C_SEQ_NO)	' �ִ�SEQ_NO�� ���ؿ´�.
			
		ggoSpread.InsertRow ,imRow	' �׸��� �� �߰�(����� ��� ����)
		SetSpreadColor iRow, iRow + imRow - 1	' �׸��� ���󺯰� 
		
		For i = iRow to iRow + imRow - 1	' �߰��� �׸����� SEQ_NO�� �����Ѵ�.
			.Row = i
			.Col = C_SEQ_NO
			.Text = iSeqNo
			iSeqNo = iSeqNo + 1		' SEQ_NO�� �����Ѵ�.
			If i = 1 Then
				.Col = C_W_TYPE	: .text = "��ǥ�̻�"
			Else
				.Col = C_W_TYPE	: .text = "������"
			End If
		Next				
		.ReDraw = True	

		''SetSpreadColor .vspdData.ActiveRow    
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement  
    
    lgBlnFlgChgValue = True     
End Function


'========================================================================================
Function FncDeleteRow()
     On Error Resume Next
      Dim lDelRows


	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		lDelRows = ggoSpread.DeleteRow
	End With
	 Set gActiveElement = document.ActiveElement  
	lgBlnFlgChgValue = True
End Function


'========================================================================================
Function FncPrint()
     On Error Resume Next
    parent.FncPrint()
End Function


'========================================================================================
Function FncPrev()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    ElseIf lgPrevNo = "" then
		Call DisplayMsgBox("900011", "X", "X", "X")
	End IF

    response.write lgPrevNo

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtco_cd =" & lgPrevNo

	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
Function FncNext()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						  '��: �����Ͻ� ó�� ASP�� ���°� 
    strVal = strVal & "&txtco_cd=" & lgNextNo

	Call RunMyBizASP(MyBizASP, strVal)
End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLE)												'��: ȭ�� ���� 
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtco_cd=" & Trim(frm1.txtco_cd.value)				'��: ���� ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
End Function


'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False
    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
 '   Call LayerShowHide(0)
End Function

'========================================================================================
Function DbQueryOk()

	' �������� ���� : ���ߵǸ� ���ȴ�.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	Call InitVariables
	
	If wgConfirmFlg = "Y" Then

		Call SetToolbar("1110000000000111")	
	Else
		Call SetToolbar("1110100000000111")
		lgBlnFlgChgValue = False
		Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
		Call SetFieldAtt()
		Call ClickTab1
		lgIntFlgMode = parent.OPMD_UMODE
	End If
	
End Function

'========================================================================================
Function DbSave() 
	
    Err.Clear
	DbSave = False

    Dim strVal, lMaxRows, lMaxCols, lRow, strDel, lCol

    Call LayerShowHide(1) 

	With Frm1
	
		With frm1.vspdData
		
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			For lRow = 1 To lMaxRows
    
		    .Row = lRow
		    .Col = 0
		 
 			     Select Case .Text
			         Case  ggoSpread.InsertFlag                                      '��: Insert
			                                            strVal = strVal & "C"  &  Parent.gColSep
			         Case  ggoSpread.UpdateFlag                                      '��: Update
			                                            strVal = strVal & "U"  &  Parent.gColSep
			         Case  ggoSpread.DeleteFlag                                      '��: Delete
			                                            strDel = strDel & "D"  &  Parent.gColSep
			     End Select
			       
				' ��� �׸��� ����Ÿ ����     
				If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Then
				  	For lCol = 1 To lMaxCols
				  		.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				  	Next
				  	strVal = strVal & Trim(.Text) &  Parent.gRowSep
				Elseif .Text = ggoSpread.DeleteFlag then
				    For lCol = 1 To lMaxCols
				  		.Col = lCol : strDel = strDel & Trim(.Text) &  Parent.gColSep
				  	Next
				  	strDel = strDel & Trim(.Text) &  Parent.gRowSep
				End If  
			Next

      
		End With

       .txtSpread.value        =  strDel & strVal

       strDel = ""	: strVal = ""
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()
    frm1.txtCO_CD.value = frm1.txtCO_CD_Body.value
    frm1.txtFISC_YEAR.text = frm1.txtFISC_YEAR_Body.text  
    lgBlnFlgChgValue = False
    Call MainQuery()
End Function

Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE,EBR_RPT_ID,EBR_RPT_ID2
	Dim StrUrl  , i
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim intCnt,IntRetCD

	EBR_RPT_ID	    = "WB101OA1"

    If Not chkField(Document, "1") Then					'��: This function check indispensable field
       Exit Function
    End If
    

    Call SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE)
   
    StrUrl = StrUrl & "varCo_Cd|"			& varCo_Cd
	StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
	StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE

     ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
     if  strPrintType = "VIEW" then
	 Call FncEBRPreview(ObjName, StrUrl)
     else
	 Call FncEBRPrint(EBAction,ObjName,StrUrl)
     end if	
	
	
	call CommonQueryRs("ISNULL(Count(SEQ_NO),0)"," TB_AGENT_INFO "," CO_CD= '" & varCo_Cd & "' AND FISC_YEAR='" & varFISC_YEAR & "' AND REP_TYPE='" & varREP_TYPE & "' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if Trim(replace(lgF0,chr(11),"")) > 5 then
      	 EBR_RPT_ID	    = "WB101OA11"
      	 ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
           if  strPrintType = "VIEW" then
			   Call FncEBRPreview(ObjName, StrUrl)
		   else
			   Call FncEBRPrint(EBAction,ObjName,StrUrl)
		   end if	
 
    end if
	

End Function 

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���� ���� ���� ����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���� �븮�� ����</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:GetRef">�������� �ҷ�����</A>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCO_CD" MAXLENGTH="10" SIZE=10 ALT ="�����ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCompanyInfo(frm1.txtco_cd.value,0)"> <INPUT NAME="txtCO_FULLNM" MAXLENGTH="30" SIZE=30 ALT ="���θ�" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>�������</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/wb101ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�Ű���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="11XXXU"></SELECT></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=100% valign=top>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCO_CD_Body" ALT="�����ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN:Left" tag = "23XXXU"></TD>
											<TD CLASS=TD5 NOWRAP>���θ�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCO_NM" ALT="���θ�" MAXLENGTH="25" STYLE="TEXT-ALIGN:left" tag="22"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���μ�����</TD>
											<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtCO_ADDR" ALT="���μ�����" MAXLENGTH="60" SIZE="103" STYLE="TEXT-ALIGN:left"  tag="22"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>����ڵ�Ϲ�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_I755569916_txtOWN_RGST_NO.js'></script></TD>
											<TD CLASS=TD5 NOWRAP>���ε�Ϲ�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_I391897511_txtLAW_RGST_NO.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��ǥ�ڸ�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtREPRE_NM" ALT="��ǥ�ڸ�" MAXLENGTH="25" STYLE="TEXT-ALIGN:left" tag ="22"></TD>
											<TD CLASS=TD5 NOWRAP>��ǥ���ֹι�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_OBJECT1_txtREPRE_RGST_NO.js'></script>
											<!--<INPUT NAME="txtREPRE_RGST_NO2" ALT="��ǥ���ֹι�ȣ��ȣ" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag ="2">--></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�������ȭ��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTEL_NO" ALT="��ȭ��ȣ" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="2"></TD>
											<TD CLASS=TD5 NOWRAP>���Ҽ�����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTAX_OFFICE" ALT="���Ҽ�����" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenTaxOfficeInfo(frm1.txtTAX_OFFICE.value,0)">
											<INPUT NAME="txtTAX_OFFICE_NM" ALT="���Ҽ�����" SIZE="20" tag = "24" ></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�߼ұ������</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCOMP_TYPE1" ALT="�߼ұ������" STYLE="WIDTH: 220" tag="22" onchange="ChangeEvents()"></SELECT></TD>
											<TD CLASS=TD5 NOWRAP>���Աݹ��</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboDEBT_MULTIPLE" ALT="���Ա� ���" STYLE="WIDTH: 220" tag="22" onchange="ChangeEvents()"></SELECT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���������ش翩��</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCOMP_TYPE2" ALT="���������ش翩��" STYLE="WIDTH: 220" tag="22" onchange="ChangeEvents()"></SELECT></TD>
											<TD CLASS=TD5 NOWRAP>����ȸ���ش翩��</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboHOLDING_COMP_FLG" ALT="����ȸ���ش翩��" STYLE="WIDTH: 220" tag="22" onchange="ChangeEvents()"></SELECT></TD>
										</TR>							
										<TR>
											<TD CLASS=TD5 NOWRAP>����</TD>								
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIND_CLASS" ALT="����" SIZE="20" MAXLENGTH=50 tag = "22"></TD>
											<TD CLASS=TD5 NOWRAP>����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIND_TYPE" ALT="����" SIZE="20" MAXLENGTH=50 tag = "22"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�־����ڵ�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtHOME_TAX_MAIN_IND" ALT="�־����ڵ�" MAXLENGTH="7" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenTaxMainInd(frm1.txtHOME_TAX_MAIN_IND.value,0)">
											<INPUT NAME="txtHOME_TAX_MAIN_IND_NM" ALT="�־����ڵ�" SIZE="40" tag = "24" ></TD>
											<TD CLASS=TD5 NOWRAP>���������</TD>								
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_txtFOUNDATION_DT_txtFOUNDATION_DT.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>E-mail</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtHOME_TAX_EMAIL" ALT="E-mail" MAXLENGTH="30" SIZE="30" STYLE="TEXT-ALIGN:left"  tag="2" ></TD>
											<TD CLASS=TD5 NOWRAP>Ȩ�ؽ� ID</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtHOME_TAX_USR_ID" ALT="HOME TAXID" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag="22"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% valign=top>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>�������</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_txtFISC_YEAR_Body_txtFISC_YEAR_Body.js'></script></TD>
											<TD CLASS=TD5 NOWRAP>�Ű���</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboREP_TYPE_Body" ALT="�Ű���" STYLE="WIDTH: 220" tag="23X" onchange="ChangeEvents()"></SELECT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>����������</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_fpDateTime1_txtFISC_START_DT.js'></script></TD>
											<TD CLASS=TD5 NOWRAP>�����������</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_fpDateTime2_txtFISC_END_DT.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���ýŰ��������</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_fpDateTime3_txtHOME_ANY_START_DT.js'></script></TD>
											<TD CLASS=TD5 NOWRAP>���ýŰ���������</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_fpDateTime4_txtHOME_ANY_END_DT.js'></script></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% valign=top>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>ȯ�ް�������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBANK_CD" ALT="�����ڵ�" MAXLENGTH="2" SIZE="10" STYLE="TEXT-ALIGN:left" tag="2"  ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenBankCD(frm1.txtBANK_CD.value,0)">
											<INPUT NAME="txtBANK_NM" ALT="�־����ڵ�" SIZE="20" tag = "24" ></TD>
											<TD CLASS=TD5 NOWRAP>ȯ�ް��� ������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBANK_BRANCH" ALT="ȯ�ޱݰ��� ������" MAXLENGTH="15" STYLE="TEXT-ALIGN:left" tag="2" > (��)����</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>ȯ�ް��� ��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBANK_DPST" ALT="ȯ�ޱݰ��� ��������" MAXLENGTH="10" STYLE="TEXT-ALIGN:left" tag  ="2" > ����</TD>
											<TD CLASS=TD5 NOWRAP>ȯ�� ���¹�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBANK_ACCT_NO" ALT="ȯ�ޱݰ��¹�ȣ" MAXLENGTH="30" STYLE="TEXT-ALIGN:left" tag="2" ></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% valign=top>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ܺ���������</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboEX_RECON_FLG" ALT="�ܺ���������" STYLE="WIDTH: 220" tag="22" onchange="ChangeEvents()"><OPTION VALUE="N">�ƴϿ�<OPTION VALUE="Y">��</SELECT></TD>
											<TD CLASS=TD5 NOWRAP>�ֽĺ����ڷ� <br>��ü�� ���⿩��</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboEX_54_FLG" ALT="�ֽĺ����ڷ��ü�����⿩��" STYLE="WIDTH: 220" tag="22" onchange="ChangeEvents()"><OPTION VALUE="N">�ƴϿ�<OPTION VALUE="Y">��</SELECT></TD>
										</TR>

										<TR>
											<TD CLASS=TD5 NOWRAP>�Ű���������</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_txtINCOM_DT_txtINCOM_DT.js'></script></TD>
											<TD CLASS=TD5 NOWRAP>�Ű��ۼ�����</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_txtHOME_TAX_MAKE_DT_txtHOME_FILE_MAKE_DT.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����ڱ���</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboSUBMIT_FLG" ALT="�����ڱ���" STYLE="WIDTH: 220" tag="22" onchange="ChangeEvents()"><OPTION VALUE="1">�����븮��<OPTION VALUE="2" SELECTED>������</SELECT></TD>
											<TD CLASS=TD5 NOWRAP>�������</TD>
											<TD CLASS=TD6 NOWRAP><SELECT NAME="cboUSE_FLG" ALT="�������" STYLE="WIDTH: 220" tag="22" onchange="ChangeEvents()"><OPTION VALUE="Y">���<OPTION VALUE="N" SELECTED>�̻��</SELECT></TD>
										</TR>

									</TABLE>
								</TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=100% valign=top HEIGHT=10>1. �����븮�� �⺻���� 
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% valign=top HEIGHT=20%>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����븮�μ���</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAGENT_NM" ALT="�����븮�μ���" MAXLENGTH="30" STYLE="TEXT-ALIGN:left" tag  ="25" ></TD>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������ ��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_OBJECT1_txtRECON_BAN_NO.js'></script></TD>
											<TD CLASS=TD5 NOWRAP>������ ������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_OBJECT1_txtRECON_MGT_NO.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����븮����ȭ��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAGENT_TEL_NO" ALT="�����븮����ȭ��ȣ" MAXLENGTH="14" STYLE="TEXT-ALIGN:left" tag ="25"  ></TD>
											<TD CLASS=TD5 NOWRAP>�����븮�λ���ڹ�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_OBJECT1_txtAGENT_RGST_NO.js'></script></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% valign=top HEIGHT=10>2. �����ݿ� ���� ���� 
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% valign=top HEIGHT=60%>
									<script language =javascript src='./js/wb101ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% valign=top HEIGHT=15%>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>��û����</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_txtREQUEST_DT_txtREQUEST_DT.js'></script></TD>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_txtAPPO_DT_txtAPPO_DT.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/wb101ma1_OBJECT1_txtAPPO_NO.js'></script></TD>
											<TD CLASS=TD5 NOWRAP>�����ǳ��</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtREVISION_YM" ALT="�����ǳ��" MAXLENGTH="10" STYLE="TEXT-ALIGN:left" tag="24" ></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtAPPO_DESC" ALT="��������" MAXLENGTH="50" SIZE="103" STYLE="TEXT-ALIGN:left;" tag ="25"  ></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtFlgMode" tag="24" tabindex="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex="-1"></TEXTAREA>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" tabindex="-1"></iframe>
</DIV>

</BODY>
</HTML>
