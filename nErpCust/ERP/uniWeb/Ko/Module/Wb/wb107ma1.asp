<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���μ� ���� 
'*  3. Program ID           : WB107MA1
'*  4. Program Name         : WB107MA1.asp
'*  5. Program Desc         : ��51ȣ �߼ұ�� ���ذ���ǥ 
'*  6. Modified date(First) : 2005/02/14
'*  7. Modified date(Last)  : 2005/02/14
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  �α����� ������ �����ڵ带 ����ϱ� ����  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<STYLE>
	.RADIO {
		BORDER: 0
	}
</STYLE>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "WB107MA1"
Const BIZ_PGM_ID		= "WB107mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "WB107mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "WB107OA1"

' -- �׸��� �÷� ���� 
Dim C_W01	
Dim C_W04	
Dim C_W07	
Dim C_W02	
Dim C_W05	
Dim C_W08
Dim C_W06	
Dim C_W09
Dim C_W_SUM	
Dim C_W19	
Dim C_W10	
Dim C_W11	
Dim C_W12
Dim C_W13	
Dim C_W14	
Dim C_W15	
Dim C_W16	
Dim C_W17	
Dim C_W20	
Dim C_W21
Dim C_W18
Dim C_W22
Dim C_W23

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2

Dim IsRunEvents	' �Ф� �����̺�Ʈ�ݺ��� ���� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W01		= 0	' HTML���� ���� 
	C_W04		= 1
	C_W07		= 2	
	C_W02		= 3
	C_W05		= 4
	C_W08		= 5
	C_W06		= 6
	C_W09		= 7
	C_W_SUM		= 8
	C_W19		= 9
	C_W10		= 11
	C_W11		= 12
	C_W12		= 13
	C_W13		= 14
	C_W14		= 15
	C_W15		= 16
	C_W16		= 17
	C_W17		= 18
	C_W20		= 19
	C_W21		= 20
	C_W18		= 21 
	C_W22		= 22
	C_W23		= 10
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1

    IsRunEvents = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �Ű��� �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	' ��ȸ����(����)
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

	Call InitSpreadPosVariables

End Sub

Sub InitSpreadSheet()

	Call AppendNumberPlace("6","5","1")
	Call AppendNumberPlace("7","5","0")
	Call AppendNumberPlace("8","4","0")
	
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()
	With frm1
	
	.txtFISC_YEAR.text = "<%=wgFISC_YEAR%>"
    .txtCO_CD.value = "<%=wgCO_CD%>"
    .txtCO_NM.value = "<%=wgCO_NM%>"
    .cboREP_TYPE.value = "<%=wgREP_TYPE%>"
    
    .txtW19(0).checked = true
    .txtW20(0).checked = true
    .txtW21(0).checked = true
    .txtW22(0).checked = true
    .txtW23(0).checked = true
    
    Call GetCompanyInfo
    
    Call InitVariables
    End With
End Sub

Sub InitSpreadComboBox()

End Sub

'============================== ���۷��� �Լ�  ========================================

Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 

	Call window.open("WB107MA2.txt", BIZ_MNU_ID, _
	"Width=600px,Height=450px,center= Yes,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes")

End Function

' ��� ���� 
Sub SetHeadReCalc()	
	Dim dblSum, dblW07, dblW08, dblW09
	
	If IsRunEvents Then Exit Sub	' �Ʒ� .vlaue = ���� �̺�Ʈ�� �߻��� ����Լ��� ���°� ���´�.
	
	IsRunEvents = True
	
	With frm1
		dblW07 = UNICDbl(.txtData(C_W07).value)
		dblW08 = UNICDbl(.txtData(C_W08).value)
		dblW09 = UNICDbl(.txtData(C_W09).value)
		dblSum = dblw07 + dblW08 + dblW09
		.txtData(C_W_SUM).value = dblSum
	End With

	lgBlnFlgChgValue= True ' ���濩�� 
	IsRunEvents = False	' �̺�Ʈ �߻������� ������ 
End Sub

Function  OpenPopUp(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd, strCode
	
	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 1
			strCode = frm1.txtData(C_W04).value
		Case 2
			strCode = frm1.txtData(C_W05).value
	End Select
	
	arrParam(0) = "ǥ�ؼҵ���"								' �˾� ��Ī 
	arrParam(1) = "tb_std_income_rate" 								' TABLE ��Ī 
	arrParam(2) = Trim(strCode)										' Code Condition
	arrParam(3) = ""												' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "ǥ�ؼҵ���"									' �����ʵ��� �� ��Ī 
            
	arrField(0) = "STD_INCM_RT_CD"									' Field��(0)
	arrField(1) = "BUSNSECT_NM"									' Field��(1)
	arrField(2) = "DETAIL_NM"									' Field��(1)
	arrField(3) = "FULL_DETAIL_NM"									' Field��(1)
			
	arrHeader(0) = " ��ȣ"									' Header��(0)
	arrHeader(1) = "����"									' Header��(1)
	arrHeader(2) = "����"									' Header��(1)
	arrHeader(3) = "������"									' Header��(1)
	
	IsOpenPop = True
			
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=750px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function

Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtData(C_W04).value = arrRet(0)
			Case 2
				.txtData(C_W05).value = arrRet(0)
		End Select
	End With
	
	lgBlnFlgChgValue = True

End Function

Sub QueryRadio()
	' -- ������ ������ ���� ������ ������ư�� ��ó���Ѵ�.
	With frm1
		If .txtData(C_W19).value = "1" Then
			.txtW19(0).checked = true
		Else
			.txtW19(1).checked = true
		End If
		
		If .txtData(C_W20).value = "1" Then
			.txtW20(0).checked = true
		Else
			.txtW20(1).checked = true
		End If
		
		If .txtData(C_W21).value = "1" Then
			.txtW21(0).checked = true
		Else
			.txtW21(1).checked = true
		End If
		
		If .txtData(C_W22).value = "1" Then
			.txtW22(0).checked = true
		Else
			.txtW22(1).checked = true
		End If
		
		If .txtData(C_W23).value = "1" Then
			.txtW23(0).checked = true
		Else
			.txtW23(1).checked = true
		End If
		
	End With
End Sub

Sub SaveRadio()
	' -- ���̺�� ���� ���õ� ������ư�� �������� ó���Ѵ�.
	With frm1
		If .txtW19(0).checked = true Then
			.txtData(C_W19).value = "1"
		Else
			.txtData(C_W19).value = "2"
		End If
		
		If .txtW20(0).checked = true Then
			.txtData(C_W20).value = "1"
		Else
			.txtData(C_W20).value = "2"
		End If
		
		If .txtW21(0).checked = true Then
			.txtData(C_W21).value = "1"
		Else
			.txtData(C_W21).value = "2"
		End If
		
		If .txtW22(0).checked = true Then
			.txtData(C_W22).value = "1"
		Else
			.txtData(C_W22).value = "2"
		End If
		
		.txtW23(0).checked = false
		.txtW23(1).checked = false
			
		' -- �������� 
		If .txtData(C_W19).value = "1" And .txtData(C_W20).value = "1" And _
			.txtData(C_W21).value = "1" And .txtData(C_W22).value = "1" Then
			.txtW23(0).checked = true
			.txtData(C_W23).value = "1"
		ElseIf .txtData(C_W20).value = "2" And .txtData(C_W22).value = "1" Then
			.txtW23(0).checked = true
			.txtData(C_W23).value = "1"
		Else
			.txtW23(1).checked = true
			.txtData(C_W23).value = "2"
		End If
	
	End With
End Sub

Sub GetCompanyInfo()	' ��û������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, sFISC_START_DT, sFISC_END_DT, datMonCnt, i, datNow
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("IND_CLASS, HOME_TAX_MAIN_IND"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    With frm1
		IsRunEvents = True
		If lgF0 <> "" Then 
			.txtData(C_W01).value = Replace(lgF0, chr(11), "")
			.txtData(C_W04).value = Replace(lgF1, chr(11), "")
		Else
			.txtData(C_W01).value = ""
			.txtData(C_W04).value = ""
		End If
		IsRunEvents = False
	End With
End Sub

Sub RadioClicked()
	lgBlnFlgChgValue = True
End Sub
'====================================== �� �Լ� =========================================

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 
	Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call ggoOper.FormatDate(frm1.txtData(C_W18), parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
	Call InitData 
	'
    Call fncQuery() 
    
End Sub


'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	Call GetCompanyInfo
End Sub

Sub txtFISC_YEAR_Change()
	Call GetCompanyInfo
End Sub
'============================================  �׸��� �̺�Ʈ   ====================================

'============================================  �������� �Լ�  ====================================

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                               <%'Protect system from crashing%>

	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>  
    If lgBlnFlgChgValue Or blnChange Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call InitVariables													<%'Initializes local global variables%>
    'Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, i, sMsg
    
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	    

    If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

' ----------------------  ���� -------------------------
Function  Verification()

	Verification = False

	
	Verification = True	
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

    Call SetToolbar("1111100000000011")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

 	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 

End Function

Function FncInsertRow(ByVal pvRowCnt) 

End Function

Function FncDeleteRow() 

End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'��: ȭ�� ���� %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'��:ȭ�� ����, Tab ���� %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    If lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
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

'============================================  DB �＼�� �Լ�  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key   
        'strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	'-----------------------
	'Reset variables area
	'-----------------------
	' �������� ���� : ���ߵǸ� ���ȴ�.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ : �׸��� �� 
	If wgConfirmFlg = "N" Then
		lgIntFlgMode = parent.OPMD_UMODE
	
		' ���� �ڵ� ȯ�氪�� ����� ���� ���� 
		With frm1

		End With
		Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>
	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("1100000000000111")										<%'��ư ���� ���� %>
	End If
	
	'lgvspdData(lgCurrGrid).focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    Call SaveRadio	' -- ������ư ó�� 
    
	With frm1
	
		For i = C_W01 To C_W22	
			If i = C_W18 Then
				strVal = strVal & .txtData(i).text & Parent.gColSep
			Else
				strVal = strVal & .txtData(i).value & Parent.gColSep
			End If
		Next 

	End With

	Frm1.txtSpread.value      =  strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	frm1.txtHeadMode.value	  =  lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' ���� ������ ���� ���� %>
	Call InitVariables
	
    Call MainQuery()
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT LANGUAGE=javascript FOR=txtData EVENT=Change>
<!--
	if (this.WithEvent == "1") {
		SetHeadReCalc();
	} else if (this.WithEvent == "2") {
		RadioClicked();
	}
//-->
</SCRIPT>
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	

	
		<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">�߼ұ���⺻������� ��ǥ1</A>  
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">�������</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/wb107ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�Ű���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE <%=LR_SPACE_TYPE_20%> border="0" height=100% width="100%">
						   <TR>
								<TD>
									<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
									 <TR HEIGHT=25>
										   <TD CLASS="TD51" width="10%" COLSPAN=2>(1) �� ��</TD>
										   <TD CLASS="TD51" width="52%">(2) �� �� �� ��</TD>
										   <TD CLASS="TD51" width="8%">(3)���տ���</TD>
										   <TD CLASS="TD51" width="8%">(4)��������</TD>
									</TR>
									<TR>
									       <TD CLASS="TD61" width="10%" COLSPAN=2 ALIGN=CENTER title = "������(����Ư�����ѹ������Ģ��2����1���� ���������� ����), ����, �Ǽ���, �����Ͼ���, �������, �ؿ���� ���Ѽ��ڰ�����, ������� ������۾�, ���, ���ž�, �Ҹž�, ������ž�, ���� �� ���߾�, ��۾�, ����ó�� �� ��Ÿ��ǻ�Ϳ���þ�, �ڵ��������, �Ƿ��, ��⹰ó����, ���ó����, �д�����ÿ���, �۹�����, ����, ���йױ�����񽺾�, �����������, ��ȭ���, �������, ���������ξ�, ����������, �����, �������û��, ��������о��п�, �������(ī����, ���������������� �� �ܱ��� ���������������� ����), ���κ������� ���� ���κ����ü����, �����ȭ��" >(101)<br> �� ��<br>�� ��</TD>
										   <TD CLASS="TD61">
										   <TABLE <%=LR_SPACE_TYPE_20%> border="1" height=90% width="90%">
											<TR>
												<TD CLASS="TD51" width="30%">���º�&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;����</TD>
												<TD CLASS="TD51" width="30%">���ذ�����ڵ�</TD>
												<TD CLASS="TD51" width="40%">������Աݾ�</TD>
											</TR>
											<TR>
												<TD CLASS="TD61">(01) (<INPUT TYPE=text id="txtData" name=txtData size=15 maxlength=10 tag="25X" WithEvent="2">)��</TD>
												<TD CLASS="TD61">(04) <INPUT TYPE=text id="txtData" name=txtData size=10 maxlength=6 tag="25X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopUp(1)"></TD>
												<TD CLASS="TD61">(07) <script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script></TD>
											</TR>
											<TR>
												<TD CLASS="TD61">(02) (<INPUT TYPE=text id="txtData" name=txtData size=15 maxlength=10 tag="25X" WithEvent="2">)��</TD>
												<TD CLASS="TD61">(05) <INPUT TYPE=text id="txtData" name=txtData size=10 maxlength=6 tag="25X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopUp(2)"></TD>
												<TD CLASS="TD61">(08) <script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script></TD>
											</TR>
											<TR>
												<TD CLASS="TD61">(03) ��Ÿ���</TD>
												<TD CLASS="TD61">(06) <INPUT TYPE=text id="txtData" name=txtData size=10 maxlength=6 tag="25X" WithEvent="2"></TD>
												<TD CLASS="TD61">(09) <script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script></TD>
											</TR>
											<TR>
												<TD CLASS="TD61" ALIGN=CENTER>��</TD>
												<TD CLASS="TD61">&nbsp;</TD>
												<TD CLASS="TD61"><script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script></TD>
											</TR>
										   </TABLE>
										   </TD>
										   <TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE>(19)<br> <INPUT TYPE=RADIO NAME=txtW19 ID=txtW19 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>����<br>(Y)<br><br><INPUT TYPE=RADIO NAME=txtW19 ID=txtW19 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>������<br>(N)
										   <INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="24"></TD>
										   <TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE ROWSPAN=4>(23)<br> <INPUT TYPE=RADIO NAME=txtW23 ID=txtW23 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>��<br>(Y)<br><br><br><INPUT TYPE=RADIO NAME=txtW23 ID=txtW23 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>��<br>(N)
										   <INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="23"></TD>
									</TR>
									<TR>										   
										   <TD CLASS="TD61" ALIGN=CENTER COLSPAN=2 title = "�� �Ʒ� ��� ��,�踦 ���ÿ� ������ ��                  �� ��� ��� �����������ں��ݤ������ �� �ϳ� �̻��� �߼ұ���⺻������� ��ǥ1�� �Ը���� �̳��ϰ� ����������(����������� 1õ�� �̸�, �ڱ��ں� 1õ�� �̸�, ����� 1õ�� �̸�)�̳��� ��" >(102) ��������<br>���ں���<br>�������<br>���ڱ��ں����ڻ�<br>���� </TD>
										   <TD CLASS="TD61">
										   <TABLE <%=LR_SPACE_TYPE_20%> border="0" height=100% width="100%">
											<TR>
												<TD CLASS="TD61" COLSPAN=4>&nbsp;��. ��� ��������(������ο�)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=2 width=30%>(1) �� ȸ��(10)</TD>
												<TD CLASS="TD61" width=*>(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>��)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=3>(2) �߼ұ���⺻������� ��ǥ1��</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" >�Ը����(11)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>��)�̸�</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" COLSPAN=4>&nbsp;��. �ں���</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=2>(1) �� ȸ��(12)</TD>
												<TD CLASS="TD61">(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>��)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=3>(2) �߼ұ���⺻������� ��ǥ1��</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" >�Ը����(13)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>��)����</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" COLSPAN=4>&nbsp;��. �����</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=2>(1) �� ȸ��(14)</TD>
												<TD CLASS="TD61">(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>��)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=3>(2) �߼ұ���⺻������� ��ǥ1��</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" >�Ը����(15)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>��)����</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" COLSPAN=3>&nbsp;��. �ڱ��ں�(16)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>��)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" COLSPAN=4>&nbsp;��. ����.��ȸ��Ϲ����� ���</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10></TD>
												<TD CLASS="TD61" width=10></TD>
												<TD CLASS="TD61" >�Ը����(17)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>��)</TD>
											</TR>
										   </TABLE>
										   </TD>
										   <TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE>(20)<br> <INPUT TYPE=RADIO NAME=txtW20 ID=txtW20 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>����<br>(Y)<br><br><INPUT TYPE=RADIO NAME=txtW20 ID=txtW20 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>������<br>(N)
										   <INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS="TD61"  ALIGN=CENTER COLSPAN=2 title = "�߼ұ���⺻������� ��3�� ��2ȣ�� ���ؿ� ���� ������ ����">(103) ������<br>�濵��<br>������  </TD>
												<TD CLASS="TD61">&nbsp;���ڻ��Ѿ� 5,000��� �̻��� ������ �����ֽ��� 30%�̻� �����ϰ� �ִ� ������ �ƴ� ��<BR>
												&nbsp;�۵������� �� �����ŷ��� ���� ������ ���� ��ȣ�������ѱ�����ܿ� ������ ������</TD>
												<TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE>(21)<INPUT TYPE=RADIO NAME=txtW21 ID=txtW21 tag="21" CLASS="RADIO" onclick="RadioClicked()">����(Y)<br><INPUT TYPE=RADIO NAME=txtW21 ID=txtW21 tag="21" CLASS="RADIO" onclick="RadioClicked()">������(N)
												<INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS="TD61"  ALIGN=CENTER COLSPAN=2 title = "01.1���� ���ÿ��� �� �߼ұ�� ������ �Ը� �ʰ��ϴ� ������ �ʰ������� �� �� 3�Ⱓ �߼ұ������ ���� �� �Ŀ��� �ų⸶�� �Ǵ�">(104) �����Ⱓ</TD>
												<TD CLASS="TD61">&nbsp;�� �ʰ�����(18) (<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>)�� 
												&nbsp;&nbsp;* 2001����</TD>
												<TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE>(22)<INPUT TYPE=RADIO NAME=txtW22 ID=txtW22 tag="21" CLASS="RADIO" onclick="RadioClicked()">����(Y)<br><INPUT TYPE=RADIO NAME=txtW22 ID=txtW22 tag="21" CLASS="RADIO" onclick="RadioClicked()">������(N)
												<INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="24"></TD>
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
	</TR>
		<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
	
		
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHeadMode" tag="24">
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="strUrl" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

