<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : �޿�����Ÿ��� 
*  3. Program ID           : H2001ma1
*  4. Program Name         : H2001ma1
*  5. Program Desc         : �޿�����Ÿ��� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/04
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : Hwang Jeong-Won
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"> </SCRIPT>

<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Const BIZ_PGM_ID = "h6002mb1.asp"												'�����Ͻ� ���� ASP�� 
Const TAB1 = 1
Const TAB2 = 2

Dim gSelframeFlg                                                       '���� TAB�� ��ġ�� ��Ÿ���� Flag %>
Dim lsConcd
Dim IsOpenPop

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
End Sub

'===========================================  2.3.1 Tab Click ó��  =====================================
'=	Name : Tab Click																					=
'=	Description : Tab Click�� �ʿ��� ����� �����Ѵ�.													=
'========================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
		
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
End Function

'======================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=======================================================================================================
Sub InitComboBox()
	Dim iCodeArr
    Dim iNameArr
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = 'H0038' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboMedType, iCodeArr, iNameArr, Chr(11))
    
    ' �޿����� 
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = 'H0005' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboPayCd, iCodeArr, iNameArr, Chr(11))
    
    ' ���ױ��� 
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = 'H0006' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboTaxCd, iCodeArr, iNameArr, Chr(11))
    
    iCodeArr = "Y" & Chr(11) & "N" & Chr(11)
    iNameArr = "Y" & Chr(11) & "N" & Chr(11)
    Call SetCombo2(frm1.cboSpouseAllow, iCodeArr, iNameArr, Chr(11))
    
End Sub

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmp()
    
    Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmpNo.value			' Code Condition
	arrParam(1) = ""'frm1.txtEmpNm.value			' Name Cindition
    arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmpNo.focus	
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmpNo.value = arrRet(0)
		.txtEmpNm.value = arrRet(1)
		Call ggoOper.ClearField(Document, "2")					 '��: Clear Contents  Field
		.txtEmpNo.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'======================================================================================================
'	Name : OpenBank()
'	Description : Bank PopUp
'=======================================================================================================
Function OpenBank(Byval flgs)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����ڵ� �˾�"			<%' �˾� ��Ī %>
	arrParam(1) = "B_Bank"				 		<%' TABLE ��Ī %>
	arrParam(2) = frm1.txtBank.value			<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "�����ڵ�"			
	
    arrField(0) = "bank_cd"					<%' Field��(0)%>
    arrField(1) = "bank_nm"				<%' Field��(1)%>
    
    arrHeader(0) = "�����ڵ�"						<%' Header��(0)%>
    arrHeader(1) = "�����"					<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBank(arrRet,flgs)
	End If	

End Function

'======================================================================================================
'	Name : SetBank()
'	Description : Bank Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetBank(Byval arrRet ,Byval flgs)
	With frm1

		Select Case flgs      
		    Case 1
				.txtBank.value = arrRet(0)
				.txtBankNm.value = arrRet(1)

		    Case 2
				.txtBank2.value = arrRet(0)
				.txtBankNm2.value = arrRet(1)

		    Case 3
		    	.txtBank3.value = arrRet(0)
				.txtBankNm3.value = arrRet(1)
	
		End Select
	
	End With
	
lgBlnFlgChgValue = True

End Function
'======================================================================================================
'	Name : OpenCode()
'	Description : Grade PopUp
'=======================================================================================================
Function OpenCode(Byval flgs)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case flgs
        
	    Case 1

			arrParam(0) = "�ǰ������� �˾�"			<%' �˾� ��Ī %>
			arrParam(1) = "hdb010t"				 			<%' TABLE ��Ī %>
			arrParam(2) = frm1.txtInsurGrade.value			<%' Code Condition%>
			arrParam(3) = ""								<%' Name Cindition%>
			arrParam(4) = " insur_type = '1' and insur_area = '*'"							<%' Where Condition%>
			arrParam(5) = "�ǰ�������"			
	    Case 2
			arrParam(0) = "���ο��ݵ�� �˾�"			<%' �˾� ��Ī %>
			arrParam(1) = "hdb010t"				 			<%' TABLE ��Ī %>
			arrParam(2) = frm1.txtAnutGrade.value			<%' Code Condition%>
			arrParam(3) = ""								<%' Name Cindition%>
			arrParam(4) = " insur_type = '2' and insur_area = '*'"							<%' Where Condition%>
			arrParam(5) = "���ο��ݵ��"				    		

	End Select

			arrField(0) = "ED7"  & Parent.gColSep &"grade"					<%' Field��(0)%>
			arrField(1) = "F213" & Parent.gColSep & "std_strt_amt"			<%' Field��(1)%>
			arrField(2) = "F213" & Parent.gColSep & "std_end_amt"			<%' Field��(1)%>
			arrField(3) = "F211" & Parent.gColSep & "std_amt"				<%' Field��(1)%>
			arrField(4) = "F211" & Parent.gColSep & "insur_amt"				<%' Field��(1)%>
			arrField(5) = "ED7"  & Parent.gColSep & "insur_rate"			<%' Field��(1)%>
    
			arrHeader(0) = "���"				<%' Header��(0)%>
			arrHeader(1) = "����ǥ�غ�������"	<%' Header��(1)%>
			arrHeader(2) = "����ǥ�غ�������"	<%' Header��(1)%>
			arrHeader(3) = "ǥ�غ�������"		<%' Header��(1)%>
			arrHeader(4) = "�����"				<%' Header��(1)%>
			arrHeader(5) = "�����"				<%' Header��(1)%>

	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1

			Select Case flgs      
			    Case 1
					.txtInsurGrade.value = arrRet(0) 
			    Case 2
					.txtAnutGrade.value = arrRet(0) 					
			End Select
			lgBlnFlgChgValue = True

		End With
	End If	

End Function


Function CookiePage(ByVal flgs)
End Function

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ 
'=======================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field%>                         
    
    Call AppendNumberPlace("6", "2", "0")                                   'Format Numeric Contents Field%>
    Call AppendNumberPlace("7", "4", "0")                                   'Format Numeric Contents Field%>
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")

    Call InitVariables                                                      'Initializes local global variables%>
    
    gSelframeFlg = TAB1
	Call SetToolBar("1100100000001111")												'��: Set ToolBar

    Call changeTabs(TAB1)
    
    frm1.txtEmpNo.focus() 
    gIsTab     = "Y" ' <- "Yes"�� ���� Y(����) �Դϴ�.[V(����)�ƴմϴ�]
    gTabMaxCnt = 2   ' Tab�� ������ ���� �ּ���    

    Call InitComboBox()
    
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub
'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False    
    Err.Clear                                                               <%'Protect system from crashing%>

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    If  txtEmpNo_Onchange() then
        Exit Function
    End If

    If frm1.txtEmpNo.value = "" Then frm1.txtEmpNm.value = ""
        
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
        Call RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '��: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '��: Lock  Field
    
    Call SetToolbar("11001000000011")
    Call InitVariables                                                        '��: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True
End Function
   
'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '��: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")                        '��: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
	Call DisableToolBar(Parent.TBC_DELETE)
    If DbDelete = False Then
        Call RestoreToolBar()
        Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    
    FncDelete = True                                                            '��: Processing is OK
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
    Dim intRetCD
    
    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If
    
	If frm1.txtEmpNo.value <> frm1.txtHEmpNo.value Then
		IntRetCD = DisplayMsgBox("800603", Parent.VB_YES_NO , frm1.txtHEmpNm.value&"("&frm1.txtHEmpNo.value&")" ,"x")  '��: Data is changed.  Do you want to display it? 
		If IntRetCD <> vbYes Then
		    lgBlnFlgChgValue = False
'		    Call FncQuery() 
			Exit Function
		End If
    End If

    If frm1.cboPayCd.value = "" Then
		intRetCD = DisplayMsgBox("970021","x", frm1.cboPayCd.Alt,"x")
	    Call changeTabs(TAB2)
        gSelframeFlg = TAB2
        frm1.cboPayCd.value = ""
		frm1.cboPayCd.focus 
        Set gActiveElement = document.activeElement        'focus �̵� 
		Exit Function
	ElseIf frm1.cboTaxCd.value = "" Then
		intRetCD = DisplayMsgBox("970021","x", frm1.cboTaxCd.Alt,"x")
	    Call changeTabs(TAB2)
        gSelframeFlg = TAB2
        frm1.cboTaxCd.value = ""
		frm1.cboTaxCd.focus
        Set gActiveElement = document.activeElement        'focus �̵� 
		Exit Function
    End If
    
    If frm1.txtAccntNo.value <> "" And frm1.txtBankMaster.value = "" Then
        frm1.txtBankMaster.value = frm1.txtHEmpNm.value
    End If
    If frm1.txtAccntNo2.value <> "" And frm1.txtBankMaster2.value = "" Then
        frm1.txtBankMaster2.value = frm1.txtHEmpNm.value
    End If
    If frm1.txtAccntNo3.value <> "" And frm1.txtBankMaster3.value = "" Then
        frm1.txtBankMaster3.value = frm1.txtHEmpNm.value
    End If           
    
    
    If frm1.txtInsurGrade.value <> "" Then
		intRetCD = CommonQueryRs(" grade "," hdb010t "," insur_type = '1' and grade = " & _
		           FilterVar(Trim(frm1.txtInsurGrade.value), "''", "S")  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If intRetCD = false Then
            Call DisplayMsgBox("800257","X","X","X")
	        Call changeTabs(TAB1)
            gSelframeFlg = TAB1
            frm1.txtInsurGrade.value = ""
            frm1.txtInsurGrade.focus
            Set gActiveElement = document.activeElement        'focus �̵� 
            Exit Function
        End If
    End If
    
    If frm1.txtAnutGrade.value <> "" Then
		intRetCD = CommonQueryRs(" grade "," hdb010t "," insur_type = '2'  and grade = " & _
		           FilterVar(Trim(frm1.txtAnutGrade.value), "''", "S")  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If intRetCD = false Then
            Call DisplayMsgBox("800138","X","X","X")
	        Call changeTabs(TAB1)
            gSelframeFlg = TAB1
            frm1.txtAnutGrade.value = ""
            frm1.txtAnutGrade.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End If
    End If
        
    
    If CompareDateByFormat(frm1.txtAnutAcqDt.text,frm1.txtAnutlossDt.text,frm1.txtAnutAcqDt.Alt,frm1.txtAnutLossDt.Alt,_
      "970025", Parent.gDateFormat,Parent.gComDateType,True) = False Then
		Call changeTabs(TAB1)
		gSelframeFlg = TAB1
		frm1.txtanutAcqDt.focus()
		Set gActiveElement = document.activeElement   
		Exit Function
	End If
	
	If CompareDateByFormat(frm1.txtMedAcqDt.text,frm1.txtMedlossDt.text,frm1.txtMedAcqDt.Alt,frm1.txtMedLossDt.Alt,_
      "970025", Parent.gDateFormat,Parent.gComDateType,True) = False Then
		Call changeTabs(TAB1)
		gSelframeFlg = TAB1
		frm1.txtMedAcqDt.focus()
		Set gActiveElement = document.activeElement   
		Exit Function
	End If
    
	Call DisableToolBar(Parent.TBC_SAVE)
    If DbSave = False Then
        Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True
    
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()

End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 

End Function
'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '��: Processing is OK
    Err.Clear                                                                    '��: Clear err status
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '��: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
    Call ggoOper.ClearField(Document, "2")										 '��: Clear Contents Area
    
    Call InitVariables														 '��: Initializes local global variables

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0001                   '��: Query
    strVal = strVal     & "&txtEmpNo="         & frm1.txtEmpNo.value         '��: Query Key
    strVal = strVal     & "&txtInternal="      & lgUsrIntCd			         '��: Query Key
    strVal = strVal     & "&txtPrevKey="       & "P"	                         '��: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '��: Run Biz 

    FncPrev = True
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '��: Processing is OK
    Err.Clear                                                                    '��: Clear err status
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '��: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '��: Clear Contents Area
    
    Call InitVariables														 '��: Initializes local global variables

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If


    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0001                       '��: Query
    strVal = strVal     & "&txtEmpNo="         & frm1.txtEmpNo.value             '��: Query Key
    strVal = strVal     & "&txtInternal="      & lgUsrIntCd			         '��: Query Key
    strVal = strVal     & "&txtPrevKey="       & "N"	                         '��: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '��: Run Biz 

    FncNext = True
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 <%'��: ȭ�� ���� %>
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call Parent.FncFind(Parent.C_SINGLE, True)
End Function

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'��: Protect system from crashing%>

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal
    
    With frm1
    
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						<%'���� �˻��������� Query%>
		strVal = strVal & "&txtEmpNo=" & .txtEmpNo.value
		strVal = strVal & "&txtInternal=" & lgUsrIntCd
		strVal = strVal & "&txtPrevKey=" & lgStrPrevKey	
    
	Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
        
    End With
    
    DbQuery = True
    
End Function

'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================
Function DbQueryOk()													'��ȸ ������ ������� 
	
    lgIntFlgMode = Parent.OPMD_UMODE
    lgBlnFlgChgValue = False    
    Call ggoOper.LockField(Document, "Q")								'This function lock the suitable field    
	Call SetToolbar("1100100011011111")									'��ư ���� ���� 
	
	If frm1.txtSexCd.value = "1" Then
		frm1.chkLadyFlg.checked = False
		frm1.chkLadyFlg.disabled = True
		
'		If frm1.chkSpouseFlg.checked = True Then		
'			frm1.txtchild.text = "0"
'			frm1.txtchild.enabled = False
'		Else		
'			frm1.txtchild.enabled = True
'		End If
	Else
		frm1.chkLadyFlg.disabled = False
		frm1.txtchild.enabled = True
	End If
	If gSelframeFlg = TAB1 Then		
		frm1.txtInsurGrade.focus
	else
		frm1.cboPayCd.focus
	end if
	
End Function

'======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbSave()     
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	With Frm1
		.txtMode.value        = Parent.UID_M0002                                        '��: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtEmpNo.Value		  = .txtEmpNo.value                                       '��: Save Key
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                            
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
    Call InitVariables
	Call MainQuery()
End Function

'======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbDelete = False			                                                 '��: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0003                       '��: Query
    strVal = strVal     & "&txtEmpNo="         & frm1.txtEmpNo.value             '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '��: Direction
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '��: Run Biz logic
	
	DbDelete = True                                                              '��: Processing is NG
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables
	Call MainNew()	
End Function

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
Sub txtBank_OnChange()

    If  frm1.txtBank.value <> "" Then
        if  CommonQueryRs(" bank_nm "," B_BANK "," bank_cd = " & FilterVar(frm1.txtBank.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtBankNm.value = ""
            Call DisplayMsgBox("800137", Parent.VB_INFORMATION,"x","x")            
	        frm1.txtBank.focus
	        Set gActiveElement = document.ActiveElement
	    Else
	        frm1.txtBankNm.value = Replace(lgF0, Chr(11), "")
	    End If
	ELSE
        frm1.txtBankNm.value = ""
    End If

End Sub
Sub txtBank2_OnChange()

    If  frm1.txtBank2.value <> "" Then
        if  CommonQueryRs(" bank_nm "," B_BANK "," bank_cd = " & FilterVar(frm1.txtBank2.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtBankNm2.value = ""
            Call DisplayMsgBox("800137", Parent.VB_INFORMATION,"x","x")            
	        frm1.txtBank2.focus
	        Set gActiveElement = document.ActiveElement
	    Else
	        frm1.txtBankNm2.value = Replace(lgF0, Chr(11), "")
	    End If
	ELSE
        frm1.txtBankNm2.value = ""
    End If

End Sub
Sub txtBank3_OnChange()

    If  frm1.txtBank3.value <> "" Then
        if  CommonQueryRs(" bank_nm "," B_BANK "," bank_cd = " & FilterVar(frm1.txtBank3.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtBankNm3.value = ""
            Call DisplayMsgBox("800137", Parent.VB_INFORMATION,"x","x")            
	        frm1.txtBank3.focus
	        Set gActiveElement = document.ActiveElement
	    Else
	        frm1.txtBankNm3.value = Replace(lgF0, Chr(11), "")
	    End If
	ELSE
        frm1.txtBankNm3.value = ""
    End If

End Sub

'========================================================================================================
'   Event Name : txtEmpNo_Onchange           
'   Event Desc :
'========================================================================================================
Function txtEmpNo_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

	frm1.txtEmpNm.value = ""

    If  frm1.txtEmpNo.value = "" Then
		frm1.txtEmpNm.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmpNo.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            end if
            Call ggoOper.ClearField(Document, "2")
            call InitVariables()
            frm1.txtEmpNo.focus
            Set gActiveElement = document.ActiveElement
            txtEmpNo_Onchange = true
        Else
            frm1.txtEmpNm.value = strName
        End if 
    End if
    
End Function
'-------------------------------------------------------------------------------------

Sub txtInsurGrade_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtMedInsurNo_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtMedAcqDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMedLossDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSuppCnt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSupp_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtAnutGrade_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtAnutNo_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtAnutAcqDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtAnutLossDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtAnnualSal_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSalary_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtBonusSalary_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtBankMaster_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtBankMaster2_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtBankMaster3_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtAccntNo_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtAccntNo2_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtAccntNo3_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtOld_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtYoung_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtParia_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtOldCnt1_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtOldCnt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtChild_Change()
	lgBlnFlgChgValue = True
End Sub

Sub cboMedType_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboSpouseAllow_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboPayCd_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboTaxCd_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub chkPayFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub chkEmpInsurFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtForeign_separate_tax_yn_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtForeign_no_tax_yn_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub chkYearFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub chkRetireFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub chkTaxFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub chkYearTaxFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub chkSpouseFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub chkLadyFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : Radio OnClick()
'   Event Desc : Radio Button Click�� lgBlnFlgChgValue ó�� / Value
'==========================================================================================
Sub rdoUnionFlag1_OnClick()
	lgBlnFlgChgValue = True	
End Sub

Sub rdoUnionFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPressFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPressFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoOverseaFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoOverseaFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoResFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoResFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtMedAcqDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtMedAcqDt.Action = 7
        frm1.txtMedAcqDt.focus
    End If
End Sub

Sub txtMedLossDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtMedLossDt.Action = 7
        frm1.txtMedLossDt.focus
    End If
End Sub

Sub txtAnutAcqDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtAnutAcqDt.Action = 7
        frm1.txtAnutAcqDt.focus
    End If
End Sub

Sub txtAnutLossDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtAnutLossDt.Action = 7
        frm1.txtAnutLossDt.focus
    End If
End Sub


Sub rdoBankFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoBankFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoBankFlag3_OnClick()
	lgBlnFlgChgValue = True
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
									<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�⺻��������</font></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
							    </TR>
							</TABLE>
						</TD>
						<TD CLASS="CLSMTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
								<TR>
									<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
									<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�޿����û���</font></td>
									<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
							    </TR>
							</TABLE>
						</TD>						
						<TD WIDTH=*>&nbsp;</TD>
					</TR>					
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5">���</TD>
								<TD CLASS="TD656">
									<INPUT TYPE=TEXT NAME="txtEmpNo" SIZE=13 MAXLENGTH=13 tag="12XXXU"  ALT="���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">
									<INPUT TYPE=TEXT NAME="txtEmpNm" SIZE=20 tag="14X">
								</TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
					<!-- ù��° �� ���� -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS=TD5 NOWRAP>���</TD>   
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtHEmpNo" ALT="���" TYPE="Text" SiZE=20 tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>����</TD>
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtHEmpNm" ALT="����" TYPE="Text" SiZE=20  tag="24"></TD>
						</TR>					
						<TR>
							<TD CLASS=TD5 NOWRAP>�μ�</TD>   
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptNm" ALT="�μ�" TYPE="Text" SiZE=20 tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>�Ի籸��</TD>
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="cboEntrCd" ALT="�Ի籸��" TYPE="Text" SiZE=20  tag="24"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>����</TD>
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="cboRollPstn" ALT="����" TYPE="Text" SiZE=20  tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>�׷��Ի���</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtGroupEntrDt NAME="txtGroupEntrDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="�׷��Ի���"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>����</TD>
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="cboOcptType" ALT="����" TYPE="Text" SiZE=20  tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>����Ի���</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtEntrDt NAME="txtEntrDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="����Ի���"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>����</TD>
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="cboFuncCd" ALT="����" TYPE="Text" SiZE=20  tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>����������</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtInternDt NAME="txtInternDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="����������"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>��å</TD>
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="cboRoleCd" ALT="��å" TYPE="Text" SiZE=20  tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>��������</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtRestDt NAME="txtRestDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="��������"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>��ȣ</TD>
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="cboPay_grd1" ALT="��ȣ" TYPE="Text" SiZE=20  tag="24">
							            	     <INPUT NAME="txtPay_grd2" TYPE=TEXT SIZE="5" TAG="24" ALT="ȣ��">ȣ��</TD>
							<TD CLASS=TD5 NOWRAP>��°�����</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtCareer NAME="txtCareer" CLASS=FPDS115 tag="24X7Z" ALT="��°�����" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP><�ǰ�����></TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>���</TD>   
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInsurGrade" ALT="���" TYPE="Text" SiZE=20 MAXLENGTH=2 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInsurGrade" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode(1)"></TD>
							<TD CLASS=TD5 NOWRAP>����</TD>
							<TD CLASS=TD6 NOWRAP><SELECT NAME="cboMedType" tag="21" CLASS ="cbonormal" ALT="����"><OPTION value=""></OPTION></SELECT></TD>							
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>��ȣ</TD>   
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMedInsurNo" ALT="��ȣ" TYPE="Text" SIZE=20 MAXLENGTH=20 tag="21XXXU"></TD>
							<TD CLASS=TD5 NOWRAP>�ξ���(��������)</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSuppCnt NAME="txtSuppCnt" CLASS=FPDS115 tag="21X6Z" ALT="�ξ���(��������" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>�����</TD>   
					    	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtMedAcqDt NAME="txtMedAcqDt" CLASS=FPDTYYYYMMDD tag="21" Title="FPDATETIME" ALT="�ǰ����������"></OBJECT>');</SCRIPT></TD>
							<TD CLASS=TD5 NOWRAP>�����</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtMedLossDt NAME="txtMedLossDt" CLASS=FPDTYYYYMMDD tag="21" Title="FPDATETIME" ALT="�ǰ���������"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP><��������></TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>�����</TD>   
					    	<TD CLASS=TD6 NOWRAP><SELECT NAME="cboSpouseAllow" tag="21" CLASS ="cbonormal" ALT="�����"><OPTION value=""></OPTION></SELECT></TD>
							<TD CLASS=TD5 NOWRAP>�ξ���</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSupp NAME="txtSupp" CLASS=FPDS115 tag="21X6Z" ALT="�ξ���" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP><���ο���></TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>���ο��ݵ��</TD>   
					    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAnutGrade" ALT="�μ�" TYPE="Text" SiZE=20 MAXLENGTH=2 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAnutGrade" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode(2)"></TD>
							<TD CLASS=TD5 NOWRAP>��ȣ</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAnutNo" ALT="�Ի籸��" TYPE="Text" SiZE=20 MAXLENGTH=13 tag="21XXXU"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>�����</TD>   
					    	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtAnutAcqDt NAME="txtAnutAcqDt" CLASS=FPDTYYYYMMDD tag="21" Title="FPDATETIME" ALT="���ο��������"></OBJECT>');</SCRIPT></TD>
							<TD CLASS=TD5 NOWRAP>�����</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtAnutLossDt NAME="txtAnutLossDt" CLASS=FPDTYYYYMMDD tag="21" Title="FPDATETIME" ALT="���ο��ݻ����"></OBJECT>');</SCRIPT></TD>
						</TR>						
					</TABLE>
					</DIV>

					<DIV ID="TabDiv" SCROLL=no>
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS=TD5 NOWRAP>�޿�����</TD>
							<TD CLASS=TD6 NOWRAP><SELECT NAME="cboPayCd" tag="22" CLASS ="cbonormal" ALT="�޿�����"><OPTION value=""></OPTION></SELECT></TD>
							<TD CLASS=TD5 NOWRAP>����(������)</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtAnnualSal NAME="txtAnnualSal" CLASS=FPDS140 tag="21X2Z" ALT="����(������)" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>�⺻��(����)</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSalary NAME="txtSalary" CLASS=FPDS140 tag="21X2Z" ALT="�⺻��(����)" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							<TD CLASS=TD5 NOWRAP>�󿩱��ر�(����)</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtBonusSalary NAME="txtBonusSalary" CLASS=FPDS140 tag="21X2Z" ALT="�󿩱��ر�(����)" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>�����������뱸��</TD>
							<TD CLASS=TD6 NOWRAP><SELECT NAME="cboTaxCd" tag="22" CLASS ="cbonormal" ALT="���ױ���"><OPTION value=""></OPTION></SELECT></TD>
							<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>���� 1</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBank" ALT="����1" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditChkType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBank(1)">&nbsp;<INPUT NAME="txtBankNm" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>���¹�ȣ/������ 1</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAccntNo" ALT="���¹�ȣ1" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="21XXXU">
							                     &nbsp;/&nbsp;<INPUT NAME="txtBankMaster" ALT="������1" TYPE="Text" SiZE=20  tag="21">
							                     &nbsp;/&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoBankFlag" TAG="21X" VALUE="1" CHECKED ID="rdoBankFlag1"><LABEL FOR="rdoBankFlag1"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>���� 2</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBank2" ALT="����2" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditChkType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBank(2)">&nbsp;<INPUT NAME="txtBankNm2" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>���¹�ȣ/������ 2</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAccntNo2" ALT="���¹�ȣ2" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="21XXXU">
							                     &nbsp;/&nbsp;<INPUT NAME="txtBankMaster2" ALT="������2" TYPE="Text" SiZE=20  tag="21">
							                     &nbsp;/&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoBankFlag" TAG="21X" VALUE="2" ID="rdoBankFlag2"><LABEL FOR="rdoBankFlag2"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>���� 3</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBank3" ALT="����3" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditChkType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBank(3)">&nbsp;<INPUT NAME="txtBankNm3" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>���¹�ȣ/������ 3</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAccntNo3" ALT="���¹�ȣ3" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="21XXXU">
							                     &nbsp;/&nbsp;<INPUT NAME="txtBankMaster3" ALT="������3" TYPE="Text" SiZE=20  tag="21">
							                     &nbsp;/&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoBankFlag" TAG="21X" VALUE="3" ID="rdoBankFlag3"><LABEL FOR="rdoBankFlag3"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>���ֱ���</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoResFlag" TAG="21X" VALUE="Y" CHECKED ID="rdoResFlag1"><LABEL FOR="rdoResFlag1">������</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoResFlag" TAG="21X" VALUE="N" ID="rdoResFlag2"><LABEL FOR="rdoResFlag2">�������</LABEL>			
							</TD>							
							<TD CLASS=TD5 NOWRAP>���ڱ���</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPressFlag" TAG="21X" VALUE="Y" CHECKED ID="rdoPressFlag1"><LABEL FOR="rdoPressFlag1">����</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPressFlag" TAG="21X" VALUE="N" ID="rdoPressFlag2"><LABEL FOR="rdoPressFlag2">�����</LABEL>			
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>���ܱٷ��ڱ���</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoOverseaFlag" TAG="21X" VALUE="Y" CHECKED ID="rdoOverseaFlag1"><LABEL FOR="rdoOverseaFlag1">���ܱٷ���</LABEL>&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoOverseaFlag" TAG="21X" VALUE="N" ID="rdoOverseaFlag2"><LABEL FOR="rdoOverseaFlag2">�����ٷ���</LABEL>			
							<TD CLASS=TD5 NOWRAP>��������</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoUnionFlag" TAG="21X" VALUE="Y" CHECKED ID="rdoUnionFlag1"><LABEL FOR="rdoUnionFlag1">������</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoUnionFlag" TAG="21X" VALUE="N" ID="rdoUnionFlag2"><LABEL FOR="rdoUnionFlag2">�������</LABEL>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="txtForeign_separate_tax_yn" ID="txtForeign_separate_tax_yn">
								<LABEL FOR="txtForeign_separate_tax_yn">�ܱ��αٷ��ںи��������뿩��</LABEL>
							</TD>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="txtForeign_no_tax_yn" ID="txtForeign_no_tax_yn">
								<LABEL FOR="txtForeign_no_tax_yn">�ܱ��αٷ��ڸ鼼����</LABEL>
							</TD>								
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkPayFlg" ID="chkPayFlg">
								<LABEL FOR="chkPayFlg">�ӱ����޴�󿩺�</LABEL>
							</TD>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkEmpInsurFlg" ID="chkEmpInsurFlg">
								<LABEL FOR="chkEmpInsurFlg">��뺸�迩��</LABEL>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkYearFlg" ID="chkYearFlg">
								<LABEL FOR="chkYearFlg">���������޴��</LABEL>
							</TD>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkRetireFlg" ID="chkRetireFlg">
								<LABEL FOR="chkRetireFlg">���������޴��</LABEL>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkTaxFlg" ID="chkTaxFlg">
								<LABEL FOR="chkTaxFlg">���װ����</LABEL>
							</TD>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkYearTaxFlg" ID="chkYearTaxFlg">
								<LABEL FOR="chkYearTaxFlg">��������Ű���</LABEL>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP><�ҵ����></TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
						</TR>						
						<TR>
							<TD CLASS=TD5 NOWRAP>�ξ���(��)</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtOld NAME="txtOld" CLASS=FPDS115 tag="21X6Z" ALT="�ξ���(��)" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>��</TD>
							<TD CLASS=TD5 NOWRAP>�ξ���(��)</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtYoung NAME="txtYoung" CLASS=FPDS115 tag="21X6Z" ALT="�ξ���(��)" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>��</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>�����</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtParia NAME="txtParia" CLASS=FPDS115 tag="21X6Z" ALT="�����" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>��</TD>
							<TD CLASS=TD5 NOWRAP>�����(65���̻�)</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtOldCnt1 NAME="txtOldCnt1" CLASS=FPDS115 tag="21X6Z" ALT="�����1" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>��</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>�ڳ������</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtChild NAME="txtChild" CLASS=FPDS115 tag="21X6Z" ALT="�ڳ������" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>��</TD>
							<TD CLASS=TD5 NOWRAP>�����(70���̻�)</TD>
							<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtOldCnt2 NAME="txtOldCnt2" CLASS=FPDS115 tag="21X6Z" ALT="�����2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>��</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkSpouseFlg" ID="chkSpouseFlg">
								<LABEL FOR="chkSpouseFlg">�����</LABEL>
							</TD>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkLadyFlg" ID="chkLadyFlg">
								<LABEL FOR="chkLadyFlg">�γ���</LABEL>
							</TD>
						</TR>
					</TABLE>
					</DIV>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
      <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    <TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="h6002mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>    
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSexCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtResNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

