  <%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A6102BA1
'*  4. Program Name         : �ΰ����Ű���ϻ���
'*  5. Program Desc         : �ΰ����Ű���ϻ��� ��ġ
'*  6. Component List       : +
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Hee Jung, Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																		'��: indicates that All variables must be declared in advance 

'==========================================================================================================

Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
EndDate     =   "<%=GetSvrDate%>"

Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
StartDate   = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth-3, "01")
EndDate     = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth-1, strDay) 

Const BIZ_PGM_ID = "a6102bb1.asp"													'��: �����Ͻ� ���� ASP��
Const BIZ_PGM_ID2 = "a6102bb2.asp"													'��: �����Ͻ� ���� ASP��
Const BIZ_PGM_ID3 = "a6102bb3.asp"	
Const BIZ_PGM_ID4 = "a6102bb4.asp"	

Const TAB1 = 1																		'��: Tab�� ��ġ
Const TAB2 = 2
										 '��: �����Ͻ� ���� ASP��
 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ ��
'========================================================================================================= 
Dim lgBlnFlgConChg																	'��: Condition ���� Flag
Dim lgBlnFlgChgValue																'��: Variable is for Dirty flag
Dim lgIntFlgMode																	'��: Variable is for Operation Status

'==========================================  1.2.3 Global Variable�� ����  ===============================

Dim lgMpsFirmDate, lgLlcGivenDt														'��: �����Ͻ� ���� ASP���� �����ϹǷ� 

Dim  lgCurName()																	'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim  cboOldVal          
Dim  IsOpenPop          
Dim  gSelframeFlg

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE												'��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False														'��: Indicates that no value changed

    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False																'��: ����� ���� �ʱ�ȭ
    lgMpsFirmDate=""
    lgLlcGivenDt=""
    
End Sub

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtIssueDT1.Text = StartDate
	frm1.txtIssueDT2.Text = EndDate
	frm1.txtReportDt.Text = EndDate
	frm1.txtYear.text	= strYear
	'frm1.txtBizAreaCD.value	= gBizArea
	frm1.txtBizAreaNM.value	= ""

	frm1.txtIssueDT3.Text = StartDate
	frm1.txtIssueDT4.Text = EndDate
		
	frm1.txtIssueDT5.Text = StartDate
	frm1.txtIssueDT6.Text = EndDate
		
	'frm1.txtBizAreaCD2.value	= gBizArea
	frm1.txtBizAreaNM2.value	= ""
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	'//���ϱ���
	Call SetCombo(frm1.cbofileGubun, "A", "�����Ϻ�")
    Call SetCombo(frm1.cbofileGubun, "B", "������")
	Call SetCombo(frm1.cbofileGubun, "C", "�����Ϻ�+������")
    
	'//�ⱸ��
	Call SetCombo(frm1.cboGiGubun, "1", "1��")
	Call SetCombo(frm1.cboGiGubun, "2", "2��")
	'//Call SetCombo(frm1.cboGiGubun, "3", "��ü")
	
	'//�Ű���
	Call SetCombo(frm1.cboSingoGubun, "1", "����")
	Call SetCombo(frm1.cboSingoGubun, "2", "Ȯ��")
End Sub

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
		Case 0,1
			arrParam(0) = "���ݽŰ����� �˾�"					' �˾� ��Ī
			arrParam(1) = "B_TAX_BIZ_AREA"	 						' TABLE ��Ī
			arrParam(2) = strCode									' Code Condition
			arrParam(3) = ""										' Name Cindition
			arrParam(4) = ""										' Where Condition
			arrParam(5) = "���ݽŰ������ڵ�"					' �����ʵ��� �� ��Ī

			arrField(0) = "TAX_BIZ_AREA_CD"							' Field��(0)
			arrField(1) = "TAX_BIZ_AREA_NM"							' Field��(0)
    
			arrHeader(0) = "���ݽŰ������ڵ�"					' Header��(0)
			arrHeader(1) = "���ݽŰ������"					' Header��(0)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' �����
				frm1.txtBizAreaCD.focus
			Case 1		' �����
				frm1.txtBizAreaCD2.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetPopUp()  -------------------------------------------------
'	Name : SetPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' �����
				.txtBizAreaCD.focus
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)
			Case 1		' �����
				.txtBizAreaCD2.focus
				.txtBizAreaCD2.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM2.value = arrRet(1)	
		End Select
	End With
End Function

'======================================================================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=======================================================================================================
Function ClickTab1()
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar("1000000000001111")										'��: ��ư ���� ����
	Else                 
	    Call SetToolbar("1000000000001111")										'��: ��ư ���� ����
	End If
	
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ ù��° Tab 
	gSelframeFlg = TAB1
	Call SetDefaultVal()

	frm1.txtBizAreaCD.focus
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ �ι�° Tab 
	gSelframeFlg = TAB2
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolbar("1000000000001111")
	ELSE                 
		Call SetToolbar("1000000000001111")
	END IF	
	Call SetDefaultVal()
	frm1.txtBizAreaCD2.focus

End Function

'========================================================================================================= 
Sub Form_Load()

    Call InitVariables							'��: Initializes local global variables
    Call LoadInfTB19029							'��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")		'��: Lock  Suitable  Field
    Call ggoOper.FormatDate(frm1.txtYear, parent.gDateFormat, 3)
    '----------  Coding part  -------------------------------------------------------------
    Call ClickTab1()
   '// Call SetDefaultVal : ClickTab1�ȿ��� ȣ����
	Call InitComboBox()
	Call Radio3_Click
    Call SetToolbar("1000000000001111")										'��: ��ư ���� ����
   
    gIsTab     = "Y" 
	gTabMaxCnt = 2     
	'//msgbox "��ȭ���� ���� �׽�Ʈ���Դϴ�." & vbcrlf & "-- �̳���"
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt1.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt1_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt2.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt2_Change()
	frm1.cbofileGubun.value = ""
	call cbofileGubun_onChange()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReportDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtReportDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtYear_DblClick(Button)
    If Button = 1 Then
        frm1.txtYear.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYear.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportDt_Change()
    'lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt3_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt3.Action = 7
 		Call SetFocusToDocument("M")
		frm1.txtIssueDt3.Focus
   End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt3_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt4_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt4.Action = 7
 		Call SetFocusToDocument("M")
		frm1.txtIssueDt4.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt4_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt5_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt5_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt5.Action = 7
 		Call SetFocusToDocument("M")
		frm1.txtIssueDt5.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt6_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt6_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt6.Action = 7
 		Call SetFocusToDocument("M")
		frm1.txtIssueDt6.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : cbofileGubun_OnChange()
'   Event Desc : ���ϱ��м��ý� �����а�꼭�������� �����Ѵ�.
'=======================================================================================================
Sub cbofileGubun_onChange()
	Select Case Trim(frm1.cbofileGubun.value)
		Case "C"	
   		    call  ExtractDateFrom(frm1.txtIssueDT2.Text,  parent.gDateFormatYYYYMM, parent.gComDateType,strYear, strMonth, strDay)   	                

		    IF  strMonth = "06" or strMonth = "12" Then
				Call ElementVisible(frm1.txtIssueDT5, "1")
		    	call ElementVisible(frm1.txtissueDT6, "1")	            
				spndate.innerHTML = "�����й�����"
		        spnSign.innerHTML = "~"		
		    Else
				DisplayMsgBox "115116","X" , frm1.txtFileName.Alt, frm1.txtIssueDt2.Alt
				frm1.cbofileGubun.value=""
			End If
		Case Else		
			Call ElementVisible(frm1.txtIssueDT5, "0")
			call ElementVisible(frm1.txtissueDT6, "0")	            
			spndate.innerHTML = ""
		    spnSign.innerHTML = ""	
	End Select
End Sub

'===========================================================================================================
'	Event Name :Radio3_Click
'	Event Desc : ���ݰ�꼭, ��꼭���� ������ư ���ý�
'===========================================================================================================
Sub Radio3_Click()
	frm1.txtFileName.value = ""
	If gSelFrameFlg = Tab1 Then
		If frm1.Rb_TA1.checked = True Then
			Call ElementVisible(frm1.txtIssueDT5, "0")
			call ElementVisible(frm1.txtissueDT6,"0")
			Call ElementVisible(frm1.txtYear,"0")
			Call ElementVisible(frm1.cboGiGubun,"0")
			Call ElementVisible(frm1.cboSingoGubun,"0")
			Call ElementVisible(frm1.chkDari,"0")
			
			spndate.innerHTML = ""
			spnSign.innerHTML = ""
			spnYear.innerHTML = ""
			spnGiGubun.innerHTML = ""
			spnSingoGubun.innerHTML = ""
			spnDari.innerHTML = ""
			'frm1.txtFileName.className = "Required"
			frm1.txtFileName.className = "protected"
			frm1.txtFileName.readonly = false
			frm1.cbofileGubun.className="Required"
			Call ggoOper.SetReqAttr(frm1.cbofileGubun, "N")
		ElseIf frm1.Rb_TA2.checked = True Then
			Call ElementVisible(frm1.txtIssueDT5, "0")
			call ElementVisible(frm1.txtissueDT6,"0")
			Call ElementVisible(frm1.txtYear,"1")
			Call ElementVisible(frm1.cboGiGubun,"1")
			Call ElementVisible(frm1.cboSingoGubun,"1")
			Call ElementVisible(frm1.chkDari,"1")
			frm1.txtFileName.className = "protected"
			frm1.txtFileName.readonly = True
			frm1.cbofileGubun.value = "A"
			frm1.cbofileGubun.className = "protected"
			Call ggoOper.SetReqAttr(frm1.cbofileGubun, "Q")

			spndate.innerHTML = ""
			spnSign.innerHTML = ""
			spnYear.innerHTML = "�ͼӳ⵵"
			spnGiGubun.innerHTML = "�ⱸ��"
			spnSingoGubun.innerHTML = "�Ű���"
			spnDari.innerHTML = "�ϰ��븮����"
		Else
			Call ElementVisible(frm1.txtIssueDT5, "0")
			call ElementVisible(frm1.txtissueDT6,"0")
			Call ElementVisible(frm1.txtYear,"1")
			Call ElementVisible(frm1.cboGiGubun,"1")
			Call ElementVisible(frm1.cboSingoGubun,"1")
			Call ElementVisible(frm1.chkDari,"1")
			frm1.txtFileName.className = "protected"
			frm1.txtFileName.readonly = True			
			
			spndate.innerHTML = ""
			spnSign.innerHTML = ""
			spnYear.innerHTML = "�ͼӳ⵵"
			spnGiGubun.innerHTML = "�ⱸ��"
			spnSingoGubun.innerHTML = "�Ű���"
			spnDari.innerHTML = "�ϰ��븮����"

			frm1.cbofileGubun.className="Required"
			Call ggoOper.SetReqAttr(frm1.cbofileGubun, "N")
		End If
	End If	
	frm1.txtBizAreaCD.focus
	
End Sub

 '#########################################################################################################
'												4. Common Function��
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ�
'######################################################################################################### 
Function subVatDisk() 
	Dim RetFlag
	Dim strVal
	Dim intRetCD
	Dim intI, strFileName, intChrChk	'Ư������ Check
	Dim strYear1,strMonth1, strDay1, strDate1
	Dim strYear2,strMonth2, strDay2	, strDate2
	Dim strMsg
	
    '-----------------------
    'Check content area
    '-----------------------
	If gSelFrameFlg = Tab1 Then	
	'ȭ�ϸ����� ����� �� ���� Ư������ \/:*?"<>|&. ���Կ��� Ȯ��
		'2006.11.01
		'�ڵ����ι߻������.lee wol san
		
		'If frm1.Rb_TA1.checked = True Then		
		'	strFileName = frm1.txtFileName.value
		'	For intI = 1 To Len(strFileName)
		''		intChrChk = ASC(Mid(strFileName, intI, 1))
		'		If intChrChk = ASC("\") Or intChrChk = ASC("/") Or intChrChk = ASC(":") Or intChrChk = ASC("*") Or _
		'			intChrChk = ASC("?") Or intChrChk = 34 Or intChrChk = ASC("<") Or intChrChk = ASC(">") Or _
		'			intChrChk = ASC("|") OR intChrChk = ASC("&") OR intChrChk = ASC(".") Then
		'				intRetCD =  DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, frm1.txtIssueDt2.Alt)
		'				Exit Function
		'		End If
		'	Next
		' End IF	
		 
		 
		'//�Ʒ��� �ڵ带 �ּ����� ���Ƴ��� ������ �ǿ����� üũ�ؾ��� �׸��� �ٸ��⶧���� ����
		' Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
		  ' ChkField(pDoc, pStrGrp) As Boolean
		'If Not chkField(Document, "1") Then        '��: Check contents area
		'  Exit Function
		'End If
		
		'*************************************************************************
		'//�ʼ��׸� üũ : �ǿ����� üũ�ؾ��� �׸��� �ٸ��⶧���� ����
		'*************************************************************************
		If Trim(frm1.txtIssueDt1.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt1.Alt, "X") 	
			Exit Function
		End If
		If Trim(frm1.txtIssueDt2.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt2.Alt, "X") 	
			Exit Function
		End If
		
		If Trim(frm1.txtBizAreaCD.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD.Alt, "X") 	
			Exit Function
		End If

		If Trim(frm1.txtReportDt.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtReportDt.Alt, "X") 	
			Exit Function
		End If
		
		If frm1.Rb_TA1.checked = True Then		
			If Trim(frm1.txtFileName.value) = "" Then
				'RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, "X") 	
				'Exit Function
			End If
		ElseIf frm1.Rb_TA2.checked = True Then		
		    frm1.txtFileName.value = ""
			If Trim(frm1.txtYear.text) = "" Then
				RetFlag = DisplayMsgBox("970029","X" , frm1.txtYear.Alt, "X") 		'�ͼӳ⵵�� Ȯ���ϼ���
				Exit Function
			End If
		
			If Trim(frm1.cboGiGubun.value) = "" Then			
				RetFlag = DisplayMsgBox("970029","X" , frm1.cboGiGubun.Alt, "X") 	'�ⱸ���� Ȯ���ϼ���
				Exit Function
			End If
			If Trim(frm1.cboSingoGubun.value) = "" Then								'�Ű� ������ �������� �ʾ������
				RetFlag = DisplayMsgBox("970029","X" , frm1.cboSingoGubun.Alt, "X")
				Exit Function
			End If
		End If	

		If frm1.cbofileGubun.value = "C" Then
		    If Trim(frm1.txtIssueDt5.text) = "" Then
			    RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt5.Alt, "X") 	
			    Exit Function
		    End If
		    If Trim(frm1.txtIssueDt6.text) = "" Then
			    RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt6.Alt, "X") 	
			    Exit Function
		    End If
		Else
		    frm1.txtIssueDt5.text = ""
		    frm1.txtIssueDt6.text = ""
		End If
		
		If CompareDateByFormat(frm1.txtIssueDt1.text,frm1.txtIssueDt2.text,frm1.txtIssueDt1.Alt,frm1.txtIssueDt2.Alt, _
	     	               "970025",frm1.txtIssueDt1.UserDefinedFormat,parent.gComDateType, True) = False Then
		   frm1.txtIssueDt1.focus
		   Exit Function
		End If

		If CompareDateByFormat(frm1.txtIssueDt5.text,frm1.txtIssueDt6.text,frm1.txtIssueDt5.Alt,frm1.txtIssueDt6.Alt, _
	     	               "970025",frm1.txtIssueDt5.UserDefinedFormat,parent.gComDateType, True) = False Then
		   frm1.txtIssueDt1.focus
		   Exit Function
		End If

		'������ �������ڴ� ������ �������ں��� �ݵ�� �������� �̾�� ��. (2005-04-14 JYK)		
		If CompareDateByFormat(frm1.txtIssueDt6.text,frm1.txtIssueDt1.text,frm1.txtIssueDt6.Alt,frm1.txtIssueDt1.Alt, _
	     	               "970024",frm1.txtIssueDt5.UserDefinedFormat,parent.gComDateType, True) = False Then
		   frm1.txtIssueDt1.focus
		   Exit Function
		End If		
	ElseIf gSelFrameFlg = Tab2 Then	 '//�����
		If Trim(frm1.txtIssueDt3.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt3.Alt, "X") 	
			Exit Function
		End If
		If Trim(frm1.txtIssueDt4.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt4.Alt, "X") 	
			Exit Function
		End If
		If Trim(frm1.txtBizAreaCD2.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD2.Alt, "X") 	
			Exit Function
		End If
		
		If CompareDateByFormat(frm1.txtIssueDt3.text,frm1.txtIssueDt4.text,frm1.txtIssueDt3.Alt,frm1.txtIssueDt4.Alt, _
	     	               "970025",frm1.txtIssueDt3.UserDefinedFormat,parent.gComDateType, True) = False Then
		   frm1.txtIssueDt3.focus
		   Exit Function
		End If
	Else
		Exit Function
	End If
		
	RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"x","x")   '�� �ٲ�κ�
	'RetFlag = Msgbox("�۾��� ���� �Ͻðڽ��ϱ�?", vbOKOnly + vbInformation, "����")
	If RetFlag = VBNO Then
		Exit Function
	End IF

    Err.Clear                                                               '��: Protect system from crashing

	Call LayerShowHide(1)
    dim chkYn 
    

    
    With frm1
    
	if frm1.chkYN(0).checked then 
		chkYn="N"
    else
		chkYn="Y"
    end if
    
    
		If gSelFrameFlg = Tab1 Then
			If .Rb_TA1.checked = True Then
				strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
				strVal = strVal & "&txtIssueDt1=" & Trim(.txtIssueDt1.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtIssueDt2=" & Trim(.txtIssueDt2.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtReportDt=" & Trim(.txtReportDt.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'��: ó�� ���� ����Ÿ
				strVal = strVal & "&cbofileGubun=" & Trim(.cbofileGubun.value)          '��: ����,����������,�����и�
				strVal = strVal & "&txtIssueDT5=" & Trim(.txtIssueDT5.text)             '��: �������ǰ�꼭������From
				strVal = strVal & "&txtIssueDT6=" & Trim(.txtIssueDT6.text)		    	'��: �������ǰ�꼭������To	
			ElseIf .Rb_TA2.checked = True Then
				strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
				strVal = strVal & "&txtIssueDt1=" & Trim(.txtIssueDt1.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtIssueDt2=" & Trim(.txtIssueDt2.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtReportDt=" & Trim(.txtReportDt.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtYear=" & Trim(.txtYear.text)						'��: ó�� ���� ����Ÿ
				strVal = strVal & "&cboGiGubun=" & Trim(.cboGiGubun.value)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&cboSingoGubun=" & Trim(.cboSingoGubun.value)		'��: ó�� ���� ����Ÿ
				If .chkDari.checked = True Then
					strVal = strVal & "&chkDaeri=" & "Y"								'��: ó�� ���� ����Ÿ
				Else
					strVal = strVal & "&chkDaeri=" & "N"								'��: ó�� ���� ����Ÿ
				End If
				strVal = strVal & "&rdoGubun=" & "1"									'��: ó�� ���� ����Ÿ
			ElseIf 	.Rb_TA7.checked = True Then
				strVal = BIZ_PGM_ID4 & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
				strVal = strVal & "&txtIssueDt1=" & Trim(.txtIssueDt1.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtIssueDt2=" & Trim(.txtIssueDt2.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtReportDt=" & Trim(.txtReportDt.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtYear=" & Trim(.txtYear.text)						'��: ó�� ���� ����Ÿ
				strVal = strVal & "&cboGiGubun=" & Trim(.cboGiGubun.value)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&cboSingoGubun=" & Trim(.cboSingoGubun.value)		'��: ó�� ���� ����Ÿ								
				strVal = strVal & "&cbofileGubun=" & Trim(.cbofileGubun.value)          '��: ����,����������,�����и�
				strVal = strVal & "&txtIssueDT5=" & Trim(.txtIssueDT5.text)             '��: �������ǰ�꼭������From
				strVal = strVal & "&txtIssueDT6=" & Trim(.txtIssueDT6.text)		    	'��: �������ǰ�꼭������To
				If .chkDari.checked = True Then
					strVal = strVal & "&chkDaeri=" & "Y"								'��: ó�� ���� ����Ÿ
				Else
					strVal = strVal & "&chkDaeri=" & "N"								'��: ó�� ���� ����Ÿ
				End If										
			End If
		Else
				strVal = BIZ_PGM_ID3 & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
				strVal = strVal & "&txtIssueDt3=" & Trim(.txtIssueDt3.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtIssueDt4=" & Trim(.txtIssueDt4.text)				'��: ó�� ���� ����Ÿ
				strVal = strVal & "&txtBizAreaCD2=" & UCase(Trim(.txtBizAreaCD2.value))	'��: ó�� ���� ����Ÿ
				If frm1.Rb_TA3.checked = True Then
					strVal = strVal & "&rdoGubun=" & "3"								'��: ó�� ���� ����Ÿ
				ElseIf frm1.Rb_TA4.checked = True Then
					strVal = strVal & "&rdoGubun=" & "4"								'��: ó�� ���� ����Ÿ
				ElseIf 	frm1.Rb_TA8.checked = True Then
					strVal = strVal & "&rdoGubun=" & "6"								'��: ó�� ���� ����Ÿ				
				End If	
				If frm1.Rb_TA5.checked = True Then
					strVal = strVal & "&rdofileGubun=" & "A"							'��: ����
				ElseIf frm1.Rb_TA6.checked = True Then
					strVal = strVal & "&rdofileGubun=" & "B"							'��: ������
				End If	
		End If	

		strVal = strVal & "&chkYn=" & chkYn

		Call RunMyBizASP(MyBizASP, strVal)												'��: �����Ͻ� ASP �� ����
	End With
    
End Function

Function subVatDiskOK(ByVal pFileName) 
	Dim strVal
    Err.Clear																			'��: Protect system from crashing

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002								'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtFileName=" & pFileName										'��: ��ȸ ���� ����Ÿ
	Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ����
End Function

Function subVatDiskOK2(ByVal strVal) 
    Err.Clear
	On Error Resume Next
	Dim IntRetCD

	If strVal = "OK" Then
		IntRetCD = DisplayMsgBox("183114", "X", "X", "X")
	End If
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, True)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ����
'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ����
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű�
'========================================================================================

Function DbQueryOk()							'��: ��ȸ ������ �������
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ���
'========================================================================================

Function DbSave() 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű�
'========================================================================================

Function DbSaveOk()			'��: ���� ������ ���� ����
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag��
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ΰ������ϻ���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ΰ����������</font></td>
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
						<!--ù��° TAB  -->
						<DIV ID="TabDiv"  SCROLL="no">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ݰ�꼭����</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio3 ID=Rb_TA1 Checked onclick="Radio3_Click()" value="0"><LABEL FOR=Rb_TA1>���ݰ�꼭</LABEL>&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio3 ID=Rb_TA2 onclick="Radio3_Click()" value="1"><LABEL FOR=Rb_TA2>��꼭</LABEL>&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio3 ID=Rb_TA7 onclick="Radio3_Click()" value="2"><LABEL FOR=Rb_TA7>�ſ�ī��</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��꼭������</TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��꼭������(From)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
												  &nbsp; ~ &nbsp;
												  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt2 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��꼭������(To)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								
						
								<TR>
									<TD CLASS=TD5 NOWRAP>���հ�������</TD>
									<TD CLASS=TD6>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="chkYN" TAG="1X" VALUE="N" CHECKED ID="chkYN0"><LABEL FOR="chkYN0">����庰</LABEL>&nbsp;
				        	        <INPUT TYPE="RADIO" CLASS="RADIO" NAME="chkYN" TAG="1X" VALUE="Y"  ID="chkYN1"><LABEL FOR="chkYN1">����</LABEL>
				
									 </TD>
								</TR>
								
								<TR>
									<TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="���ݽŰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" tag="14X" ALT="���ݽŰ�����"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�Ű�����</TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtReportDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�Ű�����" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ȭ�ϸ�</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtFileName" NAME="txtFileName" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="11X" ALT="ȭ�ϸ�"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ϱ���</TD>
									<TD CLASS="TD6"><SELECT ID="cbofileGubun" NAME="cbofileGubun" ALT="���ϱ���" STYLE="WIDTH: 130px" tag="12X"><OPTION VALUE="" selected></OPTION></SELECT></TD>									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnDate">�����й�����</span></TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt5 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�����й�����(From)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
												  &nbsp; <span id="spnSign">~</span> &nbsp;
												  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt6 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�����й�����(To)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnYear">�ͼӳ⵵</span></TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtYear CLASS=FPDTYYYY title=FPDATETIME ALT="�ͼӳ⵵" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnGiGubun">�ⱸ��</span></TD>
									<TD CLASS="TD6"><SELECT ID="cboGiGubun" NAME="cboGiGubun" ALT="�ⱸ��" STYLE="WIDTH: 98px" tag="12X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnSingoGubun">�Ű���</span></TD>
									<TD CLASS="TD6"><SELECT ID="cboSingoGubun" NAME="cboSingoGubun" ALT="�Ű���" STYLE="WIDTH: 98px" tag="12X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnDari">�ϰ��븮����</span></TD>
									<TD CLASS="TD6"><input type="checkbox" class = "check" name="chkDari" value="Y"></TD>
								</TR>																		
								<TR>
									<TD CLASS=TD5></TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
							</TABLE>
						</div>
						<!--�ι�° TAB  -->
						<DIV ID="TabDiv"  SCROLL="no">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ݰ�꼭����</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio4 ID=Rb_TA3 Checked  value="0"><LABEL FOR=Rb_TA3>���ݰ�꼭</LABEL>&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio4 ID=Rb_TA4  value="1"><LABEL FOR=Rb_TA4>��꼭</LABEL>&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio4 ID=Rb_TA8  value="2"><LABEL FOR=Rb_TA8>�ſ�ī��</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ϱ���</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio5 ID=Rb_TA5 Checked  value="A"><LABEL FOR=Rb_TA5>����</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio5 ID=Rb_TA6  value="B"><LABEL FOR=Rb_TA6>������</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��꼭������</TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt3 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��꼭������(From)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
												  &nbsp; ~ &nbsp;
												  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt4 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��꼭������(To)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtBizAreaCD2" NAME="txtBizAreaCD2" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="121XXU" ALT="���ݽŰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD2.Value, 1)">&nbsp;<INPUT TYPE=TEXT ID="txtBizAreaNM2" NAME="txtBizAreaNM2" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" tag="14X" ALT="���ݽŰ�����"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>	
							</TABLE>
						</div>
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
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExecute" CLASS="CLSMBTN" OnClick="VBScript:Call subVatDisk()" Flag=1>�� ��</BUTTON>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>

