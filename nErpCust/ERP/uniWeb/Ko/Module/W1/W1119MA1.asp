<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �� ���� ���� 
'*  3. Program ID           : W1119MA1
'*  4. Program Name         : W1119MA1.asp
'*  5. Program Desc         : ��3ȣ 
'*  6. Modified date(First) : 2005/01/07
'*  7. Modified date(Last)  : 2005/
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'*  ������ ���� --  �Ϲݹ��� : ǥ�ؼ��Ͱ�꼭(��)  ��������(���ս�) <>  06 or 35	����	WC0025	���Ͱ�꼭�� ��������(���ս�) �ݾװ� �����׿���(��ձ�)ó�� ��꼭�� ��������(���ս�) �ݾ��� ��ġ���� �ʽ��ϴ�.
'					�������� : ǥ�ؼ��Ͱ�꼭(��)  ��������(���ս�) <>  06 or 35	����	WC0025	���Ͱ�꼭�� ��������(���ս�) �ݾװ� �����׿���(��ձ�)ó�� ��꼭�� ��������(���ս�) �ݾ��� ��ġ���� �ʽ��� 
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

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W1119MA1"
Const BIZ_PGM_ID		= "W1119MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "W1119MB2.asp"
Const EBR_RPT_ID		= "W1119OA1"


Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

'============================================  �ʱ�ȭ �Լ�  ====================================

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgRefMode = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �Ű��� �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
    
End Sub



'============================== ���۷��� �Լ�  ========================================
'����ڰ� 1,�����׿���ó�а�꼭�� 2��ձ�ó����꼭�� 			
'�������� �ƴ��ϰ� �ݾ׺ҷ����⸦ �����ϴ� ���	����(W10001 : ��꼭 ������ �������� �ƴ��Ͽ����ϴ�. ���� ��꼭 ������ �����Ͽ� �ֽʽÿ�)
'
'02		��������ǥ�� �̿������׿���(�Ǵ� �̿���ձ�)�� �����ܾ��� �Է���.		
'06		ǥ�ؼ��Ͱ�꼭�� ��������(���ս�)�� �Է���.		
'
'31		��������ǥ�� �̿������׿���(�Ǵ� �̿���ձ�)�� �����ܾ��� �Է���.		
'35		ǥ�ؼ��Ͱ�꼭�� ��������(���ս�) �� (-1)�� �Է���.		

Function GetRef()	' �ݾ׺ҷ����� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim sMesg
	DIm BackColor_w,BackColor_g

	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
    If NOT Frm1.chkW_TYPE1.checked And NOT Frm1.chkW_TYPE2.checked Then
        Call DisplayMsgBox("W10003", "X", "X", "X")
	    Frm1.chkW_TYPE1.focus
	    Exit Function
    End If

	sMesg = wgRefDoc & vbCrLf & vbCrLf
	
	' ȭ��� ������ ����Ÿ�� ������ ǥ���Ѵ�.
	If Frm1.chkW_TYPE1.checked Then
		BackColor_g = frm1.txtW1.BackColor
		BackColor_w = frm1.txtW2.BackColor
		Frm1.txtW2.BackColor = &H009BF0A2&
		Frm1.txtW6.BackColor = &H009BF0A2&
	Else
		BackColor_g = frm1.txtW35.BackColor
		BackColor_w = frm1.txtW31.BackColor
		Frm1.txtW31.BackColor = &H009BF0A2&
		Frm1.txtW35.BackColor = &H009BF0A2&
	End If
	
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"

	If Frm1.chkW_TYPE1.checked Then
		Frm1.txtW2.BackColor = BackColor_w
		Frm1.txtW6.BackColor = BackColor_g
	Else
		Frm1.txtW31.BackColor = BackColor_w
		Frm1.txtW35.BackColor = BackColor_g
	End If

	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtCO_CD="      	 & Frm1.txtCO_CD.Value	      '��: Query Key        
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
    	
End Function

Sub GetRefOK()
	If Frm1.chkW_TYPE1.checked Then
		Frm1.txtW2.Value = Frm1.txtRW1.Value
		Frm1.txtW6.Value = Frm1.txtRW2.Value
	Else
		Frm1.txtW31.Value = Frm1.txtRW1.Value
		Frm1.txtW35.Value = unicdbl(Frm1.txtRW2.Value) * -1
	End If
	Call SetAllTxtChkCalc
End Sub


'============================================  ��ȸ���� �Լ�  ====================================
Sub CheckFISC_DATE()	' ��û������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, sFISC_START_DT, sFISC_END_DT, datMonCnt, i, datNow
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		sFISC_START_DT = CDate(lgF0)
	Else
		sFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		sFISC_END_DT = CDate(lgF1)
	Else
		sFISC_END_DT = ""
	End if
	
End Sub

'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100111100101011")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call AppendNumberRange("0","","")
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData

    call fncquery()
    
End Sub


'============================================  �̺�Ʈ �Լ�  ====================================

<%
'==========================================================================================
'   Event Name : chkW_TYPE ...
'   Event Desc : üũ�ڽ� Value Change
'==========================================================================================
%>
Sub chkW_TYPE1_OnClick()
	If Frm1.chkW_TYPE1.checked Then
		Call SetW_TYPE(1)
	Else
		Frm1.chkW_TYPE1.checked = True
	End If
End Sub

Sub chkW_TYPE2_OnClick()	
	If Frm1.chkW_TYPE2.checked Then
		Call SetW_TYPE(2)
	Else
		Frm1.chkW_TYPE2.checked = True
	End If
End Sub

Sub SetW_TYPE(ByVal chkNum)
	If chkNum = 1 Then
		With Frm1
			.chkW_TYPE1.checked = True
			.chkW_TYPE2.checked = False
			.txtW30.Value = 0 :	call ggoOper.SetReqAttr(.txtW30, "Q")
			.txtW31.Value = 0 :	call ggoOper.SetReqAttr(.txtW31, "Q")
			.txtW32.Value = 0 :	call ggoOper.SetReqAttr(.txtW32, "Q")
			.txtW33.Value = 0 :	call ggoOper.SetReqAttr(.txtW33, "Q")
			.txtW34.Value = 0 :	call ggoOper.SetReqAttr(.txtW34, "Q")
			.txtW35.Value = 0 :	call ggoOper.SetReqAttr(.txtW35, "Q")
			.txtW40.Value = 0 :	call ggoOper.SetReqAttr(.txtW40, "Q")
			.txtW41.Value = 0 :	call ggoOper.SetReqAttr(.txtW41, "Q")
			.txtW42.Value = 0 :	call ggoOper.SetReqAttr(.txtW42, "Q")
			.txtW43.Value = 0 :	call ggoOper.SetReqAttr(.txtW43, "Q")
			.txtW44.Value = 0 :	call ggoOper.SetReqAttr(.txtW44, "Q")
			.txtW50.Value = 0 :	call ggoOper.SetReqAttr(.txtW50, "Q")

'			call ggoOper.SetReqAttr(.txtW1, "D")
			call ggoOper.SetReqAttr(.txtW2, "D")
			call ggoOper.SetReqAttr(.txtW3, "D")
			call ggoOper.SetReqAttr(.txtW4, "D")
			call ggoOper.SetReqAttr(.txtW5, "D")
'			call ggoOper.SetReqAttr(.txtW6, "D")
			call ggoOper.SetReqAttr(.txtW8, "D")
'			call ggoOper.SetReqAttr(.txtW10, "D")
'			call ggoOper.SetReqAttr(.txtW11, "D")
			call ggoOper.SetReqAttr(.txtW12, "D")
			call ggoOper.SetReqAttr(.txtW13, "D")
			call ggoOper.SetReqAttr(.txtW14, "D")
'			call ggoOper.SetReqAttr(.txtW15, "D")
			call ggoOper.SetReqAttr(.txtW16, "D")
			call ggoOper.SetReqAttr(.txtW17, "D")
			call ggoOper.SetReqAttr(.txtW18, "D")
			call ggoOper.SetReqAttr(.txtW19, "D")
			call ggoOper.SetReqAttr(.txtW20, "D")
'			call ggoOper.SetReqAttr(.txtW25, "D")

			call ggoOper.SetReqAttr(.txtW26, "D")
			call ggoOper.SetReqAttr(.txtW27, "D")
			call ggoOper.SetReqAttr(.txtW28, "D")

			.txtW2.Value = .txtRW1.Value
			.txtW6.Value = .txtRW2.Value
		End With
	Else
		With Frm1
			.chkW_TYPE1.checked = False
			.chkW_TYPE2.checked = True
			.txtW1.Value = 0 :	call ggoOper.SetReqAttr(.txtW1, "Q")
			.txtW2.Value = 0 :	call ggoOper.SetReqAttr(.txtW2, "Q")
			.txtW3.Value = 0 :	call ggoOper.SetReqAttr(.txtW3, "Q")
			.txtW4.Value = 0 :	call ggoOper.SetReqAttr(.txtW4, "Q")
			.txtW5.Value = 0 :	call ggoOper.SetReqAttr(.txtW5, "Q")
			.txtW6.Value = 0 :	call ggoOper.SetReqAttr(.txtW6, "Q")
			.txtW8.Value = 0 :	call ggoOper.SetReqAttr(.txtW8, "Q")
			.txtW10.Value = 0 :	call ggoOper.SetReqAttr(.txtW10, "Q")
			.txtW11.Value = 0 :	call ggoOper.SetReqAttr(.txtW11, "Q")
			.txtW12.Value = 0 :	call ggoOper.SetReqAttr(.txtW12, "Q")
			.txtW13.Value = 0 :	call ggoOper.SetReqAttr(.txtW13, "Q")
			.txtW14.Value = 0 :	call ggoOper.SetReqAttr(.txtW14, "Q")
			.txtW15.Value = 0 :	call ggoOper.SetReqAttr(.txtW15, "Q")
			.txtW16.Value = 0 :	call ggoOper.SetReqAttr(.txtW16, "Q")
			.txtW17.Value = 0 :	call ggoOper.SetReqAttr(.txtW17, "Q")
			.txtW18.Value = 0 :	call ggoOper.SetReqAttr(.txtW18, "Q")
			.txtW19.Value = 0 :	call ggoOper.SetReqAttr(.txtW19, "Q")
			.txtW20.Value = 0 :	call ggoOper.SetReqAttr(.txtW20, "Q")
			.txtW25.Value = 0 :	call ggoOper.SetReqAttr(.txtW25, "Q")
			
			.txtW26.Value = 0 :	call ggoOper.SetReqAttr(.txtW26, "Q")
			.txtW27.Value = 0 :	call ggoOper.SetReqAttr(.txtW27, "Q")
			.txtW28.Value = 0 :	call ggoOper.SetReqAttr(.txtW28, "Q")

'			call ggoOper.SetReqAttr(.txtW30, "D")
			call ggoOper.SetReqAttr(.txtW31, "D")
			call ggoOper.SetReqAttr(.txtW32, "D")
			call ggoOper.SetReqAttr(.txtW33, "D")
			call ggoOper.SetReqAttr(.txtW34, "D")
'			call ggoOper.SetReqAttr(.txtW35, "D")
'			call ggoOper.SetReqAttr(.txtW40, "D")
			call ggoOper.SetReqAttr(.txtW41, "D")
			call ggoOper.SetReqAttr(.txtW42, "D")
			call ggoOper.SetReqAttr(.txtW43, "D")
			call ggoOper.SetReqAttr(.txtW44, "D")
'			call ggoOper.SetReqAttr(.txtW50, "D")
			.txtW31.Value = .txtRW1.Value
			.txtW35.Value = unicdbl(.txtRW2.Value) * -1
		End With
	End If
    Call SetAllTxtChkCalc()
End Sub

Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub txtW2_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw3_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw4_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw5_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw8_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw12_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw13_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw14_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw16_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw17_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw18_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw19_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw20_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw31_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw32_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw33_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw34_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw41_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw42_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw43_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw44_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub


Sub SetAllTxtChkCalc()
	'10 < 0  &  15 = 0	����	W10003	ó���� �����׿��ݰ� ���������� ���Ծ��� �հ�ݾ��� 0���� �����ϴ�. ��ձ�ó����꼭�� �ۼ��Ͽ� �ֽʽÿ� 
	'30 < 0	����	W10002	ó������ձ��� �ݾ��� 0���� �����ϴ�. �����׿���ó�а�꼭�� �ۼ��Ͽ� �ֽʽÿ� 
	'50 < 0	����	W10002	�����̿���ձ��� �ݾ��� 0���� �����ϴ�. �����׿���ó�а�꼭�� �ۼ��Ͽ� �ֽʽÿ� 

	
	'01		02+03+04-05+06 �� ����Ͽ� �Է���.		=> 02+03+04+05+06 �� ����Ͽ� �Է��� 
	'10		01 + 08 �� ����Ͽ� �Է���.	
	'11		12+13+14+15+18+19+20 �� ����Ͽ� �Է���.	(+26+27+28 �߰� 2006.03����)	
	'15		16+17 �� ����Ͽ� �Է���.	
	'25		10 - 11 �� ����Ͽ� �Է���.	
	'30		(31+32+33-34-35) �� (-1) �� ����Ͽ� �Է���.		
	'35		ǥ�ؼ��Ͱ�꼭�� ��������(���ս�) �� (-1)�� �Է���.	�ҷ����� 
	'40		41+42+43+44 �� ����Ͽ� �Է���.		
	'50		30 - 40  �� ����Ͽ� �Է���.	

    Frm1.txtW1.value = unicdbl(Frm1.txtw2.value) + unicdbl(Frm1.txtw3.value) + unicdbl(Frm1.txtw4.value) - unicdbl(Frm1.txtw5.value) + unicdbl(Frm1.txtw6.value)

    Frm1.txtW10.value = unicdbl(Frm1.txtw1.value) + unicdbl(Frm1.txtw8.value)
	
    Frm1.txtW15.value = unicdbl(Frm1.txtw16.value) + unicdbl(Frm1.txtw17.value)
	
    Frm1.txtW11.value = unicdbl(Frm1.txtw12.value) + unicdbl(Frm1.txtw13.value) + unicdbl(Frm1.txtw14.value) + unicdbl(Frm1.txtw15.value) + unicdbl(Frm1.txtw18.value) + unicdbl(Frm1.txtw19.value) + unicdbl(Frm1.txtw20.value) + unicdbl(Frm1.txtw26.value) + unicdbl(Frm1.txtw27.value) + unicdbl(Frm1.txtw28.value)

    Frm1.txtW25.value = unicdbl(Frm1.txtw10.value) - unicdbl(Frm1.txtw11.value)
	
    Frm1.txtW30.value = (unicdbl(Frm1.txtw31.value) + unicdbl(Frm1.txtw32.value) + unicdbl(Frm1.txtw33.value) - unicdbl(Frm1.txtw34.value) - unicdbl(Frm1.txtw35.value)) * -1

    Frm1.txtW40.value = unicdbl(Frm1.txtw41.value) + unicdbl(Frm1.txtw42.value) + unicdbl(Frm1.txtw43.value) + unicdbl(Frm1.txtw44.value)

    Frm1.txtW50.value = unicdbl(Frm1.txtW30.value) - unicdbl(Frm1.txtW40.value)

End Sub

'==========================================================================================
Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	Call CheckFISC_DATE
End Sub


'============================================  �������� �Լ�  ====================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

'    Call DbQuery
    FncQuery = True
End Function

Function FncSave() 
	Dim IntRetCD
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
	    Exit Function
	End If

	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "A") Then                             '��: Check contents area
	   Exit Function
	End If

    If NOT Frm1.chkW_TYPE1.checked And NOT Frm1.chkW_TYPE2.checked Then
        Call DisplayMsgBox("W10001", "X", "X", "X")
	    Frm1.chkW_TYPE1.focus
	    Exit Function
    End If
    
	If Verification = False Then Exit Function

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function


' ---------------------- ���ĳ� ���� -------------------------
Function  Verification()
	
	Verification = False

    If unicdbl(Frm1.txtw10.value) < 0 And unicdbl(Frm1.txtw15.value) = 0 Then
        Call DisplayMsgBox("W10005", "X", "X", "X")
	    Frm1.chkW_TYPE2.focus
	    Exit Function
    End If
    
    If unicdbl(Frm1.txtw30.value) < 0 Then
        Call DisplayMsgBox("W10004", "X", "ó������ձ�", "X")
	    Frm1.chkW_TYPE1.focus
	    Exit Function
    End If

    If unicdbl(Frm1.txtw50.value) < 0 Then
        Call DisplayMsgBox("W10004", "X", "�����̿���ձ�", "X")
	    Frm1.chkW_TYPE1.focus
	    Exit Function
    End If

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
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables

    Call SetToolbar("1100100000001011")
    lgIntFlgMode = parent.OPMD_CMODE

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
	exit Function
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

	
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
     On Error Resume Next
End Function

Function FncInsertRow(ByVal pvRowCnt) 
     On Error Resume Next
End Function

Function FncDeleteRow() 
     On Error Resume Next
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
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

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
	    strVal = strVal 	& "&txtCo_Cd=" 			 & Trim(frm1.txtCo_Cd.value)				'��: ��ȸ ���� ����Ÿ 
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key   
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Call SetToolbar("1101100000010111")
    lgBlnFlgChgValue = False
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

    lgIntFlgMode = parent.OPMD_UMODE
	Call SetW_TYPE(Frm1.txtW_TYPE.Value)

    		
End Function

Function DbQueryFalse()													<%'��ȸ ������ ������� %>
	
    Call SetToolbar("1101100000010111")
    Frm1.chkW_TYPE1.checked = true
    Call chkW_TYPE1_OnClick()
    Call InitVariables   
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
     Err.Clear
	DbSave = False

    Dim strVal

    Call LayerShowHide(1) 

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode
		
		If .chkW_TYPE1.checked Then
			.txtW_TYPE.Value = "1"
		Else
			.txtW_TYPE.Value = "2"
		End If

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True                                                         
End Function


Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
    lgBlnFlgChgValue = False
    'FncQuery
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal 	& "&txtCo_Cd=" 			 & Trim(frm1.txtCo_Cd.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function

Sub txtW_DT_DblClick(Button)
    If Button = 1 Then
       frm1.txtW_DT.Action = 7                                    ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("M")	
       frm1.txtW_DT.Focus
    End If
End Sub


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">�ݾ׺ҷ�����</A>
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
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="�������" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
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
						<DIV ID="ViewDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=1>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP>
                                   <FIELDSET CLASS="CLSFLD">
                                   <BR>
                                   <TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   <TR>
									       <TD CLASS="TD51" width="50%" ALIGN=CENTER Colspan="3"><LABEL FOR="chkBpTypeC">1. �����׿���ó�а�꼭</LABEL>
									       		<INPUT TYPE=CHECKBOX NAME="chkW_TYPE1" ID="chkW_TYPE1" tag="21" Class="Check">
									       </TD>
									       <TD CLASS="TD51" width="50%" ALIGN=CENTER Colspan="3"><LABEL FOR="chkBpTypeC">2. ��ձ�ó����꼭</LABEL>
									       		<INPUT TYPE=CHECKBOX NAME="chkW_TYPE2" ID="chkW_TYPE2" tag="21" Class="Check">
								           </TD>
									  </TR>
									   <TR>
									       <TD CLASS="TD51" width="25%" ALIGN=CENTER>����</TD>
									       <TD CLASS="TD51" width="7%" ALIGN=CENTER>�ڵ�</TD>
									       <TD CLASS="TD51" width="18%" ALIGN=CENTER>�ݾ�</TD>
									       <TD CLASS="TD51" width="25%" ALIGN=CENTER>����</TD>
									       <TD CLASS="TD51" width="7%" ALIGN=CENTER>�ڵ�</TD>
									       <TD CLASS="TD51" width="18%" ALIGN=CENTER>�ݾ�</TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT><B>I. ó���������׿���</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>01</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT><B>I. ó������ձ�</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>30</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW30" name=txtW30 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;1. �����̿������׿���(�Ǵ�<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����̿���ձ�)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>02</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2" name=txtW2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;1. �����̿������׿���(�Ǵ�<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����̿���ձ�)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>31</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW31" name=txtW31 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;2. ȸ�躯���� ����ȿ��</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>03</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3" name=txtW3 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;2. ȸ�躯���� ����ȿ��</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>32</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW32" name=txtW32 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;3. ���������������(�Ǵ�<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������������ս�)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>04</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;3. ���������������(�Ǵ�<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������������ս�)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>33</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW33" name=txtW33 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;4. �߰�����</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>05</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;4. �߰�����</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>34</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW34" name=txtW34 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;5. ��������(�Ǵ�<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����ս�)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>06</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;5. ��������(�Ǵ�<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����ս�)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>35</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW35" name=txtW35 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT><B>II. ���������� ���� ���Ծ�</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>08</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT><B>II. ��ձ� ó����</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>40</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW40" name=txtW40 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=CENTER>��     ��</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>10</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW10" name=txtW10 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;1. �������������Ծ�</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>41</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW41" name=txtW41 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT><B>III. �����׿��� ó�о�</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>11</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW11" name=txtW11 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;2. ��Ÿ�������������Ծ�</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>42</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW42" name=txtW42 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;1. �����غ��</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>12</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12" name=txtW12 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;3. �����غ�����Ծ�</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>43</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW43" name=txtW43 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;2. ��Ÿ����������</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>13</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW13" name=txtW13 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;4. �ں��׿������Ծ�</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>44</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW44" name=txtW44 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;3. �ֽ����ι������ݻ󰢾�</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>14</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW14" name=txtW14 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT><B>III. �����̿���ձ�</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>50</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW50" name=txtW50 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;4. ����</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>15</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW15" name=txtW15 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD61" colspan="3"></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;&nbsp;&nbsp;��. ���ݹ��</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>16</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW16" name=txtW16 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT><B>ó��(ó��) Ȯ����</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER></TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtW_DT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="ó��(ó��) Ȯ����" tag="22X1" id=txtW_DT style="width: 100%"></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;&nbsp;&nbsp;��. �ֽĹ��</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>17</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW17" name=txtW17 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD61" colspan="3" rowspan="8"></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;5. ����ó�п� ���� �󿩱�</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>26</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW26" name=txtW26 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;6. ���Ȯ��������</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>18</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW18" name=txtW18 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;7. ��ä������</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>19</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW19" name=txtW19 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;8. ��Ÿ������</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>20</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW20" name=txtW20 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;9. ����Ư�����ѹ� �� �غ�� �� ������</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>27</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW27" name=txtW27 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;10. ��Ÿ �׿���ó�о�</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>28</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW28" name=txtW28 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT><B>IV. �����̿������׿���</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>25</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW25" name=txtW25 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
								  </table>
								  </FIELDSET>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtW_TYPE" tag="24">
<INPUT TYPE=hidden NAME="txtRW1" tag="24" tabindex="-1" value="0">
<INPUT TYPE=hidden NAME="txtRW2" tag="24" tabindex="-1" value="0">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
