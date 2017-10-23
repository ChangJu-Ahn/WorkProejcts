<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ÿ���� 
'*  3. Program ID           : w9101mA1
'*  4. Program Name         : w9101mA1.asp
'*  5. Program Desc         : ��47ȣ �ֿ��������(��)
'*  6. Modified date(First) : 2005/02/23
'*  7. Modified date(Last)  : 2005/02/23
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : 
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

Const BIZ_MNU_ID		= "w9101mA1"
Const BIZ_PGM_ID		= "w9101mB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "w9101mB2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "w9101OA1"


' -- �׸��� �÷� ���� 
Dim C_W1
Dim C_W1_NM1
Dim C_W1_NM2
Dim C_W2
Dim C_W2_CD
Dim C_W3
Dim C_W4
Dim C_W5

Dim C_01
Dim C_02
Dim C_03
Dim C_04
Dim C_05
Dim C_06
Dim C_07
Dim C_08
Dim C_09
Dim C_10
Dim C_11
Dim C_12
Dim C_13
Dim C_14
Dim C_15
Dim C_16
Dim C_17
Dim C_18
Dim C_19
Dim C_20
Dim C_21
'Dim C_22
'Dim C_23

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgFISC_START_DT, lgFISC_END_DT 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W1		= 1
	C_W1_NM1	= 2
	C_W1_NM2	= 3
	C_W2		= 4
	C_W2_CD		= 5
	C_W3		= 6
	C_W4		= 7
	C_W5		= 8
	
	C_01 = 1
	C_02 = 2
	C_03 = 3
	C_04 = 4
	C_05 = 5
	C_06 = 6
	C_07 = 7
	C_08 = 8
	C_09 = 9
	C_10 = 10
	C_11 = 11
	C_12 = 12
	C_13 = 13
	C_14 = 14
	C_15 = 15
	C_16 = 16
	C_17 = 17
	C_18 = 18
	C_19 = 19
	C_20 = 20
	C_21 = 21
	'C_22 = 22
	'C_23 = 23
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
    lgRefMode = False

End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	' ��ȸ����(����)
	Dim IntRetCD1
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
	
End Sub


Sub InitSpreadComboBox()
    Dim IntRetCD1

End Sub


Sub InitSpreadSheet()
	Dim ret, iRow
	
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")
 	
	' 1�� �׸��� 

	With Frm1.vspdData
				
		ggoSpread.Source = Frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20041222_0" ,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W5 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_W1,		"(1)����",			5,,,10,1
		ggoSpread.SSSetEdit		C_W1_NM1,	"(1)����",			11,2,,50,1
		ggoSpread.SSSetEdit		C_W1_NM2,	"(1)����",			22,2,,50,1
		ggoSpread.SSSetEdit  	C_W2,		"(2)�ٰŹ� ����"		, 20,,,100,1	' 
		ggoSpread.SSSetEdit		C_W2_CD,	"�ڵ�",			5,2,,10,1
	    ggoSpread.SSSetFloat	C_W3,		"(3)ȸ����ݾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	    ggoSpread.SSSetFloat	C_W4,		"(4)���������ݾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
		ggoSpread.SSSetFloat	C_W5,		"(5)�������ݾ�" & vbCrLf & "((3)-(4))",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 

	    ret = .AddCellSpan(C_W1_NM1, 0, 2, 1)
		.rowheight(0) = 20	' ���� ������ 

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W1, C_W1, True)

		Call FncNew()
		Call SetSpreadLock()

		.ReDraw = true	
				
	End With 

 
	Call InitSpreadComboBox
	
					
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call GetFISC_DATE
	
	'Exit Sub
		
End Sub


Sub SetSpreadLock()

	With Frm1.vspdData
	
		ggoSpread.Source = Frm1.vspdData

		ggoSpread.SpreadLock C_W1,   -1, C_W2_CD
		ggoSpread.SpreadUnLock C_W3, -1, C_W5	' ��ü ���� 

		ggoSpread.SpreadLock C_W3,   C_13, C_W4, C_13
		ggoSpread.SpreadLock C_W3,   C_16, C_W4, C_16
		ggoSpread.SpreadLock C_W3,   C_20, C_W4, C_21
		'ggoSpread.SpreadLock C_W3,   C_23, C_W4, C_23

		ggoSpread.SpreadLock C_W5,   C_01, C_W5, C_12
		ggoSpread.SpreadLock C_W5,   C_14, C_W5, C_15
		ggoSpread.SpreadLock C_W5,   C_17, C_W5, C_19
		'ggoSpread.SpreadLock C_W5,   C_22, C_W5, C_22

	End With	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow

	With Frm1.vspdData
		ggoSpread.Source = Frm1.vspdData
		For iRow = pvStartRow To pvEndRow
			.Col = C_SEQ_NO
			.Row = iRow
			If .Text = 999999 Then
				ggoSpread.SpreadLock C_W1,   iRow, C_W5, iRow
			Else
				ggoSpread.SpreadUnLock C_W1, iRow, C_W_DESC, iRow	' ��ü ���� 
				ggoSpread.SSSetRequired C_W1, iRow, iRow
				ggoSpread.SSSetRequired C_W1, iRow, iRow
				ggoSpread.SSSetRequired C_W1_NM, iRow, iRow
				ggoSpread.SpreadLock C_W5,   iRow, C_W5
			End If
		Next
			
	End With	
End Sub

Sub SetSpreadTotalLine()
	Dim ret
		
	ggoSpread.Source = Frm1.vspdData
	With Frm1.vspdData
		If .MaxRows > 0 Then
			.Row = .MaxRows
			ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' ���� 2�� ��ħ 
			.Col	= C_W1	:	.CellType = 1	:	.Text	= "��"	:	.TypeHAlign = 2
			SetSpreadColor 1, .MaxRows

		End If
	End With
End Sub 

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W1		= iCurColumnPos(2)
            C_W1_BT		= iCurColumnPos(3)
            C_W1_NM		= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W3		= iCurColumnPos(6)
            C_W4		= iCurColumnPos(7)
            C_W5		= iCurColumnPos(8)
            C_W_DESC	= iCurColumnPos(9)
    End Select    
End Sub

Sub InitSpreadRow()
	Dim ret, iRow

	With Frm1.vspdData
		
		ggoSpread.Source = Frm1.vspdData
		
		'patch version
		If .MaxRows = 0 Then	.MaxRows = C_21

	    ret = .AddCellSpan(C_W1_NM1, C_01, 1, 5)
	    ret = .AddCellSpan(C_W1_NM1, C_06, 1, 4)
	    ret = .AddCellSpan(C_W1_NM1, C_11, 1, 5)
	    ret = .AddCellSpan(C_W1_NM1, C_16, 1, 3)
	    ret = .AddCellSpan(C_W1_NM1, C_19, 2, 1)
	    ret = .AddCellSpan(C_W1_NM1, C_21, 1, 2)
	   ' ret = .AddCellSpan(C_W1_NM1, C_22, 1, 2)
       
	    ' ù��° ��� ��� ���� 
		.Col = C_W1_NM1
		.Row = C_01	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "��" & VbCrlf & "��" & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & "��"
		.Row = C_06	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "��" & VbCrlf & "��" & vbCrLf & "��" & vbCrLf & "��"
		.Row = C_10	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "�ͱݺ�" & vbCrLf & "���Ժ�"
		.Row = C_11	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "��" & VbCrlf & VbCrlf & "��" & vbCrLf & VbCrlf & "��"
		.Row = C_16	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "��" & VbCrlf & "��" & vbCrLf & "��"
		.Row = C_19	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(119) ��ȭ�ڻ����ä�򰡼���"
		.Row = C_21	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "���������ε���" & VbCrlf & "� ������" & VbCrlf & "���Ա�����"
	    
		' �ι�° ��� ��� ���� 
		.Col = C_W1_NM2
		.Row = C_01	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(101) ������������غ��"
		.Row = C_02	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(102) �����޿�����"
		.Row = C_03	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(103) ���������"
		.Row = C_04	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(104) �������"
		.Row = C_05	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(105) ��ձ�"
		.Row = C_06	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(106) �պ�������"
		.Row = C_07	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(107) ����������"
		.Row = C_08	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(108) ���������ڻ�絵����"
		.Row = C_09	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(109) ��ȯ�ڻ�絵����"
		.Row = C_10	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(110) ä�������͵� �̿���ձ� ������"
		.Row = C_11	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(111) �翬�ձݱ�α�"
		.Row = C_12	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(112) 50% �ձݱ�α�"
		.Row = C_13	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(113) ������α��ѵ���"
		.Row = C_14	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(114) ������α�"
		.Row = C_15	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(115) ��Ÿ��α�"
		.Row = C_16	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(116) ������ѵ���"
		.Row = C_17	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(117) �����" & vbCrlf & "(118 ����)"
		.Row = C_18	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(118) 5���� (�������� "& vbCrLf &"10����) �ʰ� �����"
		.Row = C_21	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(120) ���������ε��� ��"
		
		' �ٰŹ� ���� �ֱ� 
		.Col = C_W2
		.Row = C_01	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.Text = "���μ��� ��29��           ����Ư�����ѹ� ��74��"
		.Row = C_02	:	.TypeHAlign = 0	:			.Text = "���μ��� ��33��"
		.Row = C_03	:	.TypeHAlign = 0	:			.Text = "���μ�������� ��44����2"
		.Row = C_04	:	.TypeHAlign = 0	:			.Text = "���μ��� ��34��"
		.Row = C_05	:	.TypeHAlign = 0	:			.Text = "���μ��� ��34��"
		.Row = C_06	:	.TypeHAlign = 0	:			.Text = "���μ��� ��44��"
		.Row = C_07	:	.TypeHAlign = 0	:			.Text = "���μ��� ��46��"
		.Row = C_08	:	.TypeHAlign = 0	:			.Text = "���μ��� ��47��"
		.Row = C_09	:	.TypeHAlign = 0	:			.Text = "���μ��� ��50��"
		.Row = C_10	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "���μ��� ��18����8ȣ"
		.Row = C_11	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.Text = "���μ��� ��24����2��   ����Ư�����ѹ� ��73�� ��1�� ��1ȣ"
		.Row = C_12	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.Text = "����Ư�����ѹ�  ��73�� ��1�� ��2ȣ ���� ��14ȣ"
		.Row = C_13	:	.TypeHAlign = 0	:			.Text = "���μ��� ��24�� ��1��"
		.Row = C_14	:	.TypeHAlign = 0	:			.Text = "���μ��� ��24�� ��1��"
		.Row = C_15	:	.TypeHAlign = 0	:			.Text = "���μ��� ��24�� ��1��"
		.Row = C_16	:	.TypeHAlign = 0	:			.Text = "���μ��� ��24�� ��1��"
		.Row = C_17	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "���μ��� ��24�� ��1��"
		.Row = C_18	:	.TypeHAlign = 0	:			.Text = "���μ��� ��24�� ��2��"
		.Row = C_19	:	.TypeHAlign = 0	:			.Text = "���μ��� ��42��"
		.Row = C_21	:	.TypeVAlign = 2	:				.Text = "���μ��� ��28�� ��1��"

		
		' �⺻�ڵ尪�Է��ϱ� 
		.Col = C_W1
		.Row = C_01	:	.TypeHAlign = 0	:			.Text = "101"
		.Row = C_02	:	.TypeHAlign = 0	:			.Text = "102"
		.Row = C_03	:	.TypeHAlign = 0	:			.Text = "103"
		.Row = C_04	:	.TypeHAlign = 0	:			.Text = "104"
		.Row = C_05	:	.TypeHAlign = 0	:			.Text = "105"
		.Row = C_06	:	.TypeHAlign = 0	:			.Text = "106"
		.Row = C_07	:	.TypeHAlign = 0	:			.Text = "107"
		.Row = C_08	:	.TypeHAlign = 0	:			.Text = "108"
		.Row = C_09	:	.TypeHAlign = 0	:			.Text = "109"
		.Row = C_10	:	.TypeHAlign = 0	:			.Text = "110"
		.Row = C_11	:	.TypeHAlign = 0	:			.Text = "111"
		.Row = C_12	:	.TypeHAlign = 0	:			.Text = "112"
		.Row = C_13	:	.TypeHAlign = 0	:			.Text = "113"
		.Row = C_14	:	.TypeHAlign = 0	:			.Text = "114"
		.Row = C_15	:	.TypeHAlign = 0	:			.Text = "115"
		.Row = C_16	:	.TypeHAlign = 0	:			.Text = "116"
		.Row = C_17	:	.TypeHAlign = 0	:			.Text = "117"
		.Row = C_18	:	.TypeHAlign = 0	:			.Text = "118"
		.Row = C_19	:	.TypeHAlign = 0	:			.Text = "119"
		.Row = C_20	:	.TypeHAlign = 0	:			.Text = "1191"
		.Row = C_21	:	.TypeHAlign = 0	:			.Text = "120"




		' Ȩ�ؽ��ڵ尪�Է��ϱ� 
		.Col = C_W2_CD
		.Row = C_01	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "53"
		.Row = C_02	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "12"
		.Row = C_03	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "71"
		.Row = C_04	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "13"
		.Row = C_05	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "72"
		.Row = C_06	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "55"
		.Row = C_07	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "56"
		.Row = C_08	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "57"
		.Row = C_09	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "58"
		.Row = C_10	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "59"
		.Row = C_11	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "41"
		.Row = C_12	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "64"
		.Row = C_13	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "66"
		.Row = C_14	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "42"
		.Row = C_15	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "73"
		.Row = C_16	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "49"
		.Row = C_17	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "65"
		.Row = C_18	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "61"
		.Row = C_19	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "74"
		.Row = C_20	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "75"
		.Row = C_21	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "76"


		.rowheight(C_01) = 20	' ���� ������ 
		.rowheight(C_10) = 20	' ���� ������ 
		.rowheight(C_11) = 30	' ���� ������ 
		.rowheight(C_12) = 20	' ���� ������ 
		.rowheight(C_17) = 20	' ���� ������ 
		.rowheight(C_18) = 20	' ���� ������ 
		.rowheight(C_20) = 20	' ���� ������ 
		.rowheight(C_21) = 30	' ���� ������ 


		.Col = C_W3	:	.Row = C_01	:	.TypeVAlign = 2	:	.Row = C_10	:	.TypeVAlign = 2	:	.Row = C_11	:	.TypeVAlign = 2	:	.Row = C_17	:	.TypeVAlign = 2	:	.Row = C_20	:	.TypeVAlign = 2 : .Row = C_21	:	.TypeVAlign = 2: .Row = C_18	:	.TypeVAlign = 2
		.Col = C_W4	:	.Row = C_01	:	.TypeVAlign = 2	:	.Row = C_10	:	.TypeVAlign = 2	:	.Row = C_11	:	.TypeVAlign = 2	:	.Row = C_17	:	.TypeVAlign = 2	:	.Row = C_20	:	.TypeVAlign = 2 : .Row = C_21	:	.TypeVAlign = 2: .Row = C_18	:	.TypeVAlign = 2
		.Col = C_W5	:	.Row = C_01	:	.TypeVAlign = 2	:	.Row = C_10	:	.TypeVAlign = 2	:	.Row = C_11	:	.TypeVAlign = 2	:	.Row = C_17	:	.TypeVAlign = 2	:	.Row = C_20	:	.TypeVAlign = 2: .Row = C_21	:	.TypeVAlign = 2: .Row = C_18	:	.TypeVAlign = 2
		
		

		.Row = C_20	: .RowHidden = True
'		.Row = C_22	: .RowHidden = True
'		.Row = C_23	: .RowHidden = True

    	
	End With 

End Sub

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �·ε�� ���۷����޽��� �����´�.
     wgRefDoc = GetDocRef(sCoCd,sFiscYear, sRepType, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData

	Call InitSpreadRow()
	Call SetSpreadLock
			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
End Function

Function GetRefOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr, iSeqNo, iLastRow, iRow
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	Call Fn_GridCalc()
	lgBlnFlgChgValue = True
	Frm1.vspdData.focus			
End Function

Function ChangeRowFlg(iObj)
	Dim iRow
	
	With iObj
		ggoSpread.Source = iObj
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
		Next
	End With
End Function



Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, iCnt
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		lgFISC_START_DT = CDate(lgF0)
	Else
		lgFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		lgFISC_END_DT = CDate(lgF1)
	Else
		lgFISC_END_DT = ""
	End if
		
End Sub

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
  
	' �����Ѱ� 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	

	Call InitComboBox	' �����ؾ� �Ѵ�. ����� ȸ��������� �о���� ���� 

	Call InitData

	Call FncQuery()
	
     

End Sub

'============================================  ����� �Լ�  ====================================
Function Fn_GridCalc()
	Dim iRow, dblSum
	Dim dblW3, dblW4, dblW5
	
    ggoSpread.Source = Frm1.vspdData

	With Frm1.vspdData
		For iRow = C_01 To C_21
		' (5) = (3) - (4) : 113, 116, 120, 121, 123���� 
			If iRow <> C_13 And iRow <> C_16 And iRow <> C_20 And iRow <> C_21  Then ' ������������ ���ŵ� 
				.Row = iRow	:	.Col = C_W3	:	dblW3 = UNICdbl(.Text)
				.Row = iRow	:	.Col = C_W4	:	dblW4 = UNICdbl(.Text)
				.Row = iRow	:	.Col = C_W5	:	dblW5 = dblW3 - dblW4
				.Text = dblW5
			End If
		Next
			
	End With

End Function


'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	Call GetFISC_DATE
End Sub

Sub txtw124_Change( )
    lgBlnFlgChgValue = True
End Sub

Sub txtw125_Change( )
    lgBlnFlgChgValue = True
End Sub

'============================================  �׸��� �̺�Ʈ   ====================================


'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIdx, iRow, sW3, sW4, dblW2

	With Frm1.vspdData
		Select Case Col
			Case C_W3_NM
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col +1
				.Value = iIdx
			Case C_W3
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
		End Select
		
		

	End With
End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	lgBlnFlgChgValue= True ' ���濩�� 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.UpdateRow Row

	' --- �߰��� �κ� 
	Call Fn_GridCalc()


	
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = Frm1.vspdData
   
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    	Exit Sub
       ggoSpread.Source = Frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	Frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = Frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If Frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	ggoSpread.Source = Frm1.vspdData
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = Frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
'    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With Frm1.vspdData
		If Row > 0 And Col = C_W1_BT Then
		    .Row = Row
		    .Col = C_W1_BT

		    Call OpenAdItem()
		End If
    End With
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

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
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
'    Call InitVariables													<%'Initializes local global variables%>
'    Call InitData                              
    															
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

    'If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

' ---------------------- ���ĳ� ���� -------------------------
Function  Verification()
	Dim dblW11, dblW12, dblW16, dblW14, dblW15, dblW13
	
	Verification = False

	Verification = True	
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD  , iRow

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
    Call InitSpreadRow
    Call SetSpreadLock
    Call InitVariables
    Call InitData

    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
    lgIntFlgMode = parent.OPMD_CMODE
    With Frm1.vspdData
		For iRow = 1 To .MaxRows
				.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
		Next
    End With
	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
	Dim iActiveRow
	
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	With Frm1.vspdData
	    ggoSpread.Source = Frm1.vspdData
	    iActiveRow = .ActiveRow

		If .ActiveRow > 0 and .ActiveRow <> .MaxRows Then
			.focus
			.ReDraw = False
		
			ggoSpread.CopyRow
			.Col = C_W1
			.Text = ""

			Call SetDefaultVal(iActiveRow + 1, 1)
			SetSpreadColor iActiveRow, iActiveRow + 1
			.ReDraw = True
			
			Call Fn_GridCalc()
    
		End If
	End With


    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows, iActiveRow, dblSum

	With Frm1.vspdData
		.focus
		iActiveRow = .ActiveRow
		ggoSpread.Source = Frm1.vspdData
		If CheckTotalRow(Frm1.vspdData, .ActiveRow) = True Then
			MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
			Exit Function
		Else
			lDelRows = ggoSpread.EditUndo
		End If
		
	End With

	Call Fn_GridCalc()

End Function

' -- �հ� ������ üũ(Header Grid)
Function CheckTotalRow(Byref pObj, Byval pRow) 
	CheckTotalRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = "999999" And pObj.MaxRows > 1 Then	 ' �հ� �� 
		CheckTotalRow = True
	End If
End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo
    Dim ret

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

	With Frm1.vspdData
	
		.focus
		ggoSpread.Source = Frm1.vspdData
	
		iSeqNo = .MaxRows+1
	
		if .MaxRows = 0 then
		
			ggoSpread.InsertRow  imRow 
			.Col	= C_SEQ_NO	:	.Text	= 1
			SetSpreadColor 1, 1
			
			ggoSpread.InsertRow  imRow 
			.Row = .MaxRows
			.Col	= C_SEQ_NO	:	.Text	= 999999
			ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' ���� 2�� ��ħ 
			.Col	= C_W1	:	.CellType = 1	:	.Text	= "��"	:	.TypeHAlign = 2
			SetSpreadColor .MaxRows, .MaxRows
			.Row  = 1
			.ActiveRow = 1

		else
			iRow = .ActiveRow

			If iRow = .MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
				iRow = iRow - 1
				ggoSpread.InsertRow iRow, imRow 
				SetSpreadColor iRow, iRow + imRow + 1

				Call SetDefaultVal(iRow + 1, imRow)
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor iRow, iRow + imRow + 1

				Call SetDefaultVal(iRow + 1, imRow)
			End If   
			.vspdData.Row  = iRow + 1
			.vspdData.ActiveRow = iRow +1
			
        End if 	
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With Frm1.vspdData
	
		If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
			.Row = iRow
			.Value = MaxSpreadVal(Frm1.vspdData, C_SEQ_NO, iRow)
		Else
			iSeqNo = MaxSpreadVal(Frm1.vspdData, C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
			
			For i = iRow to iRow + iAddRows -1
				.Row = i	:	.Col = C_SEQ_NO
				If .Text <> 999999 Then
					: .Value = iSeqNo : iSeqNo = iSeqNo + 1
				End If
			Next
		End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows, iActiveRow, dblSum

	With Frm1.vspdData
		.focus
		iActiveRow = .ActiveRow
		ggoSpread.Source = Frm1.vspdData
		If CheckTotalRow(Frm1.vspdData, .ActiveRow) = True Then
			MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
			Exit Function
		Else
			lDelRows = ggoSpread.DeleteRow
		End If
		
	End With

	Call Fn_GridCalc()

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
	
    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Then
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
        'strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryFalse()
	Call FncNew()
End Function
		
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	If Frm1.vspdData.MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		

	    Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>
	End If
	
	Call InitSpreadRow()
	Call SetSpreadLock
'	Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
	Frm1.vspdData.focus			
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
    
	With Frm1.vspdData

		ggoSpread.Source = Frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
		
		' ----- 1��° �׸��� 
		For lRow = 1 To .MaxRows

	       .Row = lRow
	       .Col = 0
	    
	       Select Case .Text
	           Case  ggoSpread.InsertFlag                                      '��: Insert
	                                              strVal = strVal & "C"  &  Parent.gColSep
	           Case  ggoSpread.UpdateFlag                                      '��: Update
	                                              strVal = strVal & "U"  &  Parent.gColSep
		       Case  ggoSpread.DeleteFlag                                      '��: Delete
		                                          strVal = strVal & "D"  &  Parent.gColSep
		       Case Else
		                                          strVal = strVal & ""  &  Parent.gColSep
	       End Select
	       
		  ' ��� �׸��� ����Ÿ ����     
'		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = 1 To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
'		  End If  
		Next
	
	End With

	Frm1.txtSpread.value      = strVal
	strVal = ""

	Frm1.txtMode.value		=	Parent.UID_M0002
	Frm1.txtFlgMode.Value	=	lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' ���� ������ ���� ���� %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
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

Function ProgramJump
    Call PgmJump(JUMP_PGM_ID)
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
					<TD WIDTH=* align=right><A href="vbscript:GetRef">�ݾ׺ҷ�����</A></TD>
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
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="�������" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD >
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">
											<TABLE <%=LR_SPACE_TYPE_60%> border="1" height=100% width="100%">
												<TR>
													<TD CLASS="TD51" width="10%" ALIGN=CENTER>
														�󿩹��� 
													</TD>
													<TD CLASS="TD51" width="26%" ALIGN=CENTER>
														(121)�ҵ�ó�� �ݾ�   (���μ����������106��)
													</TD>
													<TD CLASS="TD51" width="4%" ALIGN=CENTER>
														97
													</TD>
													<TD CLASS="TD51" width="17%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW124" name=txtW124 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
													<TD CLASS="TD51" width="22%">
														(122)����ó�� �ݾ�      (�����462����)
													</TD>
													<TD CLASS="TD51" width="4%" ALIGN=CENTER>
														98
													</TD>
													<TD CLASS="TD51" width="17%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW125" name=txtW125 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
												</TR>
											</TABLE>

										</TD>
									</TR>
								</TABLE>
							    </TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
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
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
