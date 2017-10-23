<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ÿ���� 
'*  3. Program ID           : W6125MA1
'*  4. Program Name         : W6125MA1.asp
'*  5. Program Desc         : �� 8ȣ(��) �������鼼�� ���� 
'*  6. Modified date(First) : 2005/02/04
'*  7. Modified date(Last)  : 2005/02/04
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

Const BIZ_MNU_ID		= "W6125MA1"
Const BIZ_PGM_ID		= "W6125MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "W6125OA1"


Const TYPE_1	= 0		' �׸��带 �������� ���� ��� 
Const TYPE_2	= 1		
Const TYPE_3	= 2		

' -- �׸��� �÷� ���� 
Dim	C_W_TYPE
Dim C_W_SPAN
Dim C_W1_CD
Dim C_W1
Dim C_W2
Dim C_W2_1
Dim C_W3
Dim C_W4
Dim C_W7

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgCurrGrid, lgvspdData(2), IsRunEvents


'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W_TYPE	= 1
	C_W1_CD		= 2
	C_W_SPAN	= 3
	C_W1		= 4
	C_W2		= 5
	C_W2_1		= 6
	C_W3		= 7
	C_W4		= 8
	C_W7		= 9
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

    lgCurrGrid = TYPE_1
    IsRunEvents = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	' ��ȸ����(����)
	Dim IntRetCD1
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," ufn_TB_MINOR('W1018', '" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

End Sub


Sub InitSpreadComboBox()

End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1)		= frm1.vspdData0
	Set lgvspdData(TYPE_2)		= frm1.vspdData1
	Set lgvspdData(TYPE_3)		= frm1.vspdData2

		
    Call initSpreadPosVariables()  

	'Call AppendNumberPlace("6","3","2")	' -- ����(����)
	
	' 1�� �׸��� 

	With lgvspdData(TYPE_1)
				
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W4 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_W_TYPE,	"����", 10,,,15,1
		ggoSpread.SSSetEdit		C_W1_CD,	"�ڵ�", 10,,,15,1
		ggoSpread.SSSetEdit		C_W_SPAN,	"����", 5,,,50,1
		ggoSpread.SSSetEdit		C_W1,		"(1)�� ��", 35,,,50,1
		ggoSpread.SSSetEdit		C_W2,		"(2)�ٰŹ�����", 35,,,50,1
		ggoSpread.SSSetEdit		C_W2_1,		"�ڵ�"	, 7, 2,,10,1
		ggoSpread.SSSetFloat	C_W3,		"(3)��󼼾�"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W4,		"(4)��������"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	

		ret = .AddCellSpan(C_W_SPAN	, 0, 2, 1)	
		'.rowheight(-998) = 15	' ���� ������	(2���� ���, 1���� 15)
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W1_CD,True)
		
		'Call InitSpreadComboBox
		
		.ReDraw = true	
				
	End With 

 
	' 2�� �׸��� 
	With lgvspdData(TYPE_2)
				
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W4 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_W_TYPE,	"����", 10,,,15,1
		ggoSpread.SSSetEdit		C_W1_CD,	"�ڵ�", 10,,,15,1
		ggoSpread.SSSetEdit		C_W_SPAN,	"����", 5,,,50,1
		ggoSpread.SSSetEdit		C_W1,		"(1)�� ��", 35,,,50,1
		ggoSpread.SSSetEdit		C_W2,		"(2)�ٰŹ�����", 35,,,50,1
		ggoSpread.SSSetEdit		C_W2_1,		"�ڵ�"	, 7, 2,,10,1
		ggoSpread.SSSetFloat	C_W3,		"(3)��󼼾�"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W4,		"(4)��������"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	

		ret = .AddCellSpan(C_W_SPAN	, 0, 2, 1)	
		'.rowheight(-998) = 15	' ���� ������	(2���� ���, 1���� 15)
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W1_CD,True)
		
				
		.ReDraw = true	
				
	End With 

	' 3�� �׸��� 
	With lgvspdData(TYPE_3)
				
		ggoSpread.Source = lgvspdData(TYPE_3)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W7 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_W_TYPE,	"����", 10,,,15,1
		ggoSpread.SSSetEdit		C_W1_CD,	"�ڵ�", 10,,,15,1
		ggoSpread.SSSetEdit		C_W_SPAN,	"����", 5,,,50,1
		ggoSpread.SSSetEdit		C_W1,		"(1)�� ��", 35,,,50,1
		ggoSpread.SSSetEdit		C_W2,		"(2)�ٰŹ�����", 20,,,50,1
		ggoSpread.SSSetEdit		C_W2_1,		"�ڵ�"	, 7, 2,,10,1
		ggoSpread.SSSetFloat	C_W3,		"(5)�����̿���"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W4,		"(6)���߻���"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W7,		"(7)��������"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	

		ret = .AddCellSpan(C_W_SPAN	, 0, 2, 1)	
		'.rowheight(-998) = 15	' ���� ������	(2���� ���, 1���� 15)
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W1_CD,True)
		
				
		.ReDraw = true	
				
	End With 
	
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	'Call GetFISC_DATE

End Sub

Sub SpreadInitData()
    ' �׸��� �ʱ� ����Ÿ���� 
    Dim arrW1_CD, arrW1, arrW2, arrW2_1, iMaxRows, iRow, iMinorCnt, iType, sMinorCd, ret

	For iType = TYPE_1 To TYPE_3
		Select Case iType
			Case TYPE_1
				sMinorCd = "W1066"
			Case TYPE_2
				sMinorCd = "W1067"
			Case TYPE_3
				sMinorCd = "W1068"
		End Select
		
		call CommonQueryRs("MINOR_CD, MINOR_NM, REFERENCE_1, REFERENCE_2","ufn_TB_Configuration('" & sMinorCd & "','" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		arrW1_CD	= Split(lgF0, Chr(11))
		arrW1		= Split(lgF1, Chr(11))
		arrW2		= Split(lgF2, Chr(11))
		arrW2_1		= Split(lgF3, Chr(11))
    
 		iMaxRows = UBound(arrW1_CD)
	
		With lgvspdData(iType)
			.Redraw = False
			
			ggoSpread.Source = lgvspdData(iType)
			
			ggoSpread.InsertRow , iMaxRows

			Select Case iType
				Case TYPE_1
					' �׸��� ��� ��ħ 
					ret = .AddCellSpan(C_W_SPAN	, 1, 1, 20)	
					.BlockMode = True
					.Col = C_W_SPAN : .Row = 1 : .Col2 = C_W_SPAN : .Row2 = 20
					.TypeHAlign = 2 : .TypeVAlign = 2 : .TypeEditMultiLIne = True
					.BlockMode = False
					.Col = C_W_SPAN	: .Row = 1  : .Value = "��" & vbCrLf & vbCrLf & "��" & vbCrLf & vbCrLf & "��" & vbCrLf & vbCrLf & "��"

					ret = .AddCellSpan(C_W_SPAN	, 21, 1, 27)	
					.BlockMode = True
					.Col = C_W_SPAN : .Row = 21 : .Col2 = C_W_SPAN : .Row2 = 27
					.TypeHAlign = 2 : .TypeVAlign = 2 : .TypeEditMultiLIne = True
					.BlockMode = False
					.Col = C_W_SPAN	: .Row = 21  : .Value = "��" & vbCrLf & vbCrLf & "��" & vbCrLf & vbCrLf & "��" & vbCrLf & vbCrLf & "��"
					
					ret = .AddCellSpan(C_W_SPAN	, 28, 2, 1)	
				Case TYPE_2
					' �׸��� ��� ��ħ 
					ret = .AddCellSpan(C_W_SPAN	, 1, 1, .MaxRows)	
					.BlockMode = True
					.Col = C_W_SPAN : .Row = -1  : .Col2 = C_W_SPAN : .Row2 = -1
					.TypeHAlign = 2 : .TypeVAlign = 2 : .TypeEditMultiLIne = True
					.BlockMode = False
					.Col = C_W_SPAN	: .Row = 1 : .Value = "��" & vbCrLf & vbCrLf & "��" & vbCrLf & vbCrLf & "��" & vbCrLf & vbCrLf & "��"

				Case TYPE_3
					' �׸��� ��� ��ħ 
					ret = .AddCellSpan(C_W_SPAN	, 1, 1, .MaxRows-4)	
					.BlockMode = True
					.Col = C_W_SPAN : .Row = -1  : .Col2 = C_W_SPAN : .Row2 = -1
					.TypeHAlign = 2 : .TypeVAlign = 2 : .TypeEditMultiLIne = True
					.BlockMode = False
					ret = .AddCellSpan(C_W_SPAN	, .MaxRows-3, 2, 1)	
					ret = .AddCellSpan(C_W_SPAN	, .MaxRows-2, 2, 1)	
					ret = .AddCellSpan(C_W_SPAN	, .MaxRows-1, 2, 1)	
					ret = .AddCellSpan(C_W_SPAN	, .MaxRows  , 2, 1)	
					.Col = C_W_SPAN	: .Row = 1 : .Value = "��" & vbCrLf & vbCrLf & "��" & vbCrLf & vbCrLf & "��" & vbCrLf & vbCrLf & "��"

			End Select
						
			' �迭�� �׸��忡 ���� 
			For iRow = 1 To iMaxRows
				
				.Row = iRow
				.Col = C_W_TYPE	: .value = iType
				.Col = C_W1_CD	: .value = arrW1_CD(iRow-1)
				
				Select Case arrW2_1(iRow-1)
					Case "50", "51", "83", "89"
						.Col = C_W_SPAN	: .value = arrW1(iRow-1)
					Case Else
						.Col = C_W1		: .value = arrW1(iRow-1)
				End Select

				.Col = C_W2		: .value = arrW2(iRow-1)
				.Col = C_W2_1	: .value = arrW2_1(iRow-1)
				
			Next
			.Redraw = True
		End With
		
		Call SetSpreadLock(iType)

	Next

End Sub

Sub SetSpreadLock(pType)

	With lgvspdData(pType)
	
		ggoSpread.Source = lgvspdData(pType)	

		Select Case pType
			Case TYPE_1 
				ggoSpread.SpreadLock C_W_TYPE, 1, C_W2_1, 18	' ��ü ���� 
				ggoSpread.SpreadLock C_W_TYPE, 20, C_W4,  20	' ��ü ���� 

				ggoSpread.SpreadLock C_W_TYPE, 21, C_W2_1, 25	' ��ü ���� 
				ggoSpread.SpreadLock C_W_TYPE, 27, C_W4,  28	' ��ü ���� 
				
				'ggoSpread.SpreadLock C_W_TYPE, 19, C_W1_CD, 19	' ��ü ���� 
				ggoSpread.SpreadLock C_W2_1, -1, C_W2_1, -1	' ��ü ���� 
				
			Case TYPE_2
				ggoSpread.SpreadLock C_W_TYPE, 1, C_W2_1, 13	' ��ü ���� 
				ggoSpread.SpreadLock C_W_TYPE, .MaxRows, C_W4,  .MaxRows	' ��ü ���� 
				ggoSpread.SpreadLock C_W_TYPE, 14, C_W1_CD, 14	' ��ü ���� 
				ggoSpread.SpreadLock C_W2_1, 14, C_W2_1, 14	' ��ü ���� 
				
			Case TYPE_3
				ggoSpread.SpreadLock C_W_TYPE, 1, C_W2_1, 16	' ��ü ���� 
				ggoSpread.SpreadLock C_W_TYPE, 18, C_W2_1, 18	' ��ü ���� 
				ggoSpread.SpreadLock C_W2_1, 17, C_W2_1, 17	' ��ü ���� 
				ggoSpread.SpreadLock C_W_TYPE, 17, C_W1_CD, 17	' ��ü ���� 
				ggoSpread.SpreadLock C_W_TYPE, .MaxRows-4, C_W7,  .MaxRows-2	' ��ü ���� 
				ggoSpread.SpreadLock C_W_TYPE, .MaxRows-2, C_W2_1,  .MaxRows	' ��ü ���� 
				
		End Select
		
	End With	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(pType)	

			
	End With	
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case pvSpdNo
       Case TYPE_1
            ggoSpread.Source = frm1.vspdData0
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_W1_CD		= iCurColumnPos(1)
            C_W1		= iCurColumnPos(2)
            C_W2		= iCurColumnPos(3)
            C_W2_1		= iCurColumnPos(4)
            C_W3		= iCurColumnPos(5)
            C_W4		= iCurColumnPos(6)
 
        Case TYPE_2
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_W1_CD		= iCurColumnPos(1)
            C_W1		= iCurColumnPos(2)
            C_W2		= iCurColumnPos(3)
            C_W2_1		= iCurColumnPos(4)
            C_W3		= iCurColumnPos(5)
            C_W4		= iCurColumnPos(6)      

        Case TYPE_3
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_W1_CD		= iCurColumnPos(1)
            C_W1		= iCurColumnPos(2)
            C_W2		= iCurColumnPos(3)
            C_W2_1		= iCurColumnPos(4)
            C_W3		= iCurColumnPos(5)
            C_W4		= iCurColumnPos(6)   
            C_W7		= iCurColumnPos(7)   
    End Select    
End Sub


Sub SetSpreadTotalLine()
	Dim iRow, i

	For i = TYPE_1 To TYPE_2
		ggoSpread.Source = lgvspdData(i)
		With lgvspdData(i)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W1 : .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				'ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
			End If
		End With
	Next
End Sub 

' �ش� �׸��忡�� ����Ÿ�������� 
Function GetGrid(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : GetGrid = UNICDbl(.Value)
	End With
End Function

' �ش� �׸��忡�� ����Ÿ�������� 
Function GetGridText(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : GetGridText = .text
	End With
End Function

' �ش� �׸��忡�� ����Ÿ�������� 
Function PutGrid(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : .Value = pVal
		
		ggoSpread.Source = lgvspdData(pType)
		ggoSpread.UpdateRow pRow
		
	End With
End Function

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �׸���1�� �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW2, arrW3, arrW4, arrW7, iMaxRows, sTmp, iType
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf

	' ����� ��ġ�� �˷��� 
	Dim iCol, iRow, i

	IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
		
	If IntRetCD = vbNo Then
		 Exit Function
	End If

	IntRetCD = CommonQueryRs("W1, W2, W3, W4, W7"," dbo.ufn_TB_8A_GetRef_" & C_REVISION_YM & "('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		arrW3		= Split(lgF2, chr(11))
		arrW4		= Split(lgF3, chr(11))
		arrW7		= Split(lgF4, chr(11))
		iMaxRows	= UBound(arrW1)
		
		For i = 0 To iMaxRows -1
			iType	= CDbl(arrW1(i))
			iRow	= CDbl(arrW2(i))
			
			lgvspdData(iType).Row = iRow
			lgvspdData(iType).Col = C_W3	: lgvspdData(iType).value = arrW3(i)
			lgvspdData(iType).Col = C_W4	: lgvspdData(iType).value = arrW4(i)
			
			If iType = TYPE_3 Then
				lgvspdData(iType).Col = C_W7	: lgvspdData(iType).value = arrW7(i)
			End If

		Next
		
		Call CalSum
		
	End If
	
	
End Function

Function GetRowByW1_CD(Byval pType, Byval pW1_CD)
	Dim iRow, iMaxRows
	With lgvspdData(pType)
		iMaxRows = .MaxRows
		.Col = C_W1_CD
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			If .Text = CStr(pW1_CD) Then 
				GetRowByW1_CD = iRow
				Exit Function
			End If
		Next
	End With
End Function

Function GetRowByW2_1(Byval pType, Byval pW2_1)
	Dim iRow, iMaxRows
	With lgvspdData(pType)
		iMaxRows = .MaxRows
		.Col = C_W1_CD
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			If .Text = CStr(pW2_1) Then 
				GetRowByW2_1 = iRow
				Exit Function
			End If
		Next
	End With
End Function

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
 
	Call InitComboBox	
	Call InitData
	Call SpreadInitData
	
    Call FncQuery
    
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
	'Call GetFISC_DATE
End Sub


'============================================  �׸��� �̺�Ʈ   ====================================
' -- 0�� �׸��� 
Sub vspdData0_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_1
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData0_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_1
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData0_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_1
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_GotFocus()
	lgCurrGrid = TYPE_1
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData0_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_1
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData0_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_1
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData0_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_1
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData0_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_1
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 1�� �׸��� 
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_2
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_2
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 1�� �׸��� 
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_3
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_3
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_3
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_GotFocus()
	lgCurrGrid = TYPE_3
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_3
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_3
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_3
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_3
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, dblSum141, i170Row, i160Row, i140Row, i120Row, i180Row
	
	lgBlnFlgChgValue= True ' ���濩�� 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col

    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(lgvspdData(Index).text) < UNICDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row

	' --- �߰��� �κ� 
	With lgvspdData(Index)

	If Index = TYPE_1 Then	'1�� �׸� 
		Select Case Col
			Case C_W3, C_W4

				'Call FncSumSheet(lgvspdData(TYPE_1), Col, 1, .MaxRows-1, true, .MaxRows, Col, "V")	' ���� ���� �հ� 
				'ggoSpread.UpdateRow .MaxRows
		End Select
	ElseIf Index = TYPE_2 Then	'2�� �׸� 
		Select Case Col
			Case C_W3, C_W4

				'Call FncSumSheet(lgvspdData(TYPE_2), Col, 1, .MaxRows-1, true, .MaxRows, Col, "V")	' ���� ���� �հ� 
				'ggoSpread.UpdateRow .MaxRows
		End Select

	ElseIf Index = TYPE_3 Then	'3�� �׸� 
		
		Select Case Col
			Case C_W3, C_W4, C_W7
				'i160Row = GetRowByW1_CD(TYPE_3, "160")
				' 141+160�� �� 
				'Call FncSumSheet(lgvspdData(TYPE_3), Col, 1, i160Row-1, true, i160Row, Col, "V")	' ���� ���� �հ� 
				'ggoSpread.UpdateRow i160Row
		End Select

	End If
	
	' �հ� 
	'i120Row = GetRowByW1_CD(TYPE_1, "120")
	'i140Row = GetRowByW1_CD(TYPE_2, "140")
	'i160Row = GetRowByW1_CD(TYPE_3, "160")
	'i170Row = GetRowByW1_CD(TYPE_3, "170")
	'i180Row = GetRowByW1_CD(TYPE_3, "180")
	
	' 170�� ��� : C_W3 = 160) �� ����Ͽ� �Է��� 
	'Call PutGrid(TYPE_3, C_W3, i170Row, GetGrid(TYPE_3, C_W3, i160Row))
	
	' 170�� ��� : C_W4 = (140��160) �� ����Ͽ� �Է���.
	'dblSum = GetGrid(TYPE_2, C_W3, i140Row) + GetGrid(TYPE_3, C_W4, i160Row)
	'Call PutGrid(TYPE_3, C_W4, i170Row, dblSum)
	' 170�� ��� : C_W7 = (140��160) �� ����Ͽ� �Է��� 
	'dblSum = GetGrid(TYPE_2, C_W4, i140Row) + GetGrid(TYPE_3, C_W7, i160Row)
	'Call PutGrid(TYPE_3, C_W7, i170Row, dblSum)

	' 180�� ��� : C_W3 = 170) �� ����Ͽ� �Է��� 
	'Call PutGrid(TYPE_3, C_W3, i180Row, GetGrid(TYPE_3, C_W3, i170Row))
	' 180�� ��� : C_W4 = (120��170) �� ����Ͽ� �Է���.
	'dblSum = GetGrid(TYPE_1, C_W3, i120Row) + GetGrid(TYPE_3, C_W4, i170Row)
	'Call PutGrid(TYPE_3, C_W4, i180Row, dblSum)
	' 180�� ��� : C_W7 = (120��170) �� ����Ͽ� �Է���.
	'dblSum = GetGrid(TYPE_1, C_W4, i120Row) + GetGrid(TYPE_3, C_W7, i170Row)
	'Call PutGrid(TYPE_3, C_W7, i180Row, dblSum)

	'ggoSpread.UpdateRow i170Row	' -- PutGrid �Լ����� �̵� 
	'ggoSpread.UpdateRow i180Row
	
	Call CalSum

	End With
	
End Sub

Sub CalSum()
	Dim dblSum, dbl10(1), dbl50(2)
	
    ggoSpread.Source = lgvspdData(TYPE_1)
	
	' -- �׸���1 
	Call FncSumSheet(lgvspdData(TYPE_1), C_W3, 1, 19, true, 20, C_W3, "V")	' ���� ���� �հ� 
	Call FncSumSheet(lgvspdData(TYPE_1), C_W4, 1, 19, true, 20, C_W4, "V")	' ���� ���� �հ� 

    ggoSpread.UpdateRow 20
	
	Call FncSumSheet(lgvspdData(TYPE_1), C_W3,21, 26, true, 27, C_W3, "V")	' ���� ���� �հ� 
	Call FncSumSheet(lgvspdData(TYPE_1), C_W4,21, 26, true, 27, C_W4, "V")	' ���� ���� �հ� 

    ggoSpread.UpdateRow 27
	
	' -- �հ� = �Ұ� + �Ұ� 
	dbl10(0) = GetGrid(TYPE_1, C_W3, 20) + GetGrid(TYPE_1, C_W3, 27)
	Call PutGrid(TYPE_1, C_W3, 28, dbl10(0))

	dbl10(1) = GetGrid(TYPE_1, C_W4, 20) + GetGrid(TYPE_1, C_W4, 27)
	Call PutGrid(TYPE_1, C_W4, 28, dbl10(1))

    ggoSpread.UpdateRow 28
	
	ggoSpread.Source = lgvspdData(TYPE_2)
	
	' -- �׸���2�� ��ȭ����		
	Call FncSumSheet(lgvspdData(TYPE_2), C_W3, 1, lgvspdData(TYPE_2).MaxRows-1, true, lgvspdData(TYPE_2).MaxRows, C_W3, "V")	' ���� ���� �հ� 
	Call FncSumSheet(lgvspdData(TYPE_2), C_W4, 1, lgvspdData(TYPE_2).MaxRows-1, true, lgvspdData(TYPE_2).MaxRows, C_W4, "V")	' ���� ���� �հ� 
	
	ggoSpread.UpdateRow lgvspdData(TYPE_2).MaxRows
	
	ggoSpread.Source = lgvspdData(TYPE_3)
	
	Call FncSumSheet(lgvspdData(TYPE_3), C_W3, 1, 17, true, 18, C_W3, "V")	' ���� ���� �հ� 
	Call FncSumSheet(lgvspdData(TYPE_3), C_W4, 1, 17, true, 18, C_W4, "V")	' ���� ���� �հ� 
	Call FncSumSheet(lgvspdData(TYPE_3), C_W7, 1, 17, true, 18, C_W7, "V")	' ���� ���� �հ� 

	ggoSpread.UpdateRow 18

	' -- �հ�(50) = �Ұ�(49)
	dbl50(0) = GetGrid(TYPE_3, C_W3, 18)
	Call PutGrid(TYPE_3, C_W3, 19, dbl50(0))

	' -- �հ�(50) = �Ұ�(30) + �Ұ�(49)
	dbl50(1) = GetGrid(TYPE_2, C_W3, 15) + GetGrid(TYPE_3, C_W4, 18)
	Call PutGrid(TYPE_3, C_W4, 19, dbl50(1))

	' -- �հ�(50) = �Ұ�(30) + �Ұ�(49)
	dbl50(2) = GetGrid(TYPE_2, C_W4, 15) + GetGrid(TYPE_3, C_W7, 18)
	Call PutGrid(TYPE_3, C_W7, 19, dbl50(2))

	ggoSpread.UpdateRow 19

	
	' -- �������鼼���հ�(51 = 10 + 50)
	Call PutGrid(TYPE_3, C_W3, 20, dbl50(0))
	
	dblSum = dbl10(0) + dbl50(1)
	Call PutGrid(TYPE_3, C_W4, 20, dblSum)

	dblSum = dbl10(1) + dbl50(2)
	Call PutGrid(TYPE_3, C_W7, 20, dblSum)

	ggoSpread.UpdateRow 20
	
	lgBlnFlgChgValue = True
End Sub

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
'    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(Index)
   
    If lgvspdData(Index).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(Index)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	lgvspdData(Index).Row = Row
End Sub

Sub vspdData_ColWidthChange(Index, ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(Index, ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If lgvspdData(Index).MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus(Index)
    ggoSpread.Source = lgvspdData(Index)
    lgCurrGrid = Index
End Sub

Sub vspdData_MouseDown(Index, Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	lgCurrGrid = Index
	ggoSpread.Source = lgvspdData(Index)
End Sub    

Sub vspdData_ScriptDragDropBlock(Index, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos(Index)
End Sub

Sub vspdData_TopLeftChange(Index, ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if lgvspdData(Index).MaxRows < NewTop + VisibleRowCnt(lgvspdData(Index),NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(Index, ByVal Col, ByVal Row, Byval ButtonDown)

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
    If lgBlnFlgChgValue Then
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
    Call InitVariables		
    Call SpreadInitData											<%'Initializes local global variables%>
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
	For i = TYPE_1 To TYPE_3
		ggoSpread.Source = lgvspdData(i)
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
			Exit For
		End If
    Next
    
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

' ---------------------- ���ĳ� ���� -------------------------
Function  Verification()
	Dim iRow, iMaxRows, dblW3, dblW4, dblW7
	
	Verification = False

	With lgvspdData(TYPE_3)
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_W3	: dblW3 = UNICDbl(.value)
			.Col = C_W4	: dblW4 = UNICDbl(.value)
			.Col = C_W7	: dblW7 = UNICDbl(.value)
			
			If dblW3 + dblW4 < dblW7 Then
				Call DisplayMsgBox("WC0015", parent.VB_INFORMATION, GetGridText(TYPE_3, C_W7, 0), GetGridText(TYPE_3, C_W3, 0) & "��" & GetGridText(TYPE_3, C_W4, 0)) 
				.SetActiveCell C_W3, iRow
				Exit Function
			End If
		Next
	
	End With

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
	Call SpreadInitData
	
    Call SetToolbar("1100100000000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
  
	
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
Dim IntRetCD, iRow
	
	FncExit = False
    If lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
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
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	If lgIntFlgMode <> parent.OPMD_UMODE  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1101100000001111")										<%'��ư ���� ���� %>

		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100000000000111")										<%'��ư ���� ���� %>
		End If

	End If
	
	'Call SetSpreadLock(TYPE_1)
	'Call SetSpreadLock(TYPE_2)
	'Call SetSpreadLock(TYPE_3)
	
	'Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
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
    Dim strVal, strDel, sTmp
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    For i = TYPE_1 To TYPE_3	' ��ü �׸��� ���� 
    
		With lgvspdData(i)
	
			ggoSpread.Source = lgvspdData(i)
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
			' ----- 1��° �׸��� 
			For lRow = 1 To .MaxRows

    
				.Row = lRow	: sTmp = "" : .Col = 0
		    
				  ' ��� �׸��� ����Ÿ ����     
				  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
						For lCol = 1 To lMaxCols
							Select Case lCol
								'Case C_W31
								'	.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
								Case Else
									.Col = lCol : sTmp = sTmp & Trim(.Value) &  Parent.gColSep
							End Select
						Next
						sTmp = sTmp & Trim(.Text) &  Parent.gRowSep
				  End If  

				.Col = 0
				Select Case .Text
					Case  ggoSpread.InsertFlag                                      '��: Insert
				                                       strVal = strVal & "C"  &  Parent.gColSep & sTmp
				    Case  ggoSpread.UpdateFlag                                      '��: Update
				                                       strVal = strVal & "U"  &  Parent.gColSep & sTmp
				    Case  ggoSpread.DeleteFlag                                      '��: Update
				                                       strDel = strDel & "D"  &  Parent.gColSep & sTmp
				End Select

			Next
							   
		End With

		
	Next
		
	Frm1.txtSpread.value      = strDel & strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow
	
	Call InitVariables
	
	For iRow = TYPE_1 To TYPE_3
	
		lgvspdData(iRow).MaxRows = 0
		ggoSpread.Source = lgvspdData(iRow)
		ggoSpread.ClearSpreadData
	Next
    	
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">�ݾ׺ҷ�����</A></TD>
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
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=100%>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;1. �����Ѽ� �������� �������鼼��					
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="30%">
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=455 tag="23" TITLE="SPREAD" id=vspdData0 Index=0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">&nbsp;2. �����Ѽ� ������ �������鼼��							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="30%">
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=250 tag="23" TITLE="SPREAD" id=vspdData1 Index=1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="35%">
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=370 tag="23" TITLE="SPREAD" id=vspdData2 Index=2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
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
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

