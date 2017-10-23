<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ÿ�������� 
'*  3. Program ID           : w7105mA1
'*  4. Program Name         : w7105mA1.asp
'*  5. Program Desc         : ��50ȣ �ں��ݰ� ��������������(��)
'*  6. Modified date(First) : 2005/02/21
'*  7. Modified date(Last)  : 2005/02/21
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/JpQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "w7105mA1"
Const BIZ_PGM_ID		= "w7105mB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "w7105mB2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		 = "w7105OA1"

' -- �׸��� �÷� ���� 
Dim C_H1
Dim C_H2
Dim C_W_CD
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
DIm C_W_DESC

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
Dim C_20
Dim C_21
Dim C_22
Dim C_23
Dim C_30
Dim C_31

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgFISC_START_DT, lgFISC_END_DT 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_H1		= 1
	C_H2		= 2
	C_W1		= 3
	C_W_CD		= 4
	C_W2		= 5
	C_W3		= 6
	C_W4		= 7
	C_W5		= 8
	C_W_DESC	= 9
	
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
	C_20 = 14
	C_21 = 15
	C_22 = 16
	C_23 = 17
	C_30 = 18
	C_31 = 19
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

Function OpenAdItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�������� �˾�"					<%' �˾� ��Ī %>
	arrParam(1) = "TB_ADJUST_ITEM"					<%' TABLE ��Ī %>
	

		frm1.vspdData.Col = C_W1
	    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>

	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = " USE_YN = '1' "							<%' Where Condition%>
	arrParam(5) = "��������"						<%' �����ʵ��� �� ��Ī %>
	
    arrField(0) = "ITEM_CD"					<%' Field��(0)%>
    arrField(1) = "ITEM_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "��������"					<%' Header��(0)%>
    arrHeader(1) = "�����"						<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAdItem(arrRet)
	End If	
	
End Function

Function SetAdItem(byval arrRet)
    With frm1
			.vspdData.Col = C_W1
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_W1_NM
			.vspdData.Text = arrRet(1)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			lgBlnFlgChgValue = True
	End With
End Function

Function GetAdItem(ByVal pCol, ByVal pRow)
	Dim arrRet(2), sWhere, bRet

	If pCol = C_W1 Then
		sWhere = " ITEM_CD LIKE '%"
	ElseIf pCol = C_W1_NM Then
		sWhere = " ITEM_NM LIKE '%"
	Else
		Exit Function
	End If
	
	With frm1.vspdData
		.Col = pCol
		If .Text <> "" Then
			sWhere = sWhere & .Text & "%' "		<%' Code Condition%>
		
			bRet = CommonQueryRs("top 1 ITEM_CD,ITEM_NM"," TB_ADJUST_ITEM ",sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			arrRet(0) = Replace(lgF0, chr(11), "")
			arrRet(1) = Replace(lgF1, chr(11), "")
		Else
			arrRet(0) = ""
			arrRet(1) = ""
		End If
	End With
	
	Call SetAdItem(arrRet)
	
End Function


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

		.MaxCols = C_W_DESC + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'����� 2�ٷ�    
	    .ColHeaderRows = 2
	    
		ggoSpread.SSSetEdit     C_H1,	     ""		, 5,,,50,1	' 
		ggoSpread.SSSetEdit  	C_H2,		 ""		, 20,,,50,1	' 
		ggoSpread.SSSetEdit		C_W1,		"",			10,,,50,1
		ggoSpread.SSSetEdit		C_W_CD,		"�ڵ�",			5,2,,10,1
		ggoSpread.SSSetFloat	C_W2,		"(2)�����ܾ�",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	    ggoSpread.SSSetFloat	C_W3,		"(3)����",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	    ggoSpread.SSSetFloat	C_W4,		"(4)����",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
		ggoSpread.SSSetFloat	C_W5,		"�⸻�ܾ�" & vbCrLf & "(�ͱ�������)",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
		ggoSpread.SSSetEdit		C_W_DESC,	"���",	15,,,100,1

	    ret = .AddCellSpan(0, -1000, 1, 2)
	    ret = .AddCellSpan(C_H1, -1000, 3, 2)
	    ret = .AddCellSpan(C_W_CD, -1000, 1, 2)
	    ret = .AddCellSpan(C_W2, -1000, 1, 2)
	    ret = .AddCellSpan(C_W3, -1000, 2, 1)
	    ret = .AddCellSpan(C_W5, -1000, 1, 2)
	    ret = .AddCellSpan(C_W_DESC, -1000, 1, 2)

	    ' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_H1
		.Text = "(1)���� �Ǵ� ����"
		.Col = C_W3
		.Text = "���������"
	
		' �ι�° ��� ��� ���� 
		.Row = -999
		.Col = C_W3
		.Text = "(3)����"
		.Col = C_W4
		.Text = "(4)����"

		.rowheight(-999) = 15	' ���� ������ 
	

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	    Call ggoSpread.SSSetColHidden(C_03 , C_03 , True)
	    ' Call ggoSpread.SSSetRowHidden(10 , 14 , True)
		
		Call InitSpreadRow()

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

		ggoSpread.SpreadLock C_H1,   -1, C_W_CD
		ggoSpread.SpreadUnLock C_W2, -1, C_W_DESC	' ��ü ���� 
		
		ggoSpread.SpreadLock C_W5,   C_01, C_W5, C_31
		ggoSpread.SpreadLock C_W5,   C_01, C_W5, C_31
		ggoSpread.SpreadLock C_W2,   C_20, C_W5, C_21
		ggoSpread.SpreadLock C_W2,   C_30, C_W5, C_30
		ggoSpread.SpreadLock C_W2,   C_31, C_W5, C_31
		ggoSpread.SpreadLock C_W3,   C_10, C_W3, C_13

		' ���� �޾ƾ� �� ���� 
		ggoSpread.SpreadUnLock C_W1, C_02, C_W1, C_05
		ggoSpread.SpreadUnLock C_W1, C_06, C_W1, C_09
		
		ggoSpread.SpreadUnLock C_H2, C_05, C_H2, C_09
		

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
	Dim ret

	With Frm1.vspdData
				
		ggoSpread.Source = Frm1.vspdData
		'patch version
		If .MaxRows = 0 Then	.MaxRows = C_31

	    ret = .AddCellSpan(C_H1, C_01, 1, 14)
	    ret = .AddCellSpan(C_H1, C_21, 2, 1)
	    ret = .AddCellSpan(C_H1, C_22, 1, 3)
	    ret = .AddCellSpan(C_H1, C_31, 2, 1)

	    ' ù��° ��� ��� ���� 
		.Col = C_H1
		.Row = C_01	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "��" & VbCrlf & "��" & vbCrLf & "��" & vbCrLf & " " & vbCrLf & "��" & vbCrLf & " " & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & " " & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & " " & vbCrLf & "��" & vbCrLf & "��"
		.Row = C_22	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "����" & vbCrLf & "�̰�" & vbCrLf & "���" & vbCrLf & "�μ�" & vbCrLf & "��"
		.ColWidth(C_H1) = 5
	
		.Row = C_21
		.Text = "6.�ں��ݰ������ݰ�꼭(��)��(��)"
		.Row = C_31
		.Text = "10.��������(����������)"
		
		' �ι�° ��� ��� ���� 
		.Col = C_H2
		.Row = C_01	:	.TypeHAlign = 0	:			.Text = "1. �ں���"
		.Row = C_02	:	.TypeHAlign = 0	:			.Text = "2. �ں��׿���"
		.Row = C_03	:	.TypeHAlign = 0	:			.Text = "3. �����׿���"
		.Row = C_04	:	.TypeHAlign = 0	:			.Text = "4. �ں�����"
		.Row = C_05	:	.TypeHAlign = 0	:			.Text = ""
		.Row = C_06	:	.TypeHAlign = 0	:			.Text = ""
		.Row = C_07	:	.TypeHAlign = 0	:			.Text = ""
		.Row = C_08	:	.TypeHAlign = 0	:			.Text =  ""
		.Row = C_09	:	.TypeHAlign = 0	:			.Text =  ""
		.Row = C_10	:	.TypeHAlign = 0	:			.Text =  ""
		.Row = C_11	:	.TypeHAlign = 0	:			.Text = ""
		.Row = C_12	:	.TypeHAlign = 0	:			.Text =  ""
		.Row = C_13	:	.TypeHAlign = 0	:			.Text =  ""
		.Row = C_20	:	.TypeHAlign = 0	:			.Text = "5.   �� "
		
		.Row = C_22	:	.TypeHAlign = 0	:			.Text = "7. ���μ�"
		.Row = C_23	:	.TypeHAlign = 0	:			.Text = "8. �ֹμ�"
		.Row = C_30	:	.TypeHAlign = 0	:			.Text = "9.    �� (��)"
		.rowheight(-1) = 15	' ���� ������ 
		.ColWidth(C_H2) = 20
		
		' �⺻�ڵ尪�Է��ϱ� 
		.Col = C_W_CD
		.Row = C_01	:	.Text = "01"
		.Row = C_02	:	.Text = "02"
		.Row = C_03	:	.Text = "14"
		.Row = C_04	:	.Text = "15"
		.Row = C_05	:	.Text = "16"
		.Row = C_06	:	.Text = "17"
		.Row = C_07	:	.Text = ""
		.Row = C_08	:	.Text = ""
		.Row = C_09	:	.Text = ""
		.Row = C_10	:	.Text = ""
		.Row = C_11	:	.Text = ""
		.Row = C_12	:	.Text = ""
		.Row = C_13	:	.Text = ""
		.Row = C_20	:	.Text = "20"
		.Row = C_21	:	.Text = "21"
		.Row = C_22	:	.Text = "22"
		.Row = C_23	:	.Text = "23"
		.Row = C_30	:	.Text = "30"
		.Row = C_31	:	.Text = "31"
		.Row = C_10	: .RowHidden = True
		.Row = C_11	: .RowHidden = True
		.Row = C_12	: .RowHidden = True
		.Row = C_13	: .RowHidden = True

	End With 

End Sub


'============================== ���۷��� �Լ�  ========================================
Function GetRef_old()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg
      wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	ggoSpread.Source = Frm1.vspdData
    'ggoSpread.ClearSpreadData
			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
End Function


'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	 dim istrSql
	 dim arrW1
	 dim arrW2
	 dim arrW3
	 dim arrW4

	dim w1,w2,w3,w4,w5,w6
	 
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg
      wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	 '
	 '�ں����߰�,�ں��׿��� �߰� 2007.05.30
	with frm1.vspdData

 
	if  CommonQueryRs("W1,W2,W3,W4,W5,W6","dbo.ufn_TB_7105ma1_GetRef_200703('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then

	   w1 = cDbl(REPLACE(lgF0, chr(11),""))
	   w2 = cDbl(REPLACE(lgF1, chr(11),""))
	   w3 = cDbl(REPLACE(lgF2, chr(11),""))
	   w4 = cDbl(REPLACE(lgF3, chr(11),""))
	   w5 = cDbl(REPLACE(lgF4, chr(11),""))
	   w6 = cDbl(REPLACE(lgF5, chr(11),""))
	   
	   
	   .Row=C_01 : .Col = 0: .value =ggoSpread.UpdateFlag 
       .Row=C_01 : .Col = C_W2 : .value = w1
       if (w2-w1) > 0 then
		 .Row=C_01 : .Col = C_W4 : .value =w2-w1
       else
         .Row=C_01 : .Col = C_W3 : .value =w2-w1
       end if
      
       .Row=C_02 : .Col = 0: .value =ggoSpread.UpdateFlag 
       .Col = C_W2 : .value = w3
        
       if (w4-w3) > 0 then
		 .Col = C_W4 : .value = w4-w3
       else
         .Col = C_W3 : .value = w4-w3
       end if
       
       
       .Row=C_03 : .Col = 0: .value =ggoSpread.UpdateFlag 
       .Col = C_W2 : .value = w5
        
       if ( w6-w5) > 0 then
		 .Col = C_W4 : .value = w6-w5
       else
         .Col = C_W3 : .value = w6-w5
       end if
      
	else
	end if
	

	 'call CommonQueryRs("W1, W2, W3, W4, W5, W_DESC","dbo.ufn_TB_50A_GetRef"&C_REVISION_YM&"('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	call CommonQueryRs("W2,W3,W4,W5","TB_50B", "CO_CD='" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "' AND SEQ_NO='999999'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 = "" Then	 Exit Function
	lgBlnFlgChgValue= TRUE	
	    arrW1 = REPLACE(lgF0, chr(11),"")
	    arrW2 = REPLACE(lgF1, chr(11),"")
		arrW3 = REPLACE(lgF2, chr(11),"")
		arrW4 = REPLACE(lgF3, chr(11),"")	

	
	
	   .Row=C_21 : .Col = 0: .value =ggoSpread.UpdateFlag 
       .Row=C_21 : .Col = C_W2 : .value =arrW1
	   .Row=C_21 : .Col = C_W3: .value =arrW2
	   .Row=C_21 : .Col = C_W4 : .value =arrW3
	   .Row=C_21 : .Col = C_W5 : .value =arrW4

	End with   
	call GetRefOk
End Function

Function GetRefOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr, iSeqNo, iLastRow, iRow
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	If Frm1.vspdData.MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
	    lgIntFlgMode = parent.OPMD_CMODE
		

	    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
'		Call ChangeRowFlg(frm1.vspdData)
		For iRow = 1 To Frm1.vspdData.MaxRows
			Call Fn_GridCalc(C_W2, iRow)
		Next
	End If
	
	'Call InitSpreadRow()
	'Call SetSpreadLock
'	Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
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

'	Call DBQuery()
	Call FncQuery
     

End Sub

'============================================  ����� �Լ�  ====================================
Function Fn_GridCalc(ByVal pCol, ByVal pRow)
	Dim iRow, dblSum
	Dim dblW2, dblW3, dblW4, dblW5
	
	If Frm1.vspdData.MaxRows <= 0 Then Exit Function

    ggoSpread.Source = Frm1.vspdData
    iRow = pRow

	With Frm1.vspdData
		' (5) = (2) - (3) + (4)
		If pRow = 0 Then iRow = Frm1.vspdData.ActiveRow
		If iRow <= C_09 Then
			.Row = iRow	:	.Col = C_W2	:	dblW2 = UNICdbl(.Text)
			.Row = iRow	:	.Col = C_W3	:	dblW3 = UNICdbl(.Text)
			.Row = iRow	:	.Col = C_W4	:	dblW4 = UNICdbl(.Text)
			.Row = iRow	:	.Col = C_W5	:	dblW5 = dblW2 - dblW3 + dblW4
			.Text = dblW5
		ElseIf iRow <= C_13 Then
			.Row = iRow	:	.Col = C_W2	:	dblW2 = UNICdbl(.Text)
			.Row = iRow	:	.Col = C_W3	:	.Text = dblW2
			.Row = iRow	:	.Col = C_W4	:	dblW4 = UNICdbl(.Text)
			.Row = iRow	:	.Col = C_W5	:	.Text = dblW4
		ElseIf iRow = C_22 Or iRow = C_23 Then
			.Row = iRow	:	.Col = C_W2	:	dblW2 = UNICdbl(.Text)
			.Row = iRow	:	.Col = C_W3	:	dblW3 = UNICdbl(.Text)
			.Row = iRow	:	.Col = C_W4	:	dblW4 = UNICdbl(.Text)
			.Row = iRow	:	.Col = C_W5	:	dblW5 = dblW2 - dblW3 + dblW4
			.Text = dblW5
		End If

		dblSum = FncSumSheet(Frm1.vspdData, C_W2, 1, C_13, true, C_20, C_W2, "V")	' �հ� 
		dblSum = FncSumSheet(Frm1.vspdData, C_W3, 1, C_13, true, C_20, C_W3, "V")	' �հ� 
		dblSum = FncSumSheet(Frm1.vspdData, C_W4, 1, C_13, true, C_20, C_W4, "V")	' �հ� 
		dblSum = FncSumSheet(Frm1.vspdData, C_W5, 1, C_13, true, C_20, C_W5, "V")	' �հ� 
		
		dblSum = FncSumSheet(Frm1.vspdData, C_W2, C_22, C_23, true, C_30, C_W2, "V")	' �հ� 
		dblSum = FncSumSheet(Frm1.vspdData, C_W3, C_22, C_23, true, C_30, C_W3, "V")	' �հ� 
		dblSum = FncSumSheet(Frm1.vspdData, C_W4, C_22, C_23, true, C_30, C_W4, "V")	' �հ� 
		dblSum = FncSumSheet(Frm1.vspdData, C_W5, C_22, C_23, true, C_30, C_W5, "V")	' �հ� 
		
		.Row = C_20
		.Col = C_W2	:	dblW2 = UNICdbl(.Text)
		.Col = C_W3	:	dblW3 = UNICdbl(.Text)
		.Col = C_W4	:	dblW4 = UNICdbl(.Text)
		.Col = C_W5	:	dblW5 = UNICdbl(.Text)
	
		.Row = C_21
		.Col = C_W2	:	dblW2 = dblW2 + UNICdbl(.Text)
		.Col = C_W3	:	dblW3 = dblW3 + UNICdbl(.Text)
		.Col = C_W4	:	dblW4 = dblW4 + UNICdbl(.Text)
		.Col = C_W5	:	dblW5 = dblW5 + UNICdbl(.Text)
	
		.Row = C_30
		.Col = C_W2	:	dblW2 = dblW2 - UNICdbl(.Text)
		.Col = C_W3	:	dblW3 = dblW3 - UNICdbl(.Text)
		.Col = C_W4	:	dblW4 = dblW4 - UNICdbl(.Text)
		.Col = C_W5	:	dblW5 = dblW5 - UNICdbl(.Text)
	
		.Row = C_31
		.Col = C_W2	:	.Text = dblW2
		.Col = C_W3	:	.Text = dblW3
		.Col = C_W4	:	.Text = dblW4
		.Col = C_W5	:	.Text = dblW5
			
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
	if Col=c_w5 then exit sub

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.UpdateRow Row

	' --- �߰��� �κ� 
	Call Fn_GridCalc(Col, Row)

'	Call GetAdItem(Col, Row)			' �������� ��������		ǥ�ؿ��� ��÷� ��Ī�� �Ȱ����´�.

	
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


    Call PageLink("vspdData", Col, Row, "S")
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
    Call InitSpreadRow
    Call SetSpreadLock
    Call InitVariables
    Call InitData

    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
    lgIntFlgMode = parent.OPMD_CMODE

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
			
			Call Fn_GridCalc(C_W2, iActiveRow + 1)
    
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

	Call Fn_GridCalc(0,0)

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

	Call Fn_GridCalc(0,0)

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
    If ggoSpread.SSCheckChange = True Then
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

		
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr,i
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
	'for i=1 to
	SetSpreadValue frm1.vspdData,C_H2,C_05,GetSpreadText(Frm1.vspdData,C_W1,C_05,"X","X"),"","" 
	SetSpreadValue frm1.vspdData,C_H2,C_06,GetSpreadText(Frm1.vspdData,C_W1,C_06,"X","X"),"","" 
	SetSpreadValue frm1.vspdData,C_H2,C_07,GetSpreadText(Frm1.vspdData,C_W1,C_07,"X","X"),"","" 
	SetSpreadValue frm1.vspdData,C_H2,C_08,GetSpreadText(Frm1.vspdData,C_W1,C_08,"X","X"),"",""
	SetSpreadValue frm1.vspdData,C_H2,C_09,GetSpreadText(Frm1.vspdData,C_W1,C_09,"X","X"),"",""	
	
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
    
	dim tmpSeq,tmpSub
	
	tmpSub=1		
			
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
	                                              strVal = strVal & "C"  &  Parent.gColSep
		       Case  ggoSpread.DeleteFlag                                      '��: Delete
		                                          strVal = strVal & "D"  &  Parent.gColSep
		       Case Else
		                                          strVal = strVal & "C"  &  Parent.gColSep
	       End Select
	       
		  ' ��� �׸��� ����Ÿ ����     
'		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
			
			
			if lRow=C_05 or lRow=C_06 or lRow=C_07 or lRow=C_08 or lRow=C_09 then
				SetSpreadValue frm1.vspdData,C_W1,lRow,GetSpreadText(Frm1.vspdData,C_H2,lRow,"X","X"),"","" 
            end if 		
          
			if GetSpreadText(Frm1.vspdData,C_W_CD,lRow,"X","X")="" then	
				tmpSeq=GetSpreadText(Frm1.vspdData,C_W_CD,6,"X","X") & tmpSub
				SetSpreadValue frm1.vspdData,C_W_CD,lRow,tmpSeq,"","" 
				tmpSub = tmpSub + 1
            end if			
				For lCol = 1 To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
			'end if	
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
										<TD HEIGHT="10">&nbsp; 1. �ں��ݰ� ������ ��꼭 
										</TD>
									</TR>
									<TR>
										<TD >
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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