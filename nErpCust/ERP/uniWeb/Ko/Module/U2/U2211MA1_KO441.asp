<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : S3322MA1_KO412
'*  4. Program Name         : ǰ�Ǽ���������(S)
'*  5. Program Desc         : ǰ�Ǽ���������(S)
'*  6. Component List       : 
'*  7. Modified date(First) : 2007/07/04
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee wol san
'* 10. Modifier (Last)      : Lee Ho Jun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" --> 

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit											

Const BIZ_PGM_ID = "U2211MB1_KO441.asp"
'Const BIZ_PGM_REG_ID = "U2211MA1_KO441"
<!-- #Include file="../../inc/lgvariables.inc" -->	

'@Grid_Column
Dim C_H_BP_CD		'��ǰó�ڵ�
Dim C_H_DLVY_NO		'��ǰ������ȣ
Dim C_H_DOCUMENT_NO			
Dim C_H_TITLE
Dim C_H_INS_USER
Dim C_H_INS_DT
Dim C_H_DOCUMENT_ABBR


Dim C_HH_DLVY_NO
Dim C_HH_PO_NO
Dim C_HH_PO_SEQ_NO
Dim C_HH_BP_CD 
Dim C_HH_ITEM_CD
Dim C_HH_ITEM_NM
Dim C_HH_SPEC
Dim C_HH_BASIC_UNIT               
Dim C_HH_PLAN_DVRY_DT              
Dim C_HH_PLAN_DVRY_QTY             
Dim C_HH_D_BP_CD               
Dim C_HH_SL_NM               
Dim C_HH_SPLIT_SEQ_NO              
Dim C_HH_PO_UNIT               
Dim C_HH_TRACKING_NO               
                   
                   
                   
'@Global_Var       
Dim lgSortKey1     
Dim IsOpenPop      
Dim lgitem_lvl     
Dim EndDate, StartDate
Dim lgAcct_item_cd
Dim lgAcct_kind_cd
Dim LocSvrDate

LocSvrDate = "<%=GetSvrDate%>"

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------

StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
	lgPageNo = ""
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	
	frm1.txtDvFrDt.text = UniConvDateAToB(UNIDateAdd ("D", -7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtDvToDt.text = UniConvDateAToB(UNIDateAdd ("D", 7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)		
	Call SetToolBar("110000010011111")				'��ư ���� ���� 

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	Select Case UCase(pvSpdNo)
		Case "A"
	
			With frm1.vspdData1
				.ReDraw = false
		
				ggoSpread.Source = frm1.vspdData1
		        ggoSpread.Spreadinit "V20050103",, parent.gAllowDragDropSpread
			
			   .MaxCols = C_HH_TRACKING_NO + 1
			   '.MaxRows = 0
		
				Call GetSpreadColumnPos("A")
				
				ggoSpread.SSSetEdit		C_HH_DLVY_NO		, "��ǰ������ȣ", 15
				ggoSpread.SSSetEdit		C_HH_PO_NO          , "���ֹ�ȣ",       15
				ggoSpread.SSSetEdit		C_HH_PO_SEQ_NO      , "���",	  		10
				ggoSpread.SSSetEdit		C_HH_BP_CD          , "��ü",	  		12
				ggoSpread.SSSetEdit		C_HH_ITEM_CD        , "ǰ��",   		12
				ggoSpread.SSSetEdit		C_HH_ITEM_NM        , "ǰ���", 		12
				ggoSpread.SSSetEdit		C_HH_SPEC           , "�԰�",   		20
				ggoSpread.SSSetEdit     C_HH_BASIC_UNIT     , "����",   		12
                ggoSpread.SSSetDate     C_HH_PLAN_DVRY_DT   , "��ǰ��������"	,13    ,2                 ,parent.gDateFormat   ,-1
                ggoSpread.SSSetFloat    C_HH_PLAN_DVRY_QTY  , "��ǰ��������"	,15    , Parent.ggQtyNo   ,ggStrIntegeralPart 	 ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec ,       ,      ,"P" 
                ggoSpread.SSSetEdit     C_HH_D_BP_CD        , "��ǰâ��",    	12
                ggoSpread.SSSetEdit     C_HH_SL_NM          , "��ǰâ���",  	15
                ggoSpread.SSSetEdit     C_HH_SPLIT_SEQ_NO   , "���ҹ�ȣ",    	10
                ggoSpread.SSSetEdit     C_HH_PO_UNIT        , "����",        	10
                ggoSpread.SSSetEdit     C_HH_TRACKING_NO  	, "Tracking No", 	10

		
				call ggoSpread.SSSetColHidden(C_HH_BP_CD,C_HH_BP_CD,True)
'20080604::hanc::hidden ����				call ggoSpread.SSSetColHidden(C_HH_DLVY_NO,C_HH_DLVY_NO,True)
				call ggoSpread.SSSetColHidden(C_HH_TRACKING_NO,C_HH_TRACKING_NO,True)
				Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)
					
				.ReDraw = True
		    End With
	
    		Call SetSpreadLock(pvSpdNo)
    				
		Case "B"
		
			With frm1.vspdData
				.ReDraw = false
		
				ggoSpread.Source = frm1.vspdData
		        ggoSpread.Spreadinit "V20050103",, parent.gAllowDragDropSpread
		
			   .MaxCols = C_H_DOCUMENT_ABBR + 1
			   '.MaxRows = 0
		
				Call GetSpreadColumnPos("B")
				
				ggoSpread.SSSetEdit		C_H_BP_CD,			"��ü", 10
				ggoSpread.SSSetEdit		C_H_DLVY_NO,		"��ǰ������ȣ", 15
				ggoSpread.SSSetEdit		C_H_DOCUMENT_NO,	"������ȣ",	13
				ggoSpread.SSSetEdit		C_H_TITLE,			"����",	30
				ggoSpread.SSSetEdit		C_H_INS_USER,		"�����", 15
				ggoSpread.SSSetEdit		C_H_INS_DT,			"�����", 15
				ggoSpread.SSSetEdit		C_H_DOCUMENT_ABBR,	"��༳��", 50,,,100
				
				Call ggoSpread.MakePairsColumn(C_H_DOCUMENT_NO, C_H_TITLE, "1")
		
				call ggoSpread.SSSetColHidden(C_H_BP_CD,C_H_BP_CD,True)
				call ggoSpread.SSSetColHidden(C_H_DLVY_NO,C_H_DLVY_NO,True)
				call ggoSpread.SSSetColHidden(C_H_DOCUMENT_NO,C_H_DOCUMENT_NO,True)
				Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)
					
				.ReDraw = True
		    End With
	
    		Call SetSpreadLock(pvSpdNo)
    		
    End Select
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	Select Case UCase(pvSpdNo)
		Case "A"
		    With frm1.vspdData1
				.ReDraw = False
		    
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.SpreadLock		-1,			-1
		   		
				.ReDraw = True
		    End With  
		Case "B"
		    With frm1.vspdData
				.ReDraw = False
		    
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadLock		-1,			-1
		   		
				.ReDraw = True
		    End With  
	End Select		    		
	
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData
		.ReDraw = False

  		ggoSpread.Source = frm1.vspdData
  
   		ggoSpread.SpreadUnLock		1, pvStartRow, ,pvEndRow
		ggoSpread.SSSetRequired  C_H_DOCUMENT_ABBR,			pvStartRow,	pvEndRow

	    .ReDraw = True
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)

	Select Case pvSpdNo
		Case "A"
			C_HH_DLVY_NO		= 1
			C_HH_PO_NO          = 2
			C_HH_PO_SEQ_NO      = 3
			C_HH_BP_CD          = 4
			C_HH_ITEM_CD        = 5
			C_HH_ITEM_NM        = 6
			C_HH_SPEC           = 7
			C_HH_BASIC_UNIT     = 8      
			C_HH_PLAN_DVRY_DT   = 9      
			C_HH_PLAN_DVRY_QTY  = 10      
			C_HH_D_BP_CD        = 11 
			C_HH_SL_NM          = 12
			C_HH_SPLIT_SEQ_NO   = 13      
			C_HH_PO_UNIT        = 14  
			C_HH_TRACKING_NO  	= 15
			
		Case "B"
			C_H_BP_CD			= 1
			C_H_DLVY_NO  		= 2
			C_H_DOCUMENT_NO		= 3
			C_H_TITLE			= 4
			C_H_INS_USER      	= 5
			C_H_INS_DT			= 6
			C_H_DOCUMENT_ABBR	= 7
			
	End Select 

End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
 
    Select Case UCase(pvSpdNo)
    	
		Case "A"
   		
			ggoSpread.Source = frm1.vspdData1
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_HH_DLVY_NO		=	iCurColumnPos(1)
			C_HH_PO_NO          =	iCurColumnPos(2)
			C_HH_PO_SEQ_NO      =	iCurColumnPos(3)
			C_HH_BP_CD          =	iCurColumnPos(4)
			C_HH_ITEM_CD        =	iCurColumnPos(5)
			C_HH_ITEM_NM        =	iCurColumnPos(6)
			C_HH_SPEC           =	iCurColumnPos(7)	    	
    	    C_HH_BASIC_UNIT     =	iCurColumnPos(8)
    	    C_HH_PLAN_DVRY_DT   =	iCurColumnPos(9)
    	    C_HH_PLAN_DVRY_QTY  =	iCurColumnPos(10)
    	    C_HH_D_BP_CD        =	iCurColumnPos(11)
    	    C_HH_SL_NM          =	iCurColumnPos(12)
    	    C_HH_SPLIT_SEQ_NO   =	iCurColumnPos(13)
    	    C_HH_PO_UNIT        =	iCurColumnPos(14)
    	    C_HH_TRACKING_NO  	=	iCurColumnPos(15)
    	    
		Case "B"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_H_BP_CD			=	iCurColumnPos(1)
			C_H_DLVY_NO			=	iCurColumnPos(2)
			C_H_DOCUMENT_NO		=	iCurColumnPos(3)
			C_H_TITLE			=	iCurColumnPos(4)
			C_H_INS_USER		=	iCurColumnPos(5)
			C_H_INS_DT			=	iCurColumnPos(6)
			C_H_DOCUMENT_ABBR	=	iCurColumnPos(7)	
	End Select    
End Sub


'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()	'###�׸��� ������ ���Ǻκ�###
	Call LoadInfTB19029                                                         'Load table , B_numeriC_H_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
	Call InitVariables
	Call SetDefaultVal

	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")
	
	'Call InitComboBox
	Call SetToolbar("11000000000111")
	'Call dbQuery()
	
End Sub

'==========================================  2.2.6 InitComboBox()  ========================================
' Name : InitComboBox()
' Desc : Combo Display
'==========================================================================================================
Sub InitComboBox()

    Dim strCboCd
    Dim strCboNm

	'// ����
	Call CommonQueryRs(" UD_MINOR_CD,UD_MINOR_NM "," B_USER_DEFINED_MINOR ", " UD_MAJOR_CD = " & FilterVar("SX006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_H_DOCUMENT_NO
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_H_TITLE
   
	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = Col
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)        '  <------����� ǥ�� ���� 

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_H_DOCUMENT_NO         ' �ý��۱���
			intIndex = .value
			.col = C_H_TITLE
			.value = intindex					
		Next	
	End With

End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###�׸��� ������ ���Ǻκ�###
	
	Dim sDhBpCd, sDhDlvyNo, sDhDocumentNo
	Dim strval
	Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
 	
 	gMouseClickStatus = "SPC" 
 	
 	'sDhBpCd		= frm1.txtBpCd.value  
 	'sDhDlvyNo	= frm1.txtDlvyNo.value
	
 	with frm1.vspddata
		.Row = Row
 		.Col = C_H_BP_CD
 		'hdnBpCD = .text 
 		sDhBpCd = .text
 		
		.Row = Row
 		.Col = C_H_DLVY_NO
 		'hdnDlvyNo = .text  		
 		sDhDlvyNo = .text
 	End With

 	sDhDocumentNo = GetSpreadText(frm1.vspdData,C_H_DOCUMENT_NO,Row,"X","X")
 	
 	Set gActiveSpdSheet = frm1.vspdData
 	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		Exit Sub
 	End If

' 	with frm1
' 	    strVal = BIZ_PGM_ID & "?txtMode=view"
'	    strVal = strVal & "&dlvy_no=" & sDhDlvyNo
'	    strVal = strVal & "&Document_no=" & sDhDocumentNo
'	    strVal = strVal & "&lgPageNo="	 & lgPageNo						'��: Next key tag 
'	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
'	End with 
''MSGBOX strVal
''MSGBOX "U2211RA1_KO441.asp?txtBpCd=" & sDhBpCd & "&DLVY_NO=" & sDhDlvyNo & "&document_no=" & sDhDocumentNo
	  MyBizASP1.location.href = "U2211RA1_KO441.asp?txtBpCd=" & sDhBpCd & "&DLVY_NO=" & sDhDlvyNo & "&document_no=" & sDhDocumentNo

 	//Call RunMyBizASP(MyBizASP, strVal)	//zerry �߾ȵ�..�����Ұ�.	

End Sub



'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)	'###�׸��� ������ ���Ǻκ�###

	Dim sDhBpCd, sDhDlvyNo, sDhDocumentNo
	Dim strval
	Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
 	
 	gMouseClickStatus = "SPC" 
 	
	If frm1.vspdData1.MaxRows = 0 Then Exit Sub

 	Call DbQueryDtl()


End Sub





'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     

    If OldLeft <> NewLeft Then Exit Sub
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '��: ������ üũ 
		If Trim(lgPageNo) = "" Then Exit Sub
		If lgPageNo > 0 Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery1 = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End If
End Sub

'========================================================================================
' Function Name : vspdData_ButtonClicked
' Function Desc : �˾���ư ���ý� 
'========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

  
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
'    Select Case Col
'        Case C_H_Cost
'            Call EditModeCheck(frm1.vspdData, Row, C_H_Curr, C_H_Cost,    "C" ,"I", Mode, "X", "X")
'    End Select
End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("A")
    Call InitSpreadSheet("B")
    'Call InitData()
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'======================================================================================================
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================
Sub txtDlvyNo_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
  
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() '###�׸��� ������ ���Ǻκ�###
    Dim IntRetCD     
    FncQuery = False                                                        
	
	If ggoSpread.SSCheckChange = True Then 'lgBlnFlgChgValue = True Or lgBtnClkFlg = True Or 
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")		'��: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	ggoSpread.Source = frm1.vspdData1	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData
    Call InitVariables															'Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
  

    If Not chkField(Document, "1") Then         
       Exit Function
    End If
    
    If ValidDateCheck(frm1.txtDvFrDt, frm1.txtDvToDt) = False Then Exit Function

	If DbQuery = False then	Exit Function
	      
    FncQuery = True
    	
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData

    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If

    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal

    FncNew = True  
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    FncSave = False                                                         

    If frm1.vspdData.maxrows < 1 then exit function    

	'-----------------------
    'Precheck area
    '-----------------------
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           
        Exit Function
    End If
 
    '----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then Exit Function

   	Call SetToolbar("10000000000111")
   	
'    If CompareDateByFormat(frm1.txtFromInsrtDt.text,frm1.txtToInsrtDt.text,frm1.txtFromInsrtDt.Alt,frm1.txtToInsrtDt.Alt, _
'        	               "970024",frm1.txtFromInsrtDt.UserDefinedFormat,parent.gComDateType, true) = False Then
'	   frm1.txtFromInsrtDt.focus
'	   Exit Function
'	End If
    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then Exit Function
'msgbox "FncSave(50)"	  
    FncSave = True                                                       
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    With frm1.vspdData
		If .maxrows < 1 then exit function
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor .ActiveRow, .ActiveRow
		.Row = .ActiveRow
		.Col = C_H_DOCUMENT_ABBR
		.Text = ""
	
		.ReDraw = True
	End With
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

    Dim iDx

	On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCancel = False                                                             '��: Processing is NG

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
     Frm1.vspdData.Row = frm1.vspdData.ActiveRow
     
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '��: Processing is OK
    End If
    

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	
	'On Error Resume Next
	
	FncInsertRow = False
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
   
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End If 
	
    With frm1.vspdData	
		.ReDraw = False
		.focus
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow .ActiveRow, imRow
		SetSpreadColor .ActiveRow, .ActiveRow + imRow - 1
		.ReDraw = True
    End With

    Set gActiveElement = document.ActiveElement
    
    If Err.number = 0 Then FncInsertRow = True
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows 
	Dim lTempRows 

	If frm1.vspdData.maxrows < 1 then exit function
	
 '----------  Coding part  ------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData	
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
End Function
'========================================================================================
' Function Name : FncDelete
' Function Desc : 
'========================================================================================
Function FncDelete()	
	If frm1.vspdData.maxRows >= 1 then
		If DisplayMsgBox("210034", parent.VB_YES_NO, "x", "x") = vbYes Then '�����Ͻðڽ��ϱ�?
		 
		End If   
	End If

    'MyBizASP.location.href = "S3322MA1_KO412_frwriteBiz.asp?txtMode=" & UID_M0003  & "&txtKeyNo=" & MyBizAsp.frTitle.intKeyNo
    //MyBizASPForDelete.location.href = "S3322MA1_KO412_frwriteBiz.asp?txtMode=" & UID_M0003  & "&txtKeyNo=" & MyBizAsp.frTitle.intKeyNo
	    
End Function
	
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 	Call parent.FncExport(Parent.C_MULTI)		
 End Function
 
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()

	FncExit = False
	
	Dim IntRetCD
	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If
    
    FncExit = True    
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
	Dim strVal

	DbQuery = False                                                             

	Call LayerShowHide(1)

	With frm1
	
'		If lgIntFlgMode = "txtMode" Then
'		
'		    strVal = BIZ_PGM_ID & "?txtMode=" & "head"
'		    strVal = strVal & "&txtBpCd=" 	& .hdnBpCD.value
'		    strVal = strVal & "&txtDlvyNo=" & .hdnDlvyNo.value
'		    strVal = strVal & "&txtItemCd=" & .hdnItemCd.value	   
'		    strVal = strVal & "&txtDvFrDt=" & .hdnDvFrDt.value	  
'		    strVal = strVal & "&txtDvToDt=" & .hdnDvToDt.value		
'		    strVal = strVal & "&lgPageNo="	& lgPageNo						'��: Next key tag 
'		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
'	  
'	    Else	
	  	
		    strVal = BIZ_PGM_ID & "?txtMode=" & "head"
		    strVal = strVal & "&txtBpCd=" 	& .txtBpCD.value	   
		    strVal = strVal & "&txtDlvyNo=" & .txtDlvyNo.value
		    strVal = strVal & "&txtItemCd=" & .txtItemCd.value	   
		    strVal = strVal & "&txtDvFrDt=" & .txtDvFrDt.text	  
		    strVal = strVal & "&txtDvToDt=" & .txtDvToDt.text		      
		    strVal = strVal & "&lgPageNo="	& lgPageNo						'��: Next key tag 
		    strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		    
'	    End If 

	End With

	Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 

	DbQuery = True
End Function



'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQueryDtl() 
	Dim strVal
	Dim strBpCd, strDlvyNo
	

	DbQueryDtl = False                                                             

	Call LayerShowHide(1)

	With frm1
		.vspdData.MAXROWS = 0
		
		.vspdData1.Row = .vspdData1.ActiveRow
		.vspdData1.Col = C_HH_BP_CD
		strBpCd = .vspdData1.text
			
		.vspdData1.Col = C_HH_DLVY_NO
		strDlvyNo = .vspdData1.text

'MsgBox "DbQueryDtl(10)"
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&txtBpCd=" 	& strBpCd
		    strVal = strVal & "&txtDlvyNo=" & strDlvyNo
		    strVal = strVal & "&lgPageNo="	& lgPageNo						'��: Next key tag 
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	  
	    Else	
	  	
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&txtBpCd=" & strBpCd
		    strVal = strVal & "&txtDlvyNo=" & strDlvyNo
		    strVal = strVal & "&lgPageNo="	 & lgPageNo						'��: Next key tag 
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    End If 
'MsgBox "DbQueryDtl(20)"
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 
	
	DbQueryDtl = True
End Function



'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk()
	Dim ii
	
    lgIntFlgMode = Parent.OPMD_UMODE				'��: Indicates that current mode is Update mode
	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("110000000001111")				'��ư ���� ���� 

    If frm1.vspdData1.MaxRows > 0 Then
		Call SetToolBar("110010110001111")
		frm1.vspddata1.focus
		
		Call DbQueryDtl()
		
	End If
	
	'Call InitData()
	Set gActiveElement = document.activeElement

	frm1.txtDlvyNo.focus()
	
End Function



'=======================================================================================================
' Function Name : DbQueryDtlOk
' Function Desc : DbQueryDtl�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryDtlOk()
	Dim ii
    lgIntFlgMode = Parent.OPMD_UMODE							'��: Indicates that current mode is Update mode
	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("110000000001111")				'��ư ���� ���� 
    If frm1.vspdData.MaxRows > 0 Then
		Call SetToolBar("110010110001111")
		frm1.vspddata.focus
		MyBizASP1.location.href = "U2211RA1_KO441.asp?dlvy_no=" & ""
	End If
	Set gActiveElement = document.activeElement
	call vspdData_click (1,1)
	frm1.txtDlvyNo.focus()
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display =>���۾���
'========================================================================================================
Function DbSave() 
    Dim lRow
	Dim lGrpCnt     
	Dim strVal,strDel
	Dim ColSep, RowSep
	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size
	Dim iTmpDBuffer          '������ ���� [����] 
	Dim iTmpDBufferCount     '������ ���� Position
	Dim iTmpDBufferMaxCount  '������ ���� Chunk Size
	Dim ii

	ColSep = parent.gColSep               
	RowSep = parent.gRowSep               

    DbSave = False                                                          '��: Processing is NG
	Call LayerShowHide(1)

	frm1.txtMode.value = Parent.UID_M0002
	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 0

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]

	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]

	iTmpCUBufferMaxCount = -1 
	iTmpDBufferMaxCount = -1 

	With frm1
		.txtMode.value = parent.UID_M0002
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	ggoSpread.source = frm1.vspdData

      For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		'��: C=Create
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		'��: U=Update
				Case ggoSpread.DeleteFlag
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep		'��: U=Delete
						
			End Select			
 
		    Select Case .vspdData.Text 
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'��: �ű�, ���� 
	
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_H_DLVY_NO,lRow,"X","X"))  & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_H_DOCUMENT_NO,lRow,"X","X"))  & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_H_DOCUMENT_ABBR,lRow,"X","X")) & ColSep
					lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag							'��: ���� 
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_H_BP_CD,lRow,"X","X")) & ColSep					
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_H_DLVY_NO,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_H_DOCUMENT_NO,lRow,"X","X")) & RowSep

  		            lGrpCnt = lGrpCnt + 1
		    End Select
		 
		Next
		
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			'��: �����Ͻ� ASP �� ���� 
	DbSave = True                                                      
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
End Function


'=======================================================================================================
' Function Name : FncWrite
' Function Desc : 
'========================================================================================================
Function FncWrite()

		Dim arrRet,parm
		Dim sBpCd, sDlvyNo

	'sBpCd			=	GetSpreadText(frm1.vspdData,1,frm1.vspdData.ActiveRow,"X","X")	'1: BP_CD
	'sDlvyNo		=	GetSpreadText(frm1.vspdData,2,frm1.vspdData.ActiveRow,"X","X")	'2: DLVY_NO
	'sDocument_no	=	GetSpreadText(frm1.vspdData,3,frm1.vspdData.ActiveRow,"X","X")	'3: document_no
	
	
	With frm1
		
		.vspdData1.Row = .vspdData1.ActiveRow
		.vspdData1.Col = C_HH_BP_CD
		sBpCd = .vspdData1.text
			
		.vspdData1.Col = C_HH_DLVY_NO
		sDlvyNo = .vspdData1.text

	End With	

	
	'arrParam(0) = sBpCd
	'arrParam(1) = sDlvyNo
	'arrParam(2) = sDocument_no

'		If UCase(Trim(frm1.txtBpCd.value)) = "" Then
'			Call DisplayMsgBox("900002", "x", "x", "x")		 '��: "Will you destory previous data"
'			frm1.txtBpCd.Focus
'			Exit Function
'		End If
'		
'		If UCase(Trim(frm1.txtDlvyNo.value)) = "" Then
'			Call DisplayMsgBox("900002", "x", "x", "x")		 '��: "Will you destory previous data"
'			frm1.txtDlvyNo.Focus
'			Exit Function
'		End If
		
		If IsOpenPop = True Then Exit Function
         reDim parm(3)
		IsOpenPop = True

		arrRet = window.showModalDialog ("U2211PA1_KO441.asp?strMode=" & parent.UID_M0001 & "&bp_cd=" & sBpCd & "&dlvy_no=" & sDlvyNo, Array(window.parent,parm(0),parm(1)), _
       "dialogWidth=600px; dialogHeight=470px; center: Yes; help: No; resizable: No; status: No;")	

		If arrRet = True Then
			call dbsaveOK()
			MyBizASP1.location.reload						
		End If
				
		IsOpenPop = False

	End Function

'=======================================================================================================
' Function Name : FncModify
' Function Desc : 
'========================================================================================================
Function FncModify()

	Dim arrRet, sBpCd, sDlvyNo, sDocument_no
	Dim arrParam(3)
		
	If IsOpenPop = True Then Exit Function
	
'	If UCase(Trim(frm1.txtDlvyNo.value)) = "" Then
'		Call DisplayMsgBox("900002", "x", "x", "x")		 '��: "Will you destory previous data"
'		Exit Function
'	End If
	
	if frm1.vspdData.maxRows = 0 then
		 call DisplayMsgBox("900025", "X", "X", "X") 
		 Exit Function
	end if
	
	sBpCd			=	GetSpreadText(frm1.vspdData,1,frm1.vspdData.ActiveRow,"X","X")	'1: BP_CD
	sDlvyNo			=	GetSpreadText(frm1.vspdData,2,frm1.vspdData.ActiveRow,"X","X")	'2: DLVY_NO
	sDocument_no	=	GetSpreadText(frm1.vspdData,3,frm1.vspdData.ActiveRow,"X","X")	'3: document_no
	
	arrParam(0) = sBpCd
	arrParam(1) = sDlvyNo
	arrParam(2) = sDocument_no
	
		
	IsOpenPop = True
	arrRet = window.showModalDialog ("U2211PA1_KO441.asp?strMode=" & parent.UID_M0002 & "&bp_cd=" & sBpCd & "&dlvy_no=" & sDlvyNo & "&Document_no=" & sDocument_no, Array(window.parent, arrParam), _
	"dialogWidth=600px; dialogHeight=470px; center: Yes; help: No; resizable: No; status: No;")	

	If arrRet = True Then
		call dbsaveOK()
		MyBizASP1.location.reload
	End If
		
	IsOpenPop = False

End Function

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(lRow)
	With frm1.vspdData
		.Row = lRow
		.Col = C_H_DOCUMENT_ABBR
		.Action = 0
		Call SetFocusToDocument("M") 
		.focus
	End With
End Function


Function Jump()	

    Dim iRet
    Dim iRet2
    Dim strVal
    Dim iArr
    Dim eisWindow ,strPgmID
    'On Error Resume Next
  
	CookiePage("")
    PgmJump(BIZ_PGM_REG_ID)

End Function



'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenDlvyNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenDlvyNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
	Dim IntRetCD

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(Trim(frm1.txtbpcd.Value), "", "S"), lgF0) = False Then
		Call DisplayMsgBox("971012", "X", "����ó", "X")
		frm1.txtbpNM.VALUE = ""
		frm1.txtbpcd.focus
    	Exit Function
    Else
		lgF0 = replace(lgF0,chr(12),"")
		frm1.txtbpnm.value = replace(lgF0,chr(11),"")
	End If
		
	if Trim(frm1.txtbpcd.Value) = "" then
		Call DisplayMsgBox("205152", "X", "��ü", "X")
		frm1.txtbpcd.focus
    	Exit Function
    End if
			
	If IsOpenPop = True Or UCase(frm1.txtDlvyNo.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
		
	IsOpenPop = True
		
	arrParam(0) = Trim(frm1.txtbpcd.Value)

	iCalledAspName = AskPRAspName("U2122PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "U2122PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.txtDlvyNo.value = strRet(0)
		frm1.txtDlvyNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function


'================================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��ü"										' �˾� ��Ī 
	arrParam(1) = "B_Biz_Partner"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBpCd.Value)						' Code Condition
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"	' Where Condition
	arrParam(5) = "��ü"										' TextBox ��Ī 
	
    arrField(0) = "BP_CD"										' Field��(0)
    arrField(1) = "BP_NM"										' Field��(1)
    
    arrHeader(0) = "��ü"										' Header��(0)
    arrHeader(1) = "��ü��"									' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
		
	End If	
End Function


'================================================================================================================================
Function OpenItemInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "ǰ���˾�"
	arrParam(1) = "(			SELECT	DISTINCT ITEM_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_HDR B "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.SPLIT_SEQ_NO = 0 AND B.BP_CD = '" & frm1.txtBpCd.value & "') A, B_ITEM B"
	arrParam(2) = Trim(frm1.txtItemCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD "
	arrParam(5) = "ǰ��"
	 
    arrField(0) = "A.ITEM_CD"												' Field��(0)
    arrField(1) = "B.ITEM_NM"												' Field��(1)
    
    arrHeader(0) = "ǰ��"													' Header��(0)
    arrHeader(1) = "ǰ���"													' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function





'=================================================================================================
'   Event Name :vspddata_ComboSelChange
'   Event Desc :Combo Change Event
'==================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row

		Select Case Col

			'// �ý��� ����
			Case C_H_DOCUMENT_NO
				.Col = Col
				intIndex = .Value
				.Col = C_H_DOCUMENT_NO
				.Value = intIndex

			Case C_H_TITLE
				.Col = Col
				intIndex = .Value
				.Col = C_H_TITLE
				.Value = intIndex

		End Select
    End With
End Sub

'=====================================================================================================
'   Event Name : txtDlvyNo_OnChange
'   Event Desc :
'=====================================================================================================
Sub txtDlvyNo_OnChange()

	Call CommonQueryRs(" PROJECT_NM "," pms_project ", " dlvy_no = " & FilterVar(frm1.txtDlvyNo.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	frm1.txtProjectNm.value = Replace(Trim(lgF0), Chr(11), "")
	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%> >
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
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strAspMnuMnunm")%></font></td>
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						</TR>
					</TABLE>
					</TD>
				
					<TD WIDTH=* Align=right>
					<A onclick="vbscript:FncWrite()">���</A>&nbsp;|&nbsp;<A onclick="vbscript:FncModify()">����</A>
					<A onclick="vbscript:FncDelete()"></A></TD>
					<TD WIDTH=10>&nbsp;</TD>
					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=65%>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD" >
					<TABLE <%=LR_SPACE_TYPE_40%>>
					   <TR>
							<TD CLASS=TD5 NOWRAP>��ü</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="��ü"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Openbpcd()">&nbsp;
												 <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14" ALT="��ü��"></TD>
							<TD CLASS=TD5 NOWRAP>���������ȣ</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDlvyNo" ALT="���������ȣ" TYPE="Text" MAXLENGTH=18 SiZE=18 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDlvyno" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDlvyNo()"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>��ǰ������</TD>
							<TD CLASS="TD6">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDvFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��ǰ����������" id=OBJECT1></OBJECT>');</SCRIPT>
								&nbsp;~&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDvToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��ǰ����������" id=OBJECT2></OBJECT>');</SCRIPT>
							</TD>
							<TD CLASS=TD5 NOWRAP>ǰ��</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
						</TR>
					</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>				
			</TABLE>
		</TD>
	</TR>
		<TD <%=HEIGHT_TYPE_02%> ></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0% FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
			<IFRAME NAME="MyBizASP1" SRC="U2211RA1_KO441.asp" WIDTH=100% HEIGHT=90% FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
			<IFRAME NAME="MyBizASPForDelete" SRC="../../blank.htm" WIDTH=10% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
		</TD>
	</TR>	
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDlvyNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDvFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDvToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hCurrency" tag="24">
</FORM>
</BODY>
</HTML>
 
