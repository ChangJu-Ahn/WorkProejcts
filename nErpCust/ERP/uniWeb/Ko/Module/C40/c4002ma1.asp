<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : C/C�� ��α�Ģ ��� 
'*  3. Program ID           : c4002ma1.asp
'*  4. Program Name         : C/C�� ��α�Ģ ��� 
'*  5. Program Desc         : C/C�� ��α�Ģ ��� 
'*  6. Modified date(First) : 2005-08-30
'*  7. Modified date(Last)  : 2005-08-30
'*  8. Modifier (First)     : choe0tae 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4002mb1.asp"                               'Biz Logic ASP
Const BIZ_COPY_PGM_ID = "c4002mb2.asp"                               'Biz Logic ASP

' -- �׸���1�� �÷� ���� 
Dim C_SEQ_NO			' -- ��� Ű 
Dim C_DSTB_ORDER		' -- ��μ��� 
Dim C_COST_CD_LEVEL		' -- C/C ���� 
Dim C_COST_CD_LEVEL_POP	
Dim C_SENDER_COST_CD	' -- SEND C/C
Dim C_SENDER_COST_CD_POP
Dim C_SENDER_COST_NM	
Dim C_GP_LEVEL		
Dim C_GP_LEVEL_POP
Dim C_GP_CD				' -- �����׷� 
Dim C_GP_CD_POP
Dim C_GP_NM
Dim C_ACCT_CD			' -- ���� 
Dim C_ACCT_CD_POP
Dim C_ACCT_NM
Dim C_DI_FLAG
Dim C_DI_FLAG_NM
Dim C_ACTL_DSTB_FCTR_CD		' - ��ο�� 
Dim C_ACTL_DSTB_FCTR_CD_POP
Dim C_ACTL_DSTB_FCTR_NM
Dim C_ACTL_DSTB_FCTR_RATE
Dim C_ACTL_DSTB_FCTR_CD2		' - ��ο�� 
Dim C_ACTL_DSTB_FCTR_CD_POP2
Dim C_ACTL_DSTB_FCTR_NM2
Dim C_ACTL_DSTB_FCTR_RATE2
Dim C_ACTL_DSTB_FCTR_CD3		' - ��ο�� 
Dim C_ACTL_DSTB_FCTR_CD_POP3
Dim C_ACTL_DSTB_FCTR_NM3
Dim C_ACTL_DSTB_FCTR_RATE3
Dim C_ACTL_DSTB_FCTR_CD4		' - ��ο�� 
Dim C_ACTL_DSTB_FCTR_CD_POP4
Dim C_ACTL_DSTB_FCTR_NM4
Dim C_ACTL_DSTB_FCTR_RATE4
Dim C_ACTL_DSTB_FCTR_CD5		' - ��ο�� 
Dim C_ACTL_DSTB_FCTR_CD_POP5
Dim C_ACTL_DSTB_FCTR_NM5
Dim C_ACTL_DSTB_FCTR_RATE5
Dim C_STD_DSTB_FCTR_CD		' - ��ο�� 
Dim C_STD_DSTB_FCTR_CD_POP
Dim C_STD_DSTB_FCTR_NM

' -- �׸���2�� ���̴� �÷� ���� 
Dim C_SUB_SEQ_NO			' -- ������ Ű 
Dim C_RECV_COST_CD_LEVEL	' -- C/C ���� 
Dim C_RECV_COST_CD_LEVEL_POP
Dim C_RECV_COST_CD			' -- RECV C/C
Dim C_RECV_COST_CD_POP
Dim C_RECV_COST_NM

' -- �׸���2�� �����÷� : �׸���1�� Ű�÷� 
Dim C_COST_CD_LEVEL_PARENT
Dim C_SENDER_COST_CD_PARENT
Dim C_GP_CD_PARENT
Dim C_ACCT_CD_PARENT
Dim C_DI_FLAG_PARENT

Const GRID_1	= 1
Const GRID_2	= 2
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          
Dim lgCurrGrid
Dim lgCopyVersion
Dim lgErrRow, lgErrCol
Dim lgRowCnt


'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	
	' -- �׸���1�� �÷� ���� 
	 C_SEQ_NO					= 1		' -- ���� 
	 C_DSTB_ORDER				= 2		' -- ��μ��� 
	 C_COST_CD_LEVEL			= 3		' -- C/C ���� 
	 C_COST_CD_LEVEL_POP		= 4			
	 C_SENDER_COST_CD			= 5		' -- SEND C/C
	 C_SENDER_COST_CD_POP		= 6		
	 C_SENDER_COST_NM			= 7		
	 C_GP_LEVEL					= 8		
	 C_GP_LEVEL_POP				= 9		
	 C_GP_CD					= 10		' -- �����׷� 
	 C_GP_CD_POP				= 11		
	 C_GP_NM					= 12		
	 C_ACCT_CD					= 13	' -- ���� 
	 C_ACCT_CD_POP				= 14		
	 C_ACCT_NM					= 15	
	 C_DI_FLAG					= 16
	 C_DI_FLAG_NM				= 17
	 
	 C_ACTL_DSTB_FCTR_CD		= 18	' - ��ο�� 
	 C_ACTL_DSTB_FCTR_CD_POP	= 19	
	 C_ACTL_DSTB_FCTR_NM		= 20
	 C_ACTL_DSTB_FCTR_RATE		= 21
	 
	 C_ACTL_DSTB_FCTR_CD2		= 22	' - ��ο�� 
	 C_ACTL_DSTB_FCTR_CD_POP2	= 23	
	 C_ACTL_DSTB_FCTR_NM2		= 24
	 C_ACTL_DSTB_FCTR_RATE2		= 25

	 C_ACTL_DSTB_FCTR_CD3		= 26	' - ��ο�� 
	 C_ACTL_DSTB_FCTR_CD_POP3	= 27
	 C_ACTL_DSTB_FCTR_NM3		= 28
	 C_ACTL_DSTB_FCTR_RATE3		= 29

	 C_ACTL_DSTB_FCTR_CD4		= 30	' - ��ο�� 
	 C_ACTL_DSTB_FCTR_CD_POP4	= 31
	 C_ACTL_DSTB_FCTR_NM4		= 32
	 C_ACTL_DSTB_FCTR_RATE4		= 33
	 
	 C_ACTL_DSTB_FCTR_CD5		= 34	' - ��ο�� 
	 C_ACTL_DSTB_FCTR_CD_POP5	= 35
	 C_ACTL_DSTB_FCTR_NM5		= 36
	 C_ACTL_DSTB_FCTR_RATE5		= 37
	 
	 C_STD_DSTB_FCTR_CD			= 38	' - ��ο�� 
	 C_STD_DSTB_FCTR_CD_POP		= 39	
	 C_STD_DSTB_FCTR_NM			= 40

	' -- �׸���2�� ���̴� �÷� ���� 
	 C_SUB_SEQ_NO				= 2		' -- �θ�seq_no���ԵǾ� 2������		
	 C_RECV_COST_CD_LEVEL		= 3		' -- C/C ���� 
	 C_RECV_COST_CD_LEVEL_POP	= 4		
	 C_RECV_COST_CD				= 5		' -- RECV C/C
	 C_RECV_COST_CD_POP			= 6		
	 C_RECV_COST_NM				= 7		

	' -- �׸���2�� �����÷� : �׸���1�� Ű�÷� 
	 C_COST_CD_LEVEL_PARENT		= 8		
	 C_SENDER_COST_CD_PARENT	= 9		
	 C_GP_CD_PARENT				= 10		
	 C_ACCT_CD_PARENT			= 11		
	 C_DI_FLAG_PARENT			= 12
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0  
    lgRowCnt=0
    
    lgStrPrevKey = ""	
    lgLngCurRows = 0 
	lgSortKey = 1
	lgCurrGrid = GRID_1
	lgCopyVersion = ""
	lgErrRow = 0 : lgErrCol = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'   `     3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	' -- �׸��� �÷� ��ġ �ʱ�ȭ 
	Call initSpreadPosVariables()    
	
	Call AppendNumberPlace("6","3","0")
	Call AppendNumberPlace("7","2","0")
	' -- �׸��� 1 ���� 
	With frm1.vspdData
	
	.MaxCols = C_STD_DSTB_FCTR_NM+1

	.Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122",,parent.gForbidDragDropSpread 'gAllowDragDropSpread 

	Call ggoSpread.ClearSpreadData()    '��: Clear spreadsheet data 

	.ReDraw = false

	Call GetSpreadColumnPos("A")
	
    ggoSpread.SSSetEdit		C_SEQ_NO			,"��ȣ",7,1
    ggoSpread.SSSetFloat	C_DSTB_ORDER,		"��μ���"	, 7,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"1","" 
    ggoSpread.SSSetFloat	C_COST_CD_LEVEL		,"C/C Level"	, 10,		"7",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	ggoSpread.SSSetButton	C_COST_CD_LEVEL_POP    
    ggoSpread.SSSetEdit		C_SENDER_COST_CD	,"Sender" & vbCrLf & "C/C"	,10,,,10,2
    ggoSpread.SSSetButton	C_SENDER_COST_CD_POP    
    ggoSpread.SSSetEdit		C_SENDER_COST_NM	,"Sender" & vbCrLf & "C/C��",15
    ggoSpread.SSSetFloat	C_GP_LEVEL		,"�����׷�" & vbCrLf & "Level"	, 10,		"7",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	ggoSpread.SSSetButton	C_GP_LEVEL_POP    
    ggoSpread.SSSetEdit		C_GP_CD				,"�����׷�" ,10,,, 20,2
    ggoSpread.SSSetButton	C_GP_CD_POP    
    ggoSpread.SSSetEdit		C_GP_NM				,"�����׷��",15
    ggoSpread.SSSetEdit		C_ACCT_CD			,"����" ,10,,, 20,2
    ggoSpread.SSSetButton	C_ACCT_CD_POP    
    ggoSpread.SSSetEdit		C_ACCT_NM			,"������",15
	ggoSpread.SSSetCombo	C_DI_FLAG		,"��/����", 10
    ggoSpread.SSSetCombo	C_DI_FLAG_NM		,"��/����", 10
    
    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_CD		,"��ο��1" ,7,,,, 2
    ggoSpread.SSSetButton	C_ACTL_DSTB_FCTR_CD_POP    
    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_NM		,"��ο�Ҹ�1",20
    ggoSpread.SSSetFloat	C_ACTL_DSTB_FCTR_RATE		,"Rate1" , 7,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 

    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_CD2		,"��ο��2" ,7,,,, 2
    ggoSpread.SSSetButton	C_ACTL_DSTB_FCTR_CD_POP2    
    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_NM2		,"��ο�Ҹ�2",20
    ggoSpread.SSSetFloat	C_ACTL_DSTB_FCTR_RATE2		,"Rate2" , 7,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 

    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_CD3		,"��ο��3" ,7,,,, 2
    ggoSpread.SSSetButton	C_ACTL_DSTB_FCTR_CD_POP3    
    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_NM3		,"��ο�Ҹ�3",20
    ggoSpread.SSSetFloat	C_ACTL_DSTB_FCTR_RATE3		,"Rate3" , 7,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 

    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_CD4		,"��ο��4" ,7,,,, 2
    ggoSpread.SSSetButton	C_ACTL_DSTB_FCTR_CD_POP4    
    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_NM4		,"��ο�Ҹ�4",20
    ggoSpread.SSSetFloat	C_ACTL_DSTB_FCTR_RATE4		,"Rate4" , 7,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 

    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_CD5		,"��ο��5" ,7,,,, 2
    ggoSpread.SSSetButton	C_ACTL_DSTB_FCTR_CD_POP5    
    ggoSpread.SSSetEdit		C_ACTL_DSTB_FCTR_NM5		,"��ο�Ҹ�5",20
    ggoSpread.SSSetFloat	C_ACTL_DSTB_FCTR_RATE5		,"Rate5" , 7,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 

        
    ggoSpread.SSSetEdit		C_STD_DSTB_FCTR_CD		,"ǥ�ؿ�����ο��" ,10,,,, 2
    ggoSpread.SSSetButton	C_STD_DSTB_FCTR_CD_POP    
    ggoSpread.SSSetEdit		C_STD_DSTB_FCTR_NM		,"ǥ�ؿ�����ο�Ҹ�",20


	Call ggoSpread.SSSetColHidden(C_DI_FLAG,C_DI_FLAG,True)

	call ggoSpread.MakePairsColumn(C_DI_FLAG,C_DI_FLAG_NM)
	
	Call ggoSpread.SSSetColHidden(C_STD_DSTB_FCTR_CD,C_STD_DSTB_FCTR_CD,True)
	Call ggoSpread.SSSetColHidden(C_STD_DSTB_FCTR_CD_POP,C_STD_DSTB_FCTR_CD_POP,True)
	Call ggoSpread.SSSetColHidden(C_STD_DSTB_FCTR_NM,C_STD_DSTB_FCTR_NM,True)
	


	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
	
	.rowheight(-1000) = 20	' ���� ������ 

	.ReDraw = true
	
    Call SetSpreadLock 
    Call InitComboBox
    
    End With
    
    
    ' -- �׸��� 2 ���� 
    With frm1.vspdData2
	
	.MaxCols = C_DI_FLAG_PARENT+1

	.Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.Spreadinit "V20021122_2",,parent.gForbidDragDropSpread 

	Call ggoSpread.ClearSpreadData()    '��: Clear spreadsheet data 

	.ReDraw = false

	Call GetSpreadColumnPos("B")
	
	ggoSpread.SSSetEdit		C_SEQ_NO					,"��ȣ",7,1
	ggoSpread.SSSetEdit		C_SUB_SEQ_NO				,"��ȣ",7,1
	ggoSpread.SSSetFloat	C_RECV_COST_CD_LEVEL		,"C/C Level"	, 8,		"7",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	ggoSpread.SSSetButton	C_RECV_COST_CD_LEVEL_POP    
    ggoSpread.SSSetEdit		C_RECV_COST_CD				,"Receiver" & vbCrLf & "C/C"	,12,,,10,2
    ggoSpread.SSSetButton	C_RECV_COST_CD_POP    
    ggoSpread.SSSetEdit		C_RECV_COST_NM				,"Receiver" & vbCrLf & "C/C��"	,20

    ggoSpread.SSSetEdit		C_COST_CD_LEVEL_PARENT		,"C/C Level"	,10
    ggoSpread.SSSetEdit		C_SENDER_COST_CD_PARENT		,"Sender" & vbCrLf & "C/C"	,12, 10
	ggoSpread.SSSetEdit		C_GP_CD_PARENT				,"�����׷�" ,12, 20
    ggoSpread.SSSetEdit		C_ACCT_CD_PARENT			,"����" ,12, 20
	ggoSpread.SSSetEdit		C_DI_FLAG_PARENT			,"��/����" ,12, 20
   'call ggoSpread.MakePairsColumn(C_DSTB_ORDER,C_COST_CD_LEVEL)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SUB_SEQ_NO,True)
	Call ggoSpread.SSSetColHidden(C_COST_CD_LEVEL_PARENT,C_DI_FLAG_PARENT,True)
	
	.rowheight(-1000) = 20	' ���� ������ 

	.ReDraw = true
	
    Call SetSpreadLock2 
    'Call InitComboBox
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
    
    ggoSpread.Source = frm1.vspdData
    
    .ReDraw = False
    ggoSpread.SpreadLock		C_SEQ_NO			,-1	,C_SEQ_NO
    ggoSpread.SSSetRequired		C_DSTB_ORDER		,-1	,-1
	ggoSpread.SSSetRequired		C_COST_CD_LEVEL		,-1	,-1
	ggoSpread.SSSetRequired		C_SENDER_COST_CD	,-1	,-1
	ggoSpread.SSSetRequired		C_ACTL_DSTB_FCTR_CD	 ,-1	,-1
	ggoSpread.SSSetRequired		C_ACTL_DSTB_FCTR_RATE	 ,-1	,-1
	ggoSpread.SpreadLock		C_SENDER_COST_NM	,-1	,C_SENDER_COST_NM
	ggoSpread.SpreadLock		C_GP_NM				,-1	,C_GP_NM
	ggoSpread.SpreadLock		C_ACCT_NM			,-1	,C_ACCT_NM
	ggoSpread.SpreadLock		C_ACTL_DSTB_FCTR_NM		,-1	,C_ACTL_DSTB_FCTR_NM
	ggoSpread.SpreadLock		C_ACTL_DSTB_FCTR_NM2		,-1	,C_ACTL_DSTB_FCTR_NM2
	ggoSpread.SpreadLock		C_ACTL_DSTB_FCTR_NM3		,-1	,C_ACTL_DSTB_FCTR_NM3	
	ggoSpread.SpreadLock		C_ACTL_DSTB_FCTR_NM4		,-1	,C_ACTL_DSTB_FCTR_NM4
	ggoSpread.SpreadLock		C_ACTL_DSTB_FCTR_NM5		,-1	,C_ACTL_DSTB_FCTR_NM5
	ggoSpread.SpreadLock		C_STD_DSTB_FCTR_NM		,-1	,C_ACTL_DSTB_FCTR_NM
    .ReDraw = True

    End With
End Sub

Sub SetSpreadLock2()
    With frm1.vspdData2
    
    ggoSpread.Source = frm1.vspdData2
    
    .ReDraw = False
    ggoSpread.SpreadLock		C_SEQ_NO			,-1	,C_SUB_SEQ_NO
    ggoSpread.SSSetRequired		C_RECV_COST_CD		,-1	,-1
	ggoSpread.SpreadLock		C_RECV_COST_NM		,-1	,C_ACCT_CD_PARENT
    .ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False
								      'Col          Row				Row2    
	ggoSpread.SSSetProtected	C_SEQ_NO			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetRequired		C_DSTB_ORDER		,pvStartRow		,pvEndRow
	ggoSpread.SSSetRequired		C_COST_CD_LEVEL		,pvStartRow		,pvEndRow
	ggoSpread.SSSetRequired		C_SENDER_COST_CD	,pvStartRow		,pvEndRow
	ggoSpread.SSSetRequired		C_ACTL_DSTB_FCTR_CD	,pvStartRow		,pvEndRow
	ggoSpread.SSSetRequired		C_ACTL_DSTB_FCTR_RATE	,pvStartRow		,pvEndRow
	ggoSpread.SSSetProtected	C_SENDER_COST_NM	,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_GP_NM				,pvStartRow		,pvEndRow    	
	ggoSpread.SSSetProtected	C_ACCT_NM			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ACTL_DSTB_FCTR_NM		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ACTL_DSTB_FCTR_NM2		,pvStartRow		,pvEndRow    	
	ggoSpread.SSSetProtected	C_ACTL_DSTB_FCTR_NM3		,pvStartRow		,pvEndRow    	
	ggoSpread.SSSetProtected	C_ACTL_DSTB_FCTR_NM4		,pvStartRow		,pvEndRow    	
	ggoSpread.SSSetProtected	C_ACTL_DSTB_FCTR_NM5		,pvStartRow		,pvEndRow    				
	ggoSpread.SSSetProtected	C_STD_DSTB_FCTR_NM		,pvStartRow		,pvEndRow    
	
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub SetSpreadColor2(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    ggoSpread.Source = frm1.vspdData2
    .vspdData2.ReDraw = False
									      'Col          Row				Row2
	ggoSpread.SSSetProtected	C_SEQ_NO			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_sub_SEQ_NO		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetRequired		C_RECV_COST_CD		,pvStartRow		,pvEndRow
	ggoSpread.SSSetProtected	C_RECV_COST_NM		,pvStartRow		,pvEndRow
	
	ggoSpread.SSSetProtected	C_COST_CD_LEVEL_PARENT		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_SENDER_COST_CD_PARENT		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_GP_CD_PARENT				,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ACCT_CD_PARENT			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_DI_FLAG_PARENT			,pvStartRow		,pvEndRow    
    .vspdData2.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx, oGrid, j, iSeqNo, iSubSeqNo
    Dim iRow
    If iPosArr = "" Then Exit Sub
    iPosArr = Split(iPosArr,Parent.gColSep)		' ���Ϲ��ڿ�: �׸���n/gColSep/�����÷���/gColSep/�������ȣ(C:SEQ_NO��ȣ)/gColSep/SUB_SEQ_NO
    If IsNumeric(iPosArr(0)) Then
       iDx = CDbl(iPosArr(2))	' ���ȣ/SEQ_NO��ȣ 
       
		If iPosArr(0) = "1" Then	' �׸���n ���� 
			Set oGrid = frm1.vspdData
		Else
			Set oGrid = frm1.vspdData2
		End If
       
		With oGrid
		
		For iRow = 1 To  .MaxRows 
		    .Col = 0
		    .Row = iRow
		    
			If iPosArr(0) = "1" Then	' -- �׸���1�� ��� 
				.Col = C_SEQ_NO	: iSeqNo = UNICDbl(.value)
				If iSeqNo = iDx Then	' -- �������ȣ�� SEQ_NO�� ���ٸ� 
					Call ClickGrid1(iSeqNo)
					Exit Sub
				End If
				' -- �������ȣ�� SEQ_NO�� �ٸ��Ƿ� ���� For�� ���� 
			Else
				' -- �׸���2 �� ��� 
				.Col = C_SEQ_NO		: iSeqNo	= UNICDbl(.value)
				.Col = C_SUB_SEQ_NO	: iSubSeqNo = UNICDbl(.value)
				If iSeqNo = iDx And iSubSeqNo = UNICDbl(iPosArr(3)) Then	' -- �������ȣ�� SEQ_NO�� ���ٸ� 
					.Col = C_RECV_COST_CD	: .Action  = 0	
					lgErrRow = iRow		' -- ������ ������ 
					Call ClickGrid1(iSeqNo)
					Exit Sub
				End If
			End If
					
		Next
        
        End With 
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			' -- �׸���1�� �÷� ���� 
			 C_SEQ_NO					= iCurColumnPos(1)	
			 C_DSTB_ORDER				= iCurColumnPos(2)	' -- ��μ��� 
			 C_COST_CD_LEVEL			= iCurColumnPos(3)	' -- C/C ���� 
			 C_COST_CD_LEVEL_POP		= iCurColumnPos(4)		
			 C_SENDER_COST_CD			= iCurColumnPos(5)	' -- SEND C/C
			 C_SENDER_COST_CD_POP		= iCurColumnPos(6)		
			 C_SENDER_COST_NM			= iCurColumnPos(7)		
			 C_GP_LEVEL					= iCurColumnPos(8)		
			 C_GP_LEVEL_POP				= iCurColumnPos(9)		
			 C_GP_CD					= iCurColumnPos(10)	' -- �����׷� 
			 C_GP_CD_POP				= iCurColumnPos(11)		
			 C_GP_NM					= iCurColumnPos(12)		
			 C_ACCT_CD					= iCurColumnPos(13)	' -- ���� 
			 C_ACCT_CD_POP				= iCurColumnPos(14)		
			 C_ACCT_NM					= iCurColumnPos(15)	
			 C_DI_FLAG					= iCurColumnPos(16)	
			 C_DI_FLAG_NM				= iCurColumnPos(17)	
			 
			 C_ACTL_DSTB_FCTR_CD		= iCurColumnPos(18)	' - ��ο�� 
			 C_ACTL_DSTB_FCTR_CD_POP	= iCurColumnPos(19)	
			 C_ACTL_DSTB_FCTR_NM		= iCurColumnPos(20)	
			 C_ACTL_DSTB_FCTR_RATE		= iCurColumnPos(21)	

			 C_ACTL_DSTB_FCTR_CD2		= iCurColumnPos(22)	' - ��ο�� 
			 C_ACTL_DSTB_FCTR_CD_POP2	= iCurColumnPos(23)	
			 C_ACTL_DSTB_FCTR_NM2		= iCurColumnPos(24)	
			 C_ACTL_DSTB_FCTR_RATE2		= iCurColumnPos(25)	

			 C_ACTL_DSTB_FCTR_CD3		= iCurColumnPos(26)	' - ��ο�� 
			 C_ACTL_DSTB_FCTR_CD_POP3	= iCurColumnPos(27)	
			 C_ACTL_DSTB_FCTR_NM3		= iCurColumnPos(28)	
			 C_ACTL_DSTB_FCTR_RATE3		= iCurColumnPos(29)	

			 C_ACTL_DSTB_FCTR_CD4		= iCurColumnPos(30)	' - ��ο�� 
			 C_ACTL_DSTB_FCTR_CD_POP4	= iCurColumnPos(31)	
			 C_ACTL_DSTB_FCTR_NM4		= iCurColumnPos(32)	
			 C_ACTL_DSTB_FCTR_RATE4		= iCurColumnPos(33)	

			 C_ACTL_DSTB_FCTR_CD5		= iCurColumnPos(34)	' - ��ο�� 
			 C_ACTL_DSTB_FCTR_CD_POP5	= iCurColumnPos(35)	
			 C_ACTL_DSTB_FCTR_NM5		= iCurColumnPos(36)	
			 C_ACTL_DSTB_FCTR_RATE5		= iCurColumnPos(37)	

		 
			 C_STD_DSTB_FCTR_CD			= iCurColumnPos(38)	' - ��ο�� 
			 C_STD_DSTB_FCTR_CD_POP		= iCurColumnPos(39)	
			 C_STD_DSTB_FCTR_NM			= iCurColumnPos(40)	

		Case "B"

            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			' -- �׸���2�� ���̴� �÷� ���� 
			C_SUB_SEQ_NO				= iCurColumnPos(2)		
			C_RECV_COST_CD_LEVEL		= iCurColumnPos(3)		' -- C/C ���� 
			C_RECV_COST_CD_LEVEL_POP	= iCurColumnPos(4)		
			C_RECV_COST_CD				= iCurColumnPos(5)		' -- RECV C/C
			C_RECV_COST_CD_POP			= iCurColumnPos(6)		
			C_RECV_COST_NM				= iCurColumnPos(7)		

			' -- �׸���2�� �����÷� : �׸���1�� Ű�÷� 
			C_COST_CD_LEVEL_PARENT		= iCurColumnPos(8)		
			C_SENDER_COST_CD_PARENT	= iCurColumnPos(9)		
			C_GP_CD_PARENT				= iCurColumnPos(10)		
			C_ACCT_CD_PARENT			= iCurColumnPos(11)		
			C_DI_FLAg_PARENT			= iCurColumnPos(12)	
    End Select    
End Sub

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
	Dim sCd
	Dim sNm
    

	sCd = "*" & vbTab & "D" & vbTab & "I"
	sNm = "*" & vbTab & "����" & vbTab & "����"
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo sCd, C_DI_FLAG			'COLM_DATA_TYPE
    ggoSpread.SetCombo sNm, C_DI_FLAG_NM
    

     
	
End Sub

' -- Version �˾���.
Function OpenVersion(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 1

			If Not chkField(Document, "1") Then
			   Exit Function
			End If

			Dim IntRetCD , blnChange1, blnChange2
    
			Err.Clear
    
			ggoSpread.Source = frm1.vspdData
			blnChange1 = ggoSpread.SSCheckChange

			ggoSpread.Source = frm1.vspdData2
			blnChange2 = ggoSpread.SSCheckChange
    
			If blnChange1 = True Or blnChange2 = True Then
				IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")
				If IntRetCD = vbNo Then
			      	Exit Function
				End If
			End If
	End Select

	IsOpenPop = True

	arrParam(0) = "Version �˾�"
	arrParam(1) = "C_MFC_DSTB_RULE_BY_CC_S"	
	
	arrParam(3) = ""		
	If frm1.txtVER_CD.value <> "" Then	
		arrParam(4) = "ver_cd <> " & FilterVar(frm1.txtVER_CD.value, "''", "S")
	End If

	arrParam(5) = "Version"
	
    arrField(0) = "ver_cd"
    
    arrHeader(0) = "Version"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtVER_CD.focus
		Exit Function
	Else
		Call SetVersion(arrRet, iWhere)
	End If
		
End Function

Function OpenErrorPopup(Byval pErrDesc)
	Dim arrRet, sTmp, iWidth
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	With frm1
	
	iWidth = 500	' -- �˾�Width

	arrParam(0) = "��α�Ģ ����"
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = ""		
	arrParam(4) = ""
	arrParam(5) = "" 

	arrField(0) = "ED10" & Parent.gColSep & "A1"	
	arrField(1) = "ED15" & Parent.gColSep & "A2"
	arrField(2) = "ED15" & Parent.gColSep & "A3"		
    
	arrHeader(0) = "��μ���"	
	arrHeader(1) = "Sender C/C"
	arrHeader(2) = "Receiver C/C"

	arrRet = window.showModalDialog("c4002ma2.asp", Array(arrParam, arrField, arrHeader, pErrDesc), _
		"dialogWidth=" & CStr(iWidth) & "px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	End With
End Function

' -- �׸���1���� �˾� Ŭ���� 
Function OpenPopUp(Byval iWhere, Byval strCode, Byval strCode1)
	Dim arrRet, sTmp, iWidth
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	iWidth = 500	' -- �˾�Width
	
	Select Case iWhere
		Case C_COST_CD_LEVEL_POP
			arrParam(0) = "C/C Level�˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('1')"	
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = ""
			arrParam(5) = "C/C Level" 

			arrField(0) = "LEVEL_CD"		
    
			arrHeader(0) = "C/C Level"	
			
		Case C_SENDER_COST_CD_POP
			arrParam(0) = "Sender C/C �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('2')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			
			sTmp = GetGridTxt(.vspdData, C_COST_CD_LEVEL, .vspdData.ActiveRow)
			If 	sTmp <> "" Then
				arrParam(4) = "LEVEL_CD=" & FilterVar(sTmp, "''", "S")
			End If
			arrParam(5) = "Sender C/C" 

			arrField(0) = "ED15" & Parent.gColSep & "CODE"
			arrField(1) = "ED25" & Parent.gColSep & "CD_NM"		
			arrField(2) = "ED15" & Parent.gColSep & "LEVEL_CD"	
    
			arrHeader(0) = "Sender C/C"
			arrHeader(1) = "Sender C/C��"
			arrHeader(2) = "C/C Level"	

		Case C_GP_LEVEL_POP
			arrParam(0) = "�����׷� Level �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('3')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "�����׷� Level" 

			arrField(0) = "LEVEL_CD"	
			arrHeader(0) = "�����׷� Level"	
		
		Case C_GP_CD_POP
			iWidth = 640
			arrParam(0) = "�����׷� �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('4')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			
			sTmp = GetGridTxt(.vspdData, C_GP_LEVEL, .vspdData.ActiveRow)
			If 	sTmp <> "" Then
				arrParam(4) = "LEVEL_CD=" & FilterVar(sTmp, "''", "S")
			End If
			arrParam(5) = "�����׷�" 

			arrField(0) = "ED15" & Parent.gColSep & "CODE"
			arrField(1) = "ED25"  & Parent.gColSep &"CD_NM"		
			arrField(2) = "ED15"  & Parent.gColSep &"LEVEL_CD"	

			arrHeader(0) = "�����׷�"
			arrHeader(1) = "�����׷��"
			arrHeader(2) = "�����׷� Level"

		Case C_ACCT_CD_POP
			iWidth = 800
			arrParam(0) = "���� �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('5')"	
			arrParam(2) = strCode
			arrParam(3) = ""	

			sTmp = GetGridTxt(.vspdData, C_GP_CD, .vspdData.ActiveRow)
			If 	sTmp <> "" Then
				arrParam(4) = "TEMP_CD1=" & FilterVar(sTmp, "''", "S")
			End If


			sTmp = GetGridTxt(.vspdData, C_GP_LEVEL, .vspdData.ActiveRow)
			If 	sTmp <> "" Then
				IF GetGridTxt(.vspdData, C_GP_CD, .vspdData.ActiveRow) <> "" Then
					arrParam(4) = arrParam(4) & " AND LEVEL_CD=" & FilterVar(sTmp, "''", "S")
				ELSE
					arrParam(4) = arrParam(4) & " LEVEL_CD=" & FilterVar(sTmp, "''", "S")
				END IF 
			End If

			arrParam(5) = "����" 

			arrField(0) = "ED15" & Parent.gColSep & "CODE"
			arrField(1) = "ED25" & Parent.gColSep & "CD_NM"					
    
			arrHeader(0) = "����"	
			arrHeader(1) = "������"
							
		Case C_ACTL_DSTB_FCTR_CD_POP
			arrParam(0) = "��ο��1 �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('6')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "��ο��1" 

			arrField(0) = "CODE"
			arrField(1) = "CD_NM"		
    
			arrHeader(0) = "��ο��1"	
			arrHeader(1) = "��ο�Ҹ�1"
			
		Case C_ACTL_DSTB_FCTR_CD_POP2
			arrParam(0) = "��ο��2 �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('6')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "��ο��2" 

			arrField(0) = "CODE"
			arrField(1) = "CD_NM"		
    
			arrHeader(0) = "��ο��2"	
			arrHeader(1) = "��ο�Ҹ�2"
			
		Case C_ACTL_DSTB_FCTR_CD_POP3
			arrParam(0) = "��ο��3 �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('6')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "��ο��3" 

			arrField(0) = "CODE"
			arrField(1) = "CD_NM"		
    
			arrHeader(0) = "��ο��3"	
			arrHeader(1) = "��ο�Ҹ�3"
		Case C_ACTL_DSTB_FCTR_CD_POP4
			arrParam(0) = "��ο��4 �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('6')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "��ο��4" 

			arrField(0) = "CODE"
			arrField(1) = "CD_NM"		
    
			arrHeader(0) = "��ο��4"	
			arrHeader(1) = "��ο�Ҹ�4"
		Case C_ACTL_DSTB_FCTR_CD_POP5
			arrParam(0) = "��ο��5 �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('6')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "��ο��5" 

			arrField(0) = "CODE"
			arrField(1) = "CD_NM"		
    
			arrHeader(0) = "��ο��5"	
			arrHeader(1) = "��ο�Ҹ�5"									
			

		Case C_STD_DSTB_FCTR_CD_POP
			arrParam(0) = "ǥ�ؿ�����ο�� �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('6')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "ǥ�ؿ�����ο��" 

			arrField(0) = "CODE"
			arrField(1) = "CD_NM"		
    
			arrHeader(0) = "��ο��"	
			arrHeader(1) = "��ο�Ҹ�"

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=" & CStr(iWidth) & "px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

	End With
End Function


Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim sTmp
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		Select Case iWhere
		
			Case C_COST_CD_LEVEL_POP
				.Col = C_COST_CD_LEVEL	: .Text = arrRet(0)
		
				Call vspdData_Change(C_COST_CD_LEVEL, .ActiveRow)
				
			Case C_SENDER_COST_CD_POP
				.Col = C_SENDER_COST_CD	: .Text = arrRet(0)
				.Col = C_SENDER_COST_NM	: .Text = arrRet(1)
				.Col = C_COST_CD_LEVEL	: .Text = arrRet(2)
				
				If arrRet(0) = "ALL" Then
					Call ChangeColorByAll(C_COST_CD_LEVEL, frm1.vspdData.ActiveRow, True)
				Else
					Call ChangeColorByAll(C_COST_CD_LEVEL, frm1.vspdData.ActiveRow, False)
				End If

			Case C_GP_LEVEL_POP
				.Col = C_GP_LEVEL	: .Text = arrRet(0)
			
			Case C_GP_CD_POP
				.Col = C_GP_CD			: .Text = arrRet(0)
				.Col = C_GP_NM			: .Text = arrRet(1)
				.Col = C_GP_LEVEL		: .Text = arrRet(2)
				
				.Col = C_ACCT_CD		: .Text = ""
				.Col = C_ACCT_NM		: .Text = ""
			
			Case C_ACCT_CD_POP
				.Col = iWhere - 1		: .Text = arrRet(0)
				.Col = iWhere + 1		: .Text = arrRet(1)
			Case C_ACTL_DSTB_FCTR_CD_POP, C_STD_DSTB_FCTR_CD_POP,C_ACTL_DSTB_FCTR_CD_POP2,C_ACTL_DSTB_FCTR_CD_POP3,C_ACTL_DSTB_FCTR_CD_POP4,C_ACTL_DSTB_FCTR_CD_POP5
				.Col = iWhere - 1		: .Text = arrRet(0)
				.Col = iWhere + 1		: .Text = arrRet(1)
				
				
		End Select
		
		Call vspddata_Change(.ActiveCol, .ActiveRow)
		
		lgBlnFlgChgValue = True
	End With
	
End Function

' -- �׸���2���� �˾� Ŭ���� 
Function OpenPopUp2(Byval iWhere, Byval strCode, Byval strCode1)
	Dim arrRet, sTmp
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	Select Case iWhere
		Case C_RECV_COST_CD_LEVEL_POP
			arrParam(0) = "C/C Level �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('7')"	
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = ""
			arrParam(5) = "C/C Level" 

			arrField(0) = "LEVEL_CD"	
    
			arrHeader(0) = "C/C Level"	
			arrHeader(1) = "Sender C/C"
			arrHeader(2) = "Sender C/C��"
			
		Case C_RECV_COST_CD_POP
			arrParam(0) = "Receiver C/C �˾�"
			arrParam(1) = "dbo.ufn_c_getListOfPopup_C4002MA1('8')"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			
			sTmp = GetGridTxt(.vspdData2, C_RECV_COST_CD_LEVEL, .vspdData2.ActiveRow)
			If 	sTmp <> "" Then
				arrParam(4) = "LEVEL_CD=" & FilterVar(sTmp, "''", "S")
			End If
			arrParam(5) = "Sender C/C" 

			arrField(0) = "ED15" & Parent.gColSep & "CODE"
			arrField(1) = "ED25" & Parent.gColSep & "CD_NM"		
			arrField(2) = "ED15" & Parent.gColSep & "LEVEL_CD"	
    
			arrHeader(0) = "Receiver C/C"
			arrHeader(1) = "Receiver C/C��"
			arrHeader(2) = "C/C Level"	

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp2(arrRet, iWhere)
	End If	

	End With
End Function

Function SetPopUp2(Byval arrRet, Byval iWhere)
	Dim sTmp
	
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		Select Case iWhere
		
			Case C_RECV_COST_CD_LEVEL_POP
				.Col = C_COST_CD_LEVEL	: .Text = arrRet(0)
			
				Call vspdData2_Change(C_RECV_COST_CD_LEVEL, .ActiveRow)
			Case C_RECV_COST_CD_POP
				.Col = C_SENDER_COST_CD	: .Text = arrRet(0)
				.Col = C_SENDER_COST_NM	: .Text = arrRet(1)
				.Col = C_COST_CD_LEVEL	: .Text = arrRet(2)
				
				If arrRet(0) = "ALL" Then
					Call ChangeColorByAll(C_COST_CD_LEVEL, frm1.vspdData2.ActiveRow, True)
				Else
					Call ChangeColorByAll(C_COST_CD_LEVEL, frm1.vspdData2.ActiveRow, False)
				End If

		End Select
		
		Call vspddata2_Change(.ActiveCol, .ActiveRow)
		
		lgBlnFlgChgValue = True
	End With
	
End Function

Function SetVersion(byval arrRet, Byval iWhere)
	Select Case iWhere
		Case 0
			frm1.txtVER_CD.focus
			frm1.txtVER_CD.Value    = arrRet(0)		
		Case 1
			IF LayerShowHide(1) = False Then
				Exit Function
			END IF

			Dim strVal
	
			With frm1
				strVal = BIZ_COPY_PGM_ID & "?txtMode=" & Parent.UID_M0001
				strVal = strVal & "&txtVER_CD=" & Trim(.txtVER_CD.value)	
				strVal = strVal & "&hCopyVerCd=" & arrRet(0)
				
				Call RunMyBizASP(MyBizASP, strVal)
   
			End With
	End Select
    
End Function

' -- ���������� 
Function GetGridTxt(Byref pObj, Byval pCol, Byval pRow)
	With pObj
		.Col = pCol	: .Row = pRow
		GetGridTxt = Trim(.Text)
	End With
End Function

' -- �� ���� 
Function GetGridVal(Byref pObj, Byval pCol, Byval pRow)
	With pObj
		.Col = pCol	: .Row = pRow
		GetGridVal = .Value
	End With
End Function

Sub SetGridTxt(Byref pObj, Byval pCol, Byval pRow, Byval pVal)
	With pObj
		.Col = pCol	: .Row = pRow
		.Text = pVal
	End With
End Sub

Sub SetGridVal(Byref pObj, Byval pCol, Byval pRow, Byval pVal)
	With pObj
		.Col = pCol	: .Row = pRow
		.Value = pVal
	End With
End Sub
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet
    Call InitVariables
    
'	Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("111011010011111")	
    frm1.txtVER_CD.focus
   	Set gActiveElement = document.activeElement			    
     
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'==========================================================================================
'   Event Desc : Grid�� Max Count �� ã�´�.
'==========================================================================================
Function MaxSpreadVal(Byref objSpread, ByVal intCol, byval Row)

	Dim iRows, iMax, iTmp, iMaxRows
	Dim strFrom 
	Dim strWhere ,sSeqNo
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim sVerCd

	If lgIntFlgMode = Parent.OPMD_CMODE	Then
		sVerCd = Trim(frm1.txtVER_CD.value)
	Else
		sVerCd=Trim(frm1.hVerCd.value)
	End IF
	
	strFrom = " C_MFC_DSTB_RULE_BY_CC_S  "
	strWhere = " ver_cd= "		& filterVar(sVerCd,"","S")
		
	Call CommonQueryRs(" max(seq_no)  ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	With objSpread
		iMaxRows = .MaxRows
		IF Len(lgF0) < 1 Then 				
			.Row	= Row
			.Col	= intCol
			.text	= iMaxRows+1
			MaxSpreadVal=iMaxRows+1
			
		Else
			sSeqNo = split(lgF0,chr(11))
			MaxSpreadVal= cdbl(sSeqNo(0))+lgRowCnt
		End If
	End With

end Function

Function MaxSpreadVal2(Byref objSpread, ByVal iSeqNo, ByVal intCol, byval Row)

	Dim iRows, iMax, iTmp, iMaxRows, iTmp2

	iMax = 0

	With objSpread
		.ReDraw = False
		iMaxRows = .MaxRows
		
		For iRows = 1 to  iMaxRows
			.row = iRows
		    .col = intCol

			If iRows <> Row Then
				If .Text = "" Then
				   iTmp = 0
				Else
				   iTmp = UNICDbl(.value)	' -- intCol�� �ִ밪 
				End If

				.Col = intCol -1 : iTmp2 = UNICDbl(.value)
			
				If iTmp > iMax And iTmp2 = iSeqNo Then	' �θ��ȣ�̸鼭 �ڽ��ִ��ȣ ã�� 
				   iMax = iTmp
				End If
			End If
		Next

		iMax	= iMax + 1
		.Row	= Row
		.Col	= intCol
		.Value	= iMax
	
		MaxSpreadVal2 = iMax
		.ReDraw = True
	End With
end Function

'==========================================================================================
'   Event Desc : Grid�� Max Count �־��ش� 
'==========================================================================================
Function InsertSeqNo(Byval objSpread, Byval pSeqNo, Byval pCol, Byval pRow1, Byval pRow2)
	Dim iRow
	With objSpread
		For iRow = pRow1 To pRow2
			.Row = iRow	: .Col = pCol	: .text = pSeqNo : 	.Col = C_ACTL_DSTB_FCTR_RATE : .Value = 100
					
			pSeqNo = pSeqNo + 1
			
		Next
	End With
end Function

'==========================================================================================
'   Event Desc : Grid2�� Seq_no, Sub_Seq_no�� �׿� �ʵ带 �������ش�.
'==========================================================================================
Function InsertDefaultValToGrid2(Byval pSeqNo, Byval pSubSeqNo, Byval pCol, Byval pRow1, Byval pRow2)

	Dim sCCLvl, sCCCd, sGPCd, sAcctCd,sDiFlag ,iRow, iMaxRows, iSeqNo
	With frm1.vspdData
		.Row = .ActiveRow	
		.Col = C_COST_CD_LEVEL	: sCCLvl	= Trim(.text)
		.Col = C_SENDER_COST_CD	: sCCCd		= Trim(.text)
		.Col = C_GP_CD			: sGPCd		= Trim(.text)
		.Col = C_ACCT_CD		: sAcctCd	= Trim(.text)
		.Col = C_DI_FLAG		: sDiFlag	= Trim(.text)
	End With
	
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		
		.ReDraw = False
		For iRow = pRow1 To pRow2

			.Row = iRow	: .Col = pCol	: .Value = pSeqNo
			.Row = iRow	: .Col = pCol+1	: .Value = pSubSeqNo
		
			.Col = C_COST_CD_LEVEL_PARENT	: .Text = sCCLvl
			.Col = C_SENDER_COST_CD_PARENT	: .Text = sCCCd
			.Col = C_GP_CD_PARENT			: .Text = sGPCd
			.Col = C_ACCT_CD_PARENT			: .Text = sAcctCd
			.Col = C_DI_FLAG_PARENT			: .Text = sDiFlag	
		Next
		.ReDraw = True
	End With

end Function

'==========================================================================================
'   Event Desc : �׸��� ���̱�/����� 
'==========================================================================================
Function ShowRowHidden(Byval pSeqNo)
	Dim iRow, iMaxRows, iFirstRow, iSeqNo
	
	With frm1.vspdData2 	
		iMaxRows = .MaxRows : iFirstRow = 0
	
		.ReDraw = False	
		For iRow = 1 To iMaxRows	
			.Row = iRow	: .Col = C_SEQ_NO	: iSeqNo = Trim(.Value)			 
			If iSeqNo = pSeqNo Then
				.RowHidden = False
				If iFirstRow = 0 Then iFirstRow = iRow
			Else
				.RowHidden = True
			End If 
		Next	
		.ReDraw = True	
		ShowRowHidden = iFirstRow
	
	End With
	
End Function

'==========================================================================================
'   Event Desc : 2�� �׸��� ��ü ���� ��ƾ : FncCancel �� 
'==========================================================================================
Function CancelChildGrid2()
	Dim iCol, iRow, iMaxRows, iSeqNo, lDelRows, sFlag, iChildSeqNo
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		frm1.vspdData.Col = C_SEQ_NO : frm1.vspdData.Row = frm1.vspdData.ActiveRow: iSeqNo = UNICDbl(frm1.vspdData.text)

		For iRow = iMaxRows To 1 Step -1
			ggoSpread.Source = frm1.vspdData2
			.Col = C_SEQ_NO : .Row = iRow
			If UNICDbl(.value) = iSeqNo Then	' �θ������ ���� �� 
				.Col = 0 : sFlag = .Text
				.Col = C_SUB_SEQ_NO : iChildSeqNo = UNICDbl(.text)
				.SetActiveCell 3, iRow
				lDelRows = ggoSpread.EditUndo
			End If
		Next
		
	End With
End Function

'==========================================================================================
'   Event Desc : 2�� �׸��� ��ü ���� ��ƾ : FncDelete �� 
'==========================================================================================
Function DeleteChildGrid2()
	Dim iCol, iRow, iMaxRows, sSeqNo, sSeqNo2, lDelRows, sFlag, iChildSeqNo, i, iSelBlockRow, iSelBlockRow2
	
	With frm1.vspdData2
		iMaxRows = .MaxRows		' -- �ڽ� ���� 
		iSelBlockRow	= frm1.vspdData.SelBlockRow
		iSelBlockRow2	= frm1.vspdData.SelBlockRow2 
		
		' -- ��Ƽ �θ� ���� seq_no ��� 
		For i = iSelBlockRow To iSelBlockRow2
			frm1.vspdData.Col = C_SEQ_NO : frm1.vspdData.Row = i : sSeqNo = sSeqNo & CStr(UNICDbl(frm1.vspdData.text)) & "|"
		Next

		ggoSpread.Source = frm1.vspdData2
		For iRow = 1 To iMaxRows
			
			.Col = C_SEQ_NO : .Row = iRow : sSeqNo2 = CStr(UNICDbl(.text)) & "|"
			
			If Instr(1, sSeqNo, sSeqNo2) > 0 Then	' �θ������ ���� �� 
				.Col = 0 : sFlag = .Text	' -- ������� üũ 
				
				.Col = C_SUB_SEQ_NO : iChildSeqNo = UNICDbl(.text)
				If sFlag <> ggoSpread.DeleteFlag  Then	' ������ �ƴҰ�쿡��..
					.SetActiveCell 3, iRow
						
					lDelRows = ggoSpread.DeleteRow
				End If
			End If
		Next

	End With
End Function

'==========================================================================================
'   Event Desc : ��α�Ģ ����Ȯ�� ��ư Ŭ���� 
'==========================================================================================
Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE,EBR_RPT_ID,EBR_RPT_ID2
	Dim StrUrl  , i

	Dim intCnt,IntRetCD


    If Not chkField(Document, "1") Then					'��: This function check indispensable field
       Exit Function
    End If
    

    StrUrl = StrUrl & "VER_CD|"			& frm1.txtVER_CD.value 

     ObjName = AskEBDocumentName("C4002MA1", "ebr")
     
     if  strPrintType = "VIEW" then
		Call FncEBRPreview(ObjName, StrUrl)
     else
		Call FncEBRPrint(EBAction,ObjName,StrUrl)
     end if	
     
End Function 
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	lgCurrGrid = GRID_1
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData

    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
	lgCurrGrid = GRID_2
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData2

    
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	Dim sTemp
	
	
    If Row <> NewRow And NewRow > 0 Then
    
		Dim iLastRow	' -- ���̴� ������ �� 
		Dim iSeqNo
	
	
		With frm1.vspdData 
			.Row = NewRow

			' -- �׸���1�� Ű�� 		
			.Col = C_SEQ_NO			: iSeqNo = Trim(.text)
			.Col =0							: sTemp = Trim(.Text)
			' -- �׸���2�� �׸���1�� Ű���� �´� �ุ ���̰� �Ѵ�.
			iLastRow = ShowRowHidden(iSeqNo)
			
			If lgErrRow <> 0 Then iLastRow = lgErrRow
			frm1.vspdData2.SetActiveCell C_RECV_COST_CD, iLastRow
			'frm1.vspdData2.Focus

			If iLastRow = 0  and not( lgIntFlgMode = Parent.OPMD_CMODE) Then
				Call DBQuery2(iSeqNo)
			End If
				
		End With
    
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
	
	lgCurrGrid = GRID_1
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
	
	lgCurrGrid = GRID_2
End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim sFromSQL, sWhereSQL, sVal, sCd, sCdNm, sTmp, sLvl
	
	sFromSQL = " dbo.ufn_c_getListOfPopup_C4002MA1"
	
	With frm1.vspdData
		.Row = Row	: .Col = Col : sVal = UCase(Trim(.Value))
		
		Select Case Col
		
			Case C_COST_CD_LEVEL	' -- c/c ���� 
				sFromSQL = sFromSQL & "('1')" 
				sWhereSQL = "LEVEL_CD = " & FilterVar(sVal, "''", "S")
				
			Case C_SENDER_COST_CD	' -- c/c
				sFromSQL = sFromSQL & "('2')" 
				
				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")
				
				sTmp = GetGridTxt(frm1.vspdData, C_COST_CD_LEVEL, Row)
				If sTmp <> "" Then
					sWhereSQL = sWhereSQL & " AND LEVEL_CD = " & FilterVar(sTmp, "''", "S")
				End If

			Case C_GP_LEVEL
				sFromSQL = sFromSQL & "('3')" 
				sWhereSQL = "LEVEL_CD = " & sVal

			Case C_GP_CD
				sFromSQL = sFromSQL & "('4')" 

				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")
				

			Case C_ACCT_CD
				sFromSQL = sFromSQL & "('5')" 

				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")
				

				

			Case C_ACTL_DSTB_FCTR_CD, C_STD_DSTB_FCTR_CD,C_ACTL_DSTB_FCTR_CD2,C_ACTL_DSTB_FCTR_CD3,C_ACTL_DSTB_FCTR_CD4,C_ACTL_DSTB_FCTR_CD5
				sFromSQL = sFromSQL & "('6')" 

				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")
				

				
		End Select
	
		If sWhereSQL <> "" Then
			' -- DB �� 
			If CommonQueryRs(" TOP 1 CODE, CD_NM, LEVEL_CD ", sFromSQL , sWhereSQL, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				sCd		= Replace(lgF0, Chr(11), "")
				sCdNm	= Replace(lgF1, Chr(11), "")
				sLvl	= Replace(lgF2, Chr(11), "")
				
				' -- ����� �ڵ���� ����Ѵ�.
				Select Case Col
					Case C_SENDER_COST_CD
						.Col = Col + 2	
						.Text = sCdNm

						.Col = Col - 2	
						.Text = sLvl

						If sLvl = "0" Then
							Call ChangeColorByAll(C_COST_CD_LEVEL, Row, True)
						Else
							Call ChangeColorByAll(C_COST_CD_LEVEL, Row, False)
						End If
					Case C_GP_CD
						.Col = Col + 2	: .Text = sCdNm
						.Col = Col - 2	: .Text = sLvl
						
						.Col = Col + 3	: .Text = ""
						.Col = Col + 5	: .Text = ""

					Case C_ACCT_CD, C_ACTL_DSTB_FCTR_CD, C_STD_DSTB_FCTR_CD,C_ACTL_DSTB_FCTR_CD2,C_ACTL_DSTB_FCTR_CD3,C_ACTL_DSTB_FCTR_CD4,C_ACTL_DSTB_FCTR_CD5
						.Col = Col + 2	
						.Text = sCdNm

					Case C_COST_CD_LEVEL
						If sVal = "0" Then
							.Col = Col + 2	
							.Text = sCd
							.Col = Col + 4	
							.Text = sCdNm
							
							Call ChangeColorByAll(Col, Row, True)
						Else
							.Col = Col + 2	
							.Text = ""
							.Col = Col + 4	
							.Text = ""

							Call ChangeColorByAll(Col, Row, False)
						End If	
						
					Case C_GP_LEVEL	' -- 2005-11-23 �߰� 
							.Col = Col + 2	
							.Text = ""
							.Col = Col + 4	
							.Text = ""
									
				End Select
			Else
				' -- ������� �޽��� ó�� 
				If sVal <> "" Then
					Call DisplayMsgBox("970000", "x",sVal,"x")
					Call SetFocusToDocument("M")
					.Focus
				End If
				
				' -- �� ���� ����� 
				Select Case Col
					Case C_SENDER_COST_CD, C_ACCT_CD, C_ACTL_DSTB_FCTR_CD, C_STD_DSTB_FCTR_CD,C_ACTL_DSTB_FCTR_CD2,C_ACTL_DSTB_FCTR_CD3,C_ACTL_DSTB_FCTR_CD4,C_ACTL_DSTB_FCTR_CD5
						.Col = Col		: .Text = ""
						.Col = Col + 2	: .Text = ""
					Case C_COST_CD_LEVEL
						.Col = Col + 2	: .Text = ""
						.Col = Col + 4	: .Text = ""
						
						Call ChangeColorByAll(Col, Row, False)
					Case C_GP_LEVEL
						.Col = Col		: .Text = ""
						.Col = Col + 2	: .Text = ""
						.Col = Col + 4	: .Text = ""
						.Col = Col + 5	: .Text = ""
						.Col = Col + 7	: .Text = ""
					Case C_GP_CD
						.Col = Col		: .Text = ""
						.Col = Col + 2	: .Text = ""
						.Col = Col + 3	: .Text = ""
						.Col = Col + 5	: .Text = ""
				End Select
				
			End If
		
		End If
		
		' -- ������ ���� �׸��忡 �������ش� 
		Select Case Col		' -- ������ �׸���1 �÷� 
			Case C_COST_CD_LEVEL_POP, C_COST_CD_LEVEL, C_SENDER_COST_CD, C_SENDER_COST_CD_POP, C_GP_LEVEL, C_GP_LEVEL_POP, C_GP_CD, C_GP_CD_POP, C_ACCT_CD, C_ACCT_CD_POP,C_DI_FLAG, C_DI_FLAG_NM
				Call ChangeGrid2HiddenByGrid1	
		End Select
	End With
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

Sub vspdData2_Change(ByVal Col, ByVal Row)
    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col
	
	Call CheckMinNumSpread(frm1.vspdData2, Col, Row)

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim sFromSQL, sWhereSQL, sVal, sCd, sCdNm, sTmp, sLvl
	
	sFromSQL = " dbo.ufn_c_getListOfPopup_C4002MA1"
	
	With frm1.vspdData2
		.Row = Row	: .Col = Col : sVal = UCase(Trim(.Value))
		
		Select Case Col
		
			Case C_RECV_COST_CD_LEVEL	' -- c/c ���� 
				sFromSQL = sFromSQL & "('7')" 
				sWhereSQL = "LEVEL_CD = " & FilterVar(sVal, "''", "S")
				
			Case C_RECV_COST_CD	' -- c/c
				sFromSQL = sFromSQL & "('8')" 
				
				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")
				
				sTmp = GetGridTxt(frm1.vspdData2, C_COST_CD_LEVEL, Row)
				If sTmp <> "" Then
					sWhereSQL = sWhereSQL & " AND LEVEL_CD = " & FilterVar(sTmp, "''", "S")
				End If

		End Select

		If sWhereSQL <> "" Then	
			' -- DB �� 
			If CommonQueryRs(" TOP 1 CODE, CD_NM, LEVEL_CD ", sFromSQL , sWhereSQL, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				sCd		= Replace(lgF0, Chr(11), "")
				sCdNm	= Replace(lgF1, Chr(11), "")
				sLvl	= Replace(lgF2, Chr(11), "")
				
				' -- ����� �ڵ���� ����Ѵ�.
				Select Case Col
					Case C_RECV_COST_CD
						.Col = Col + 2	
						.Text = sCdNm

						.Col = Col - 2	
						.Text = sLvl

						If sLvl = "0" Then
							Call ChangeColorByAll(C_RECV_COST_CD_LEVEL, Row, True)
						Else
							Call ChangeColorByAll(C_RECV_COST_CD_LEVEL, Row, False)
						End If

					Case C_RECV_COST_CD_LEVEL
						If sVal = "0" Then
							.Col = Col + 2	
							.Text = sCd
							.Col = Col + 4	
							.Text = sCdNm
							
							Call ChangeColorByAll(Col, Row, True)
						Else
							.Col = Col + 2	
							.Text = ""
							.Col = Col + 4	
							.Text = ""

							Call ChangeColorByAll(Col, Row, False)
						End If					
				End Select
			Else
				' -- ������� �޽��� ó�� 
				If sVal <> "" Then
					Call DisplayMsgBox("970000", "x",sVal,"x")
					Call SetFocusToDocument("M")
					.Focus
				End If
				
				' -- �� ���� ����� 
				Select Case Col
					Case C_RECV_COST_CD
						.Col = Col		: .Text = ""
						.Col = Col + 2	: .Text = ""
					Case C_RECV_COST_CD_LEVEL
						.Col = Col + 2	: .Text = ""
						.Col = Col + 4	: .Text = ""
						
						Call ChangeColorByAll(Col, Row, False)
				End Select
				
			End If
		End If		
	End With
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

' -- �׸��� 1�� ���� �׸���2�� ����Ǿ�� �Ұ� 
Function ChangeGrid2HiddenByGrid1()
	Dim sCCLvl, sCCCd, sGPCd, sAcctCd, sDiFlag,iRow, iMaxRows, iSeqNo
	With frm1.vspdData
		.Row = .ActiveRow	
		.Col = C_COST_CD_LEVEL	: sCCLvl	= Trim(.text)
		.Col = C_SENDER_COST_CD	: sCCCd		= Trim(.text)
		.Col = C_GP_CD			: sGPCd		= Trim(.text)
		.Col = C_ACCT_CD		: sAcctCd	= Trim(.text)
		.Col = C_DI_FLAG		: sDiFlag	= Trim(.text)
		.Col = C_SEQ_NO			: iSeqNo	= UNICDbl(.text)
	End With
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData2
		
		.ReDraw = False
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_SEQ_NO
			If UNICDbl(.text) = iSeqNo And .RowHidden = False Then	' -- �θ�Ű���̰� ���̴� �ุ 
				.Col = C_COST_CD_LEVEL_PARENT	: .Text = sCCLvl
				.Col = C_SENDER_COST_CD_PARENT	: .Text = sCCCd
				.Col = C_GP_CD_PARENT			: .Text = sGPCd
				.Col = C_ACCT_CD_PARENT			: .Text = sAcctCd
				.Col = C_DI_FLAG_PARENT			: .Text = sDiFlag
				
				ggoSpread.UpdateRow iRow
			End If
		Next
		.ReDraw = True
	End With
End Function

' -- ���� 0 ���ý� �ڵ�/�ڵ���� All/??�� ���ϹǷ� ���� �ϰų� Ǭ��.
Function ChangeColorByAll(Byval pCol, Byval pRow, Byval pBlnLock)
	Select Case pCol
		Case C_COST_CD_LEVEL, C_RECV_COST_CD_LEVEL
			If pBlnLock Then
				ggoSpread.SSSetProtected	pCol + 2	,pRow		,pRow    
			Else
				ggoSpread.SpreadUnLock		pCol + 2	,pRow		,pCol + 2	,pRow
				ggoSpread.SSSetRequired		pCol + 2	,pRow		,pRow    
			End If
			
	End Select
End Function

' -- �ڵ� ����Ÿ ���� üũ 
Function CheckExists(Byval pFromSQL, Byval pWhereSQL)

	CheckExists = False
	
	If CommonQueryRs(" TOP 1 1 ", pFromSQL , pWhereSQL, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		If Trim(lgF0) = "" Then Exit Function	' -- �����Ѵٰ� ���Դµ�, ����Ÿ�� ���ٸ� False�� ���� 
		
		' -- ����� True �� ���� 
		CheckExists = True
	End If

	' -- ������� False ����	
	
End Function

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData

	    ggoSpread.Source = frm1.vspdData
		
		If Row = 0 Then Exit Sub
	
		.Row = Row
			
		Select Case Col
			Case C_GP_LEVEL
				.Col = Col
				index = .Value
				
				.Col = C_SENDER_COST_NM
				.Value = index
				
			Case C_DI_FLAG_NM
				.Col = Col							
				index = .value				
				.Col = C_DI_FLAG					
				.value = index
		
		End Select
        
	End With
	
End Sub


' -- �׸���1 �˾� ��ư Ŭ�� 
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim sCode, sCode2
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_COST_CD_LEVEL_POP, C_SENDER_COST_CD_POP, C_GP_LEVEL_POP, C_GP_CD_POP, C_ACCT_CD_POP, C_ACTL_DSTB_FCTR_CD_POP, C_STD_DSTB_FCTR_CD_POP,C_ACTL_DSTB_FCTR_CD_POP2,C_ACTL_DSTB_FCTR_CD_POP3,C_ACTL_DSTB_FCTR_CD_POP4,C_ACTL_DSTB_FCTR_CD_POP5
				.vspdData.Col = Col - 1
				.vspdData.Row = Row
				
				sCode = UCase(Trim(.vspdData.Text))
				
				Call OpenPopup(Col, sCode, sCode2)
		End Select
        Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
End Sub

' -- �׸���2 �˾� ��ư Ŭ�� 
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim sCode, sCode2
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData2
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_RECV_COST_CD_LEVEL_POP, C_RECV_COST_CD_POP
				.vspdData2.Col = Col - 1
				.vspdData2.Row = Row
				
				sCode = UCase(Trim(.vspdData2.Text))
				
				Call OpenPopup2(Col, sCode, sCode2)
		End Select
        Call SetActiveCell(.vspdData2,Col-1,.vspdData2.ActiveRow ,"M","X","X")   	
    End With
    
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKey <> "" Then
	      	DbQuery
    	End If

    End if
    
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD , blnChange1, blnChange2
    
    FncQuery = False
    
    Err.Clear
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
	ggoSpread.Source = frm1.vspdData
    blnChange1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    blnChange2 = ggoSpread.SSCheckChange
    
    If blnChange1 = True Or blnChange2 = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

    Call InitVariables 	
    Call SetToolbar("1110110100101111")

    IF DbQuery = False Then
		Exit Function
	END IF
       
    FncQuery = True		
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False 
    
    Err.Clear     

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")   
			If IntRetCD = vbNo Then
				Exit Function
			End If
    End If
    
	Call SetToolbar("111011010011111")
	
    Call ggoOper.ClearField(Document, "1") 
    Call ggoOper.ClearField(Document, "2") 
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
     
    Call ggoOper.LockField(Document, "N") 
    Call InitVariables 
    Call SetDefaultVal
    
    FncNew = True 

End Function

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

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave() 
    Dim IntRetCD , blnChange1, blnChange2, iRow, iSeqNo
    
    FncSave = False
    
    Err.Clear     

    ggoSpread.Source = frm1.vspddata
    blnChange1 = ggoSpread.SSCheckChange
    
    ggoSpread.Source = frm1.vspddata2
    blnChange2 = ggoSpread.SSCheckChange
    
    If blnChange1 = False And blnChange2 = False Then	' -- �Ѵ� �̼��� 
        IntRetCD = DisplayMsgBox("900001","x","x","x")  
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspddata
    If Not ggoSpread.SSDefaultCheck Then      
       Exit Function
    End If

    ggoSpread.Source = frm1.vspddata2
    iRow = 0
    If Not ggoSpread.SSDefaultCheck(,iRow) Then      
		If iRow <> 0 Then
			lgErrRow = iRow
			frm1.vspdData2.Row = iRow
			frm1.vspdData2.Col = C_SEQ_NO
			'iSeqNo = UNICDbl(frm1.vspdData2.Value)
			iSeqNo = UNICDbl(frm1.vspdData2.text)
			Call ClickGrid1(iSeqNo)
		End If
		Exit Function
    End If
    
    ' -- �űԸ���϶��� Version �� �ʼ��Է��̴�.
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
		If Trim(frm1.txtVER_CD.value) = "" Then
			Call DisplayMsgBox("970021","x",frm1.txtVER_CD.alt,"x")  
			Exit Function
		End If
    End If
    
    IF DbSave = False Then
		Exit function
	END IF

    FncSave = True      
    
End Function

' --- �׸��� 1 �� C_SEQ_NO�� ���� pSeqNo �̸� Ŭ�����ش� 
Function ClickGrid1(Byval pSeqNo)
	Dim iRow, iMaxRows
	
	With frm1.vspdData
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			.Row = iRow	: .Col = C_SEQ_NO
			If UNICDbl(.text) = pSeqNo Then
				.Col = C_SENDER_COST_CD	: .Action = 0
				Call vspdData_Click(frm1.vspdData.ActiveCol, iRow)
				Exit Function
			End If
		Next
	End With
End Function

' --- �׸��� 2 �� C_SEQ_NO�� ���� C_SUB_SEQ_NO���� pSeqNo, pSubSeqNo �̸� �׸���1�� pSeqNo�� Ŭ�����ش� 
Function ClickGrid2(Byval pSeqNo, Byval pSubSeqNo)
	Dim iRow, iMaxRows, iSeqNo, iSubSeqNo
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			.Row = iRow	: .Col = C_SEQ_NO		: iSeqNo	= UNICDbl(.value)
			.Row = iRow	: .Col = C_SUB_SEQ_NO	: iSubSeqNo = UNICDbl(.value)
			
			If iSeqNo = pSeqNo And iSubSeqNo = pSubSeqNo Then
				Call ClickGrid1(pSeqNo)
				Exit Function
			End If
		Next
	End With
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 
	Dim iSeqNo, iSubSeqNo, iOldCol
	
    if frm1.vspdData.maxrows = 0 then exit function 

	With frm1	
		Select Case lgCurrGrid
			Case GRID_1
				frm1.vspdData.ReDraw = False
				
				iOldCol = .vspdData.ActiveCol
				ggoSpread.Source = frm1.vspdData	
				ggoSpread.CopyRow
				SetSpreadColor frm1.vspdData.ActiveRow ,frm1.vspdData.ActiveRow

				iSeqNo = MaxSpreadVal(.vspdData, C_SEQ_NO, .vspdData.ActiveRow)
				
				frm1.vspdData.ReDraw = True

				lgCurrGrid = GRID_2
				Call FncInsertRow(1)
				lgCurrGrid = GRID_1
				Call vspdData_ScriptLeaveCell(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow-1, frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow, False)
				frm1.vspdData.SetActiveCell iOldCol, .vspdData.ActiveRow
				.vspdData.focus
				
			Case GRID_2
				.vspdData.Col = C_SEQ_NO : .vspdData.Row = .vspdData.ActiveRow : iSeqNo = UNICDbl(.vspdData.Value)
				
				frm1.vspdData2.ReDraw = False
				
				ggoSpread.Source = frm1.vspdData2	
				ggoSpread.CopyRow
				SetSpreadColor2 frm1.vspdData2.ActiveRow ,frm1.vspdData2.ActiveRow

				iSubSeqNo = MaxSpreadVal2(.vspdData2, iSeqNo, C_SUB_SEQ_NO, .vspdData2.ActiveRow)
				Call InsertDefaultValToGrid2(iSeqNo, iSubSeqNo, C_SEQ_NO, .vspdData2.ActiveRow, .vspdData2.ActiveRow)
							
				frm1.vspdData2.ReDraw = True
		End Select	
    End With
End Function


Function FncCancel() 
    Dim lDelRows, sTmp

	Select Case lgCurrGrid 
		CAse  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 

				' -- ���� �׸��� ���� ����� 
				lgCurrGrid = 2 : Call CancelChildGrid2()
				lgCurrGrid = 1
				
				ggoSpread.Source = frm1.vspdData 
				lDelRows = ggoSpread.EditUndo
				Call vspdData_ScriptLeaveCell(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow-1, frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow, False)
					
			End With
		CAse 2

			sTmp = GetGridTxt(frm1.vspdData, 0, frm1.vspdData.ActiveRow)	
			
			If sTmp = ggoSpread.DeleteFlag Then 
				Exit Function
			End If
		
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				lDelRows = ggoSpread.EditUndo
			End With    
	End Select
	
	lgBlnFlgChgValue = True
End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD, iSeqNo, iSubSeqNo, iOldCol, sTmp
    Dim imRow
    
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

	With frm1

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	If lgCurrGrid = GRID_2 And frm1.vspdData.MaxRows = 0 Then lgCurrGrid = GRID_1
	
	Select Case lgCurrGrid
		Case GRID_1
			iOldCol = .vspdData.ActiveCol
			.vspdData.focus
			
			ggoSpread.Source = .vspdData
			ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
			
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
			lgRowCnt=lgRowCnt+1
			iSeqNo = MaxSpreadVal(.vspdData, C_SEQ_NO, .vspdData.ActiveRow)
			
			Call InsertSeqNo(.vspdData, iSeqNo, C_SEQ_NO, .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1)

			If imRow = 1 Then
				lgCurrGrid = GRID_2
				Call FncInsertRow(1)
				lgCurrGrid = GRID_1
			End If
			
			'Call vspdData_Click(iOldCol, .vspdData.ActiveRow)
			Call vspdData_ScriptLeaveCell(iOldCol, .vspdData.ActiveRow+1, iOldCol, .vspdData.ActiveRow, False)
			frm1.vspdData.SetActiveCell iOldCol, .vspdData.ActiveRow
			.vspdData.focus
		Case GRID_2

			sTmp = GetGridTxt(.vspdData, 0, .vspdData.ActiveRow)	
			
			If sTmp = ggoSpread.DeleteFlag Then 
				Exit Function
			End If
		
			' -- �θ�׸�����  �������� seq_no�� �о�´�.
			.vspdData.Col = C_SEQ_NO : .vspdData.Row = .vspdData.ActiveRow : iSeqNo = UNICDbl(.vspdData.text)
			
			.vspdData2.ReDraw = False
			.vspdData2.focus
			ggoSpread.Source = .vspdData2
			ggoSpread.InsertRow  .vspdData2.ActiveRow, imRow
			SetSpreadColor2 .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1
			
			iSubSeqNo = MaxSpreadVal2(.vspdData2, iSeqNo, C_SUB_SEQ_NO, .vspdData2.ActiveRow)
			Call InsertDefaultValToGrid2(iSeqNo, iSubSeqNo, C_SEQ_NO, .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1)
			
			.vspdData2.ReDraw = True

	End Select		
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	End With
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
End Function


Function FncDeleteRow() 
    Dim lDelRows

	Select Case lgCurrGrid 
		Case  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 

				lDelRows = ggoSpread.DeleteRow					
				lgCurrGrid = 2 : Call DeleteChildGrid2()
				lgCurrGrid = 1
					
			End With
		Case 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				lDelRows = ggoSpread.DeleteRow				

			End With    
	End Select
	
	lgBlnFlgChgValue = True
End Function
Function FncPrint()
    Call parent.FncPrint() 
End Function

Function FncPrev() 
End Function

Function FncNext() 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
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

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_CMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtVER_CD=" & Trim(.txtVER_CD.value)	
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
			strVal = strVal & "&txtVER_CD=" & Trim(.hVerCd.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If
		strVal = strVal & "&WhoQuery=H"		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With    
    DbQuery = True
End Function


Function DbQuery2(Byval pSeqNo) 
	Dim strVal

    DbQuery2 = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
'    If Not chkField(Document, "1") Then
'		CAll LayerShowHide(0)
 '      Exit Function
  '  End If
    Err.Clear	
    
    With frm1    
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtVER_CD=" & Trim(.hVerCd.value)
		strVal = strVal & "&SeqNo=" & pSeqNo
		strVal = strVal & "&WhoQuery=D"		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With
    
    DbQuery2 = True

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
	
	If lgCopyVersion = "" Then
		lgIntFlgMode = Parent.OPMD_UMODE
    
		'Call ggoOper.LockField(Document, "Q")

		Call SetToolbar("111111110011111")
	Else
		' -- Ÿ����ī�� 
		Call ggoOper.ClearField(Document, "1") 
		'Call ggoOper.LockField(Document, "N")
		
		lgIntFlgMode = Parent.OPMD_CMODE
		
		Call SetToolbar("111011010011111")
		
		' -- �׸��带 ��� �Է����� �ٲ۴�.
		Call ChangeNewFlag(frm1.vspdData)
		Call ChangeNewFlag(frm1.vspdData2)
	End If

   	Call SetSpreadLock
   	Call SetSpreadLock2
	Frm1.vspdData.Focus
	
	Dim iMaxRows, i

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
		.ReDraw = False
		iMaxRows = .MaxRows
		For i = 1 To iMaxRows
			.Row = i
			.Col = C_SENDER_COST_CD
			If .Text = "ALL" Then
				Call ChangeColorByAll(C_COST_CD_LEVEL, i, True)
			End If
		Next
		.ReDraw = True
	End With

''	With frm1.vspdData2
''		ggoSpread.Source = frm1.vspdData2
		
''		.ReDraw = False
''		iMaxRows = .MaxRows
''		For i = 1 To iMaxRows
''			.Row = i
''			.Col = C_RECV_COST_CD
''			If .Text = "ALL" Then
''				Call ChangeColorByAll(C_RECV_COST_CD_LEVEL, i, True)
''			End If
''		Next
''		.ReDraw = True
''	End With
    Set gActiveElement = document.ActiveElement   
   	Call vspdData_ScriptLeaveCell(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow-1, frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow, False)
End Function

Function DBQueryOk2()
	Call vspdData_ScriptLeaveCell(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow-1, frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow, False)
End Function

' -- ī�ǹ����� �ʱ�ȭ��� 
Function ChangeNewFlag(Byref pObj)
	Dim iRow, iMaxRows
	
	With pObj
		.ReDraw = False
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = 0
			.Text = ggoSpread.InsertFlag
		Next
		.ReDraw = True
	End With
End Function

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
End Sub

Function CheckNullIs0(Byval pVal)
	If pVal = "" Then
		CheckNullIs0 = "0"
	Else
		CheckNullIs0 = pVal
	End If
End Function

Function CheckNullIsX(Byval pVal)
	If pVal = "" Then
		CheckNullIs0 = "*"
	Else
		CheckNullIs0 = pVal
	End If
End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave() 

	On Error Resume Next
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
    Dim iColSep 
    Dim iRowSep  
    Dim sSQLI1, sSQLI2, sSQLU1, sSQLU2, sSQLD1, sSQLD2, sVerCd, tmpA, tmpB, Rate
	
    DbSave = False 
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	

	lGrpCnt = 1
	strVal = ""
	strDel = ""
	
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		sVerCd = UCase(Trim(frm1.txtVER_CD.value))
	Else
		sVerCd = UCase(Trim(frm1.hVerCd.value))
		frm1.txtVER_CD.value = sVerCd
	End If


	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		
		For lRow = 1 To .MaxRows
    
			.Row = lRow	: .Col = 0
			
			Rate = 0 
        
			Select Case .Text

	            Case ggoSpread.InsertFlag	
					strVal = "C" & iColSep 
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
												: strVal = strVal & Trim(sVerCd) & iColSep
					.Col = C_SEQ_NO				: strVal = strVal & Trim(.text) & iColSep
					.Col = C_DSTB_ORDER			: strVal = strVal & Trim(.text) & iColSep
					.Col = C_COST_CD_LEVEL		: strVal = strVal & Trim(.text) & iColSep
					.Col = C_SENDER_COST_CD		: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_GP_LEVEL			: strVal = strVal & Trim(.text) & iColSep
					.Col = C_GP_CD				: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ACCT_CD			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_DI_FLAG			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ACTL_DSTB_FCTR_CD	: strVal = strVal & Trim(.Text) & iColSep						: tmpA=Trim(.Text)	
					.Col = C_ACTL_DSTB_FCTR_RATE : strVal = strVal & Trim(.Text) & iColSep						: Rate= Rate +  CDBL(Trim(.Text))
					.Col = C_ACTL_DSTB_FCTR_CD2	: strVal = strVal & Trim(.Text) & iColSep						
					IF Trim(.Text)  = "" Then
						.Col = C_ACTL_DSTB_FCTR_RATE2 : strVal = strVal &  "0" & iColSep						
					ELSE
						.Col = C_ACTL_DSTB_FCTR_RATE2 : strVal = strVal &  Trim(.Text) & iColSep						: Rate=Rate +   CDBL(Trim(.Text))
					END IF

					.Col = C_ACTL_DSTB_FCTR_CD3	: strVal = strVal & Trim(.Text) & iColSep						
					IF Trim(.Text)  = "" Then
						.Col = C_ACTL_DSTB_FCTR_RATE3 : strVal = strVal &  "0" & iColSep						
					ELSE
						.Col = C_ACTL_DSTB_FCTR_RATE3 : strVal = strVal &  Trim(.Text) & iColSep						: Rate=Rate +   CDBL(Trim(.Text))
					END IF
					
					.Col = C_ACTL_DSTB_FCTR_CD4	: strVal = strVal & Trim(.Text) & iColSep						
					IF Trim(.Text)  = "" Then
						.Col = C_ACTL_DSTB_FCTR_RATE4 : strVal = strVal &  "0" & iColSep						
					ELSE
						.Col = C_ACTL_DSTB_FCTR_RATE4 : strVal = strVal &  Trim(.Text) & iColSep						: Rate=Rate +   CDBL(Trim(.Text))
					END IF
					
					.Col = C_ACTL_DSTB_FCTR_CD5	: strVal = strVal & Trim(.Text) & iColSep						
					IF Trim(.Text)  = "" Then
						.Col = C_ACTL_DSTB_FCTR_RATE5 : strVal = strVal &  "0" & iColSep						
					ELSE
						.Col = C_ACTL_DSTB_FCTR_RATE5 : strVal = strVal &  Trim(.Text) & iColSep						: Rate=Rate +   CDBL(Trim(.Text))
					END IF					
					.Col = C_STD_DSTB_FCTR_CD		: strVal = strVal & Trim(.Text) & iColSep &  Parent.gRowSep	: tmpB=Trim(.Text)	

					sSQLI1 = sSQLI1 + strVal
					lGrpCnt = lGrpCnt + 1
					

	            Case ggoSpread.UpdateFlag		
	            
					strVal = "U" & iColSep 
					
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
												: strVal = strVal & Trim(sVerCd) & iColSep
					.Col = C_SEQ_NO				: strVal = strVal & Trim(.text) & iColSep
					.Col = C_DSTB_ORDER			: strVal = strVal & Trim(.text) & iColSep
					.Col = C_COST_CD_LEVEL		: strVal = strVal & Trim(.text) & iColSep
					.Col = C_SENDER_COST_CD		: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_GP_LEVEL			: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_GP_CD				: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ACCT_CD			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_DI_FLAG			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ACTL_DSTB_FCTR_CD		: strVal = strVal & Trim(.Text) & iColSep						: tmpA=Trim(.Text)	
					.Col = C_ACTL_DSTB_FCTR_RATE : strVal = strVal & Trim(.Text) & iColSep						: Rate= Rate +  CDBL(Trim(.Text))
					.Col = C_ACTL_DSTB_FCTR_CD2	: strVal = strVal & Trim(.Text) & iColSep						
					IF Trim(.Text)  = "" Then
						.Col = C_ACTL_DSTB_FCTR_RATE2 : strVal = strVal &  "0" & iColSep						
					ELSE
						.Col = C_ACTL_DSTB_FCTR_RATE2 : strVal = strVal &  Trim(.Text) & iColSep						: Rate=Rate +   CDBL(Trim(.Text))
					END IF
					.Col = C_ACTL_DSTB_FCTR_CD3	: strVal = strVal & Trim(.Text) & iColSep						
					IF Trim(.Text)  = "" Then
						.Col = C_ACTL_DSTB_FCTR_RATE3 : strVal = strVal &  "0" & iColSep						
					ELSE
						.Col = C_ACTL_DSTB_FCTR_RATE3 : strVal = strVal &  Trim(.Text) & iColSep						: Rate=Rate +   CDBL(Trim(.Text))
					END IF
					.Col = C_ACTL_DSTB_FCTR_CD4	: strVal = strVal & Trim(.Text) & iColSep						
					IF Trim(.Text)  = "" Then
						.Col = C_ACTL_DSTB_FCTR_RATE4 : strVal = strVal &  "0" & iColSep						
					ELSE
						.Col = C_ACTL_DSTB_FCTR_RATE4 : strVal = strVal &  Trim(.Text) & iColSep						: Rate=Rate +   CDBL(Trim(.Text))
					END IF
					.Col = C_ACTL_DSTB_FCTR_CD5	: strVal = strVal & Trim(.Text) & iColSep						
					IF Trim(.Text)  = "" Then
						.Col = C_ACTL_DSTB_FCTR_RATE5 : strVal = strVal &  "0" & iColSep						
					ELSE
						.Col = C_ACTL_DSTB_FCTR_RATE5 : strVal = strVal &  Trim(.Text) & iColSep						: Rate=Rate +   CDBL(Trim(.Text))
					END IF															
					
					.Col = C_STD_DSTB_FCTR_CD		: strVal = strVal & Trim(.Text) & iColSep &  Parent.gRowSep		: tmpB=Trim(.Text)	

					sSQLU1 = sSQLU1 + strVal
					lGrpCnt = lGrpCnt + 1
			

	            Case ggoSpread.DeleteFlag		

					strVal = "D" & iColSep 
					
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
												: strVal = strVal & Trim(sVerCd) & iColSep
					.Col = C_SEQ_NO				: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep
					
					sSQLD1 = sSQLD1 + strVal
					lGrpCnt = lGrpCnt + 1
                
	        End Select

			.Col = 0
			

			

			Select Case .Text

	            Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		
			
					If tmpA="" and tmpB="" then 	
						frm1.vspdData.focus
						.Col = 	C_ACTL_DSTB_FCTR_CD
						.Action = 0				
						Call LayerShowHide(0)
						Call DisplayMsgBox("236324", "X","X","X")
						Exit Function 
					End If
					
					
					If Rate <> 100 Then 
						frm1.vspdData.focus
						.Col = 	C_ACTL_DSTB_FCTR_RATE
						.Action = 0				
						Call LayerShowHide(0)
						Call DisplayMsgBox("236355", "X","X","X")
						Exit Function 
					End If
										
					
			End Select
		Next

	End With

	strDel = "" : strVal = ""
	
	With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2	
		
		For lRow = 1 To .MaxRows
    
			.Row = lRow	: .Col = 0
        
			Select Case .Text

	            Case ggoSpread.InsertFlag	
					strVal = "C" & iColSep 
					
					.Col = .MaxCols					: strVal = strVal & Trim(.text) & iColSep
													: strVal = strVal & Trim(sVerCd) & iColSep
					.Col = C_SEQ_NO					: strVal = strVal & Trim(.text) & iColSep
					.Col = C_SUB_SEQ_NO				: strVal = strVal & Trim(.text) & iColSep
					
					.Col = C_RECV_COST_CD_LEVEL		: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_RECV_COST_CD			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_COST_CD_LEVEL_PARENT	: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_SENDER_COST_CD_PARENT	: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_GP_CD_PARENT			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ACCT_CD_PARENT			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_DI_FLAG_PARENT			: strVal = strVal & Trim(.Text) & iColSep &  Parent.gRowSep

					sSQLI2 = sSQLI2 + strVal
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag		
	            
					strVal = "U" & iColSep 
					
					.Col = .MaxCols					: strVal = strVal & Trim(.text) & iColSep
													: strVal = strVal & Trim(sVerCd) & iColSep
					.Col = C_SEQ_NO					: strVal = strVal & Trim(.text) & iColSep
					.Col = C_SUB_SEQ_NO				: strVal = strVal & Trim(.text) & iColSep
					.Col = C_RECV_COST_CD_LEVEL		: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_RECV_COST_CD			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_COST_CD_LEVEL_PARENT	: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_SENDER_COST_CD_PARENT	: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_GP_CD_PARENT			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ACCT_CD_PARENT			: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_DI_FLAG_PARENT			: strVal = strVal & Trim(.Text) & iColSep &  Parent.gRowSep


					sSQLU2 = sSQLU2 + strVal
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag		

					strVal = "D" & iColSep 
					
					.Col = .MaxCols					: strVal = strVal & Trim(.text) & iColSep
													: strVal = strVal & Trim(sVerCd) & iColSep
					.Col = C_SEQ_NO					: strVal = strVal & Trim(.text) & iColSep
					.Col = C_SUB_SEQ_NO				: strVal = strVal & Trim(.text) & iColSep &  Parent.gRowSep

					sSQLD2 = sSQLD2 + strVal
					lGrpCnt = lGrpCnt + 1
                
	        End Select
                
		Next

	End With
		
		

			
	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtMaxRows.value = lGrpCnt-1

	frm1.txtSpreadI1.value = sSQLI1
	frm1.txtSpreadU1.value = sSQLU1
	frm1.txtSpreadD1.value = sSQLD1
	frm1.txtSpreadI2.value = sSQLI2
	frm1.txtSpreadU2.value = sSQLU2
	frm1.txtSpreadD2.value = sSQLD2
		
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
    DbSave = True    
    
End Function

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()	
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	frm1.vspdData2.MaxRows = 0
	Call MainQuery()
		
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal     & "&txtVER_CD=" & frm1.txtVER_CD.value					    '��: Query Key        
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
					<TD WIDTH=* align=right><a onclick="vbscript:Call OpenVersion(1)">Ÿ Version Copy</a>&nbsp;&nbsp;</TD>
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
									<TD CLASS="TD5">Version</TD>
									<TD CLASS="TD656"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtVER_CD" SIZE=10 MAXLENGTH=3 tag="13XXXU" ALT="Version"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDstbFctr" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenVersion(0)">
									</TD>
								</TR>    
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="50%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="*" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
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
					<TD WIDTH=10></TD>
					<TD WIDTH=10>&nbsp;<BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('VIEW')" Flag=1 style="width: 150">��α�Ģ ����Ȯ��</BUTTON></TD>
					<TD WIDTH=*></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpreadI1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadI2" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadU1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadU2" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadD1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadD2" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hVerCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCopyVerCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

