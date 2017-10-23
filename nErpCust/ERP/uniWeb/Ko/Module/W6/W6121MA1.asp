
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��8ȣ��ǥ3�������鼼�װ�꼭(3)
'*  3. Program ID           : W6121MA1
'*  4. Program Name         : W6121MA1.asp
'*  5. Program Desc         : ��8ȣ��ǥ3�������鼼�װ�꼭(3)
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2005/03/18
'*  8. Modifier (First)     : ȫ���� 
'*  9. Modifier (Last)      : ȫ���� 
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

Const BIZ_MNU_ID		= "W6121MA1"
Const BIZ_PGM_ID		= "W6121MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "W6121MB2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "W6121OA1"

Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.

' -- �׸��� �÷� ���� 
	Dim C_W101			    ' (101)�����ڵ� 
	Dim C_W101_Nm		    ' ���и� 
	Dim C_Law	    	    ' �ٰŹ� ���� 
	Dim C_W102			    ' (102)������ 
    Dim C_CODE			    ' �ڵ� 
	Dim C_W103_AMT	        ' ����(����)�ݾ� 
	Dim C_W103_RATE_VAL	    ' �ڵ�� 
	Dim C_W103_RATE		    ' ������ 
	Dim C_W103		        ' ��곻��	
	Dim C_W104			    ' �������� 
	Dim C_Limit_RATE	    ' �ѵ��� 
	Dim C_Limit_AMT		    ' �ѵ��ݾ� 
	
	Dim C_SEQ_NO        
	Dim C_W105          	   '(105) ����	
	Dim C_W105_Nm          	   '(105) ����	
	Dim C_W106		           '(106) ����⵵	
	Dim C_W107		    	   '(107) ���� 
	Dim C_W108		    	   '(108) �̿��� 
	Dim C_W109		    	   '(109) ���� 
	Dim C_W110		    	   '(110) 1���⵵ 
	Dim C_W111		    	   '(111) 2���⵵ 
	Dim C_W112		    	   '(112) 3���⵵ 
	Dim C_W113		    	   '(113) 4���⵵ 
	Dim C_W114		    	   '(114) �� 
	Dim C_W115		    	   '(115) ������ �����뿡 �ٸ� �̰����� 
	Dim C_W116		    	   '(116) ��������(114-115)
	Dim C_W117		    	   '(117) �Ҹ� 
	Dim C_W118		    	   '(118) �̿���(107 + 108 + 116 - 117)
	

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(6)	' ��Ƽ �׸��� ó�� ���� 
Dim lgblnYoon

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	
	C_W101			= 1   ' (101)�����ڵ� 
	C_W101_Nm		= 2   ' ���и� 
	C_Law	    	= 3   ' �ٰŹ� ���� 
	C_W102			= 4   ' (102)������ 
    C_CODE			= 5   ' �ڵ� 
	C_W103_AMT	    = 6   ' ����(����)�ݾ� 
	C_W103_RATE_VAL	= 7   ' �ڵ�� 
	C_W103_RATE		= 8   ' ������	
	C_W103			= 9   ' �������� 
	C_W104			= 10   ' �������� 
	C_Limit_RATE	= 11  ' �ѵ��� 
	C_Limit_AMT		= 12  ' �ѵ��ݾ� 
	
	C_SEQ_NO        =1
	C_W105          =2	   '(105) ����	
	C_W105_NM       =3	   '(105) ���и�	
	C_W106		    =4     '(106) ����⵵	
	C_W107		    =5	   '(107) ���� 
	C_W108		    =6	   '(108) �̿��� 
	C_W109		    =7	   '(109) ���� 
	C_W110		    =8	   '(110) 1���⵵ 
	C_W111		    =9	   '(111) 2���⵵ 
	C_W112		    =10	   '(112) 3���⵵ 
	C_W113		    =11	   '(113) 4���⵵ 
	C_W114		    =12	   '(114) �� 
	C_W115		    =13	   '(115) ������ �����뿡 �ٸ� �̰����� 
	C_W116		    =14	   '(116) ��������(114-115)
	C_W117		    =15	   '(117) �Ҹ� 
	C_W118		    =16	   '(118) �̿���(107 + 108 + 116 - 117)
	
	

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
	gSelframeFlg = ""
    'lgCurrGrid = TYPE_1
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



Sub InitSpreadComboBox()
	Dim IntRetCD1
	' ��ȸ����(����)
	' �������ױ���  = W1080
	
	IntRetCD1 = CommonQueryRs("reference_2, minor_nm", " dbo.ufn_TB_Configuration('w1080','" & C_REVISION_YM & "') ", " reference_1 <> ''  Order by minor_nm", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  


	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(TYPE_2) 
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W105
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W105_NM
	
    End if
   
   


End Sub


Sub InitSpreadSheet()
	Dim ret, iRow
	
	' �׸��� ���� 
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	
	lgvspdData(TYPE_1).ScriptEnhanced  = True
	lgvspdData(TYPE_2).ScriptEnhanced  = True
	
    Call initSpreadPosVariables()  

	' 1�� �׸���(��1)
	With lgvspdData(TYPE_1)
			
		ggoSpread.Source = lgvspdData(TYPE_1)	
	
					'patch version
					 ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
					 
						.ReDraw = false

					    .MaxCols = C_Limit_AMT + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
						.Col = .MaxCols														'��: ����� �� Hidden Column
						.ColHidden = True    
						       
						 .MaxRows = 0
						 ggoSpread.ClearSpreadData
						.rowheight(0) = 28
						 Call AppendNumberPlace("6","3","2")


	
	
						 ggoSpread.SSSetEdit     C_W101,			 "�����ڵ�",	 15,,,100,1
						 ggoSpread.SSSetEdit     C_W101_Nm,			 "(101)����",	 20,,,100,1
						 ggoSpread.SSSetEdit     C_Law,				 "�ٰŹ�����", 12,,,100,1
						 ggoSpread.SSSetEdit     C_W102,			 "(102)������",  30,,,100,1
						 ggoSpread.SSSetEdit     C_CODE,			 "�ڵ�",  4,2,,15,1

						 ggoSpread.SSSetFloat    C_W103_AMT,		 "����(����)�ݾ�",			12,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
						 ggoSpread.SSSetCombo    C_W103_RATE_VAL,	 "��������", 8
						 ggoSpread.SSSetCombo    C_W103_RATE,		 "������", 8
						 ggoSpread.SSSetEdit     C_W103,			 "(103)��곻��", 20,,,100,1
						 ggoSpread.SSSetFloat    C_W104,		     "(104)��������",				    12,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
						 ggoSpread.SSSetFloat    C_Limit_RATE,		 "�ѵ���",						15,	    "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
						 ggoSpread.SSSetFloat    C_Limit_AMT,		 "�ѵ��ݾ�",					12,	     Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
						 
						.RowHeight(-1) = 25
						
						Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
						Call ggoSpread.SSSetColHidden( C_W103_RATE_VAL,   C_W103_RATE_VAL, True)
		                Call ggoSpread.SSSetColHidden(C_W101,C_W101,True)
						Call ggoSpread.SSSetColHidden(C_Limit_RATE,C_Limit_RATE,True)
	
		
					
						.ReDraw = true

				 
					'Call SetSpreadLock 		
			
	End With 
 
	' 2�� �׸���(��1)
	With lgvspdData(TYPE_2)
	
	


			ggoSpread.Source = lgvspdData(TYPE_2)	
			'patch version
			ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gForbidDragDropSpread    
    
			.ReDraw = false

			.MaxCols = C_W118 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.Col = .MaxCols									'��: ����� �� Hidden Column
			.ColHidden = True    
					       
			.MaxRows = 0
			ggoSpread.ClearSpreadData

			'Call AppendNumberPlace("6","3","2")
				'����� 2�ٷ�    
			.ColHeaderRows = 2
			ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,6,1	' �����÷� 
			ggoSpread.SSSetCombo    C_W105,		"����", 8
			ggoSpread.SSSetCombo    C_W105_NM,	"(105)����", 35
      
			ggoSpread.SSSetDate		C_W106,"(106)�������",	10,		2,		Parent.gDateFormat,	-1
			ggoSpread.SSSetFloat	C_W107,		"(107) ���� "	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W108,		"(108) �̿���"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W109,		"(109) ����"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W110,		"(110) 1���⵵ ", 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W111,		"(111) 2���⵵"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W112,		"(112) 3����"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

			ggoSpread.SSSetFloat	C_W113,		"(113) 4���⵵"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W114,		"(114) ��"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W115,		"(115) ������ ������" & vbCrLf & "�� ���� �̰�����"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W116,		"(116) ��������" & vbCrLf & "(114-115) "	, 12,  Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W117,		"(117) �Ҹ�"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W118,		"(118) �̿���" & vbCrLf & "(107 + 108 - 116 - 117)"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec


		
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_W105,True)

			ret = .AddCellSpan(C_W105_NM		, -1000, 1, 2)	' 
			ret = .AddCellSpan(C_W106			, -1000, 1, 2)	' 
			ret = .AddCellSpan(C_W107			, -1000, 2, 1)	' 
			ret = .AddCellSpan(C_W109			, -1000, 6, 1)
			ret = .AddCellSpan(C_W115		, -1000, 1, 2)	' 
			ret = .AddCellSpan(C_W116		, -1000, 1, 2)	' 
			ret = .AddCellSpan(C_W117		, -1000, 1, 2)	' 
			ret = .AddCellSpan(C_W118		, -1000, 1, 2)	' 
		
			.Row = -1000
			.Col = C_W107	: .Text = "���������"
		
			.Row = -1000
			.Col = C_W109	: .Text = "��������󼼾�"
		
		
			.Row = -999
			.Col = C_W107	: .Text = "(107) ����"
			.Row = -999
			.Col = C_W108	: .Text = "(108) �̿���"
			.Row = -999
			.Col = C_W109	: .Text = "(109) ����"
			.Row = -999
			.Col = C_W110	: .Text = "(110) 1���⵵"
			.Row = -999
			.Col = C_W111	: .Text = "(111) 2���⵵"
			.Row = -999
			.Col = C_W112	: .Text = "(112) 3���⵵"
			.Row = -999
			.Col = C_W113	: .Text = "(113) 4���⵵"
			.Row = -999
			.Col = C_W114	: .Text = "(114) �հ�"
		
	    
	    
			.rowheight(-999) = 12				
			Call SetSpreadLock(TYPE_2)
					
			.ReDraw = true	
			
	End With 
     
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call CheckFISC_DATE
End Sub

Sub SpreadInitData()
    ' �׸��� �ʱ� ����Ÿ���� 
 	
 	
  Dim sFiscYear, sRepType, sCoCd, IntRetCD , iCol ,strSelect,strFrom,strWhere,ii,jj,arrVal1,arrVal2,iRow,iMaxRows,arrVal3
  Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
 		dim dblLimitAmt , dblLimitRATE , dbl3hoW16, strCode
 		
     sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value		
 		
   IntRetCD = CommonQueryRs("W16", " TB_3", "   co_cd = '" & sCoCd & "'and Fisc_year ='" & sFiscYear & "' and Rep_Type = '" & sRepType & "'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  
	If IntRetCD <> False Then
		dbl3hoW16 = replace(lgF0,chr(11),"")  '���μ� 
		frm1.txt3hoW16.value = cdbl(dbl3hoW16)							
	End if

	
	strSelect	=		 " minor_cd, minor_nm, reference_1, reference_3 ,reference_2 ,reference_4 "
	strFrom		=		 " ufn_TB_Configuration('w1080', '" & C_REVISION_YM &"')  "
	strWhere	=		 " minor_cd like '%'  order by cast( minor_cd as int)"

    ggoSpread.Source = lgvspdData(TYPE_1)	
    With  lgvspdData(TYPE_1)
       .Redraw = False	
			If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			Else 
					   arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))	
				 
					   iMaxRows = Ubound(arrVal1,1) 	
 
					   ggoSpread.InsertRow , iMaxRows
					   
					.RowHeight(3) = 35

					For iRow =1 to .MaxRows
									
								 arrVal2 = Split(arrVal1(iRow-1), chr(11))
									
								 .Row = iRow
								 .Col = C_W101 : .Value = Trim(arrVal2(1))
								 .Col = C_W101_Nm : .Value = Trim(arrVal2(2))
								 .TypeEditMultiLine = True : .TypeVAlign = 2	
						     if  .Value  <> "" then
							      ggoSpread.SpreadLock 1, iRow,  C_Code ,iRow
							 end if   
								
								.Col = C_Law : .Value = Trim(arrVal2(3))
								.TypeVAlign = 2	
								.Col = C_W102 : .Value = Trim(arrVal2(4))
								.TypeEditMultiLine = True : .TypeVAlign = 2	
										  	    
								.Col = C_Code : .Value = Trim(arrVal2(5)) : .TypeVAlign = 2	
								strCode = Trim(arrVal2(5))
							    
							    'reference_5
								IntRetCD = CommonQueryRs("reference_1, reference_2, reference_5", " dbo.ufn_TB_Configuration('" & Trim(arrVal2(6)) &"','" & C_REVISION_YM & "' ) Order By reference_1", " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

									If IntRetCD <> False Then
									
									
										Call Spread_SetCombo(TYPE_1 ,Replace( "" & chr(9) & lgF0, chr(11),  chr(9)), C_W103_RATE_VAL, iRow, iRow)
										Call Spread_SetCombo(TYPE_1 ,Replace("" &  chr(9) & lgF1, chr(11),  chr(9)), C_W103_RATE , iRow, iRow)
										    .Col = C_Limit_RATE
										    .Row = irow : .TypeEditMultiLine = True : .TypeVAlign = 2	
										    
										     arrVal3 =Split(lgF2, Chr(11))
										    .Value  = unicdbl(arrVal3(0))
										     ggoSpread.SpreadLock C_W103, iRow, C_W103 ,iRow
										    
									else
									           .Col =  C_W103_RATE_VAL : .CellType = 1
									           .Col =  C_W103_RATE : .CellType = 1
									           ggoSpread.SpreadLock C_W103_RATE_VAL, iRow, C_W103_RATE ,iRow
									         
									    	 
									End if
					            if  strCode = "15" then
								    Call AppendNumberPlace("6","3","2")
								    ggoSpread.SSSetFloat    C_W103_AMT,		 "(103)����(����)�ݾ�",	12,	   "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z","0",,iRow
								End if
								'�ѵ��� 
								  .Col = C_Limit_RATE	: .Row = iRow	: dblLimitRATE	 = Trim(.TEXT)
								
								if dblLimitRATE = 0 then
								   .Col = C_Limit_AMT	: .Row = iRow	: .text 	 = ""
								else  
		    	          		   dblLimitAmt = unicdbl(dbl3hoW16) * unicdbl(dblLimitRATE)
		    	          
		    	        		  .Col = C_Limit_AMT	: .Row = iRow	: .Value 	 = dblLimitAmt
		    	        		end if  
					
					
					           .Col =  C_W103_RATE :  .TypeEditMultiLine = True : .TypeVAlign = 2 
					           .Col =  C_W103_AMT  :  .TypeEditMultiLine = True : .TypeVAlign = 2 
					           .Col =  C_W103      :  .TypeEditMultiLine = True : .TypeVAlign = 2 
					           .Col =  C_W104      :  .TypeEditMultiLine = True : .TypeVAlign = 2 
					           .Col =  C_Limit_AMT :  .TypeEditMultiLine = True : .TypeVAlign = 2 
					
								' -- DB�÷����ǰ� �۰ԵǾ��־� ����.(���ڽŰ��÷����̴´�����) 2006.02.24���� 
								if  strCode = "13" then
									.Row = iRow
									.Col = C_W101_Nm	: .TypeMaxEditLen = 25
									.Col = C_Law		: .TypeMaxEditLen = 25
									.Col = C_W102		: .TypeMaxEditLen = 50
								End If
					Next		
								
			End If
				   	
			       ggoSpread.SpreadLock C_CODE, -1, C_CODE ,-1
			       ggoSpread.SpreadLock C_Limit_AMT, -1, C_Limit_AMT ,-1
			       
			       
		.Redraw = True		

	End With 
	
	
	
	Call SetSpreadLock(TYPE_1)
	Call SetSpreadTotalLine
End Sub

Sub Spread_SetCombo(itype,pVal, pCol1, pRow1, pRow2)

	With  lgvspdData(itype)

		.BlockMode = True
		.Col = pCol1	: .Col2 = pCol1
		.Row = pRow1	: .Row2 = pRow2
		.CellType = 8	'SS_CELL_TYPE_COMBOBOX

		.TypeComboBoxList = pVal	

		.TypeComboBoxEditable = False
		.TypeComboBoxMaxDrop = 3
		' Select the first item in the list
		'.TypeComboBoxCurSel = 0
		' Set the width to display the widest item in the list
		'.TypeComboBoxWidth = 1 
		.BlockMode = False
	End With

End Sub
Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	If pType = TYPE_2 Then	'
	    ggoSpread.SSSetRequired     C_W114 , -1, C_W114
	    ggoSpread.SSSetRequired     C_W118 , -1, C_W118
	Else

	  ggoSpread.SSSetRequired     C_W104 , -1, C_W104
	End If
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = lgvspdData(pType)

	If pType = TYPE_2 Then	'
	   
	    ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W114, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W118, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W116, pvStartRow, pvEndRow 
		ggoSpread.SSSetRequired   C_W106 , pvStartRow, pvEndRow 
	    ggoSpread.SSSetRequired  C_W105_NM , pvStartRow, pvEndRow 
	Else
	
		If pvEndRow = lgvspdData(pType).MaxRows Then
		
		End If
	End If

End Sub


Sub SetSpreadColor2()
   dim iRow
	ggoSpread.Source = lgvspdData(TYPE_2)


	   
	    ggoSpread.SSSetProtected C_SEQ_NO, -1, -1 
		ggoSpread.SSSetProtected C_W114, -1, -1 
		ggoSpread.SSSetProtected C_W114, -1, -1 
		ggoSpread.SSSetProtected C_W118,  -1, -1 
		ggoSpread.SSSetProtected C_W116, -1, -1 
		ggoSpread.SSSetRequired   C_W106 ,  -1, -1 
	    ggoSpread.SSSetRequired  C_W105_NM , -1, -1 
	
	   for iRow = 1 to lgvspdData(TYPE_2).MaxRows
	         lgvspdData(TYPE_2).Row = iRow
	         lgvspdData(TYPE_2).Col = C_SEQ_NO
	       if  lgvspdData(TYPE_2).Text= "999999" then
	          ggoSpread.SSSetProtected -1, iRow, iRow
	       end if
	       
	        lgvspdData(TYPE_2).Row = iRow
	        lgvspdData(TYPE_2).Col = C_W105_NM
	      
	      if  lgvspdData(TYPE_2).CellType= 1 then
	          ggoSpread.SSSetProtected  C_W105_NM, iRow, iRow
	       end if      
	       
	   Next 
	


End Sub




Sub SetSpreadTotalLine()
	Dim iRow
	For iRow = TYPE_1 To TYPE_2
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W101		: .CellType = 1		: .TypeHAlign = 2
				ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
			End If
		End With
	Next

End Sub 

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
           
    End Select    
End Sub

'============================== ���۷��� �Լ�  ========================================

Function GetRef2()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �������� ���� : �޽�����������.
	wgRefDoc = CheckTaxDoc(sCoCd, sFiscYear,sRepType, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, "�������װ��", "X")           '��: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
	 ggoSpread.Source = lgvspdData(TYPE_2)
	 ggoSpread.ClearSpreadData

	Call LayerShowHide(1)

	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
	
End Function



Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW2,arrW3, arrW4,arrW5, iMaxRows, sTmp,jj,arrW6
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf

	' ����� ��ġ�� �˷��� 
	Dim iCol, iRow
	

	 ggoSpread.Source = lgvspdData(TYPE_1)

     With lgvspdData(TYPE_1)
	   .Redraw = False	
	   .AddSelection C_W104, 1, C_W104, .maxrows' -- �������� ������ �߰��Ҷ� 
	

	
		IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
		Call ggoOper.LockField(Document, "N") 
		.SetSelection iCol, 1, iCol, 1
		
		If IntRetCD = vbNo Then
			 Exit Function
		End If
	.Redraw = True
	End With



	IntRetCD = CommonQueryRs("  WCODE,  W_AMT  , WRate_VAL ,  WRate , W_TAXAMT "," dbo.ufn_TB_8_3_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		arrW3		= Split(lgF2, chr(11))
		arrW4		= Split(lgF3, chr(11))
		arrW5		= Split(lgF4, chr(11))
		
		iMaxRows	= UBound(arrW1)

		ggoSpread.Source = lgvspdData(TYPE_1)

		With lgvspdData(TYPE_1)
		
				For iRow = 1 To .Maxrows -1

						For   jj = 0 to iMaxRows
		
						    .Row = iRow :.Col = C_CODE
						    if    trim(.Value)  =  Trim(arrW1(jj)) then  
						          .Row = iRow :.Col = C_W104 
						          .Col = C_W103_AMT       : .value = arrW2(jj)
								   
						          .Col = C_W103_RATE_VAL	 : .TEXT  = arrW3(jj)
						          .Col = C_W103_RATE		 : .TEXT = arrW4(jj)
						          .Col = C_W103				 : .TEXT =  formatnumber(arrW2(jj)) & "x" & arrW4(jj)
		             	          .Col = C_W104				 : .TEXT = arrW5(jj)   
						           
								Call vspdData_Change(TYPE_1, C_W104, iRow)

						    end  if
						NEXt
				Next
		
		End With
		
		'Call SetReCalc1
	End If
	
	lgBlnFlgChgValue = True
	lgvspdData(TYPE_1).focus
	
	
End Function

' ���۷������� �־����Ƿ� �Է����� ��ȯ�� �ش�.
Function ChangeRowFlg(Index)
	Dim iRow
	
	With lgvspdData(Index) 
		ggoSpread.Source = lgvspdData(Index)
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
		Next
	End With
End Function

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
	
	lgblnYoon = False
	datMonCnt = DateDiff("m", sFISC_START_DT, sFISC_END_DT)
	' ���� ������ ���Ⱓ�ȿ� ������ �ִ��� üũ�ؼ� lgblnYOON�� ��ȭ��Ų��.
	For i = 1 To datMonCnt
		datNow = DateAdd("m", i, sFISC_START_DT)
		If Month(datNow) = 2 Then	' 2���� ������ ���Ⱓ�̸� 
			lgblnYoon = CheckIntercalaryYear(Year(datNow))
			Exit For
		End If
	Next
End Sub

'====================================== �� �Լ� =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	lgCurrGrid = TYPE_1	' �⺻ �׸��� 
	'Call ShowTabLInk(TAB1)

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>
	Else
		Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
	End If
	gSelframeFlg = TAB1
	lgCurrGrid = TYPE_1
	Call ShowTabLInk(TAB1)
End Function

Function ClickTab2()	
	' Tab1 ���� üũ�� �̻������ ���� 
	If Not ChkChgTab Then Exit Function

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>
	Else
		Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>
	End If
	
	
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	lgCurrGrid = TYPE_2
	Call ShowTabLInk(TAB2)
	
	
End Function

Function ShowTabLInk(pType)
	Dim pObj1, i
	Set pObj1 = document.all("myTabRef")

	
	if pType = TAB1 then
		pObj1(2).style.display = "none"
		pObj1(1).style.display = ""
	Else	
       pObj1(1).style.display = "none"
       pObj1(2).style.display = ""
	End if
	

End Function

Function ChkChgTab()
	ChkChgTab = False
	' 1. ���� ���� �ε� üũ 
	With lgvspdData(TYPE_1)
	         .Row = 1
		     .Col = 0
		If  .text = ggoSpread.InsertFlag   Then
			Call DisplayMsgBox("W60002", "X", "�������װ����", "X")                          <%'No data changed!!%>
			Exit Function
		End If
	End With
	ChkChgTab = True
End Function




Function GridReCalc()
	
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
	Call InitComboBox

	Call InitSpreadComboBox

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData 
	

	'Call SpreadInitData
	'Call ClickTab1	
    Call FncQuery 
    
End Sub


'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
    
    Call CheckFISC_DATE
End Sub

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	Call CheckFISC_DATE
End Sub

'============================================  �׸��� �̺�Ʈ   ====================================
' -- 0�� �׸��� 
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

' -- 1�� �׸��� 
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

Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row )
	lgCurrGrid = TYPE_2
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_ComboSelChange(ByVal Col, ByVal Row )

	lgCurrGrid = TYPE_1
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)
dim iIdx,iRow
    ggoSpread.Source = lgvspdData(Index)

    With lgvspdData(Index)
        if Index = TYPE_2 then
   				Select Case Col
						Case C_W105_Nm 
				            
							.Col = Col	: .Row = Row
						
							iIdx = UNICDbl(.Value)
					       	.Col = Col -1
							.Value = iIdx
						
							Call SetGubn(Index,.Activerow)
						
			
					End Select
					
		else
		    	Select Case Col
						Case C_W103_RATE
				            
							.Col = Col	: .Row = Row
						
							iIdx = UNICDbl(.Value)
					       	.Col = Col -1
							.Value = iIdx
						
							
			
					End Select
					
		     			
		END IF			
	End With		
	

End Sub


Function SetGubn(Index,iRow)
Dim strGubun,strGubunNm
 With lgvspdData(Index)
    .Col = C_W105_NM
	 strGubunNm = .text 
	 .Col = C_W105
	 strGubun = .text 
	.Row = iRow + 1
	   Do While Not .Text ="99"				
	      .Col = C_W105
	      .Text = strGubun
	      .Col = C_W105_NM : .TypeMaxEditLen  = 100
					       
	     .Text = strGubunNm 
	      Call vspdData_Change(Index ,iRow ,C_W105)
	      iRow=  iRow  + 1
	      .Row  = iRow
	     .Col = C_W105
				         
	  Loop

  End With
End Function


Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, datW1_DOWN, datW1, iRow, iMaxRows
	Dim dbl103Rate, dbl103Amt,  dbl103Rate_val, dblLimitRATE, dblLimitAMT , dbl3hoW16 ,IntRetCD1 ,dblW114 , dblW115 ,dblW107,dblW108,dblW116,dblW117 ,IntRetCD 
    Dim sCoCd , sFiscYear , sRepType	,strW106 , strW105
    Dim strCode
	lgBlnFlgChgValue= True ' ���濩�� 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col
	
    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If uniCDbl(lgvspdData(Index).text) < uniCDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row
   
    sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value

	' --- �߰��� �κ� 
	With lgvspdData(Index)
	     if Index = Type_1 then
	        Select Case Col
				  
				
				    Case C_W103_RATE	, C_W103_AMT 
				    
				         .Col = C_W103_RATE		: .Row = Row	: dbl103Rate	 = Trim(.TEXT)
				         .Col = C_W103_RATE_VAL	: .Row = Row	: dbl103Rate_val = unicdbl(.TEXT)
		    	         .Col = C_W103_AMT		: .Row = Row	: dbl103Amt		 = Trim(.TEXT)
		    	         .Col = C_Limit_AMT  	: .Row = Row	: dblLimitAMT	 = Trim(.TEXT)
		    	         .Col = C_Limit_Rate	: .Row = Row	: dblLimitRATE	 = unicdbl(.TEXT)

		    	         if  dbl103Amt <> "" and dbl103Rate <> "" then
		    	         
		    	       
		    	           
		    							if unicdbl(dblLimitAMT) >  unicdbl(dbl103Amt) *  unicdbl(dbl103Rate_val) or  (unicdbl(dblLimitAMT) = 0  or  unicdbl(dblLimitRATE) = 0 ) then    '�ѵ��׺��� ũ�� 
		    	         
										  .Col = C_W103  		: .Row = Row	: .Value =  dbl103Amt & " x " &  dbl103Rate
										  .Col = C_W104  		: .Row = Row	: .Value =  Fix(unicdbl(dbl103Amt) *  unicdbl(dbl103Rate_val))
										else
										   .Col = C_W103  		: .Row = Row	: .Value =  Trim( UNICDBL(frm1.txt3hoW16.value) ) & " x " &  dblLimitRATE * 100 & "%"
										   .Col = C_W104  		: .Row = Row	: .Value =  unicdbl(dblLimitAmt)
				             
										 end if  
				          
				         Else  
							lgvspdData(lgCurrGrid).Row = Row
							lgvspdData(lgCurrGrid).Col = Col
							dblSum = .Value
				            .Col = C_W103  		: .Row = Row	: .Value = ""
				            .Col = C_W104	 	: .Row = Row	: .Value =  dblSum 
				         end if
				         
				          Call FncSumSheet(lgvspdData(Index), C_W103_AMT, 1, .MaxRows-1, true, .MaxRows, C_W103_AMT, "V")	' ���� ���� �հ� 
				          Call FncSumSheet(lgvspdData(Index), C_W104, 1 , .MaxRows-1, true, .MaxRows, C_W104, "V")	' ���� ���� �հ� 
				           ggoSpread.UpdateRow .MaxRows
				    
				     Case C_W104      
				         Call FncSumSheet(lgvspdData(Index), C_W104, 1 , .MaxRows-1, true, .MaxRows, C_W104, "V")	' ���� ���� �հ� 
				          ggoSpread.UpdateRow .MaxRows
				End Select
	     
	     else
	     
				Select Case Col
				
				    Case C_W105_NM 
				            .COL = C_W106 : .ROW = ROW  : strW106 =  .text
				            .COL = C_W105 : .ROW = ROW  : strW105 =  .text
				            
				         if strW106 <> "" then   
								IntRetCD = CommonQueryRs("C_W118", " TB_8_3_B", "   co_cd = '" & sCoCd & "'and Fisc_year ='" & sFiscYear-1 & "' and Rep_Type = '1' and left(W106,4) = '"& left(strW106,4)-1 &"' and W105 ='"& strW105 &"'  " , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  
									If IntRetCD <> False Then
																
										 dblW108 = unicdbl(lgF0)  '�̿��� 
										 IF dblW108 > 0 Then
										    .COL = C_W108 : .ROW = ROW  :.TEXT = dblW108
										 End if   
										
									End if
						end if			
				    
				    
			
					Case C_W106	' ������ ����� 
						iMaxRows = .MaxRows
								
						' 1. ���� �Է��� �������� �������� �����ຸ�� ũ�� ������ ����Ų��.
						If Row + 1 <> iMaxRows Then
							.Row = Row		: .Col = C_W106	: datW1 = uniCDate(.Text)
									
							' 1.1 �Ʒ����� ���� ��� 
							.Row = Row+1	: .Col = C_W106	
							If .Text <> "" and   .Text <> "�Ұ�" Then
								datW1_DOWN = uniCDate(.Text)

								If datW1 > datW1_DOWN Then ' �Ʒ��ຸ�� ��¥�� ���ĸ� ���� 
									Call DisplayMsgBox("WC0016", parent.VB_INFORMATION, "X", "X")           '��: "Will you destory previous data"
									Exit Sub						
								End If
							End If

						End If
					
					   Call vspdData_Change(Index,C_W105_NM,Row)  
					Case C_W107
						.Col = C_W105_NM : .Row = Row 
						If .Text = "" Then
							Call DisplayMsgBox("X", parent.VB_INFORMATION, "(105)������ ���� �Է��Ͻʽÿ�.", "X")	
							.Col = Col : .Row = Row : .Value = 0		    
							Exit Sub
						end If
						
						.Row = Row		: .Col = C_W107	: dblW107 = uniCdbl(.Text)
						.Row = Row		: .Col = C_W109	: .Value = unicdbl(dblW107)
				
						Call SetColSum(Index, Col)
						Call vspdData_Change(Index,C_W109,Row)	

					Case C_W107 , C_W108 , C_W109, C_W110, C_W111, C_W112 , C_W113  ,  C_W117

						.Col = C_W105_NM : .Row = Row 
						If .Text = "" Then
							Call DisplayMsgBox("X", parent.VB_INFORMATION, "(105)������ ���� �Է��Ͻʽÿ�.", "X")		
							.Col = Col : .Row = Row : .Value = 0	    
							Exit Sub
						end If

						Call SetColSum(Index, Col)

					    Call SetColSum2(Row)
					    Call SetColSum(Index, C_W118)
					
					    Call FncSumSheet(lgvspdData(lgCurrGrid), Row, C_W109, C_W113, true, Row , C_W114, "H")
					    Call vspdData_Change(Index,C_W114,Row)	
					    
					    'Call CheckReCalc()
					Case C_W114 , C_W115
						Call SetColSum(Index, Col)
						
					    Call SetColSum3(Row)
					    Call SetColSum(Index, C_W116)
						Call SetColSum2(Row)
						Call SetColSum(Index, C_W118)
						'Call CheckReCalc()
						
				End Select
			
			
			end if	
	

	End With
	
End Sub

Function SetColSum(index , byval Col)
   
' ���� �÷��� �������� �հ� ����� �� �� ����Ѵ�.

	Dim  dblSum99, dblSumCol, iRow, strGubn,  iMaxRows, iActiveRow

	With lgvspdData(TYPE_2)	' ��Ŀ���� �׸��� 
		
		ggoSpread.Source = lgvspdData(TYPE_2)
		iMaxRows = .MaxRows
		iActiveRow = .ActiveRow
		.Row = .ActiveRow	: .Col = C_W105   
	    strGubn = Trim(.text)                 '�������� ���� 
	
		For iRow = 1 To iMaxRows
		    .Row = 1	: .Col = 0 
		 
		    if .text <> ggoSpread.DeleteFlag then
					.Row = iRow	: .Col = C_W105
	
					If .Text = strGubn Then			 ' ���� ���� 
					     strGubn = .text
						.Col = Col
						dblSumCol = dblSumCol + UNICDbl(.Value)
					ElseIf .Text = "99" And iActiveRow < .Row Then
						' �Ұ� ��� 
						.Col = Col
						.Value = dblSumCol
						ggoSpread.UpdateRow iRow
						
						Call ReCalcGubunSum(Col)		
						 'dblSum99	= dblSum99 + dblSumCol               '���հ� 
					     'dblSumCol = 0
						'.Row = .MaxRows	: .Col = Col	: .Value = dblSum99
						
						'ggoSpread.UpdateRow .MaxRows
						
						'.Row = iRow	+1 : .Col = C_W105    '���� ���� 
						 'strGubn = .text
						Exit For
					End If
			End if
		Next
		
	End With
End Function

Function ReCalcGubunSum(Byval Col)
	Dim  dblSum99, dblSumCol, iRow, strGubn,  iMaxRows, iSeqNo

	With lgvspdData(TYPE_2)	' ��Ŀ���� �׸��� 
		
		ggoSpread.Source = lgvspdData(TYPE_2)
		iMaxRows = .MaxRows
	
		For iRow = 1 To iMaxRows -1
		    .Row = 1	
		    .Col = C_SEQ_NO		: iSeqNo = UNICDbl(.value)
		    .Col = 0 
		 
		    if .text <> ggoSpread.DeleteFlag then
					.Row = iRow	: .Col = C_W105
	
					If .Text = "99" Then			 ' ���� ���� 
						.Col = Col
						dblSumCol = dblSumCol + UNICDbl(.Value)
					End If
			End if
		Next

		 'dblSum99	= dblSum99 + dblSumCol               '���հ� 
		 'dblSumCol = 0
		.Row = .MaxRows	: .Col = Col	: .Value = dblSumCol
						
		ggoSpread.UpdateRow .MaxRows
		
	End With	
End Function

Function SetColSum2(byval Row)
   Dim dblW107,dblW108,dblW116,dblW117
    	With lgvspdData(TYPE_2) 
			
			.Row = Row		: .Col = C_W107	: dblW107 = uniCdbl(.Text)
			.Row = Row		: .Col = C_W108	: dblW108 = uniCdbl(.Text)
			.Row = Row		: .Col = C_W116	: dblW116 = uniCdbl(.Text)
			.Row = Row		: .Col = C_W117	: dblW117 = uniCdbl(.Text)
			.Row = Row		: .Col = C_W118	: .Value = unicdbl(dblW107) + unicdbl(dblW108) - unicdbl(dblW116)- unicdbl(dblW117)
		End With	
End function 

Function SetColSum3(byval Row)
   Dim dblW114,dblW115
    	With lgvspdData(TYPE_2) 
			   .Row = Row		: .Col = C_W114	: dblW114 = uniCdbl(.Text)
			   .Row = Row		: .Col = C_W115	: dblW115 = uniCdbl(.Text)
			   .Row = Row		: .Col = C_W116	: .Value = unicdbl(dblW114) - unicdbl(dblW115)
		End With	
End function 

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(Index)
   
    If lgvspdData(Index).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(Index)
       '
       'If lgSortKey = 1 Then
       '    ggoSpread.SSSort Col               'Sort in ascending
       '    lgSortKey = 2
       'Else
       '    ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
       '    lgSortKey = 1
       'End If
       
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
    Call GetSpreadColumnPos("A")
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
    Call ggoOper.LockField(Document, "N")
   
    ggoSpread.ClearSpreadData
    Call InitVariables													<%'Initializes local global variables%>
                                 
   																
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, i
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>

    
    For i = TYPE_1 To TYPE_2
    
		ggoSpread.Source = lgvspdData(i)
		IF ggoSpread.SSDefaultCheck = False Then								  '��: Check contents area
				Exit Function
		End If
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
		End If
	Next

	If blnChange = False And lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	
	' �����۾� 
	
	With lgvspdData(TYPE_2)
	  If lgvspdData(TYPE_2).MaxRows <> 0 then 
			.Row = .MaxRows : .Col = C_W118
			If unicdbl(.Value) < 0 Then
				Call DisplayMsgBox("WC0013", "X", "(118)�̿���", "X")                          
				Exit Function
			End If 

			.Row = .MaxRows : .Col = C_W116
			If unicdbl(.Value) < 0 Then
				Call DisplayMsgBox("WC0013", "X", "(116)��������", "X")                          
				Exit Function
			End If 
		End if	
	End With	

		
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

'========================================================================================

Sub InsertFirstRow(iRow)
	Dim iMaxRows, iType, ret

	 ' �ϵ��ڵ��Ǵ� ��� 

	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)
		.Redraw = False

	
		'Call SetSpreadLock

		if irow <= 0   then
		    iRow = 0
		    irow = irow +1
		    iMaxRows = 3
		    ggoSpread.InsertRow , iMaxRows
			.Row = iRow	 	
			Call SetSpreadColor(TYPE_2,iRow, iRow)
			.Col = C_SEQ_NO : .Value = iRow		: iRow = iRow + 1
		
			.TypeEditMultiLine = True
			.TypeHAlign = 2 : .TypeVAlign = 2
		
			.Row = iRow		
			.Col = C_SEQ_NO : .Value = "999999"	: iRow = iRow + 1
			.Col = C_W105 : .CellType = 1
			.Col = C_W105_NM : .CellType = 1
			.Col = C_W105	: .Value = "99"
			.Col = C_W106: .CellType = 1	: .Text = "�Ұ�"	: .TypeHAlign = 2	
		    ggoSpread.SpreadLock C_SEQ_NO, .Row, C_W118, .Row
	
			'ret = .AddCellSpan(C_W105	, .Row - 1, 1, 2)
			'ret = .AddCellSpan(C_W105_Nm	, .Row - 1, 1, 2)
	         

			.Row = iRow		
			.Col = C_SEQ_NO : .Value = SUM_SEQ_NO	: iRow = iRow + 1
		
			.Col = C_W105: .CellType = 1	: .Text = "�հ�"	: .TypeHAlign = 2	
			ret = .AddCellSpan(C_W105	, .Row, 3, 1)
			ggoSpread.SpreadLock C_SEQ_NO, .Row, C_W118, .Row
        else
		    iMaxRows = 2
		   	ggoSpread.InsertRow .MaxRows-1, iMaxRows
		   	
		   	.Row = .MaxRows-2	
		   	   Call SetSpreadColor(TYPE_2,.Row, .Row)
			
		       Call  SetSeqNo(TYPE_2, .Row, 1)
		     iRow = iRow + 1  
			.TypeEditMultiLine = True
			.TypeHAlign = 2 : .TypeVAlign = 2
		
			.Row =  .MaxRows -1 				
			.Col = C_SEQ_NO : .Value = "999999"	: iRow = iRow + 1
		     
			.Col = C_W105 : .CellType = 1 : .Value = "99"
	
			.Col = C_W106  : .CellType = 1	: .Text = "�Ұ�"	: .TypeHAlign = 2	
		
			ggoSpread.SpreadLock C_SEQ_NO, .Row, C_W118, .Row
			'ret = .AddCellSpan(C_W105	, .Row - 1, 1, 2)
			'ret = .AddCellSpan(C_W105_Nm	, .Row - 1, 1, 2)
			
			
	
		end if
		

		 
        
		
		.Redraw = True
	
	End With
	

        
        
	'Call SetSpreadLock(iType)
End Sub


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

    'Call SetToolbar("1100100000000111")

	Call ClickTab1()
	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If lgvspdData(lgCurrGrid).MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If lgvspdData(lgCurrGrid).ActiveRow > 0 Then
			lgvspdData(lgCurrGrid).focus
			lgvspdData(lgCurrGrid).ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor lgCurrGrid, lgvspdData(lgCurrGrid).ActiveRow, lgvspdData(lgCurrGrid).ActiveRow

			lgvspdData(lgCurrGrid).Col = C_W21
			lgvspdData(lgCurrGrid).Text = ""
    
			lgvspdData(lgCurrGrid).Col = C_W3
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).Col = C_W4
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).Col = C_W5
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
	Dim strCelltype,iRow
	If lgCurrGrid = TYPE_1 Then Exit Function

    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    lgvspdData(lgCurrGrid).Row =  lgvspdData(lgCurrGrid).ActiveRow

	lgvspdData(lgCurrGrid).Col = C_W105

	strCelltype = lgvspdData(lgCurrGrid).celltype

	lgvspdData(lgCurrGrid).Col = C_SEQ_NO

     if   Trim(lgvspdData(lgCurrGrid).TEXT) <> "999999"  and strCelltype <> 8 then
          ggoSpread.EditUndo   
     elseif    Trim(lgvspdData(lgCurrGrid).TEXT) <> "999999"  and strCelltype = 8 then
		  iRow = lgvspdData(lgCurrGrid).ActiveRow
          Do Until  lgvspdData(lgCurrGrid).text = "999999"
                     ggoSpread.EditUndo
                    'iRow = iRow + 1
                    lgvspdData(lgCurrGrid).Row = iRow
                    lgvspdData(lgCurrGrid).Col = C_SEQ_NO
                    lgvspdData(lgCurrGrid).Action = 0
          Loop          
          ggoSpread.EditUndo
     End if  
    
    
    lgBlnFlgChgValue = True
    Call CheckReCalc()				' �Ѷ����� ��ҵǸ� ���� 

End Function

' ���� 
Function CheckReCalc()
	Dim dblSum
	
    Call SetColSum(TYPE_1,C_W107)
    Call SetColSum(TYPE_1,C_W108)
    Call SetColSum(TYPE_1,C_W109)
    Call SetColSum(TYPE_1,C_W110)
    Call SetColSum(TYPE_1,C_W111)
    Call SetColSum(TYPE_1,C_W112)
    Call SetColSum(TYPE_1,C_W113)
    Call SetColSum(TYPE_1,C_W114)
    Call SetColSum(TYPE_1,C_W115)
    Call SetColSum(TYPE_1,C_W116)
    Call SetColSum(TYPE_1,C_W117)
    Call SetColSum(TYPE_1,C_W118)
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6,strActivCol
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

	If lgCurrGrid = TYPE_2 Then

			   With lgvspdData(lgCurrGrid)	' ��Ŀ���� �׸��� 
					
				ggoSpread.Source = lgvspdData(lgCurrGrid)
					
				iRow = .ActiveRow
				lgvspdData(lgCurrGrid).ReDraw = False
			
				'If iRow = .MaxRows Then Exit Function
			
				If .MaxRows = 0  or .MaxRows =iRow Then	  ' ù InsertRow�� 1��+�հ��� 
	
					Call InsertFirstRow(iRow)
				   
				Else
					.Row = iRow
					.Col = C_SEQ_NO
					If iRow <1 Then iRow = 1

					If .ActiveCol = C_W105_Nm  or .text = "999999"  Then	' �հ� �� 
					        strActivCol = C_W105_Nm
							Call InsertFirstRow(iRow)
					Else
				dim strGubnNm,strGubun
					     strActivCol = C_W106
					   
						   .Row = iRow	
						   .Col = C_W105_NM: strGubnNm = .text	
						   .Col = C_W105: strGubun = .text	
						   	
					    	ggoSpread.InsertRow iRow,imRow
						  .Col = C_W105: .CellType = 1  : .Text = strGubun
					      .Col = C_W105_NM: .CellType = 1 : .TypeMaxEditLen  = 100:.text =strGubnNm 
                          Call  SetSpreadColor(lgCurrGrid,iRow+1,iRow+1)
	   
				
						  sW1_CD = Left(.Value, 1)
						ggoSpread.SpreadLock C_W105, iRow+1, C_W105_NM, iRow +1
						
						Call  SetSeqNo(TYPE_2, iRow+1, pvRowCnt)
					
					End If
					
				End If
				
				lgvspdData(lgCurrGrid).ReDraw = True
			End With
	
	 else
	
		Exit Function		
	 End if			

	
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
		lgvspdData(TYPE_2).SetFocus
        lgvspdData(TYPE_2).Row = iRow
        lgvspdData(TYPE_2).Col = strActivCol
        lgvspdData(TYPE_2).Action = 0
    Set gActiveElement = document.ActiveElement   
    
End Function


Function InsertTotalLine(Index)
	With lgvspdData(Index)
	
	ggoSpread.Source = lgvspdData(Index)
	
	If .MaxRows > 0 Then	' ���� �߰� 
		ggoSpread.InsertRow ,1
		
		.Row = .MaxRows
		.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
		.Col = C_W1		: .CellType = 1	: .Text = "�հ�"	: .TypeHAlign = 2
		
		ggoSpread.SpreadLock C_W1, .MaxRows, C_W5, .MaxRows
	End If
	End With
End Function

' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetSeqNo(Index, iRow, iAddRows)
	
	Dim i, iSeqNo

	With lgvspdData(lgCurrGrid)	' ��Ŀ���� �׸��� 

		ggoSpread.Source = lgvspdData(lgCurrGrid)
	
		If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
			.Row = iRow
			MaxSpreadVal lgvspdData(lgCurrGrid), C_SEQ_NO, iRow
		Else
			iSeqNo = MaxSpreadVal(lgvspdData(lgCurrGrid), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
			
			For i = iRow to iRow + iAddRows -1
			    
				.Row = i
				.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
			Next
		End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows, strCelltype,iRow, iAllDel, iSeqNo

	If lgCurrGrid = TYPE_1 Then Exit Function
	
	iAllDel = True
	
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    lgvspdData(lgCurrGrid).Row =  lgvspdData(lgCurrGrid).ActiveRow

	lgvspdData(lgCurrGrid).Col = C_W105

	strCelltype = lgvspdData(lgCurrGrid).celltype

	lgvspdData(lgCurrGrid).Col = C_SEQ_NO

     if   Trim(lgvspdData(lgCurrGrid).TEXT) <> "999999"  and strCelltype <> 8 then
          lDelRows = ggoSpread.DeleteRow
     elseif    Trim(lgvspdData(lgCurrGrid).TEXT) <> "999999"  and strCelltype = 8 then
		  iRow = lgvspdData(lgCurrGrid).ActiveRow
          Do Until  lgvspdData(lgCurrGrid).text = "999999"
                     lDelRows = ggoSpread.DeleteRow
                    iRow = iRow + 1
                    lgvspdData(lgCurrGrid).Row = iRow
                    lgvspdData(lgCurrGrid).Col = C_SEQ_NO
                    lgvspdData(lgCurrGrid).Action = 0
          Loop          
          lDelRows = ggoSpread.DeleteRow
          
          With lgvspdData(lgCurrGrid)
          ' -- �հ� ���� 
			For iRow = 1 To .MaxRows
				.Row = iRow
				.Col = C_SEQ_NO : iSeqNo = UNICDbl(.value)
				.Col = 0
				If .Text <> ggoSpread.DeleteFlag And iSeqNo <> 999999 Then 
					iAllDel = False
					Exit For
				End If
			Next
			
			If iAllDel Then
				lDelRows = ggoSpread.DeleteRow(.MaxRows)
			End If
			
          End With
     End if  
    
    
    lgBlnFlgChgValue = True
    Call CheckReCalc()				' �Ѷ����� ��ҵǸ� ���� 
	
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
	
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
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
	 Call  SpreadInitData
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key   
        strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	ggoSpread.Source = lgvspdData(TYPE_1)
	
	If lgvspdData(TYPE_1).MaxRows > 0 Or _
		lgvspdData(TYPE_2).MaxRows >0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		'lgIntFlgMode = parent.OPMD_UMODE

		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg <> "Y" Then
			
			Call SetSpreadLock(TYPE_1)
			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			'Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>
		Else
		
			
			'Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
		End If
	   
		Call SetSpreadColor2
		Call SetSpreadTotalLine ' - �հ���� �籸�� 
		ggoSpread.Source = lgvspdData(TYPE_2)
		Call FncSumSheet(lgvspdData(Type_1), C_W103_AMT, 1, lgvspdData(TYPE_1).MaxRows-1, true, lgvspdData(TYPE_1).MaxRows, C_W103_AMT, "V")	' ���� ���� �հ� 
		Call FncSumSheet(lgvspdData(Type_1), C_W104, 1 , lgvspdData(TYPE_1).MaxRows-1, true, lgvspdData(TYPE_1).MaxRows, C_W104, "V")	' ���� ���� �հ� 
				    
	'Else
		'Call SetToolbar("1100110100000111")										<%'��ư ���� ���� %>
	End If
	 
	 
	 if lgCurrGrid  = TYPE_1 then
	    Call ClickTab1()
	 Else
	    Call ClickTab2()
	 End if
	
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
    
    For i = TYPE_1 To TYPE_2	' ��ü �׸��� ���� 
    
		With lgvspdData(i)
	
			ggoSpread.Source = lgvspdData(i)
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
		       End Select
		       
			  ' ��� �׸��� ����Ÿ ����     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = C_SEQ_NO To lMaxCols
						.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
			  End If  
			Next
		
		End With

		If i = TYPE_1 Then
			Frm1.txtSpread.value      = strDel & strVal
			strVal = "" :  strDel = ""
		Else
			Frm1.txtSpread2.value      = strDel & strVal
		End If
	Next

	
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow, iTab											        <%' ���� ������ ���� ���� %>

	Call InitVariables

	For iRow = TYPE_1 To TYPE_2
		lgvspdData(lgCurrGrid).MaxRows = 0
		ggoSpread.Source = lgvspdData(lgCurrGrid)
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()" width=200>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�������װ��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���������� �� �̿��� ���</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><DIV id="myTabRef">&nbsp;</DIV>
						<DIV id="myTabRef" STYLE="display:'none'"><A href="vbscript:GetRef">�ݾ׺ҷ�����</A></DIV>
						<DIV id="myTabRef" STYLE="display:'none'"><A href="vbscript:GetRef2">�ݾ׺ҷ�����</A></DIV>
						</TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="14X1"></SELECT>
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
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=15%>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="25X1" TITLE="SPREAD" id=vaSpread Index=0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</DIV>
						
								<DIV ID="TabDiv" SCROLL=no>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="25X1" TITLE="SPREAD" id=vaSpread Index=1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</DIV>

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
			<TABLE CLASS="TB3" CELLSPACING=0>
			    
		       <TR>
				        <TD WIDTH=10>&nbsp;</TD>
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><����>���������� �� �̿��� ���</LABEL>&nbsp;
				           
				        </TD>
				            
                </TR>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txt3hoW16" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

