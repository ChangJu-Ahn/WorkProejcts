<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : RECEIPT
'*  3. Program ID		    : f5122ma1
'*  4. Program Name         : ���������̵�ó�� 
'*  5. Program Desc         : ���������̵�ó�� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/10/16
'*  8. Modified date(Last)  : 2002/02/15
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Soo Min, Oh
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. �� �� �� 
'##############################################################################################################
'******************************************  1.1 Inc ����   ***************************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->			<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                              '��: indicates that All variables must be declared in advance 

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Const BIZ_PGM_ID  = "f5122mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 = "f5122mb2.asp"											 '��: �����Ͻ� ���� ASP�� : Tab1�� ADO ��ȸ��  
Const BIZ_PGM_ID3 = "f5122mb3.asp"											 '��: �����Ͻ� ���� ASP�� : Tab2�� ADO ��ȸ�� 

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

'TAB1, vspddata
Dim C_PROC_CHK
Dim C_FR_DEPT_CD
Dim C_FR_DEPT_NM
Dim C_NOTE_NO	
Dim C_NOTE_AMT
Dim C_NOTE_STS
Dim C_TO_DEPT_CD
Dim C_TO_DEPT_POP
Dim C_TO_DEPT_NM
Dim C_MOVE_DESC
Dim C_BP_CD	
Dim C_BP_NM	
Dim C_ISSUED_DT
Dim C_DUE_DT	

'TAB2, vspddata2
Dim C_CNCL_CHK	
Dim C_CNCL_TO_DEPT_CD	
Dim C_CNCL_TO_DEPT_NM	
Dim C_CNCL_NOTE_NO	
Dim C_CNCL_NOTE_AMT	
DIm C_CNCL_FR_DEPT_CD
Dim C_CNCL_FR_DEPT_NM
Dim C_CNCL_TEMP_GL_NO
Dim C_CNCL_TEMP_GL_DT
Dim C_CNCL_GL_NO	
Dim C_CNCL_GL_DT	
Dim C_CNCL_BP_CD	
Dim C_CNCL_BP_NM	
Dim C_CNCL_ISSUED_DT
Dim C_CNCL_DUE_DT	

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================

Dim lgBlnFlgConChg				'��: Condition ���� Flag
Dim  gSelframeFlg

Dim lgStrPrevKeyNoteNo	' ���� �� (CG, DG)
Dim lgStrPrevKeyTempGlNo		'���� TEmp Gl ��(DG)
Dim lgStrPrevKeyGlNo    ' ���� GL �� (DG)

Dim lgStrPrevKeyNoteNo1	' ���� �� (CG, DG)
Dim lgStrPrevKeyTempGlNo1		'���� TEmp Gl ��(DG)
Dim lgStrPrevKeyGlNo1    ' ���� GL �� (DG)

Dim IsOpenPop          

Dim lgPageNo1
Dim lstxtPlanAmtSum

'++++++++���� ���� 2002.01.10 �߰� ���� ++++++++++++++++
<%
Dim dtToday 
dtToday = GetSvrDate
%>

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE   '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '��: Indicates that no value changed
    lgIntGrpCount = 0           '��: Initializes Group View Size
    lgPageNo         = 0
	lgPageNo1        = 0
	lgStrPrevKeyNoteNo	= ""
	lgStrPrevKeyGlNo	= ""
	lgStrPrevKeyNoteNo1 = ""
	lgStrPrevKeyGlNo1	= ""
	
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'��: ����� ���� �ʱ�ȭ 
    lgSortKey = 1
    
End Sub

Sub initSpreadPosVariables(ByVal spdsep2)
	Select case spdsep2
		Case "A"
			C_PROC_CHK		= 1
			C_FR_DEPT_CD	= 2
			C_FR_DEPT_NM	= 3        
			C_NOTE_NO		= 4
			C_NOTE_AMT		= 5  
			C_NOTE_STS      = 6  
			C_TO_DEPT_CD	= 7
			C_TO_DEPT_POP	= 8
			C_TO_DEPT_NM	= 9
			C_MOVE_DESC     = 10			
			C_BP_CD			= 11	  
			C_BP_NM			= 12 
			C_ISSUED_DT		= 13              
			C_DUE_DT		= 14     
		Case "B"
			C_CNCL_CHK			= 1
			C_CNCL_TO_DEPT_CD	= 2
			C_CNCL_TO_DEPT_NM	= 3              
			C_CNCL_GL_NO		= 4
			C_CNCL_GL_DT		= 5
			C_CNCL_TEMP_GL_NO	= 6      
			C_CNCL_TEMP_GL_DT	= 7
			C_CNCL_NOTE_NO		= 8
			C_CNCL_NOTE_AMT		= 9
			C_CNCL_FR_DEPT_CD	= 10
			C_CNCL_FR_DEPT_NM	= 11        
			C_CNCL_BP_CD		= 12
			C_CNCL_BP_NM		= 13
			C_CNCL_ISSUED_DT	= 14              
			C_CNCL_DUE_DT		= 15     
	End Select 
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("A","*","NOCOOKIE","BA") %>

End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	Dim strSvrDate
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"

	frDt = UNIDateAdd("M", -1, strSvrDate,Parent.gServerDateFormat)		
	frm1.txtFromDt.Text = UNIConvDateAToB(frDt,parent.gServerDateFormat,parent.gDateFormat)  '�������� Fr
	frm1.txtToDt.Text   = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) '~ To	
	frm1.txtFrGlDt.Text = UniConvDateAToB(frDt,Parent.gServerDateFormat,Parent.gDateFormat)               '�ι�° Tab ȸ������ Fr
	frm1.txtToGlDt.Text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)                      '~ To	
	frm1.txtGLDt.text   = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)  '�̵��� 
	
	frm1.hOrgChangeId.value = Parent.gChangeOrgId
    frm1.hProcFg.value = "CG"
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet(ByVal spdsep)        
    Select case spdsep
		Case "A"
			Call initSpreadPosVariables("A")
			     
			With frm1
				.vspdData.MaxCols = C_DUE_DT
				.vspdData.Col = .vspdData.MaxCols				'��: ������Ʈ�� ��� Hidden Column
				.vspdData.ColHidden = True
				.vspdData.MaxRows = 0
				
				ggoSpread.Source = frm1.vspdData
  				
			    ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
			 
			    Call GetSpreadColumnPos("A")

				ggoSpread.SSSetCheck	C_PROC_CHK,		"����",       5, , "", True, -1
				ggoSpread.SSSetEdit		C_FR_DEPT_CD,	"���� �μ�",   8, , , 10
				ggoSpread.SSSetEdit		C_FR_DEPT_NM,	"���� �μ���", 10, , , 40				
				ggoSpread.SSSetEdit		C_NOTE_NO,		"������ȣ",   15, , , 30
				ggoSpread.SSSetFloat	C_NOTE_AMT,		"�����ݾ�",   12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetEdit		C_NOTE_STS,		"��������",   8, , , 30
				ggoSpread.SSSetEdit		C_TO_DEPT_CD,	"�̵��μ�",   8, , , 10
				ggoSpread.SSSetButton   C_TO_DEPT_POP
				ggoSpread.SSSetEdit		C_TO_DEPT_NM,	"�̵��μ���", 10, , , 40
				ggoSpread.SSSetEdit		C_MOVE_DESC,	"���"		, 15, , , 100		
				ggoSpread.SSSetEdit		C_BP_CD,		"�ŷ�ó",     10, , , 10
				ggoSpread.SSSetEdit		C_BP_NM,		"�ŷ�ó��",   15, , , 50
				ggoSpread.SSSetDate		C_ISSUED_DT,	"������",     10, 2, Parent.gDateFormat
				ggoSpread.SSSetDate		C_DUE_DT,		"������",     10, 2, Parent.gDateFormat

			    'Call ggoSpread.SSSetColHidden(C_GL_NO,C_GL_NO,True)
			    'Call ggoSpread.SSSetColHidden(C_TEMP_GL_NO,C_TEMP_GL_NO,True)
			End With
    
			Call SetSpreadLock("A")                                              '�ٲ�κ� 
		Case "B"	
			Call initSpreadPosVariables("B")

			With frm1
				.vspdData2.MaxCols = C_CNCL_DUE_DT
				.vspdData2.Col = .vspdData2.MaxCols				'��: ������Ʈ�� ��� Hidden Column
				.vspdData2.ColHidden = True
				.vspdData2.MaxRows = 0
				
				ggoSpread.Source = frm1.vspdData2
			    ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
			    
			    Call GetSpreadColumnPos("B")

				ggoSpread.SSSetCheck	C_CNCL_CHK,				"����"	  ,      5, , "", True, -1
				ggoSpread.SSSetEdit		C_CNCL_TO_DEPT_CD,	    "���� �μ�",      8, , , 10
				ggoSpread.SSSetEdit		C_CNCL_TO_DEPT_NM,	    "���� �μ���",   10, , , 40	
				ggoSpread.SSSetEdit		C_CNCL_NOTE_NO,			"������ȣ",     15, , , 30				
				ggoSpread.SSSetFloat	C_CNCL_NOTE_AMT,		"�����ݾ�",     12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec		
				ggoSpread.SSSetEdit		C_CNCL_FR_DEPT_CD,		"���μ�",    8, , , 10
				ggoSpread.SSSetEdit		C_CNCL_FR_DEPT_NM,		"���μ���", 10, , , 40		
				ggoSpread.SSSetEdit		C_CNCL_TEMP_GL_NO,		"������ǥ��ȣ", 12, , , 18		
				ggoSpread.SSSetDate		C_CNCL_TEMP_GL_DT,		"������ǥ��",	10, 2, Parent.gDateFormat		
				ggoSpread.SSSetEdit		C_CNCL_GL_NO,			"ȸ����ǥ��ȣ", 12, , , 18		
				ggoSpread.SSSetDate		C_CNCL_GL_DT,			"ȸ����ǥ��",   10, 2, Parent.gDateFormat
				ggoSpread.SSSetEdit		C_CNCL_BP_CD,			"�ŷ�ó",		10, , , 10
				ggoSpread.SSSetEdit		C_CNCL_BP_NM,			"�ŷ�ó��",		15, , , 50
				ggoSpread.SSSetDate		C_CNCL_ISSUED_DT,		"������",		10, 2, Parent.gDateFormat
				ggoSpread.SSSetDate		C_CNCL_DUE_DT,			"������",		10, 2, Parent.gDateFormat  
			End With

			Call SetSpreadLock("B")                                              '�ٲ�κ� 
	End Select 
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(ByVal spdsep1)
	Dim RowCnt
	Dim strTempGlNo
	Dim strGlNo
	
	Select case spdsep1
		Case "A"
			ggoSpread.Source = frm1.vspdData
			With frm1.vspdData
				.ReDraw = False
				ggoSpread.SpreadLock	C_FR_DEPT_CD,	-1, C_FR_DEPT_CD		' �������μ� 
				ggoSpread.SpreadLock	C_FR_DEPT_NM,	-1, C_FR_DEPT_NM		' �������μ��� 
				ggoSpread.SpreadLock	C_NOTE_NO,		-1, C_NOTE_NO			' ������ȣ 
				ggoSpread.SpreadLock	C_NOTE_AMT,		-1, C_NOTE_AMT			' �����ݾ� 
				ggoSpread.SpreadLock	C_NOTE_STS,		-1, C_NOTE_STS			' �����ݾ�				
				ggoSpread.SSSetRequired C_TO_DEPT_CD,	-1, C_TO_DEPT_CD		' �����ĺμ� 
				ggoSpread.SpreadUnLock	C_TO_DEPT_POP,	-1, C_TO_DEPT_POP		' �����ĺμ��˾� 
				ggoSpread.SpreadLock	C_TO_DEPT_NM,	-1, C_TO_DEPT_NM		' �����ĺμ��� 
				ggoSpread.SpreadUnLock	C_MOVE_DESC,	-1, C_TO_DEPT_NM		' �̵��ú�� 
				ggoSpread.SpreadLock	C_BP_CD,		-1, C_BP_CD				' �ŷ�ó�ڵ� 
				ggoSpread.SpreadLock	C_BP_NM,		-1, C_BP_NM				' �ŷ�ó�� 
				ggoSpread.SpreadLock	C_ISSUED_DT,	-1, C_ISSUED_DT			' ���������� 
				ggoSpread.SpreadLock	C_DUE_DT,		-1, C_DUE_DT			' ���������� 
				
				.ReDraw = True
			End With
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			With frm1.vspdData2
				.ReDraw = False			    		
				ggoSpread.SpreadLock C_CNCL_TO_DEPT_CD,		-1, C_CNCL_TO_DEPT_CD			' ������ȣ		
				ggoSpread.SpreadLock C_CNCL_TO_DEPT_NM,		-1, C_CNCL_TO_DEPT_NM			' ������ȣ		
				ggoSpread.SpreadLock C_CNCL_NOTE_NO,		-1, C_CNCL_NOTE_NO			' ������ȣ		
				ggoSpread.SpreadLock C_CNCL_NOTE_AMT,		-1, C_CNCL_NOTE_AMT			' ��ǥ�ݾ�		
				
				ggoSpread.SpreadLock C_CNCL_FR_DEPT_CD,		-1, C_CNCL_FR_DEPT_CD			' ������ȣ		
				ggoSpread.SpreadLock C_CNCL_FR_DEPT_NM,		-1, C_CNCL_FR_DEPT_NM			' ������ȣ	
						
				ggoSpread.SpreadLock C_CNCL_TEMP_GL_NO,		-1, C_CNCL_TEMP_GL_NO		' ������ǥ��ȣ		
				ggoSpread.SpreadLock C_CNCL_TEMP_GL_DT,		-1, C_CNCL_TEMP_GL_DT		' ������ǥ���� 
				ggoSpread.SpreadLock C_CNCL_GL_NO,			-1, C_CNCL_GL_NO			' ȸ����ǥ��ȣ 
				ggoSpread.SpreadLock C_CNCL_GL_DT,			-1, C_CNCL_GL_DT			' ��ǥ���� 

				ggoSpread.SpreadLock C_CNCL_BP_CD,			-1, C_CNCL_BP_CD			' �ŷ�ó�ڵ� 
				ggoSpread.SpreadLock C_CNCL_BP_NM,			-1, C_CNCL_BP_NM			' �ŷ�ó�� 
				ggoSpread.SpreadLock C_CNCL_ISSUED_DT,		-1, C_CNCL_ISSUED_DT			' �μ��ڵ� 
				ggoSpread.SpreadLock C_CNCL_DUE_DT,			-1, C_CNCL_DUE_DT			' �μ��� 
				
				.ReDraw = True
			End With
		Case "C"
			ggoSpread.Source = frm1.vspdData2
			With frm1.vspdData2
				.ReDraw = False			    
				For RowCnt = 1 To .MaxRows
					.Row = RowCnt
					.Col = C_CNCL_TEMP_GL_NO
					strTempGlNo = .Text
					.Col = C_CNCL_GL_NO
					strGlNo = .Text
					If strTempGlNo <> "" and strGlNo <> ""Then				
						ggoSpread.SpreadLock		C_CNCL_CHK	, RowCnt	, C_CNCL_CHK	, RowCnt				
					Else 				
						ggoSpread.SpreadUnLock	C_CNCL_CHK	, RowCnt	, C_CNCL_CHK	, RowCnt
					End If
				Next		
			End With
    End Select
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		.vspdData.ReDraw = True
    End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()	

End Sub

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8), arrField(6), arrHeader(6)
	Dim strBizAreaCd
	Dim iCalledAspName	

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	    Case 1,2
			arrParam(0) = "����� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA" 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			If iWhere = "1" Then
				' ���Ѱ��� �߰� 
				If lgAuthBizAreaCd <> "" Then
					arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
				Else
					arrParam(4) = ""
				End If
			Else
				strBizAreaCd = Trim(frm1.txtFrBizCd.value)
				
				If strBizAreaCd = "" Then
					strBizAreaCd = "%"
					arrParam(4) = ""						' Where Condition
				Else
					arrParam(4) = "BIZ_AREA_CD NOT LIKE  " & FilterVar(strBizAreaCd, "''", "S") & ""  	' Where Condition						
				End If	
			End If
			
			arrParam(5) = "������ڵ�"					' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"						' Field��(0)
			arrField(1) = "BIZ_AREA_NM"						' Field��(1)
    
			arrHeader(0) = "������ڵ�"			' Header��(0)
			arrHeader(1) = "������"
		Case 3		'���ѿ� ���� �μ��ڵ常 Popup
			iCalledAspName = AskPRAspName("DeptPopupDt")

			If Trim(iCalledAspName) = "" Then
				IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
				IsOpenPop = False
				Exit Function
			End If

			arrParam(0) = strCode						'�μ��ڵ� 
			arrParam(1) = frm1.txtGLDt.Text			'��¥(Default:������)
			arrParam(2) = "1"							'�μ�����(lgUsrIntCd)
			IsOpenPop = True

			' ���Ѱ��� �߰� 
			arrParam(5) = lgAuthBizAreaCd
			arrParam(6) = lgInternalCd
			arrParam(7) = lgSubInternalCd
			arrParam(8) = lgAuthUsrID
		Case 4,5			'�μ� 
			' ������ ����忡 ���� �μ��� PopUp
			If iWhere = "3" Then
				strBizAreaCd = Trim(frm1.txtFrBizCd.value)
			Else
				strBizAreaCd = Trim(frm1.txtToBizCd.value)
			End If

			If strBizAreaCd = "" Then
				strBizAreaCd = "%"
			End If

			arrParam(0) = "�μ��ڵ��˾�"			' �˾� ��Ī 
			arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C, A_ACCT D "    				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id" & _
			              " from b_acct_dept where org_change_dt = ( select max(org_change_dt)" & _
			              " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))" & _
			              " and C.biz_area_cd LIKE " & FilterVar(strBizAreaCd, "''", "S")  & _
			              " AND B.cost_cd = A.cost_cd " & _
			              " AND C.biz_area_cd = B.biz_area_cd AND D.REL_BIZ_AREA_CD = C.BIZ_AREA_CD"
			              
			arrParam(5) = "�μ��ڵ�"				' �����ʵ��� �� ��Ī 
			
			arrField(0) = "A.DEPT_CD"	     				' Field��(0)
			arrField(1) = "A.DEPT_NM"			    		' Field��(1)
			arrField(2) = "C.BIZ_AREA_CD"			    		' Field��(2)
			arrField(3) = "C.BIZ_AREA_NM"			    		' Field��(3)
			arrField(4) = "A.INTERNAL_CD"
    
			arrHeader(0) = "�μ��ڵ�"				' Header��(0)
			arrHeader(1) = "�μ���"				    ' Header��(1)						
			arrHeader(2) = "������ڵ�"				' Header��(2)		
			arrHeader(3) = "������"				' Header��(3)	
			arrHeader(4) = "���κμ��ڵ�"			
			
		Case 8			'������ȣ 
	'	 If frm1.txtBankCd1.className = Parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "������ȣ �˾�"						' �˾� ��Ī 
			arrParam(1) = "F_NOTE	A"		' TABLE ��Ī 
			arrParam(2) = strCode									' Code Condition
			arrParam(3) = ""										' Name Condition
			arrParam(4) = " A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.NOTE_STS = " & FilterVar("OC", "''", "S") & " "	  ' Where Condition
			arrParam(5) = "������ȣ"											' �����ʵ��� �� ��Ī 

			arrField(0) = "A.NOTE_NO"						' Field��(0)
			arrField(1) = "A.ISSUE_DT"						' Field��(1)			
			arrField(2) = "A.NOTE_AMT"						' Field��(0)			
			arrField(3) = "A.DEPT_CD"						' Field��(0)			
			
			arrHeader(0) = "������ȣ"					' Header��(0)
			arrHeader(1) = "������"					' Header��(0)
			arrHeader(2) = "�����ݾ�"						' Header��(1)	
 			arrHeader(3) = "�߻��μ�"		
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	Select Case iWhere				
		Case 1, 2
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")		
	    Case 3
			arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
					"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	    Case 4, 5
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
					 "dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	    					 	    
        Case Else 
		     arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			     	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")					 
	End Select			
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		With frm1
			Select Case iWhere				
				Case 1		'���ʻ���� 
					.txtFrBizCd.Focus
				Case 2		'�̵������ 
					.txtToBizCd.Focus
				Case 3		' �μ� 
					.txtFrDeptCd.focus
				Case 4		' ���¹�ȣ 
					.txtToDeptCd.focus
				Case 8  '������ȣ 
					.txtNoteNo.focus 			
			End Select
		End With
		Exit Function
	End If	

	With frm1
		Select Case iWhere
			Case 1		'���ʻ���� 
				.txtFrBizCd.value = arrRet(0)
				.txtFrBizNm.value = arrRet(1)
				.txtFrDeptCd.focus
			Case 2		'�̵������ 
				.txtToBizCd.value = arrRet(0)
				.txtToBizNm.value = arrRet(1)
				.txtToDeptCd.focus
			Case 3		' ���ʺμ� 
				.txtFrDeptCd.value	= Trim(arrRet(0))
				.txtFrDeptNm.value	= Trim(arrRet(1))
			Case 4		'�̵��μ� 
				.txtToDeptCd.value = Trim(arrRet(0))
				.txtToDeptNm.value = Trim(arrRet(1))

				Call fncToDeptIntoSheet(Trim(arrRet(0)),Trim(arrRet(1)))
			Case 5
       			ggoSpread.Source = frm1.vspdData

				frm1.vspdData.Row = frm1.vspdData.ActiveRow
				frm1.vspdData.Col = C_TO_DEPT_CD

				frm1.vspdData.Text = arrRet(0)

				frm1.vspdData.Col = C_TO_DEPT_NM
				frm1.vspdData.Text = arrRet(1)

				ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			Case 8
				.txtNoteNo.value  = arrRet(0)
		End Select

		lgBlnFlgChgValue = True
	End With
End Function

'============================================================
'ȸ����ǥ �˾� 
'============================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
		
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData2
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_CNCL_GL_NO
			arrParam(0) = Trim(.Text)	'ȸ����ǥ��ȣ 
			arrParam(1) = ""			'Reference��ȣ 
		End If
	End With

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'============================================================
'������ǥ �˾� 
'============================================================
Function OpenPopupTempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
		
	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData2
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_CNCL_TEMP_GL_NO
			arrParam(0) = Trim(.Text)	'ȸ����ǥ��ȣ 
			arrParam(1) = ""			'Reference��ȣ 
		End If
	End With

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function


Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDt")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode					'�μ��ڵ� 
	arrParam(1) = frm1.txtGLDt.Text			'��¥(Default:������)
	arrParam(2) = "1"						'�μ�����(lgUsrIntCd)
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	End If
	
	if iWhere = "1" then
		frm1.txtFrDeptCd.value = arrRet(0)
		frm1.txtFrDeptNm.value = arrRet(1)
		'Call txtFrDeptCD_OnChange()
		frm1.txtFrDeptCd.focus
	else	
		frm1.txtFrDeptCd.value = arrRet(0)
		frm1.txtFrDeptNm.value = arrRet(1)
		'Call txtFrDeptCD_OnChange()
		frm1.txtFrDeptCd.focus
	end if
			
	lgBlnFlgChgValue = True
End Function



'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
 '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call InitVariables							'��: Initializes local global variables
    Call LoadInfTB19029							'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.ClearField(Document, "1")      '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document, "N")		'��: Lock  Suitable  Field

    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggospread.ClearSpreadData

    Call InitSpreadSheet("A")                                                        'Setup the Spread sheet
    Call InitSpreadSheet("B")                                                        'Setup the Spread sheet
    Call SetDefaultVal
    Call ClickTab1

    gIsTab     = "Y" 
	gTabMaxCnt = 2  	

	' [Main Menu ToolBar]�� �� ��ư�� [Enable/Disable] ó���ϴ� �κ� 
	'1�޴�Ž����/2��ȸ/3�ű�/4����/5����/6���߰�/7�����/8���/9����/10����/11���ڵ庹��/12EXPORT/13�μ�/14ã��/15���� 

    Call SetToolbar("1100000000001111")										'��: ��ư ���� ���� 

	Set gActiveElement = document.activeElement
	
	' ���Ѱ��� �߰� 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' ����� 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ� 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' ���� 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing
		
End Sub


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
            
            C_PROC_CHK    = iCurColumnPos(1)
            C_FR_DEPT_CD  = iCurColumnPos(2)
            C_FR_DEPT_NM  = iCurColumnPos(3)             
            C_NOTE_NO	  = iCurColumnPos(4)
            C_NOTE_AMT    = iCurColumnPos(5)
            C_NOTE_STS    = iCurColumnPos(6)
            C_TO_DEPT_CD  = iCurColumnPos(7)
            C_TO_DEPT_POP = iCurColumnPos(8)
            C_TO_DEPT_NM  = iCurColumnPos(9)
            C_MOVE_DESC   = iCurColumnPos(10)                          
            C_BP_CD		  = iCurColumnPos(11)
            C_BP_NM	      = iCurColumnPos(12)             
            C_ISSUED_DT   = iCurColumnPos(13)             
            C_DUE_DT      = iCurColumnPos(14)
		Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
             
            C_CNCL_CHK			= iCurColumnPos(1)
            C_CNCL_TO_DEPT_CD	= iCurColumnPos(2)
            C_CNCL_TO_DEPT_NM	= iCurColumnPos(3)                                       
            C_CNCL_GL_NO		= iCurColumnPos(4)
            C_CNCL_GL_DT		= iCurColumnPos(5)
			C_CNCL_TEMP_GL_NO	= iCurColumnPos(6)
			C_CNCL_TEMP_GL_DT	= iCurColumnPos(7)
            C_CNCL_NOTE_NO		= iCurColumnPos(8)              
            C_CNCL_NOTE_AMT		= iCurColumnPos(9)              
            C_CNCL_FR_DEPT_CD	= iCurColumnPos(10)
            C_CNCL_FR_DEPT_NM	= iCurColumnPos(11)                                       
			C_CNCL_BP_CD		= iCurColumnPos(12)
            C_CNCL_BP_NM		= iCurColumnPos(13)
            C_CNCL_ISSUED_DT	= iCurColumnPos(14)             
            C_CNCL_DUE_DT		= iCurColumnPos(15)
	End Select    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
	End if
End Sub

Sub txtTodt_DblClick(Button)
	if Button = 1 then
		frm1.txtTodt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtTodt.Focus
	End if
End Sub

Sub txtGLDt_DblClick(Button)
	if Button = 1 then
		frm1.txtGLDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtGLDt.Focus
	End if
End Sub

Sub txtFrGlDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrGlDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrGlDt.Focus
	End if
End Sub

Sub txtToGlDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToGlDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToGlDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name :txtDueDt_keypress(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtFromDt.focus
		Call MainQuery
	End If   
End Sub

Sub txtToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then  
		frm1.txtToDt.focus
		Call MainQuery
	End If   
End Sub

Sub txtFrGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToGlDt.focus
	   Call MainQuery
	End If   
End Sub

Sub txtToGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtFrGlDt.focus
	   Call MainQuery
	End If   
End Sub

Sub txtDueDtEnd_Change()
End Sub

Sub txtGLDt_Change()
    Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2

	If Trim(frm1.txtToDeptCd.value) <> "" and Trim(frm1.txtGLDt.Text <> "") Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtToDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtToDeptCd.value = ""
			frm1.txtToDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtToDeptCd.value = ""
					frm1.txtToDeptNm.value = ""
				    frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
		End If
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFrDeptCd_onBlur()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFrBizCd_onBlur()
	If Trim(frm1.txtFrBizCd.value) = "" Then
		frm1.txtFrBizNm.value = ""
	End If
End Sub	

Sub txtFrDeptCd_onBlur()
	If Trim(frm1.txtFrDeptCd.value) = "" Then
		frm1.txtFrDeptNm.value = ""
	End If
End Sub	

Sub txtToBizCd_onBlur()
	If Trim(frm1.txtToBizCd.value) = "" Then
		frm1.txtToBizNm.value = ""
	End If
End Sub	

Sub txtToDeptCd_onBlur()	
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	
	If Trim(frm1.txtToDeptCd.value) = "" Then
		frm1.txtToDeptNm.value = ""		
	Else
		strSelect	= " dept_nm"    		
		strFrom		= " b_acct_dept(NOLOCK) "		
		strWhere	= " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtToDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtToDeptCd.value = ""
			frm1.txtToDeptNm.value = ""
			'frm1.hOrgChangeId.value = ""
		Else
			frm1.txtToDeptNm.value = mid(Trim(lgF2By2),2,len(lgF2By2) - 3)
			Call fncToDeptIntoSheet(UCase(Trim(frm1.txtToDeptCd.value)), Trim(frm1.txtToDeptNm.value))
		End If
	End If
End Sub	

Function fncToDeptIntoSheet(ByVal pDeptCd, pDeptNm)
	Dim IRow
	
	If frm1.vspdData.MaxRows < 1 Then
		Exit Function
	End If 

	ggoSpread.Source = frm1.vspdData
	
	For IRow = 1 To frm1.vspdData.MaxRows 
		frm1.vspdData.Row  = IRow
		frm1.vspdData.Col  = C_TO_DEPT_CD
		frm1.vspdData.Text = pDeptCd
		
		frm1.vspdData.Col  = C_TO_DEPT_NM
		frm1.vspdData.Text = pDeptNm
	Next

	lgBlnFlgChgValue = True
End Function

'======================================================================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=======================================================================================================
Function ClickTab1()	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
	    Call SetToolbar("1100000000001111")										'��: ��ư ���� ����	    
	Else                 
	    Call SetToolbar("1100000000001111")										'��: ��ư ���� ���� 
	End If

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)														'ù��° Tab 	
	
	gSelframeFlg = TAB1
	
	frm1.hProcFg.value = "CG"
End Function

Function ClickTab2()
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetToolBar("1100000000001111")
	Else                 
		Call SetToolBar("1100000000001111")
	End If	

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)														'�ι�° Tab 
	
	gSelframeFlg = TAB2
	frm1.hProcFg.value = "DG"
End Function


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")
   	gMouseClickStatus = "SPC"	'Split �����ڵ� 
	
  	Set gActiveSpdSheet = frm1.vspdData
  	
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")
   	gMouseClickStatus = "SPC"	'Split �����ڵ� 
	
  	Set gActiveSpdSheet = frm1.vspdData2

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData2
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
	End If
	
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
    End If     
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
       Exit Sub
    End If     
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
         
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If lgStrPrevKeyNoteNo <> "" Then								
			If DbQuery = False Then
				Exit Sub
			End if
    	End If
    End If    
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_PROC_CHK Or NewCol <= C_PROC_CHK Then
        Cancel = True
        Exit Sub
    End If
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_CNCL_CHK Or NewCol <= C_CNCL_CHK Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If 
    
   	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	    
		If (lgStrPrevKeyNoteNo <> "" or  lgStrPrevKeyGlNo <> "" or lgStrPrevKeyTempGlNo <> "" ) Then								
			If DbQuery = False Then
				Exit Sub
			End if
    	End If
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	'2004.05.27
	If Col = C_PROC_CHK Then
		With frm1.vspdData
			.Row = Row
			.Col = C_PROC_CHK
			
			ggoSpread.Source = frm1.vspdData
			
			If .Text = "Y" Then	
				If ButtonDown = 0 Then
					ggoSpread.UpdateRow Row
				Else
					ggoSpread.SSDeleteFlag Row,Row
				End If
			Else
				If ButtonDown = 1 Then		
					ggoSpread.UpdateRow Row  ''2004.03.19 comment ó��				
				Else
					ggoSpread.SSDeleteFlag Row,Row
					ggoSpread.SSDeleteFlag Row,Row
				End If			
			End If
		End With
	Elseif	Col = C_TO_DEPT_POP then
		With frm1.vspdData
			.Row = Row
			.Col = C_TO_DEPT_CD
			
			Call OpenPopUp(frm1.vspdData.text, 5)			
		End With
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData2
		.Row = Row
		.Col = C_PROC_CHK
		
		ggoSpread.Source = frm1.vspdData2
		
		If .Text = "Y" Then
			If ButtonDown = 0 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
				ggoSpread.SSDeleteFlag Row,Row
			End If			
		End If
	End With
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    'FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
   '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.ClearField(Document, "3")			'��: Clear Contents  Field
        
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData			'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData2
    ggospread.ClearSpreadData			'��: Clear Contents  Field
	
    If gSelframeFlg = TAB1 Then 
		Call InitSpreadSheet("A")
    Else
		Call InitSpreadSheet("B")
    End If 
       
    Call InitVariables() 
    
    frm1.vspdData.MaxRows = 0
	
    '-----------------------
    'Check condition area
    '----------------------- 
    If gSelframeFlg = TAB1 Then 
		If Not chkField(Document, "1") Then									'��: This function check indispensable field     						
			Exit Function	
		End If 
	Else
		If (frm1.txtFrGlDt.Text = "") or (frm1.txtToGlDt.Text = "") Then
			Call DisplayMsgBox("17A002", parent.VB_INFORMATION, "X", "X")
			Exit Function
		End if
		''KO      17A002 A        2        %1�� �Է��ϼ���.
	End if
	    
    If gSelframeFlg = "1" Then
		If (frm1.txtFromDt.Text <> "") And (frm1.txtTodt.Text <> "") Then
			If CompareDateByFormat(frm1.txtFromDt.Text, frm1.txtTodt.Text, frm1.txtFromDt.Alt, frm1.txtToDt.Alt, _
						"970025", frm1.txtFromDt.UserDefinedFormat, Parent.gComDateType, true) = False Then

				frm1.txtFromDt.focus											
				Exit Function
			End if	
		End If
	End If
	
    If gSelframeFlg = "2" Then
		If (frm1.txtFrGlDt.Text <> "") And (frm1.txtToGlDt.Text <> "") Then
			If CompareDateByFormat(frm1.txtFrGlDt.Text, frm1.txtToGlDt.Text, frm1.txtFrGlDt.Alt, frm1.txtToGlDt.Alt, _
						"970025", frm1.txtFrGlDt.UserDefinedFormat, Parent.gComDateType, true) = False Then
				frm1.txtFrGlDt.focus											
				Exit Function
			End if	
		End If
	End If
	
    If frm1.txtToBizCd.value  = "" Then
		frm1.txtToBizNm.value = ""
	End If
    
    If frm1.txtToDeptCd.value  = "" Then
		frm1.txtToDeptNm.value = ""
	End If
		
    Call ggoOper.LockField(Document, "N")			'��: This function lock the suitable field

    '-------------------------
    'Query function call area
    '-------------------------	   
  
    IF  DbQuery	= False Then						'��: Query db data
		Exit Function	
    End If
    
    FncQuery = True		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	dbsave()
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
     On Error Resume Next                                                   '��: Protect system from crashing
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
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)												'��: ȭ�� ���� 
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                     '��:ȭ�� ����, Tab ���� 
	Set gActiveElement = document.activeElement    
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
    
    If gSelframeFlg = TAB1 Then
		Call InitSpreadSheet("A") 
    Else
		Call InitSpreadSheet("B")      
    End If
    
    Call ggoSpread.ReOrderingSpreadData()
End Sub

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
	Dim strVal

    DbQuery = False
    Err.Clear                '��: Protect system from crashing
    
	Call LayerShowHide(1)	

	With frm1
		If gSelframeFlg = "1" Then 														'��: �ϰ�ó��(tab1) ��ȸ 
		    If lgIntFlgMode = Parent.OPMD_UMODE Then		    
				strVal = BIZ_PGM_ID2 & "?txtMode = " & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
					
				strVal = strVal & "&txtFromDt="	   & Trim(frm1.txtFromDt.Text)
				strVal = strVal & "&txtToDt="      & Trim(frm1.txtToDt.Text)
				strVal = strVal & "&txtFrBizCd="   & Trim(frm1.txtFrBizCd.value)
				
				strVal = strVal & "&txtBpCd="      & Trim(frm1.txtBpCd.value)		
				strVal = strVal & "&txtFrDeptCd=" & Trim(frm1.txtFrDeptCd.value)
				strVal = strVal & "&gChangeOrgId=" & Trim(.hOrgChangeId.Value)
				strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNo.value)
												
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo="   & lgStrPrevKeyGlNo
				strVal = strVal & "&lgPageNo="           & lgPageNo
				strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
			Else			
				strVal = BIZ_PGM_ID2 & "?txtMode = " & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
				
				strVal = strVal & "&txtFromDt="	   & Trim(frm1.txtFromDt.Text)
				strVal = strVal & "&txtToDt="      & Trim(frm1.txtToDt.Text)
				strVal = strVal & "&txtFrBizCd="   & Trim(frm1.txtFrBizCd.value)
				
				strVal = strVal & "&txtBpCd="      & Trim(frm1.txtBpCd.value)		
				strVal = strVal & "&txtFrDeptCd=" & Trim(frm1.txtFrDeptCd.value)
				strVal = strVal & "&gChangeOrgId=" & Trim(.hOrgChangeId.Value)
				strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNo.value)
				
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo="   & lgStrPrevKeyGlNo				
				strVal = strVal & "&lgPageNo="           & lgPageNo
				strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
			End If   						
		Else 																			'��: �ϰ����(tab2) ��ȸ																				
		    If lgIntFlgMode = Parent.OPMD_UMODE Then
				strVal = BIZ_PGM_ID3 & "?txtMode=" & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 				
				
				strVal = strVal & "&txtFrGlDt="	  & Trim(frm1.txtFrGlDt.Text)
				strVal = strVal & "&txtToGlDt="   & Trim(frm1.txtToGlDt.Text)				
				strVal = strVal & "&lgStrPrevKeyNoteNo="	& lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo="		& lgStrPrevKeyGlNo
				strVal = strVal & "&lgStrPrevKeyTempGlNo="	& lgStrPrevKeyTempGlNo				
				strVal = strVal & "&lgPageNo="				& lgPageNo
				strVal = strVal & "&txtMaxRows="			& .vspdData2.MaxRows
			Else
				strVal = BIZ_PGM_ID3 & "?txtMode=" & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 				

				strVal = strVal & "&txtFrGlDt=" & Trim(frm1.txtFrGlDt.Text)
				strVal = strVal & "&txtToGlDt=" & Trim(frm1.txtToGlDt.Text)
				strVal = strVal & "&lgStrPrevKeyNoteNo="	& lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo="		& lgStrPrevKeyGlNo
				strVal = strVal & "&lgStrPrevKeyTempGlNo="	& lgStrPrevKeyTempGlNo				
				strVal = strVal & "&lgPageNo="				& lgPageNo
				strVal = strVal & "&txtMaxRows="			& .vspdData2.MaxRows
			End If	
		End If		
	End With 

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ����				

	Call RunMyBizASP(MyBizASP, strVal)		'��: �����Ͻ� ASP �� ���� 
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()							'��: ��ȸ ������ ������� 
	
	If gSelframeFlg = "2" Then 					
		Call SetSpreadLock("C")
	End If 	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE	'��: Indicates that current mode is Update mode    
    
	lgBlnFlgChgValue = False
	
	' ���� Page�� From Element���� ����ڰ� �Է��� ���� ���ϰ� �ϰų� �ʼ��Է»����� ǥ���Ѵ�.
	' LockField(pDoc, pACode)
	
'   Call ggoOper.LockField(Document, "Q")		'��: This function lock the suitable field
    frm1.txtGLDt.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)	
    Call SetToolBar("1100100000001111")
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 
	Dim lRow
	Dim lGrpCnt
	Dim strVal
	Dim NoteAmtSum
	Dim ChkCnt
	Dim strGLNo
	Dim ChkFlag
	Dim BatchChk
	Dim intRetCD

	DbSave = False				'��: Processing is NG

	'2001.03.01 added
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		IntRetCD = DisplayMsgBox("900002","x","x","x")  '��ȸ�� ���� �Ͻʽÿ�.
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"x","x")	'�۾��� �����Ͻðڽ��ϱ�?

	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	'If frm1.hProcFg.value = "CG" Then
	If gSelframeFlg = TAB1 then  ''''�̵�ó�� Tab�̸�  
		If Not chkField(Document, "2") Then                                   '��: Check contents area
			Exit Function
		End If
	End If
    
	IF Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If
	
	If gSelframeFlg = TAB1 then  ''''�̵�ó�� Tab�̸� 
		If UCase(Trim(frm1.txtFrBizCd.value)) = UCase(Trim(frm1.txtToBizCd.value)) then
			IntRetCD = DisplayMsgBox("141445","x","x","x")
			Exit Function
		End If	
	End If
	
	With frm1
		.txtMode.value = Parent.UID_M0002			'��: �����Ͻ� ó�� ASP �� ���� 
		.txtInsrtUserId.value = Parent.gUsrID

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		    
		'-----------------------
		'Data manipulate area
		'-----------------------
		'If .hProcFg.value = "CG" Then										'��:�ϰ�ó�� 
		If gSelframeFlg = TAB1 then  ''''�̵�ó�� Tab�̸� 
			For lRow = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow
				.vspdData.Col = C_PROC_CHK
				
				If .vspdData.Text = "1" Then				
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData.Col = C_NOTE_NO
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ������ȣ 
					.vspdData.Col = C_TO_DEPT_CD
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' �̵��μ��ڵ�					
					.vspdData.Col = C_MOVE_DESC
					strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep		' ��� 
					lGrpCnt = lGrpCnt + 1
				End If
			Next

			.hProcFg.value = "CG"	
		ElseIf gSelframeFlg = TAB2 then  ''''�̵����  Tab�̸�  Then									 '��:�ϰ���� 
			For lRow = 1 To .vspdData2.MaxRows
				.vspdData2.Row = lRow
				.vspdData2.Col = C_CNCL_CHK
				
				If .vspdData2.Text = "1" Then
					strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData2.Col = C_CNCL_NOTE_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep		' ������ȣ				
					.vspdData2.Col = C_CNCL_TEMP_GL_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep		' ������ǥ��ȣ 
					.vspdData2.Col = C_CNCL_GL_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gRowSep		' ȸ����ǥ��ȣ				
					
					lGrpCnt = lGrpCnt + 1
				End If
			Next	
			.hProcFg.value = "DG"
		End If
			
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal

		If .txtMaxRows.value <= 0 Then
			Call DisplayMsgBox("900025","x","x","x")	'���õ� �׸��� �����ϴ�.
			Exit Function
		End If

		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end
	
		Call LayerShowHide(1)
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)				'��: �����Ͻ� ASP �� ���� 
	End With

    DbSave = True										'��: Processing is NG
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function DbSaveOk()										'��: ���� ������ ���� ���� 
	Call InitVariables
	frm1.vspdData.MaxRows = 0	
	Call MainQuery
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
End Sub

'=======================================================================================================
'   Event Name : Rowcancel() / Rowselect()
'   Event Desc :
'=======================================================================================================    
Function Rowcancel()
	Dim lRow

	If gSelframeFlg = "1" Then 						'��: �ϰ�ó��(tab1) ��ȸ 
		With Frm1.vspdData
			For lRow = 1 To .MaxRows
				.Row = lRow
				.COL = 0
				IF Trim(.TEXT) = ggoSpread.UPDATEFlag OR Trim(.TEXT) = ggoSpread.INSERTFlag THEN
					.Col = C_PROC_CHK
					.Text = "0"
					IF Trim(.TEXT) = ggoSpread.UPDATEFlag THEN
						ggoSpread.SSDeleteFlag lRow,lRow
					END IF
				END IF
			Next
		End With
	Else
		With Frm1.vspdData2
			For lRow = 1 To .MaxRows
				.Row = lRow
				.COL = 0
				IF Trim(.TEXT) = ggoSpread.UPDATEFlag OR Trim(.TEXT) = ggoSpread.INSERTFlag THEN
					.Col = C_PROC_CHK
					.Text = "0"
					IF Trim(.TEXT) = ggoSpread.UPDATEFlag THEN
						ggoSpread.SSDeleteFlag lRow,lRow
					END IF
				END IF
			Next
		End With
	End If 
End Function

Function Rowselect()
	Dim lRow
	
	If gSelframeFlg = "1" Then 						'��: �ϰ�ó��(tab1) ��ȸ 
		With Frm1.vspdData
			For lRow = 1 To .MaxRows
				.Row = lRow
				.COL = 0
				IF Trim(.TEXT) <> ggoSpread.DELETEFlag THEN
					.Col = C_PROC_CHK
					If .Lock = False Then
						.Col = C_PROC_CHK
						.Text = "1"
						ggoSpread.UpdateRow lRow
					End If
				END IF
			Next
		End With
	Else
		With Frm1.vspdData2
			For lRow = 1 To .MaxRows
				.Row = lRow
				.COL = 0
				IF Trim(.TEXT) <> ggoSpread.DELETEFlag THEN
					.Col = C_PROC_CHK
					If .Lock = False Then
						.Col = C_PROC_CHK
						.Text = "1"
						ggoSpread.UpdateRow lRow
					End If
				END IF
			Next
		End With
	End If 
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = strCode						'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""							'FrDt
	arrParam(3) = ""							'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		'Call SetPopUp(arrRet, iWhere)
		frm1.txtBpCd.value  = arrRet(0)
		frm1.txtBpNm.value  = arrRet(1)
	End If	
End Function

Sub txtFrDeptCd_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If frm1.txtFrDeptCd.value = "" Then
		frm1.txtFrDeptNm.value = ""
	End If
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtFrDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect = "dept_cd, ORG_CHANGE_ID"
		strFrom =  " B_ACCT_DEPT "
		strWhere = " ORG_CHANGE_DT >= "
		strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
		strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
		strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtTodt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
		strWhere = strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtFrDeptCd.value)), "''", "S")
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtFrDeptCd.value = ""
			frm1.txtFrDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtFrDeptCd.focus
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
		End If
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">		
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
							</TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���������̵����</font></td><td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A> &nbsp;|<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
					<TD WIDTH=10>&nbsp;</TD>					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
	
			<DIV ID="TabDiv" SCROLL="no">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=OBJECT1 name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="���۹�����"></OBJECT>');</SCRIPT>&nbsp; ~ &nbsp;
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=OBJECT1 name=txtTodt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="���������"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=10 MAXLENGTH=10  tag="1XX" ALT="�ŷ�ó�ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value)">
									                     <INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNm" NAME="txtBpNm" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="14X" ALT="�ŷ�ó��"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���� �����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFrBizCd" ALT="���� ������ڵ�" Size= "12" MAXLENGTH="10" tag="12XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtFrBizCd.value, 1)">
														 <INPUT NAME="txtFrBizNm" ALT="���� ������" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
									<TD CLASS=TD5 NOWRAP>���� �μ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFrDeptCd" ALT="���� �μ��ڵ�" Size= "10" MAXLENGTH="10" tag="11XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtFrDeptCd.value,3)">
														 <INPUT NAME="txtFrDeptNm" ALT="���� �μ���" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
								</TR>							
								<TR>
									<TD CLASS=TD5 NOWRAP>������ȣ</TD>								
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtNoteNo" NAME="txtNoteNo" SIZE=30 MAXLENGTH=30  tag="1XX" ALT="���ؾ�����ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteNo.Value, 8)"></TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>									
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�̵���</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�̵���" tag="22X1" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>					
							<TR>
								<TD CLASS=TD5 NOWRAP>�̵������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtToBizCd" ALT="�̵�������ڵ�" Size= "12" MAXLENGTH="10" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtToBizCd.value, 2)">
													 <INPUT NAME="txtToBizNm" ALT="�̵�������" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
								<TD CLASS=TD5 NOWRAP>�̵��μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtToDeptCd" ALT="�̵��μ��ڵ�" Size= "12" MAXLENGTH="10" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtToDeptCd.value, 4)">
													 <INPUT NAME="txtToDeptNm" ALT="�̵�������" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
							</TR>									
							<TR>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtNoteDesc" ALT="���" SIZE = "90" STYLE="TEXT-ALIGN: left" tag="21X"></TD></TD>
							</TR>	
							<TR>
								<TD WIDTH=100% HEIGHT="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="33" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			</DIV>

			<DIV ID="TabDiv"  SCROLL=no>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>ȸ����</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtFrGlDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="����ȸ����" tag="12X1"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 name=txtToGlDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="����ȸ����" tag="12X1" ></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>														 
								</TR>
							</TABLE>
							    <TR>
									<TD WIDTH=100% HEIGHT="100%" COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="23" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD WIDTH=100% HEIGHT="50%" colspan=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="33" ID=vspdData3> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
								</TR>						
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
  		<TD WIDTH="100%">
  			<TABLE <%=LR_SPACE_TYPE_30%>>
   				<TR>
   					<TD WIDTH=10>&nbsp;</TD>
   					<TD><BUTTON NAME="btncancel" CLASS="CLSSBTN" ONCLICK="vbscript:Rowselect()">��ü����</BUTTON>&nbsp;
						<BUTTON NAME="btnselect" CLASS="CLSSBTN" ONCLICK="vbscript:Rowcancel()">��ü���</BUTTON>
					</TD>
   				</TR>
   			</TABLE> 
  		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="2" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="2" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"		tag="1" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hToBizAreaCd"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hProcFg"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteFg1"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteFg2"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteSts"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDueDtStart"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDueDtEnd"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStsDtStart"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStsDtEnd"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtGlNo"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtGLDt"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtCRAmt"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtCRLocAmt"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDRAmt"			tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDRLocAmt"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDocCur"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtXchRate"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtOrgChangeId"	tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDeptCd"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtAcctCd"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="GtxtBankCd"		tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="CtxtBankAcctNo"	tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="DtxtNoteNo"		tag="2" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 

src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
