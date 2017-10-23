<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : RECEIPT
'*  3. Program ID		    : f5116ma1
'*  4. Program Name         : ���ұ���ī��ó�� 
'*  5. Program Desc         : ���ұ���ī��ó�� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2002/10/14
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Soo Min, Oh
'* 10. Modifier (Last)      : 
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->	
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
<%
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = GetSvrDate
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------

%>
 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Const BIZ_PGM_ID = "f5116mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 = "f5116mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID3 = "f5116mb3.asp"											 '��: �����Ͻ� ���� ASP��    

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

'TAB1, vspddata
Dim C_PROC_CHK	
Dim C_NOTE_NO	
Dim C_NOTE_AMT	
Dim C_DUE_DT	
Dim C_NOTE_STS	
Dim C_BP_CD		
Dim C_BP_NM		
Dim C_CARD_CO_CD
Dim C_CARD_CO_NM
Dim C_DEPT_CD	
Dim C_DEPT_NM	
Dim C_NOTE_ITEM_DESC
Dim C_GL_NO		
Dim C_TEMP_GL_NO		
Dim C_COL_END	


'TAB2, vspddata2
Dim C_CNCL_CHK		
Dim C_CNCL_NOTE_NO	
Dim C_CNCL_TEMP_GL_NO	
Dim C_CNCL_TEMP_GL_DT	
Dim C_CNCL_GL_NO	
Dim C_CNCL_GL_DT	
Dim C_CNCL_NOTE_AMT	
Dim C_CNCL_BP_CD	
Dim C_CNCL_BP_NM	
Dim C_CNCL_BANK_CD	
Dim C_CNCL_BANK_NM	
Dim C_CNCL_DEPT_CD	
Dim C_CNCL_DEPT_NM	
Dim C_CNCL_NOTE_ITEM_DESC
Dim C_CNCL_COL_END	


Dim  gSelframeFlg

Dim lgStrPrevKeyNoteNo	' ���� �� 
Dim lgStrPrevKeyGlNo    ' ���� GL �� (DG)
Dim lgStrPrevKeyTempGlNo		' ���� Temp GL ��(DG)

Dim IsOpenPop          
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

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
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

	lgStrPrevKeyNoteNo = ""
	
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'��: ����� ���� �ʱ�ȭ 
    
    lgSortKey = 1
    
End Sub


Sub initSpreadPosVariables(ByVal spdsep2)
      
    select case spdsep2
    case "A"
		C_PROC_CHK			= 1
		C_NOTE_NO			= 2
		C_NOTE_AMT			= 3
		C_DUE_DT			= 4
		C_NOTE_STS			= 5 
		C_BP_CD				= 6
		C_BP_NM				= 7
		C_CARD_CO_CD		= 8
		C_CARD_CO_NM		= 9
		C_DEPT_CD			= 10
		C_DEPT_NM			= 11
		C_NOTE_ITEM_DESC	= 12
		C_GL_NO				= 13
		C_TEMP_GL_NO		= 14
		C_COL_END			= 15
		     
	Case "B" 
		C_CNCL_CHK			= 1
		C_CNCL_NOTE_NO		= 2
		C_CNCL_TEMP_GL_NO	= 3
		C_CNCL_TEMP_GL_DT	= 4
		C_CNCL_GL_NO		= 5
		C_CNCL_GL_DT		= 6	
		C_CNCL_NOTE_AMT		= 7
		C_CNCL_BP_CD		= 8
		C_CNCL_BP_NM		= 9
		C_CNCL_BANK_CD		= 10
		C_CNCL_BANK_NM		= 11
		C_CNCL_DEPT_CD		= 12
		C_CNCL_DEPT_NM		= 13
		C_CNCL_NOTE_ITEM_DESC	= 14
		C_CNCL_COL_END		= 15
      End Select 
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call loadInfTB19029A("I", "*","NOCOOKIE","MA")%>
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
		
	frm1.txtDueDtEnd.Text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtStsDtStart.Text = UniConvDateAToB(frDt,Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtStsDtEnd.Text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)
	
	frm1.hOrgChangeId.value = Parent.gChangeOrgId	
	frm1.txtGLDt.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)	
	
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
    
			.vspdData.MaxCols = C_COL_END
			
			.vspdData.Col = .vspdData.MaxCols				'��: ������Ʈ�� ��� Hidden Column
			.vspdData.ColHidden = True
			.vspdData.MaxRows = 0
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
		    Call GetSpreadColumnPos("A")

			ggoSpread.SSSetCheck	C_PROC_CHK,		"����"	  , 10, , "", True, -1
			ggoSpread.SSSetEdit		C_NOTE_NO,		"���ұ���ī���ȣ", 20, , , 30
			ggoSpread.SSSetFloat		C_NOTE_AMT,	"����ī��ݾ�", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
			ggoSpread.SSSetDate		C_DUE_DT,		"��������", 10, 2, Parent.gDateFormat
			ggoSpread.SSSetEdit		C_NOTE_STS,		"ī�����", 10, , , 5
			ggoSpread.SSSetEdit		C_BP_CD,		"�ŷ�ó", 10, , , 10
			ggoSpread.SSSetEdit		C_BP_NM,		"�ŷ�ó��", 20, , , 50
			ggoSpread.SSSetEdit		C_CARD_CO_CD,	"ī���", 10, , , 10
			ggoSpread.SSSetEdit		C_CARD_CO_NM,	"ī����", 20, , , 30
			ggoSpread.SSSetEdit		C_DEPT_CD,		"�μ�", 10, , , 10
			ggoSpread.SSSetEdit		C_DEPT_NM,		"�μ���", 20, , , 40
			ggoSpread.SSSetEdit		C_NOTE_ITEM_DESC,	"���", 30, , , 128
			ggoSpread.SSSetEdit		C_GL_NO,		"��ǥ��ȣ", 15, , , 18
			ggoSpread.SSSetEdit		C_TEMP_GL_NO,	"������ǥ��ȣ", 15, , , 18

		    Call ggoSpread.SSSetColHidden(C_NOTE_STS,C_NOTE_STS,True)
		    Call ggoSpread.SSSetColHidden(C_GL_NO,C_GL_NO,True)
		    Call ggoSpread.SSSetColHidden(C_TEMP_GL_NO,C_TEMP_GL_NO,True)
		End With
		Call SetSpreadLock("A")                                              '�ٲ�κ� 
    
    Case "B"
        Call initSpreadPosVariables("B")
   
        With frm1
    
		.vspdData2.MaxCols = C_CNCL_COL_END
		
		.vspdData2.Col = .vspdData2.MaxCols				'��: ������Ʈ�� ��� Hidden Column
		.vspdData2.ColHidden = True
		
		.vspdData2.MaxRows = 0
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
		
        Call GetSpreadColumnPos("B")
				
		ggoSpread.SSSetCheck	C_CNCL_CHK,				"����"	  , 10, , "", True, -1
		ggoSpread.SSSetEdit		C_CNCL_NOTE_NO,			"���ұ���ī���ȣ", 20, , , 30
		ggoSpread.SSSetEdit		C_CNCL_TEMP_GL_NO,			"������ǥ��ȣ", 15, , , 18		
		ggoSpread.SSSetDate		C_CNCL_TEMP_GL_DT,			"������ǥ����", 10, 2, Parent.gDateFormat				
		ggoSpread.SSSetEdit		C_CNCL_GL_NO,			"��ǥ��ȣ", 15, , , 18		
		ggoSpread.SSSetDate		C_CNCL_GL_DT,			"��ǥ����", 10, 2, Parent.gDateFormat
		ggoSpread.SSSetFloat	C_CNCL_NOTE_AMT,		"����ī��ݾ�", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit		C_CNCL_BP_CD,			"�ŷ�ó", 10, , , 10
		ggoSpread.SSSetEdit		C_CNCL_BP_NM,			"�ŷ�ó��", 20, , , 50
		ggoSpread.SSSetEdit		C_CNCL_BANK_CD,			"ī���", 10, , , 10
		ggoSpread.SSSetEdit		C_CNCL_BANK_NM,			"ī����", 20, , , 30
		ggoSpread.SSSetEdit		C_CNCL_DEPT_CD,			"�μ�", 10, , , 10
		ggoSpread.SSSetEdit		C_CNCL_DEPT_NM,			"�μ���", 20, , , 40
		ggoSpread.SSSetEdit		C_CNCL_NOTE_ITEM_DESC,	"���", 30, , , 128		

     End With
     Call SetSpreadLock("B")  
    
    End select 
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(ByVal spdsep1)
	Dim RowCnt
	Dim strTempGlNo
	Dim strGlNo

	select case spdsep1
	
	Case "A"
	
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		.ReDraw = False

		ggoSpread.SpreadLock C_NOTE_NO,	-1, C_NOTE_NO			' ����ī���ȣ 
		ggoSpread.SpreadLock C_NOTE_AMT,-1, C_NOTE_AMT			' ����ī��ݾ� 
		ggoSpread.SpreadLock C_DUE_DT,	-1, C_DUE_DT			' ������ 
		ggoSpread.SpreadLock C_BP_CD,	-1, C_BP_CD				' �ŷ�ó�ڵ� 
		ggoSpread.SpreadLock C_BP_NM,	-1, C_BP_NM				' �ŷ�ó�� 
		ggoSpread.SpreadLock C_CARD_CO_CD,	-1, C_CARD_CO_CD			' ����ī��� 
		ggoSpread.SpreadLock C_CARD_CO_NM,	-1, C_CARD_CO_NM			' ����ī���� 
		ggoSpread.SpreadLock C_DEPT_CD,	-1, C_DEPT_CD			' �μ��ڵ� 
		ggoSpread.SpreadLock C_DEPT_NM,	-1, C_DEPT_NM			' �μ��� 
		ggoSpread.SpreadUnLock C_NOTE_ITEM_DESC, -1, C_NOTE_ITEM_DESC ' ��� 

		.ReDraw = True

    End With
    Case "B"
    ggoSpread.Source = frm1.vspdData2

    With frm1.vspdData2
		.ReDraw = False

		ggoSpread.SpreadLock C_CNCL_NOTE_NO,	-1, C_CNCL_NOTE_NO			' ����ī���ȣ 
		ggoSpread.SpreadLock C_CNCL_TEMP_GL_NO,	-1, C_CNCL_TEMP_GL_NO			' ��ǥ��ȣ 
		ggoSpread.SpreadLock C_CNCL_TEMP_GL_DT,	-1, C_CNCL_TEMP_GL_DT			' ��ǥ���� 		
		ggoSpread.SpreadLock C_CNCL_GL_NO,		-1, C_CNCL_GL_NO			' ��ǥ��ȣ 
		ggoSpread.SpreadLock C_CNCL_GL_DT,		-1, C_CNCL_GL_DT			' ��ǥ���� 
		ggoSpread.SpreadLock C_CNCL_NOTE_AMT,	-1, C_CNCL_NOTE_AMT			' ��ǥ�ݾ� 
		ggoSpread.SpreadLock C_CNCL_BP_CD,		-1, C_CNCL_BP_CD			' �ŷ�ó�ڵ� 
		ggoSpread.SpreadLock C_CNCL_BP_NM,		-1, C_CNCL_BP_NM			' �ŷ�ó�� 
		ggoSpread.SpreadLock C_CNCL_BANK_CD,	-1, C_CNCL_BANK_CD			' ����ī��� 
		ggoSpread.SpreadLock C_CNCL_BANK_NM,	-1, C_CNCL_BANK_NM			' ����ī���� 
		ggoSpread.SpreadLock C_CNCL_DEPT_CD,	-1, C_CNCL_DEPT_CD			' �μ��ڵ� 
		ggoSpread.SpreadLock C_CNCL_DEPT_NM,	-1, C_CNCL_DEPT_NM			' �μ��� 
		ggoSpread.SpreadLock C_CNCL_NOTE_ITEM_DESC,	-1,	C_CNCL_NOTE_ITEM_DESC	' ��� 
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
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iwhere
			Case 6		' �ŷ�ó(tab1)
				frm1.txtBpCd1.focus
			Case 7		' �ŷ�ó(tab2)
				frm1.txtBpCd2.focus
		End Select
		Exit Function
	Else
		Select Case iwhere
			Case 6		' �ŷ�ó(tab1)
				frm1.txtBpCd1.value		= arrRet(0)
				frm1.txtBpNM1.value		= arrRet(1)	
				frm1.txtBpCd1.focus
			Case 7		' �ŷ�ó(tab2)
				frm1.txtBpCd2.value		= arrRet(0)
				frm1.txtBpNM2.value		= arrRet(1)	
				frm1.txtBpCd2.focus
		End Select
	End If	
End Function

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0			'�Ա�/������� 
			arrParam(0) = "�Ա�/������� �˾�"
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD AND B.SEQ_NO = 1 AND B.REFERENCE = " & FilterVar("RP", "''", "S") & "  "
			arrParam(5) = "�Ա�/�������"
	
			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
			    
			arrHeader(0) = frm1.txtRcptType.Alt
			arrHeader(1) = frm1.txtRcptTypeNm.Alt

		Case 1,2,5
			arrParam(0) = "���� �˾�"	' �˾� ��Ī 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"						' TABLE ��Ī 
			arrParam(2) = strCode																	' Code Condition
			arrParam(3) = ""																			' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "											' Where Condition			
			arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD " 
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO " 
			arrParam(4) = arrParam(4) & "AND (C.DPST_FG = " & FilterVar("SV", "''", "S") & "  OR C.DPST_FG = " & FilterVar("ET", "''", "S") & " ) " 
			arrParam(5) = "�����ڵ�"															' �����ʵ��� �� ��Ī 

			arrField(0) = "A.BANK_CD"							' Field��(0)
			arrField(1) = "A.BANK_NM"							' Field��(1)	
			arrField(2) = "B.BANK_ACCT_NO" 				' Field��(2) 		
    
			arrHeader(0) = "�����ڵ�"					' Header��(0)
			arrHeader(1) = "�����"						' Header��(1)			
			arrHeader(2) = "���¹�ȣ" 					' Header��(2)
		
'		Case 3			'�μ� 
'			arrParam(0) = "�μ� �˾�"	' �˾� ��Ī 
'			arrParam(1) = "B_ACCT_DEPT"		 			' TABLE ��Ī 
'			arrParam(2) = strCode						' Code Condition
'			arrParam(3) = ""							' Name Cindition
'			arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(Parent.gChangeOrgId, "''", "S") & ""	' Where Condition
'			arrParam(5) = "�μ�"					' �����ʵ��� �� ��Ī 
'
'			arrField(0) = "DEPT_CD"						' Field��(0)
'			arrField(1) = "DEPT_NM"						' Field��(1)
'   
'			arrHeader(0) = "�μ��ڵ�"					' Header��(0)
'			arrHeader(1) = "�μ���"						' Header��(1)
		Case 4			'���¹�ȣ 
			arrParam(0) = "���¹�ȣ �˾�" 							' �˾� ��Ī 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, B_MINOR C, B_MINOR D, F_DPST E " 		' TABLE ��Ī 
			arrParam(2) = strCode 								' Code Condition 
			arrParam(3) = "" 									' Name Condition 
			arrParam(4) = "A.BANK_CD = B.BANK_CD " 						' Where Condition 
			arrParam(4) = arrParam(4) & "AND C.MAJOR_CD = " & FilterVar("F3011", "''", "S") & "  AND C.MINOR_CD = B.BANK_ACCT_TYPE " 
			arrParam(4) = arrParam(4) & "AND D.MAJOR_CD = " & FilterVar("F3012", "''", "S") & "  AND D.MINOR_CD = B.DPST_TYPE " 
			arrParam(4) = arrParam(4) & "AND (E.DPST_FG = " & FilterVar("SV", "''", "S") & "  OR E.DPST_FG = " & FilterVar("ET", "''", "S") & " ) " 
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = E.BANK_ACCT_NO " 
			arrParam(4) = arrParam(4) & "AND B.BANK_CD = E.BANK_CD " 
			arrParam(5) = "���¹�ȣ" 							' �����ʵ��� �� ��Ī 
				
			arrField(0) = "B.BANK_ACCT_NO" 							' Field��(0) 
			arrField(1) = "A.BANK_CD" 										' Field��(1) 
			arrField(2) = "A.BANK_NM" 										' Field��(2) 
			arrField(3) = "C.MINOR_NM" 									' Field��(3) 
			arrField(4) = "D.MINOR_NM" 									' Field��(4) 
			arrField(5) = "HH" & parent.gColSep & "C.MINOR_CD" 					' Field��(5) - Hidden 
			arrField(6) = "HH" & parent.gColSep & "D.MINOR_CD" 					' Field��(6) - Hidden  

			arrHeader(0) = "���¹�ȣ" 								' Header��(0) 
			arrHeader(1) = "�����ڵ�" 								' Header��(1) 
			arrHeader(2) = "�����" 									' Header��(2)
			arrHeader(3) = "�����ݱ���" 							' Header��(3) 
			arrHeader(4) = "����������" 							' Header��(4)
		
		Case 6, 7		' �ŷ�ó(tab1,2)	
			arrParam(0) = "�ŷ�ó �˾�"					' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER" 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�ŷ�ó"						' �����ʵ��� �� ��Ī 

			arrField(0) = "BP_CD"						' Field��(0)
			arrField(1) = "BP_NM"						' Field��(1)
    
			arrHeader(0) = "�ŷ�ó�ڵ�"					' Header��(0)
			arrHeader(1) = "�ŷ�ó��"					' Header��(1)
		
		Case 8, 10			'����ī���ȣ 
			arrParam(0) = "����ī���ȣ �˾�"			' �˾� ��Ī 
			arrParam(1) = "F_NOTE" 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "NOTE_FG = " & FilterVar("CP", "''", "S") & "  AND (NOTE_STS = " & FilterVar("OC", "''", "S") & "  OR NOTE_STS = " & FilterVar("DC", "''", "S") & "  ) "	' Where Condition
			arrParam(5) = "����ī��"						' �����ʵ��� �� ��Ī 

			arrField(0) = "NOTE_NO"						' Field��(0)			
    
			arrHeader(0) = "����ī���ȣ"					' Header��(0)
						
		Case 9			'ī��� 
			arrParam(0) = "ī��� �˾�"			' �˾� ��Ī 
			arrParam(1) = "B_CARD_CO" 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "ī���"						' �����ʵ��� �� ��Ī 

			arrField(0) = "CARD_CO_CD"						' Field��(0)
			arrField(1) = "CARD_CO_NM"						' Field��(1)
    
			arrHeader(0) = "ī����ڵ�"					' Header��(0)
			arrHeader(1) = "ī����"					' Header��(1)
			
		Case 11
			If frm1.txtNoteAcctCd.className = "protected" Then Exit Function    

			arrParam(0) = "��ݰ����˾�"								' �˾� ��Ī 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM D	"				' TABLE ��Ī 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""																' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN006", "''", "S") & "  AND D.TRANS_TYPE = " & FilterVar("FN006", "''", "S") & "  " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD  "
			arrParam(4) = arrParam(4) & " AND C.JNL_CD= D.JNL_CD AND D.SEQ = C.SEQ"
			arrParam(4) = arrParam(4) & " AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND ((C.JNL_CD = " & FilterVar("CR", "''", "S") & "  and C.DR_CR_FG =  " & FilterVar("DR", "''", "S") & " ) "
			arrParam(4) = arrParam(4) & " OR    (C.JNL_CD = " & FilterVar("CP", "''", "S") & "  and C.DR_CR_FG =  " & FilterVar("CR", "''", "S") & " )) "			
			arrParam(4) = arrParam(4) & " AND C.JNL_CD =  " & FilterVar("CP", "''", "S") & "  "	 				
			If frm1.txtRcptType.Value<>"" then
				arrParam(4) = arrParam(4) & " AND D.EVENT_CD =  " & FilterVar(UCase(frm1.txtRcptType.Value), "''", "S")
			End if
			arrParam(5) = frm1.txtNoteAcctCd.Alt							' �����ʵ��� �� ��Ī 
			
			arrField(0) = "A.ACCT_CD"									' Field��(0)
			arrField(1) = "A.ACCT_NM"									' Field��(1)
			arrField(2) = "B.GP_CD"										' Field��(2)
			arrField(3) = "B.GP_NM"					 					' Field��(3)
			
			arrHeader(0) = frm1.txtNoteAcctCd.Alt									' Header��(0)
			arrHeader(1) = frm1.txtNoteAcctNm.Alt								' Header��(1)
			arrHeader(2) = "�׷��ڵ�"									' Header��(2)
			arrHeader(3) = "�׷��"										' Header��(3)		
				
		Case 12
			If frm1.txtChargeAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "����������˾�"								' �˾� ��Ī 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM 	D	"			' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""													' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN006", "''", "S") & "  " 					' Where Condition
			arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
			arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	C.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  "
			arrParam(4) = arrParam(4) & " AND 	C.TRANS_TYPE = D.TRANS_TYPE "
			arrParam(4) = arrParam(4) & " AND 	C.JNL_CD = D.JNL_CD "  
			arrParam(4) = arrParam(4) & "	 AND 	C.DR_CR_FG = D.DR_CR_FG "
			arrParam(4) = arrParam(4) & "	 AND 	C.SEQ = D.SEQ "			
			arrParam(4) = arrParam(4) & " AND  C.JNL_CD = " & FilterVar("CP", "''", "S") & "  AND D.EVENT_CD = " & FilterVar("CC", "''", "S") & "   " 
			arrParam(5) = frm1.txtChargeAcctCd.Alt							' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"									' Field��(0)
			arrField(1) = "A.Acct_NM"									' Field��(1)
			arrField(2) = "B.GP_CD"										' Field��(2)
			arrField(3) = "B.GP_NM"										' Field��(3)
			
			arrHeader(0) = frm1.txtChargeAcctCd.Alt									' Header��(0)
			arrHeader(1) = frm1.txtChargeAcctNm.Alt								' Header��(1)
			arrHeader(2) = "�׷��ڵ�"									' Header��(2)
			arrHeader(3) = "�׷��"										' Header��(3)	


		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	If (iWhere = 1 or iWhere = 4)Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    ElseiF (iWhere = 8 or iWhere = 10)Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	
		With frm1
			Select Case iWhere
				Case 0		' �Ա�/������� 
					.txtRcptType.focus
				Case 1		' ����(rcpt_type)
					.txtBankCd.focus
				Case 2		' ��������(tab1)
					
				Case 3		' �μ� 
					.txtDeptCd.focus
				Case 4		' ���¹�ȣ 
					.txtBankAcctNo.focus
				Case 5		' ��������(tab2)
					.txtBankCd2.focus
				Case 6		' �ŷ�ó(tab1)
					.txtBpCd1.focus
				Case 7		' �ŷ�ó(tab2)
					.txtBpCd2.focus
				Case 8		' Fromī�� 
					.txtFrNoteNo.focus
				Case 9		' ī��� 
				    if gSelframeFlg = TAB1 then		' ī���(tab2)
						.txtCardCoCd1.focus
					else
						.txtCardCoCd2.focus
					end if
				
				Case 11
					.txtNoteAcctCd.focus
				Case 12	'����������ڵ� 
					.txtChargeAcctCd.focus
			End Select
		End With
		Exit Function
	End If	

	With frm1
		Select Case iWhere
			Case 0		' �Ա�/������� 
				.txtRcptType.value	= arrRet(0)
				.txtRcptTypeNm.value= arrRet(1)
				Call txtRcptType_OnChange()
				.txtRcptType.focus
			Case 1		' ����(rcpt_type)
				.txtBankCd.value	= arrRet(0)
				.txtBankNm.value	= arrRet(1)
				.txtBankAcctNo.value =  arrRet(2)
				.txtBankCd.focus
			Case 2		' ��������(tab1)
				
			Case 3		' �μ� 
				.txtDeptCd.value	= arrRet(0)
				.txtDeptNm.value	= arrRet(1)
				.txtDeptCd.focus
			Case 4		' ���¹�ȣ 
				.txtBankAcctNo.value =  arrRet(0)
				.txtBankCd.value	= arrRet(1)
				.txtBankNm.value	= arrRet(2)					
				.txtBankAcctNo.focus
			Case 5		' ��������(tab2)
				.txtBankCd2.value	= arrRet(0)
				.txtBankNM2.value	= arrRet(1)	
				.txtBankCd2.focus
			Case 6		' �ŷ�ó(tab1)
				.txtBpCd1.value		= arrRet(0)
				.txtBpNM1.value		= arrRet(1)	
				.txtBpCd1.focus
			Case 7		' �ŷ�ó(tab2)
				.txtBpCd2.value		= arrRet(0)
				.txtBpNM2.value		= arrRet(1)	
				.txtBpCd2.focus
			Case 8		' Fromī�� 
				.txtFrNoteNo.value	= arrRet(0)				
				.txtFrNoteNo.focus
			Case 9		' ī��� 
			    if gSelframeFlg = TAB1 then		' ī���(tab2)
					.txtCardCoCd1.value	= arrRet(0)
					.txtCardCoNm1.value	= arrRet(1)		
					.txtCardCoCd1.focus
				else
					.txtCardCoCd2.value	= arrRet(0)
					.txtCardCoNm2.value	= arrRet(1)		
					.txtCardCoCd2.focus
				end if
			
			Case 11
				.txtNoteAcctCd.value	= arrRet(0)
				.txtNoteAcctNm.value	= arrRet(1)
				.txtNoteAcctCd.focus
			Case 12	'����������ڵ� 

				.txtChargeAcctCd.value = arrRet(0)
				.txtChargeAcctNm.value = arrRet(1)
				.txtChargeAcctCd.focus
		End Select
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
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.PARENT,arrParam), _
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
		Else
'			Call DisplayMsgBox("900025","X","X","X")
			Exit Function
		End If
	End With						'Reference��ȣ 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.PARENT,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDt")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'�μ��ڵ� 
	arrParam(1) = frm1.txtGLDt.Text			'��¥(Default:������)
	arrParam(2) = "1"							'�μ�����(lgUsrIntCd)
	
	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID	
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	End If
	
	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	Call txtDeptCD_onBlur()
	frm1.txtDeptCd.focus
	
	lgBlnFlgChgValue = True
End Function

'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	arrParam(5) = "����� �ڵ�"			

    arrField(0) = "BIZ_AREA_CD"					' Field��(0)
    arrField(1) = "BIZ_AREA_NM"					' Field��(1)

    arrHeader(0) = "������ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If
End Function


'=======================================================================================================
'	Name : SetReturnVal()
'	Description : 
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		case 0
			frm1.txtfromBizAreaCd.Value	= arrRet(0)
			frm1.txtfromBizAreaNm.Value	= arrRet(1)
			frm1.txtfromBizAreaCd.focus
		case 1
			frm1.txttoBizAreaCd.Value = arrRet(0)
			frm1.txttoBizAreaNm.Value = arrRet(1)
			frm1.txttoBizAreaCd.focus
		case 2
			frm1.txtfromBizAreaCd1.Value = arrRet(0)
			frm1.txtfromBizAreaNm1.Value = arrRet(1)
			frm1.txtfromBizAreaCd1.focus
		case 3
			frm1.txttoBizAreaCd1.Value	= arrRet(0)
			frm1.txttoBizAreaNm1.Value	= arrRet(1)
			frm1.txttoBizAreaCd1.focus
	End Select
	
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
    Call ggoOper.ClearField(Document, "2")		'��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")		'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    
    Call InitSpreadSheet("A")                                                        'Setup the Spread sheet
    Call InitSpreadSheet("B")                                                       'Setup the Spread sheet

	
	Call txtRcptType_OnChange()
    Call SetDefaultVal
    Call ClickTab1
    gIsTab     = "Y" 
	gTabMaxCnt = 2  	

    Call SetToolbar("1100000000011111")										'��: ��ư ���� ���� 
	
	frm1.txtDueDtEnd.focus
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


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PROC_CHK			= iCurColumnPos(1)
			C_NOTE_NO			= iCurColumnPos(2)
			C_NOTE_AMT			= iCurColumnPos(3)
			C_DUE_DT			= iCurColumnPos(4)
			C_NOTE_STS			= iCurColumnPos(5)
			C_BP_CD				= iCurColumnPos(6)
			C_BP_NM				= iCurColumnPos(7)
			C_CARD_CO_CD		= iCurColumnPos(8)
			C_CARD_CO_NM		= iCurColumnPos(9)
			C_DEPT_CD			= iCurColumnPos(10)
			C_DEPT_NM			= iCurColumnPos(11)
			C_NOTE_ITEM_DESC	= iCurColumnPos(12)
			C_GL_NO				= iCurColumnPos(13)
			C_TEMP_GL_NO		= iCurColumnPos(14)
			C_COL_END			= iCurColumnPos(15)
      Case "B"
      
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_CNCL_CHK			= iCurColumnPos(1)
			C_CNCL_NOTE_NO		= iCurColumnPos(2)
			C_CNCL_TEMP_GL_NO	= iCurColumnPos(3)
			C_CNCL_TEMP_GL_DT	= iCurColumnPos(4)
			C_CNCL_GL_NO		= iCurColumnPos(5)
			C_CNCL_GL_DT		= iCurColumnPos(6)	
			C_CNCL_NOTE_AMT		= iCurColumnPos(7)
			C_CNCL_BP_CD		= iCurColumnPos(8)
			C_CNCL_BP_NM		= iCurColumnPos(9)
			C_CNCL_BANK_CD		= iCurColumnPos(10)
			C_CNCL_BANK_NM		= iCurColumnPos(11)
			C_CNCL_DEPT_CD		= iCurColumnPos(12)
			C_CNCL_DEPT_NM		= iCurColumnPos(13)
			C_CNCL_NOTE_ITEM_DESC = iCurColumnPos(14)
			C_CNCL_COL_END		= iCurColumnPos(15)
            
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
'	Event�� �浹�� �����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'=======================================================================================================
'   Event Desc : �Ա������� Set Protected/Required Fields
'=======================================================================================================
Sub txtRcptType_OnChange()
	'�����ڵ�, ���¹�ȣ Protected Setting
	Dim strval
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	strval = UCase(frm1.txtRcptType.value)
	
	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
	
			Select Case UCase(lgF0)
				Case "DP" & Chr(11)			' ������			
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "N")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "N")
				Case Else
					frm1.txtBankCd.value = ""
					frm1.txtBankNm.value = ""
					frm1.txtBankAcctNo.value = ""
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")									
			End Select
	Else
			frm1.txtBankCd.value = ""
			frm1.txtBankNm.value = ""
			frm1.txtBankAcctNo.value = ""
			Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")											
	End If 
	
End Sub

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDueDtEnd_DblClick(Button)
	if Button = 1 then
		frm1.txtDueDtEnd.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDueDtEnd.Focus
	End if
End Sub

Sub txtGLDt_DblClick(Button)
	if Button = 1 then
		frm1.txtGLDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtGLDt.Focus
	End if
End Sub

Sub txtStsDtStart_DblClick(Button)
	if Button = 1 then
		frm1.txtStsDtStart.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtStsDtStart.Focus
	End if
End Sub

Sub txtStsDtEnd_DblClick(Button)
	if Button = 1 then
		frm1.txtStsDtEnd.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtStsDtEnd.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name :txtDueDt_keypress(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDueDtEnd_KeyPress(KeyAscii)
	If KeyAscii = 13 Then  
		frm1.txtBpCd1.focus
		frm1.txtDueDtEnd.focus
		Call MainQuery
	End If   
End Sub

Sub txtStsDtStart_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtStsDtEnd.focus 
	   Call MainQuery
	End If   
End Sub

Sub txtStsDtEnd_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtStsDtStart.focus 
	   Call MainQuery
	End If   
End Sub

'=======================================================================================================
'   Event Name : txtChargeAmt_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtChargeAmt_Change()
	lgBlnFlgChgValue = True
	If unicdbl(frm1.txtChargeAmt.Text) > 0 Then	
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "N")		
	ElseIf  unicdbl(frm1.txtChargeAmt.Text) <= 0 Then			
		frm1.txtChargeAcctCd.value = ""
		frm1.txtChargeAcctNm.value = ""		
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "Q")			
	End If		
		
End Sub

'=======================================================================================================
'   Event Name : txtDueDtEnd_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDueDtEnd_Change()
    'lgBlnFlgChgValue = True
End Sub

Sub txtGLDt_Change()

	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtGLDt.Text <> "") Then
	
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtDeptCd.value = ""
					frm1.txtDeptNm.value = ""
					frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
		End If
	End If
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDeptCD_onBlur()
'   Event Desc : 
'=======================================================================================================
Sub txtDeptCD_OnBlur()
	If frm1.txtDeptCD.value = "" Then
		frm1.txtDeptNm.value = ""
		Exit Sub
	End If

	Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii

	If Trim(frm1.txtDeptCd.value) = "" And Trim(frm1.txtGLDt.Text = "") Then Exit Sub

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"			

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hOrgChangeId.value = ""
	Else 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
					
		For ii = 0 to Ubound(arrVal1,1) - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			frm1.hOrgChangeId.value = Trim(arrVal2(2))
		Next	
	End If

    lgBlnFlgChgValue = True
End Sub

Sub txtBankCd_onBlur()
	if frm1.txtBankCd.value = "" then
		frm1.txtBankNm.value = ""
	end if
End Sub	

Sub txtRcptType_onBlur()
	if frm1.txtRcptType.value = "" then
		frm1.txtRcptTypeNm.value = ""
	end if
End Sub	

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
	frm1.txtDueDtEnd.focus
						 
End Function

Function ClickTab2()

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetToolBar("1100000000001111")
	ELSE                 
		Call SetToolBar("1100000000001111")
	END IF	
	
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)														'�ι�° Tab 
	
	gSelframeFlg = TAB2
	frm1.hProcFg.value = "DG"
	frm1.txtStsDtStart.focus
	
	
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
'   Event Name : vspdData_TopLeftChange
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
    Dim iColumnName
    
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
    
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
   Dim iColumnName
    
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
       Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
    
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
    
    

End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    
    
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
    
    

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
    
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	               
    	If (lgStrPrevKeyNoteNo <> "" or  lgStrPrevKeyGlNo <> "") Then       	
      	   Call DbQuery
    	End If
    End if  
    
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

Sub vspdData_GotFocus()
    
    ggoSpread.Source = Frm1.vspdData
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

Sub vspdData2_GotFocus()
    
    ggoSpread.Source = Frm1.vspdData2
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyNoteNo <> "" Then                         
      	   Call DbQuery
    	End If
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

'	ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
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
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
				ggoSpread.SSDeleteFlag Row,Row
			End If			
		End If
		
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
				ggoSpread.UpdateRow Row
				.col = C_NOTE_AMT
				lstxtPlanAmtSum = UNIFormatNumber(UNICDbl(frm1.txtSumNoteAmt.Text) + UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				frm1.txtSumNoteAmt.Text = lstxtPlanAmtSum
			Else
				ggoSpread.SSDeleteFlag Row,Row				
				.col = C_NOTE_AMT
				lstxtPlanAmtSum = UNIFormatNumber(UNICDbl(frm1.txtSumNoteAmt.Text) - UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				frm1.txtSumNoteAmt.Text = lstxtPlanAmtSum
			End If		
		End If
		
	End With
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
    
   FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

   '-----------------------
    'Erase contents area
    '----------------------- 
    
    Call ggoOper.ClearField(Document, "2")			'��: Clear Contents  Field
    if gSelframeFlg = TAB1 Then
       Call InitSpreadSheet("A")
    Else
       Call InitSpreadSheet("B")
    end if
    
    
    Call InitVariables 
	
    frm1.vspdData.MaxRows = 0
	
    '-----------------------
    'Check condition area
    '----------------------- 
'    If gSelframeFlg = TAB1 Then
		If Not chkField(Document, "1") Then									'��: This function check indispensable field     			
		   Exit Function
		End If
'    ElseIf gSelframeFlg = TAB2 Then
'		If Not chkField(Document, "3") Then									'��: This function check indispensable field     			
'		   Exit Function
'		End If
'	End If
	
    If frm1.txtCardCoCd1.value = "" Then
		frm1.txtCardCoNm1.value = ""
	End If
	
	If frm1.txtfromBizAreaCd.value = "" Then
		frm1.txtfromBizAreaNm.value = ""
	End If
	
	If frm1.txttoBizAreaCd.value = "" Then
		frm1.txttoBizAreaNm.value = ""
	End If
	
	If frm1.txtfromBizAreaCd1.value = "" Then
		frm1.txtfromBizAreaNm1.value = ""
	End If
	
	If frm1.txttoBizAreaCd1.value = "" Then
		frm1.txttoBizAreaNm1.value = ""
	End If
	
    Call ggoOper.LockField(Document, "N")			'��: This function lock the suitable field
    '-----------------------
    'Query function call area
    '----------------------- 
    If gSelframeFlg = "1" Then 	
      if frm1.txtBpCd1.value <> "" then
    
       If CommonQueryRs(" A.BP_NM ","B_BIZ_PARTNER A","A.BP_CD = " & FilterVar(frm1.txtBpCd1.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   Call DisplayMsgBox("970000","X",frm1.txtBpCd1.alt,"X")            '�� : No data is found. 
 	    Exit Function
	   End If
	  End If
      if frm1.txtCardCoCd1.value <> "" then
	  If CommonQueryRs(" A.CARD_CO_NM ","B_CARD_CO A","A.PAY_CARD_FG = " & FilterVar("Y", "''", "S") & "  AND A.CARD_CO_CD = " & FilterVar(frm1.txtCardCoCd1.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  Call DisplayMsgBox("970000","X",frm1.txtCardCoCd1.alt,"X")            '�� : No data is found. 
	  Exit Function
	  End If
	  End If
	  
	  If Trim(frm1.txtfromBizAreaCd.value) <> "" and   Trim(frm1.txttoBizAreaCd.value) <> "" Then				
		If UCase(Trim(frm1.txtfromBizAreaCd.value)) > UCase(Trim(frm1.txttoBizAreaCd.value)) Then
				'msgbox frm1.txtfromBizAreaCd.value & " " & frm1.txttoBizAreaCd.value
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtfromBizAreaCd.Alt, frm1.txttoBizAreaCd.Alt)
			frm1.txtfromBizAreaCd.focus
			Exit Function
		End If
	  End If
	  
	  if frm1.txtfromBizAreaCd.value <> "" then
		If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtfromBizAreaCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("970000","X",frm1.txtfromBizAreaCd.alt,"X")            '�� : No data is found. 
			frm1.txtfromBizAreaCd.focus
 			Exit Function
		End If
	  End If
	  
	  if frm1.txttoBizAreaCd.value <> "" then
		If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txttoBizAreaCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("970000","X",frm1.txttoBizAreaCd.alt,"X")            '�� : No data is found. 
			frm1.txttoBizAreaCd.focus
 			Exit Function
		End If
	  End If
	  
     Else
     
      if frm1.txtBpCd2.value <> "" then
    
       If CommonQueryRs(" A.BP_NM ","B_BIZ_PARTNER A","A.BP_CD = " & FilterVar(frm1.txtBpCd2.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   Call DisplayMsgBox("970000","X",frm1.txtBpCd2.alt,"X")            '�� : No data is found. 
 	    Exit Function
	   End If
	   End If
      if frm1.txtCardCoCd2.value <> "" then
	  If CommonQueryRs(" A.CARD_CO_NM ","B_CARD_CO A","A.PAY_CARD_FG = " & FilterVar("Y", "''", "S") & "  AND A.CARD_CO_CD = " & FilterVar(frm1.txtCardCoCd2.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  Call DisplayMsgBox("970000","X",frm1.txtCardCoCd2.alt,"X")            '�� : No data is found. 
	  Exit Function
	  End If
	  End If
	  
	  If Trim(frm1.txtfromBizAreaCd1.value) <> "" and   Trim(frm1.txttoBizAreaCd1.value) <> "" Then				
		If UCase(Trim(frm1.txtfromBizAreaCd1.value)) > UCase(Trim(frm1.txttoBizAreaCd1.value)) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtfromBizAreaCd1.Alt, frm1.txttoBizAreaCd1.Alt)
			frm1.txtfromBizAreaCd1.focus
			Exit Function
		End If
	  End If
	  
	  if frm1.txtfromBizAreaCd1.value <> "" then
		If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtfromBizAreaCd1.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("970000","X",frm1.txtfromBizAreaCd1.alt,"X")            '�� : No data is found. 
			frm1.txtfromBizAreaCd1.focus
 			Exit Function
		End If
	  End If
	  
	  if frm1.txttoBizAreaCd1.value <> "" then
		If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txttoBizAreaCd1.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("970000","X",frm1.txttoBizAreaCd1.alt,"X")            '�� : No data is found. 
			frm1.txttoBizAreaCd1.focus
 			Exit Function
		End If
	  End If
	  
     End if
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
     On Error Resume Next                                                   '��: Protect system from crashing
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
	Set gActiveElement = document.activeElement    
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
Function FncSplitColumn()
		Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
	
	if gMouseClickStatus = "SPCRP" then
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
	
	End If
	
	If gMouseClickStatus = "SP2CRP" Then
	
	iColumnLimit = frm1.vspdData2.MaxCols
	
	ACol = frm1.vspdData2.ActiveCol
	ARow = frm1.vspdData2.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData2.ScrollBars = Parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData2.Col = ACol
	frm1.vspdData2.Row = ARow
	frm1.vspdData2.Action = Parent.SS_ACTION_ACTIVE_CELL
	frm1.vspdData2.ScrollBars = Parent.SS_SCROLLBAR_BOTH
	
	
	end if
End Function


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
    if gSelframeFlg = TAB1 Then
    Call InitSpreadSheet("A") 
    else
    Call InitSpreadSheet("B")      
    end if
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ������� ȣ��Ǵ� ���� Function  *******************************
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
	
	Call txtRcptType_OnChange()
		
	With frm1
		If gSelframeFlg = "1" Then 														'��: �ϰ�ó��(tab1) ��ȸ 
		    If lgIntFlgMode = Parent.OPMD_UMODE Then		    
				strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
				strVal = strVal & "&cboProcFg=" & Trim(frm1.hProcFg.value)				'��: ��ȸ ���� ����Ÿ				
				strVal = strVal & "&txtDueDtEnd=" & Trim(frm1.hDueDtEnd.value)
				strVal = strVal & "&txtBpCd=" & Trim(frm1.hBpCd1.value)
				strVal = strVal & "&txtCardCoCd=" & Trim(frm1.hCardCoCd1.value)
				strVal = strVal & "&txtBizAreaCd=" & Trim(.hfromtxtBizAreaCd.value)
				strVal = strVal & "&txtBizAreaCd1=" & Trim(.htotxtBizAreaCd.value) 
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo=" & lgStrPrevKeyGlNo
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			Else
				strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
				strVal = strVal & "&cboProcFg=" & Trim("CG")							'��: ��ȸ ���� ����Ÿ				
				strVal = strVal & "&txtDueDtEnd=" & Trim(frm1.txtDueDtEnd.Text)
				strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd1.value)
				strVal = strVal & "&txtCardCoCd=" & Trim(frm1.txtCardCoCd1.value)
				strVal = strVal & "&txtBizAreaCd=" & Trim(.txtfromBizAreaCd.value)
				strVal = strVal & "&txtBizAreaCd1=" & Trim(.txttoBizAreaCd.value)
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo=" & lgStrPrevKeyGlNo				
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			End If   			
		Else 																			'��: �ϰ����(tab2) ��ȸ																				
		    If lgIntFlgMode = Parent.OPMD_UMODE Then
				strVal = BIZ_PGM_ID3 & "?txtMode=" & Parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
				strVal = strVal & "&cboProcFg=" & Trim(frm1.hProcFg.value)				'��: ��ȸ ���� ����Ÿ 		strVal = strVal & "&txtStsDtStart=" & Trim(frm1.hStsDtStart.value)			
				strVal = strVal & "&txtStsDtEnd=" & Trim(frm1.hStsDtEnd.value)
				strVal = strVal & "&txtBpCd=" & Trim(frm1.hBpCd2.value)
				strVal = strVal & "&txtBankCd=" & Trim(frm1.hCardCoCd2.value)
				strVal = strVal & "&txtBizAreaCd=" & Trim(.hfromtxtBizAreaCd1.value) 
				strVal = strVal & "&txtBizAreaCd1=" & Trim(.htotxtBizAreaCd1.value)
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo=" & lgStrPrevKeyGlNo
				strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			Else
				strVal = BIZ_PGM_ID3 & "?txtMode=" & Parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 				
				strVal = strVal & "&cboProcFg=" & Trim("DG")							'��: ��ȸ ���� ����Ÿ 
				strVal = strVal & "&txtStsDtStart=" & Trim(frm1.txtStsDtStart.Text)
				strVal = strVal & "&txtStsDtEnd=" & Trim(frm1.txtStsDtEnd.Text)
				strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd2.value)				
				strVal = strVal & "&txtCardCoCd=" & Trim(frm1.txtCardCoCd2.value)
				strVal = strVal & "&txtBizAreaCd=" & Trim(.txtfromBizAreaCd1.value)
				strVal = strVal & "&txtBizAreaCd1=" & Trim(.txttoBizAreaCd1.value)
				strVal = strVal & "&lgStrPrevKeyNoteNo=" & lgStrPrevKeyNoteNo
				strVal = strVal & "&lgStrPrevKeyGlNo=" & lgStrPrevKeyGlNo
				strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
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
Function DbQueryOk()
	
	If gSelframeFlg = "2" or gSelframeFlg = TAB2 Then 					
		Call SetSpreadLock("C")
	End If 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE	'��: Indicates that current mode is Update mode
    
	lgBlnFlgChgValue = False
	

    Call ggoOper.LockField(Document, "Q")	'��: This function lock the suitable field
    Call txtRcptType_OnChange()
    frm1.txtGLDt.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)	
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
    'On Error Resume Next		'��: Protect system from crashing

	'2001.03.01 added
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		IntRetCD = DisplayMsgBox("900002","x","x","x")  '��ȸ�� ���� �Ͻʽÿ�.
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"x","x")	'�۾��� �����Ͻðڽ��ϱ�?

	If IntRetCD = vbNo Then
		Exit Function
	End If

	If frm1.hProcFg.value = "CG" Then
		If Not chkField(Document, "2") Then                                   '��: Check contents area
			Exit Function
		End If
	End If
    
	IF Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If
	
	' to above statement.

	'Call LayerShowHide(1)
	
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
		If .hProcFg.value = "CG" Then										'��:�ϰ�ó�� 
			For lRow = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow
				
				.vspdData.Col = C_PROC_CHK
				
				If .vspdData.Text = "1" Then
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData.Col = C_NOTE_NO
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ������ȣ 
					.vspdData.Col = C_GL_NO
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ��ǥ��ȣ 
					.vspdData.Col = C_TEMP_GL_NO
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ��ǥ��ȣ  
					.vspdData.Col = C_BP_CD
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ��ǥ��ȣ  
					.vspdData.Col = C_NOTE_ITEM_DESC
					strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep		' ��� 
				
					lGrpCnt = lGrpCnt + 1
				End If
				
			Next	
		ElseIf .hProcFg.value = "DG" Then									 '��:�ϰ���� 
			For lRow = 1 To .vspdData2.MaxRows
				.vspdData2.Row = lRow
				
				.vspdData2.Col = C_CNCL_CHK
				
				If .vspdData2.Text = "1" Then
					strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData2.Col = C_CNCL_NOTE_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep		' ������ȣ 
					.vspdData2.Col = C_CNCL_GL_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep		' ��ǥ��ȣ 
					.vspdData2.Col = C_CNCL_TEMP_GL_NO
					strVal = strVal & Trim(.vspdData2.Text) & Parent.gRowSep		' ��ǥ��ȣ  				
						
					lGrpCnt = lGrpCnt + 1
				End If
				
			Next	
		End If
		
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal
		
		If .txtMaxRows.value <= 0 Then
			Call DisplayMsgBox("900025","x","x","x")  '���õ� �׸��� �����ϴ�.
			Exit Function
		End If

		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end
		
		Call LayerShowHide(1)
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'��: �����Ͻ� ASP �� ���� 
		
	End With

    DbSave = True                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()			'��: ���� ������ ���� ���� 
   
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
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���ұ���ī�����</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>|
											<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
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
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=OBJECT1 name=txtDueDtEnd CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT></TD> 																												 
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP>							
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd1" NAME="txtBpCd1" SIZE=10 MAXLENGTH=10   tag="1XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd1.value, 6)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM1" NAME="txtBpNM1" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="14X" ALT="�ŷ�ó"> </TD>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtfromBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtfromBizAreaCd.value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtfromBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ī���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtCardCoCd1" NAME="txtCardCoCd1" SIZE=10 MAXLENGTH=10  tag="1XX" ALT="����ī����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtCardCoCd1.Value, 9)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtCardCoNm1" NAME="txtCardCoNm1" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="14X" ALT="����ī����"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txttoBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txttoBizAreaCd.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txttoBizAreaNm" SIZE=30 tag="14"></TD>
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
								<TD CLASS=TD5 NOWRAP>ȸ������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="ȸ������" tag="22X1" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" Size= "10" MAXLENGTH="10" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUpDept(frm1.txtDeptCd.value, 3)">&nbsp;<INPUT NAME="txtDeptNm" ALT="�μ���" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRcptType" ALT="��������ڵ�" SIZE="10" MAXLENGTH="2" tag="22XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnRcptType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtRcptType.value, 0)">&nbsp;<INPUT NAME="txtRcptTypeNm" ALT="���������" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>��ݰ���</TD>												
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtNoteAcctCd" ALT="��ݰ���" SIZE="10" MAXLENGTH="20"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteAcctCd.value, 11)">
													   <INPUT NAME="txtNoteAcctNm" ALT="��ݰ�����" SIZE="20" tag="24X"></TD>								
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="21XXXU" ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 1)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNm" NAME="txtBankNm" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="�����"></TD>
								<TD CLASS=TD5 NOWRAP>���¹�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankAcctNo" NAME="txtBankAcctNo" SIZE=20 MAXLENGTH=30 tag="21XXXU" ALT="���¹�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcctNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankAcctNo.Value, 4)"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=fpCharge NAME=txtChargeAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="���������" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>���������</TD>												
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtChargeAcctCd" ALT="���������" SIZE="10" MAXLENGTH="20"  tag="24X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChargeAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtChargeAcctCd.value, 12)">
													   <INPUT NAME="txtChargeAcctNm" ALT="�����������" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="23" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TDT">
									<TD CLASS="TD6">
									<TD CLASS="TD5" NOWRAP>�����Ѿ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtSumNoteAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 160px" title="FPDOUBLESINGLE" ALT="�����Ѿ�" tag="24X2Z"> </OBJECT>');</SCRIPT>&nbsp;
				                    </TD>
								</TR>
							</TABLE>
						</FIELDSET>
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
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtStsDtStart CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="����ȸ����" tag="12X1"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 name=txtStsDtEnd CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="����ȸ����" tag="12X1" ></OBJECT>');</SCRIPT></TD>																																
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>						
								</TR>
								<TR>									
									<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd2" NAME="txtBpCd2" SIZE=10 MAXLENGTH=10   tag="1XXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd2.Value, 7)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM2" NAME="txtBpNM2" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="14X" ALT="�ŷ�ó"> </TD>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtfromBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtfromBizAreaCd1.value, 2)">&nbsp;<INPUT TYPE=TEXT NAME="txtfromBizAreaNm1" SIZE=30 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ī���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtCardCoCd2" NAME="txtCardCoCd2" SIZE=10 MAXLENGTH=10  tag="1XX" ALT="ī����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtCardCoCd2.Value, 9)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtCardCoNm2" NAME="txtCardCoNM2" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="14X" ALT="ī����"></TD>								
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txttoBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txttoBizAreaCd1.value, 3)">&nbsp;<INPUT TYPE=TEXT NAME="txttoBizAreaNm1" SIZE=30 tag="14"></TD>
								</TR>
							</TABLE>
							    <TR>
									<TD WIDTH=100% HEIGHT="100%" COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="23" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD WIDTH=100% HEIGHT="50%" colspan=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% TITLE="SPREAD" tag="23" ID=vspdData3> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
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
   					<TD><BUTTON NAME="button1" CLASS="CLSMBTN" ONCLICK="vbscript:DBSave()" Flag=1>����</BUTTON>
   					</TD>
   				</TR>
   			</TABLE> 
  		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="2" Tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="2" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hProcFg"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hNoteFg"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hDueDtEnd"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hBpCd1"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hBpCd2"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hCardCoCd1"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hCardCoCd2"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"		tag="1" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hStsDtStart"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hStsDtEnd"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtGlNo"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hfromtxtBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hfromtxtBizAreaCd1" tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htotxtBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htotxtBizAreaCd1"	tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtGLDt"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtCRAmt"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtCRLocAmt"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDRAmt"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDRLocAmt"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDocCur"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtXchRate"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtOrgChangeId"	tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtDeptCd"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtAcctCd"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="GtxtBankCd"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="CtxtBankAcctNo"	tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="DtxtNoteNo"		tag="2" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
