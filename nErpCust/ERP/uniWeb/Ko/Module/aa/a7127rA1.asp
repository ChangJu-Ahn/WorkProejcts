<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : ȸ�����
'*  2. Function Name        : �ڻ����
'*  3. Program ID           : Asset Acquisition Reference Popup
'*  4. Program Name         : �ڻ꺯�� ���� �˾�(ȸ�����-�ڻ����-�����ڻ����⳻�����-�ڻ꺯����ȣ ����)
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO�� �ۼ�
'*  7. Modified date(First) : 2001/02/20
'*  8. Modified date(Last)  : 2001/03/06
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : Kim Hee Jung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2001/02/20
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. �� �� ��
'############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim lgBlnFlgChgValue                                        '��: Variable is for Dirty flag            
Dim lgStrPrevKey                                            '��: Next Key tag                          
Dim lgSortKey                                               '��: Sort���� ���庯��                      
Dim lgIsOpenPop                                             '��: Popup status                           

Dim lgSelectList                                            '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 

Dim lgTypeCD                                                '��: 'G' is for group , 'S' is for Sort    
Dim lgFieldCD                                               '��: �ʵ� �ڵ尪                           
Dim lgFieldNM                                               '��: �ʵ� ����                           
Dim lgFieldLen                                              '��: �ʵ� ��(Spreadsheet����)              
Dim lgFieldType                                             '��: �ʵ� ����                           
Dim lgDefaultT                                              '��: �ʵ� �⺻��                           
Dim lgNextSeq                                               '��: �ʵ� Pair��                           
Dim lgKeyTag                                                '��: Key ����                                

Dim lgSortFieldNm                                           '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD                                          '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark

Dim IsOpenPop        

' ���Ѱ��� �߰�
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' �����
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ����

'---------------  coding part(�������,Start)-----------------------------------------------------------

'	EndDate = GetSvrDate                                           '��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ -----
'	StartDate = UNIDateAdd("m", -1, EndDate, parent.gServerDateFormat)    '��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ -----

'    Select Case Request("PID")
'		Case "A7107MA1"	
'			Call GetAdoFiledInf("A7107RA1","S","A")                        ' 2. G is for Qroup , S is for Sort     
'		Case "A7108MA1"	
'			Call GetAdoFiledInf("A7107RA3","S","A")                        ' 2. G is for Qroup , S is for Sort     		
'		Case "A7109MA1"
'			Call GetAdoFiledInf("A7107RA2","S","A")                        '��: spread sheet �ʵ����� query   -----		         
'    End Select 
'--------------- ������ coding part(�������,End)-------------------------------------------------------------


'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Dim arrReturn
Dim arrParent
Dim arrParam					

arrParent        = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)


Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
'##
dtToday = "<%=GetSvrDate%>"
Call PopupParent.ExtractDateFrom(dtToday, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

EndDate = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
StartDate = UNIDateClientFormat(PopupParent.gFiscStart)
'StartDate = UNIDateAdd("M", -1, EndDate, PopupParent.gDateFormat)
'��: Grid Columns

Dim C_AsstChgNo
Dim C_AsstCd
Dim C_AsstNm
Dim C_ChgDt
Dim C_ChgFgCd
Dim C_ChgFgNm
Dim C_FromDeptCd
Dim C_FromDeptNm
Dim C_DocCur
Dim C_BpCdCd
Dim C_BpCdNm
Dim C_ChgTotAmt
Dim C_ChgTotLocAmt
Dim C_AsstSoldDesc


Const BIZ_PGM_ID        = "a7127rb1.asp"
Const C_SHEETMAXROWS    = 16                                   '��: Spread sheet���� �������� row
Const C_SHEETMAXROWS_D  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ�
Const C_MaxKey          = 1	

Dim lsPoNo                                                 '��: Jump�� Cookie�� ���� Grid value


	 '------ Set Parameters from Parent ASP ------ 
	 'mmmmmm

    top.document.title = PopupParent.gActivePRAspName	
    
'	top.document.title = "�ڻ꺯�������˾�"

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

 '#########################################################################################################
'												2. Function��
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

	C_AsstChgNo		=1    
	C_AsstCd		=2
	C_AsstNm		=3      
	C_ChgDt			=4        
	C_ChgFgCd		=5     
	C_ChgFgNm		=6     
	C_FromDeptCd	=7  
	C_FromDeptNm	=8  
	C_DocCur		=9   
	C_BpCdCd		=10  
	C_BpCdNm		=11 
	C_ChgTotAmt		=12 
	C_ChgTotLocAmt	=13
	C_AsstSoldDesc	=14

End Sub

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    Redim arrReturn(0)
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = 0                                  'initializes Previous Key
    lgSortKey        = 1
    
	Self.Returnvalue = arrReturn
	
	' ���Ѱ��� �߰�
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If	
End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ����
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ����
'========================================================================================================= 
Sub SetDefaultVal()

	frm1.txtFrChgDt.Text = StartDate
	frm1.txtToChgDt.Text = EndDate
	Call Radio1_onChange()

End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
   ' ------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim strCodeList
	Dim strNameList
	ggoSpread.Source = frm1.vspdData
	Call CommonQueryRs("A.MINOR_CD,A.MINOR_NM","B_MINOR A", _
					   "A.MAJOR_CD = " & FilterVar("A2001", "''", "S") & "   ", _
	                   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	'A1006

	strCodeList = Replace(lgF0, Chr(11), vbTab)
	strNameList = Replace(lgF1, Chr(11), vbTab)

	ggoSpread.SetCombo strCodeList, C_ChgFgCd
	ggoSpread.SetCombo strNameList, C_ChgFgNm

    '------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

Function OpenPopUp(Byval PopFg, Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case PopFg
	Case  "DP"
	
			arrParam(0) = "�μ� �˾�"				' �˾� ��Ī
			arrParam(1) = "B_ACCT_DEPT"    			' TABLE ��Ī
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�μ��ڵ�"					' �����ʵ��� �� ��Ī

			arrField(0) = "DEPT_CD"	     				' Field��(0)
			arrField(1) = "DEPT_NM"			    		' Field��(1)
    
			arrHeader(0) = "�μ��ڵ�"					' Header��(0)
			arrHeader(1) = "�μ���"				' Header��(1)
	
	Case "FrChgNo" 
			If frm1.txtFrChgNo.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(0) = "�ڻ��ȣ �˾�"				' �˾� ��Ī
			arrParam(1) = "A_ASSET_CHG A ,A_ASSET_CHG_MASTER B"    			' TABLE ��Ī
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.ASST_CHG_NO=B.ASST_CHG_NO "
			'-- �μ��ڵ�
			If lgInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND A.FROM_INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = arrParam(4) & ""
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND A.FROM_INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			Else
				arrParam(4) = arrParam(4) & ""
			End If
			' �����
			' ���Ѱ��� �߰�
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = arrParam(4) & " AND A.FROM_BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = arrParam(4) & ""
			End If
			' �ۼ���
			' ���Ѱ��� �߰�
			If lgAuthUsrID <> "" Then
				arrParam(4) = arrParam(4) & " AND B.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			Else
				arrParam(4) = arrParam(4) & ""
			End If			
			
			arrParam(5) = "�ڻ��ȣ"					' �����ʵ��� �� ��Ī

			arrField(0) = "ASST_CD"	     						' Field��(0)
    
			arrHeader(0) = "�ڻ��ȣ"					' Header��(0)
	
	Case "ToAsstChgNo"  ,"FrAsstChgNo"	
			If frm1.txtToAsstChgNo.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If
			
			arrParam(0) = "�Ű�����ڻ� �˾�"				' �˾� ��Ī
			
			arrParam(1) = "A_ASSET_CHG_MASTER"    			' TABLE ��Ī
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "CHG_DT BETWEEN    " & FilterVar(UniConvDate(frm1.txtFrChgDt.text), "''", "S") & " AND  " & FilterVar(UniConvDate(frm1.txtToChgDt.text), "''", "S") & "  "							' Where Condition
			'-- �μ��ڵ�
			If lgInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND FROM_INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = arrParam(4) & ""
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND FROM_INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			Else
				arrParam(4) = arrParam(4) & ""
			End If
			' �����
			' ���Ѱ��� �߰�
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = arrParam(4) & " AND FROM_BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = arrParam(4) & ""
			End If
			' �ۼ���
			' ���Ѱ��� �߰�
			If lgAuthUsrID <> "" Then
				arrParam(4) = arrParam(4) & " AND UPDT_USR_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			Else
				arrParam(4) = arrParam(4) & ""
			End If			

			arrParam(5) = "�Ű�����ȣ"					' �����ʵ��� �� ��Ī

			arrField(0) = "ASST_CHG_NO"	     				' Field��(0)
			'arrField(1) = "ASST_NM"	     				' Field��(0)
			    
			arrHeader(0) = "�Ű�����ȣ"					' Header��(0)
			'arrHeader(1) = "�ڻ��"					' Header��(0)
	end SELECT
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(PopFg,arrRet)
	End If	

End Function

Function SetPopUp(Byval PopupFg,Byval arrRet)
	
	With frm1
	select case PopupFg
		case "DP"
			.txtDeptCd.value	 = arrRet(0)
			.txtDeptNm.value	 = arrRet(1)
			
		case "FrChgNo"
			.txtFrChgNo.value = arrRet(0)			
		case "ToAsstChgNo"
			.txtToAsstChgNo.value = arrRet(0)
		case "FrAsstChgNo"	
			.txtFrAsstChgNo.value = arrRet(0)									
		end select 
	End With

End Function

'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = frm1.txtFrChgDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToChgDt.Text
	'arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' �������� ���� Condition  
	
	' ���Ѱ��� �߰�
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID	
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(PopupParent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetDept(Byval arrRet)
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtFrChgDt.text = arrRet(4)
		frm1.txtToChgDt.text = arrRet(5)
End Function
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "RA") %>  
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>  ' check
	
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================	

Function OKClick()
	Dim lgPid
	Dim intColCnt		
	
	If frm1.vspdData.ActiveRow > 0 Then 				
		Redim arrReturn(1)
		frm1.vspdData.row	= frm1.vspdData.ActiveRow
		frm1.vspdData.Col	= C_AsstChgNo	
		arrReturn(0)		= frm1.vspdData.Text
	End if			
	
	Self.Returnvalue = arrReturn
	Self.Close()		
	
					
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================

Function CancelClick()
	Self.Close()			
End Function

'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================

Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function


'========================================= 2.6 `Sheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    frm1.vspdData.operationmode = 3
    
    Call InitSpreadPosVariables()
		With frm1.vspdData
		
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20021103",,Popupparent.gAllowDragDropSpread  
			.ReDraw = false	
			
			.MaxCols = C_AsstSoldDesc + 1                               '��: �ִ� Columns�� �׻� 1�� ������Ŵ
			ggoSpread.Source = frm1.vspdData
			ggospread.ClearSpreadData		'Buffer Clear
			.Col = .MaxCols											'������Ʈ�� ��� Hidden Column
			.ColHidden = True

			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetEdit		C_AsstChgNo,	"�Ű�����ȣ",		20,		0,		-1,		40,		2
			ggoSpread.SSSetEdit		C_AsstCd,		"�ڻ��ڵ�",		15,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_AsstNm,		"�ڻ��",		15,		0,		-1,		18,		2
			ggoSpread.SSSetDate	C_ChgDt,		"����"		  ,15, 2,Popupparent.gDateFormat		
			ggoSpread.SSSetCombo	C_ChgFgCd,		"����",		12,		2,		true
			ggoSpread.SSSetCombo	C_ChgFgNm,		"����",	12,		2,		false
			ggoSpread.SSSetEdit		C_FromDeptCd,	"�μ��ڵ�",		15,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_FromDeptNm,	"�μ���",		15,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_DocCur,		"��ȭ",		15,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_BpCdCd,		"�ŷ�ó��",		15,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_BpCdNm,		"�ŷ�ó��",		15,		0,		-1,		18,		2
			Call AppendNumberPlace("6","11","0")
			ggoSpread.SSSetFloat	C_ChgTotAmt,		"�Ѻ����ݾ�",			19, Popupparent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Popupparent.gComNum1000,Popupparent.gComNumDec
			ggoSpread.SSSetFloat	C_ChgTotLocAmt,		"�Ѻ����ݾ�(�ڱ�)",	19, Popupparent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Popupparent.gComNum1000,Popupparent.gComNumDec
			ggoSpread.SSSetEdit		C_AsstSoldDesc,	"����",				20,		0,		-1,		40,		2

			Call ggoSpread.SSSetColHidden(C_ChgFgCd, C_ChgFgCd, True)
			'Call ggoSpread.MakePairsColumn(C_AsstNo,C_AsstPopup)

			.ReDraw = true
		
		End With

    Call SetSpreadLock() 
End Sub


'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1      			
		.vspdData.ReDraw = True
    End With
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
			C_AsstChgNo		=iCurColumnPos(1)
			C_AsstCd		=iCurColumnPos(2)
			C_AsstNm		=iCurColumnPos(3)
			C_ChgDt			=iCurColumnPos(4)
			C_ChgFgCd		=iCurColumnPos(5)
			C_ChgFgNm		=iCurColumnPos(6)
			C_FromDeptCd	=iCurColumnPos(7)
			C_FromDeptNm	=iCurColumnPos(8)
			C_DocCur		=iCurColumnPos(9)
			C_BpCdCd		=iCurColumnPos(10)
			C_BpCdNm		=iCurColumnPos(11)
			C_ChgTotAmt		=iCurColumnPos(12)
			C_ChgTotLocAmt	=iCurColumnPos(13)
			C_AsstSoldDesc	=iCurColumnPos(14)
	End Select    
    
End Sub

 '**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 

 '-----------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------- 

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenSortPopup()

	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Popupparent.SORTW_WIDTH & "px; dialogHeight=" & Popupparent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 

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

    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,parent.PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,parent.PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
	Call InitSpreadSheet()
    'ReDim lgPopUpR(parent.C_MaxSelList - 1,1)
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitComboBox()
	frm1.txtFrChgDt.focus

'--------------- ������ coding part(�������,Start)----------------------------------------------------
   
'--------------- ������ coding part(�������,End)------------------------------------------------------
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


'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************

Sub InitData()
    Dim intRow
    Dim intIndex 
    ggoSpread.Source = frm1.vspdData	    	
    For intRow = 1 To frm1.vspdData.MaxRows
        frm1.vspdData.Row   = intRow
        frm1.vspdData.Col   = C_ChgFgCd        ' ���뱸��
        intIndex            = frm1.vspdData.value
        frm1.vspdData.col   = C_ChgFgNm
        frm1.vspdData.value = intindex
    Next 
End Sub

Function document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function

Sub ConditionKeypress()
	If window.event.keyCode = 13 Then
		Call Search_OnClick()
	End If
End sub
Sub txtDeptCd_onBlur()
	If frm1.txtDeptCd.value = "" Then
		frm1.txtDeptNm.value = ""
	End If
End sub



'==========================================================================================
'   Event Name : txtFrChgDt
'   Event Desc :
'==========================================================================================

Sub txtFrChgDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrChgDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtFrChgDt.Focus  
	End if
End Sub

'==========================================================================================
'   Event Name : txtToChgDt
'   Event Desc :
'==========================================================================================

Sub txtToChgDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToChgDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtToChgDt.Focus  
	End if
End Sub

Sub  txtFrChgDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub  

Sub  txtToChgDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub  

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ����
			DbQuery
		End If
   End if
    
End Sub



'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻�
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
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


Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

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

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function




Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
Function Radio1_onChange	
	on error resume next
	err.clear
		
	'ggoOper.SetReqAttr frm1.txtDueDt,		 "D"    '�̼��ݸ�������			
	ggoOper.SetReqAttr frm1.txtFrChgNo,		 "Q"    '�ڻ��ȣ
	ggoOper.SetReqAttr frm1.txtFrAsstChgNo,		 "D"    '�Ű�����ȣ����
	ggoOper.SetReqAttr frm1.txtToAsstChgNo,		 "D"    '�Ű�����ȣ����
	frm1.txtFrChgNo.value = ""

	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData
	frm1.vspdData.Col = C_AsstCd
	frm1.vspdData.ColHidden = True
	frm1.vspdData.Col = C_AsstNm
	frm1.vspdData.ColHidden = True

End Function

Function Radio2_onChange
	Dim lDelRows,intRow
	Dim bMidChgVal
	

	ggoOper.SetReqAttr frm1.txtFrAsstChgNo,		 "Q"    '�Ű�����ȣ����
	ggoOper.SetReqAttr frm1.txtToAsstChgNo,		 "Q"    '�Ű�����ȣ����
	ggoOper.SetReqAttr frm1.txtFrChgNo,		 "D"    '�ڻ��ȣ
	frm1.txtToAsstChgNo.value = ""
	frm1.txtFrAsstChgNo.value = ""

	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData
	frm1.vspdData.Col = C_AsstCd
	frm1.vspdData.ColHidden = False
	frm1.vspdData.Col = C_AsstNm
	frm1.vspdData.ColHidden = False
	

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
Function FncQuery() 
	Dim IntRetCD
    Dim strFrChgDt
    Dim strToChgDt
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing   

    '-----------------------
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
	
	'---------------------------------------
	'�������� ���� Check
	'---------------------------------------
	
	strFrChgDt = UniConvDateToYYYYMMDD(frm1.txtFrChgDt.Text, PopupParent.gDateFormat,"") 
	strToChgDt = UniConvDateToYYYYMMDD(frm1.txtToChgDt.Text, PopupParent.gDateFormat,"")

	IF strToChgDt <> "" Then
		If strFrChgDt > strToChgDt Then
			Call DisplayMsgBox("970025", "X", frm1.txtFrChgDt.Alt, frm1.txtToChgDt.Alt)
			frm1.txtFrChgDt.focus
			Exit Function
		End If
	End If
	
	
	'---------------------------------------
	'�ڻ�����ȣ ���� Check
	'---------------------------------------
	If frm1.Rb_Sold.Checked =True then
		frm1.txtFrAsstChgNo.value = Trim(frm1.txtFrAsstChgNo.value)
		frm1.txtToAsstChgNo.value = Trim(frm1.txtToAsstChgNo.value)
		
		If frm1.txtFrAsstChgNo.value <> "" And frm1.txtToAsstChgNo.value <> "" Then
			If frm1.txtFrAsstChgNo.value > frm1.txtToAsstChgNo.value Then
				Call DisplayMsgBox("970025", "X", frm1.txtFrAsstChgNo.Alt, frm1.txtToAsstChgNo.Alt)
				frm1.txtFrAsstChgNo.focus 
				Exit Function
			End If
		End If
	End If
	
    '-----------------------
    'Query function call area
    '-----------------------
	'frm1.vspdData.MaxRows = 0                                                   '��: Protect system from crashing
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call DbQuery															'��: Query db data

    FncQuery = True		
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
    Call InitComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
    Call SetSpreadLock ()
End Sub

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
	Call parent.FncExport(PopupParent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(PopupParent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()

	Dim IntRetCD

	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO,,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True

End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  ******************************
'	���� : 
'********************************************************************************************************* 
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	Dim strVal
	Dim strChgFg
	Dim lgPid
	
	lgPid = "<%=Request("PID")%>"

    DbQuery = False
    
    Err.Clear            
    
	Call LayerShowHide(1)

    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtFrChgDt=" & Trim(.txtFrChgDt.Text)
		strVal = strVal & "&txtToChgDt=" & Trim(.txtToChgDt.Text)
		strVal = strVal & "&txtFrAsstChgNo=" & Trim(.txtFrAsstChgNo.value)
		strVal = strVal & "&txtToAsstChgNo=" & Trim(.txtToAsstChgNo.value)
		strVal = strVal & "&txtFrChgNo=" & Trim(.txtFrChgNo.value)
		strVal = strVal & "&txtDeptCd="   & Trim(.txtDeptCd.value)
		strVal = strVal & "&txtAsstChgDesc=" & Trim(.txtAsstChgDesc.value)
'--------------- ������ coding part(�������,End)------------------------------------------------
		If frm1.Rb_Sold.Checked =True then

			strVal = strVal & "&txtGubun=" & "A"
		Else
			strVal = strVal & "&txtGubun=" & "B"
		End If
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
        strVal = strVal & "&lgMaxCount="     & CStr(frm1.vspdData.MaxRows)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�
		strVal = strVal & "&PID="            & lgPid

		' ���Ѱ��� �߰�
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' �����
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ�
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ����

        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ����
    End With
    
    DbQuery = True

End Function

Function DbQueryOk()														'��: ��ȸ ������ �������
	Call InitData()
    lgBlnFlgChgValue = True                                                 'Indicates that no value changed
End Function


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ�
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag��
'#########################################################################################################  -->

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>	
						<TD CLASS="TD5" NOWRAP>��ȸ���</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_Sold Checked tag = 1 value="03" onclick=radio1_onchange()><LABEL FOR=Rb_Sold>�Ű�����ȣ</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_Duse tag = 1 value="04" onclick=radio2_onchange()><LABEL FOR=Rb_Duse>�ڻ�</LABEL></TD>										        							
						<TD CLASS="TD5" NOWRAP>����</TD>
						<TD CLASS="TD6" NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtFrChgDt CLASSID=<%=gCLSIDFPDT%> ALT="���ۺ�������" tag="11"> </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtToChgDt CLASSID=<%=gCLSIDFPDT%> ALT="���ắ������" tag="11"> </OBJECT>');</SCRIPT>
						</TD>
					</TR>			
					<TR>				
						<TD CLASS=TD5 NOWRAP>�Ű�����ȣ</TD>				
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtFrAsstChgNo" SIZE=15 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="���۸Ű�����ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFrAsstChgNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('FrAsstChgNo',frm1.txtFrAsstChgNo.Value)">&nbsp;~&nbsp;
							<INPUT TYPE="Text" NAME="txtToAsstChgNo" SIZE=15 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="����Ű�����ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToAsstChgNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('ToAsstChgNo',frm1.txtToAsstChgNo.Value)">
						</TD>
						<TD CLASS=TD5 NOWRAP>�ڻ��ȣ</TD>				
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtFrChgNo" SIZE=18 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="�ڻ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFrChgNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('FrChgNo',frm1.txtFrChgNo.Value)">
						</TD>						
					</TR>			
					<TR>					
						<TD CLASS=TD5 NOWRAP>�μ��ڵ�</TD>
						<TD CLASS=TD6 NOWRAP>
						<INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" MAXLENGTH="10" SIZE=15 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">&nbsp;
						<INPUT NAME="txtDeptNm" ALT="�μ���" MAXLENGTH="10" SIZE=18 STYLE="TEXT-ALIGN: left" tag="14X"></TD>		
						<TD CLASS=TD5 NOWRAP>����</TD>				
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE="Text" NAME="txtAsstChgDesc" SIZE=18 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="�����ڻ��ȣ">
						</TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% tag="2" HEIGHT=100% id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>
					</TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
									 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" src="../../blank.htm"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId"    tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

