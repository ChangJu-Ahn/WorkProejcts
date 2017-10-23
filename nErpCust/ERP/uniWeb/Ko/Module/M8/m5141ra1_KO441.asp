<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M5141RA1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open Po Ref Popup ASP														*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2001/05/08																*
'*                            2002/04/30
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : Min, HJ															*	
'*                            Kim Jae Soon
'* 11. Comment              :																			*
'* 12. Common Coding Guide  :																			*
'* 13. History              :																			*
'********************************************************************************************************
Response.Expires = -1													'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
%>
<HTML>
<HEAD>
<!--<TITLE>���ſ�û����</TITLE>-->
<TITLE></TITLE>
<%
'########################################################################################################
'#						1. �� �� ��																		#
'########################################################################################################
%>
<%
'********************************************  1.1 Inc ����  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!-- #Include file="../../inc/IncSvrVariables.inc" -->
<%
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>
<%
'============================================  1.1.2 ���� Include  ======================================
'========================================================================================================
%>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBS">

Option Explicit					<% '��: indicates that All variables must be declared in advance %>
	
	
<%
'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
%>

<%
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
%>

	
    Const BIZ_PGM_ID 		= "m5141rb1_KO441.asp"                              '��: Biz Logic ASP Name
    
    '��� �׸��� 
    Const C_IvType			=	1
    Const C_IvTypeNm		=	2
    Const C_BuildCd			=	3
    Const C_BuildNm			=	4
    Const C_PayeeCd			=	5
    Const C_PayeeNm			=	6
    Const C_SupplCd			=	7
    Const C_SupplNm			=	8
    Const C_GrpCd			=	9
    Const C_GrpNm			=	10
    Const C_BizAreaCd		=	11
    Const C_BizAreaNm		=	12
    Const C_Curr			=	13
    Const C_VatCd			=	14
    Const C_VatNm			=	15
    Const C_VatRt			=	16
    Const C_PayTermCd		=	17
    Const C_PayTermNm		=	18
    
    Const C_SpplRegNo		=	19
    Const C_SpplInvNo		=	20
    Const C_PayDur			=	21
    Const C_PayTypeCd		=	22
    Const C_PayTypeNm		=	23
    Const C_PayTermsTxt		=	24
    Const C_Remark			=	25
    
     
<%
'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
%>
	Const C_SHEETMAXROWS_D  = 100   
	                                       '��: Fetch max count at once
	Const C_MaxKey_1        = 25                                           '��: key count of SpreadSheet
	'�̼��� ���� 
	Const C_MaxKey_2		= 28
	                                    '��: key count of SpreadSheet
	Const ivType = "ST"
<%
'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
%>
<!-- #Include file="../../inc/lgvariables.inc" -->	
<%
'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================
%>


Dim lgStrPrevKey_1			'�ι�° �׸��忡�� ���Ǵ� ���� 
Dim lgPageNo_1				'�ι�° �׸��忡�� ���Ǵ� ���� 
		
Dim lgSelectList                                            '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 

Dim lgSortFieldNm                                           '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD                                           '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgPopUpR                                                '��: Orderby default ��                    

Dim lgKeyPos                                                '��: Key��ġ                               
Dim lgKeyPosVal                                             '��: Key��ġ Value                         
Dim IscookieSplit 

Dim IsOpenPop  
Dim lblnWinEvent											'��: ShowModal Dialog(PopUp) 
														    'Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
														    'PopUp Window�� ��������� ���θ� ��Ÿ�� 
Dim arrReturn												'��: Return Parameter Group
Dim arrParam
Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"


'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

<%
'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
%>
<% 
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************
%>
<%
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
%>
Function InitVariables()
		lgStrPrevKey     = ""								   'initializes Previous Key
		lgPageNo         = ""
		
		lgStrPrevKey_1     = ""								   'initializes Previous Key
		lgPageNo_1         = ""
        
        lgBlnFlgChgValue = False	                           'Indicates that no value changed
        
        lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
        frm1.vspdData2.OperationMode  = 5
        frm1.vspdData1.OperationMode = 3
        
        lgSortKey        = 1   
        
        lgIntGrpCount = 0										<%'��: Initializes Group View Size%>

        lblnWinEvent = False
       
        Redim arrReturn(0,0)        
        Self.Returnvalue = arrReturn     
End Function

<%'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 

		<% Call loadInfTB19029A("Q", "*", "NOCOOKIE", "RA") %>                                '��: 
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","RA") %>

		'------ Developer Coding part (End )   -------------------------------------------------------------- 
	End Sub

<%
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
%>
	Sub InitSpreadSheet()
	
		Call SetZAdoSpreadSheet("M5141RA1_1","S","A","V20050603",PopupParent.C_SORT_DBAGENT,frm1.vspdData1, _
									C_MaxKey_1, "X","X")
		Call SetZAdoSpreadSheet("M5141RA1_2","S","B","V20050603",PopupParent.C_SORT_DBAGENT,frm1.vspdData2, _
									C_MAXKEY_2 , "X","X")
		
		Call SetSpreadLock("A") 
		Call SetSpreadLock("B")

	End Sub



<%
'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
%>
	Sub SetSpreadLock(ByVal pOpt)
		If pOpt = "A" Then
			With frm1
				.vspdData1.ReDraw = False
				ggoSpread.Source = .vspdData1
				ggoSpread.SpreadLock 1, -1
				.vspdData1.ReDraw = True
			End With
		Else
			With frm1
				.vspdData2.ReDraw = False
				ggoSpread.Source = .vspdData2
				ggoSpread.SpreadLock 1, -1
				.vspdData2.ReDraw = True
			End With
		End If			

'		ggoSpread.Source = frm1.vspdData1
'  	    ggoSpread.SpreadLockWithOddEvenRowColor()

	End Sub	

<%
'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
%>	
	
	Function OKClick()
	
		Dim intColCnt, intRowCnt, intInsRow
		
		with frm1
		If .vspdData2.SelModeSelCount > 0 Then 
			
			intInsRow = 0
			'Redim arrReturn(frm1.vspdData2.SelModeSelCount-1, frm1.vspdData2.MaxCols-2)
			Redim arrReturn(frm1.vspdData2.SelModeSelCount, frm1.vspdData1.MaxCols+1)

			For intRowCnt = 1 To frm1.vspdData2.MaxRows

				frm1.vspdData2.Row = intRowCnt
				
				If frm1.vspdData2.SelModeSelected Then
					For intColCnt = 0 To frm1.vspdData2.MaxCols - 2						
						frm1.vspdData2.Col = GetKeyPos("B",intColCnt+1)	
						arrReturn(intInsRow, intColCnt) = frm1.vspdData2.Text
					Next
										
					intInsRow = intInsRow + 1
				End IF								
			Next
	
		arrReturn(intInsRow, 0) = frm1.hdnIvTypeCd1.value
		arrReturn(intInsRow, 1) = frm1.hdnSpplCd1.value
		arrReturn(intInsRow, 2) = frm1.hdnBuildCd1.value
		arrReturn(intInsRow, 3) = frm1.hdnPayeeCd1.value			
		arrReturn(intInsRow, 4) = frm1.hdnCurr1.value		
		arrReturn(intInsRow, 5) = frm1.hdnVatCd1.value		
		arrReturn(intInsRow, 6) = frm1.hdnGrpCd1.value
		arrReturn(intInsRow, 7) = frm1.hdnBizAreaCd1.value		
		
		arrReturn(intInsRow, 8) = frm1.hdnIvTypeNm1.value		
		arrReturn(intInsRow, 9) = frm1.hdnSpplNm1.value		
		arrReturn(intInsRow,10) = frm1.hdnBuildNm1.value		
		arrReturn(intInsRow,11) = frm1.hdnPayeeNm1.value		
		arrReturn(intInsRow,12) = frm1.hdnVatNm1.value		
		arrReturn(intInsRow,13) = frm1.hdnGrpNm1.value		
		arrReturn(intInsRow,14) = frm1.hdnBizAreaNm1.value		
		arrReturn(intInsRow,15) = frm1.hdnVatRt.value	
		arrReturn(intInsRow,16) = frm1.hdnPayTermCd1.value	
		arrReturn(intInsRow,17) = frm1.hdnPayTermNm1.value
		
		arrReturn(intInsRow,18) = frm1.hdnSpplRegNo1.value	
		arrReturn(intInsRow,19) = frm1.hdnSpplInvNo1.value	
		arrReturn(intInsRow,20) = frm1.hdnPayDur1.value	
		arrReturn(intInsRow,21) = frm1.hdnPayTypeCd1.value	
		arrReturn(intInsRow,22) = frm1.hdnPayTypeNm1.value	
		arrReturn(intInsRow,23) = frm1.hdnPayTermsTxt1.value	
		arrReturn(intInsRow,24) = frm1.hdnRemark1.value		
		
		End if		
		
		end with
		
		Self.Returnvalue = arrReturn
		Self.Close()
		
	End Function
<%
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
%>
	Function CancelClick()
		Self.Close()
	End Function
<%
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
%>
	Function MousePointer(pstr1)
	      Select case UCase(pstr1)
	            case "PON"
					window.document.search.style.cursor = "wait"
	            case "POFF"
					window.document.search.style.cursor = ""
	      End Select
	End Function
	
<% 
'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
%>

'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
'------------------------------------------  OpenIvType()  -------------------------------------------------
Function OpenIvType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If lblnWinEvent = True Then Exit Function

	lblnWinEvent = True
	
	arrHeader(0) = "��������"						' Header��(0)
    arrHeader(1) = "�������¸�"						' Header��(1)
    
    arrField(0) = "IV_TYPE_CD"							' Field��(0)
    arrField(1) = "IV_TYPE_NM"							' Field��(1)
    
	arrParam(0) = "��������"						' �˾� ��Ī 
	arrParam(1) = "M_IV_TYPE"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			' Name Cindition
	arrParam(4) = "import_flg=" & FilterVar("N", "''", "S") & " "						' Where Condition
	arrParam(5) = "��������"						' TextBox ��Ī 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			
    lblnWinEvent = False
    
    If arrRet(0) = "" Then Exit Function
    frm1.txtIvTypeCd.focus
    frm1.txtIvTypeCd.Value= arrRet(0)		
	frm1.txtIvTypeNm.Value= arrRet(1)	
	Set gActiveElement = document.activeElement
End Function
'------------------------------------------  OpenGrp()  -------------------------------------------------
'	Name : OpenGrp()
'	Description : 
'--------------------------------------------------------------------------------------------------------- %>
Function OpenGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lblnWinEvent = True Or UCase(frm1.txtGrpCd.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function

	lblnWinEvent = True

	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtGrpNm.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGroup(arrRet)
	End If	

End Function 


Function SetGroup(byval arrRet)
	frm1.txtGrpCd.Value= arrRet(0)		
	frm1.txtGrpNm.Value= arrRet(1)	
	'frm1.txtGrpCd.focus	
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenVat()  -------------------------------------------------
Function OpenVat()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If lblnWinEvent = True Then Exit Function    
	lblnWinEvent = True
	
	arrHeader(0) = "VAT����"									' Header��(0)
    arrHeader(1) = "VAT���¸�"									' Header��(1)
    arrHeader(2) = "VAT��"									    ' Header��(2)
    
    arrField(0) = "b_minor.MINOR_CD"					            ' Field��(0)
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"					    ' Field��(1)
    
	arrParam(0) = "VAT"	            							' �˾� ��Ī 
	arrParam(1) = "B_MINOR,b_configuration"
	arrParam(2) = Trim(frm1.txtVatCd.Value)						    ' Code Condition
	'arrParam(2) = Trim(frm1.txtVatCd.Value)						    ' Code Condition
	'arrParam(3) = Trim(frm1.txtVatNm.Value)						' Name Cindition
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_configuration.seq_no=1 and b_minor.major_cd=b_configuration.major_cd"
	arrParam(5) = "VAT"										    ' TextBox ��Ī 
	
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtVatCd.Value = arrRet(0)
		frm1.txtVatNm.Value = arrRet(1)		
	End If	
	
    frm1.txtVatCd.focus
    Set gActiveElement = document.activeElement
    lblnWinEvent = False
End Function

'------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenSppl()
'	Description :����ó,���ݰ�꼭����ó,����ó 
'---------------------------------------------------------------------------------------------------------
Function OpenSppl(Byval BpType)
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	arrHeader(2) = "����ڵ�Ϲ�ȣ"									' Header��(2)
    arrField(0) = "B_BIZ_PARTNER.BP_Cd"									' Field��(0)
    arrField(1) = "B_BIZ_PARTNER.BP_Nm"								    ' Field��(1)
    arrField(2) = "B_BIZ_PARTNER.BP_RGST_NO"							' Field��(2)
    
	Select Case BpType
		Case "1"  '����ó 
			If lblnWinEvent = True Then Exit Function    
			lblnWinEvent = True
			arrHeader(0) = "����ó"											' Header��(0)
			arrHeader(1) = "����ó��"										' Header��(1)

		    arrParam(0) = "����ó"											' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER "					                    ' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtSpplCd.Value)		
			'arrParam(2) = Trim(frm1.txtSpplCd.Value)							' Code Condition
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "		' Where Condition
			arrParam(5) = "����ó"											' TextBox ��Ī 
		Case "3"   '���ݰ�꼭����ó 
			If lblnWinEvent = True or Trim(frm1.txtSpplCd.Value) = "" Then 
				Call DisplayMsgBox("17a003","X","����ó","X")
				frm1.txtSpplCd.focus				
				Exit Function
			End If
			    
			lblnWinEvent = True

			arrHeader(0) = "���ݰ�꼭����ó"											' Header��(0)
			arrHeader(1) = "���ݰ�꼭����ó��" 										' Header��(1)

			arrParam(0) = "���ݰ�꼭����ó"											' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER, B_BIZ_PARTNER_FTN	"           					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtBuildCd.Value)					            		' Code Condition
			'arrParam(2) = Trim(frm1.txtBuildCd.Value)					            		' Code Condition
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER.BP_CD = B_BIZ_PARTNER_FTN.PARTNER_BP_CD  AND B_BIZ_PARTNER_FTN.BP_CD = " 				<%' Where Condition%>
            arrParam(4) = arrParam(4) & FilterVar(Trim(frm1.txtSpplCd.Value), "''", "S") & " AND  B_BIZ_PARTNER_FTN.PARTNER_FTN = " & FilterVar("MBI", "''", "S") & " "
			arrParam(5) = "���ݰ�꼭����ó"											' TextBox ��Ī 
	End Select

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Select Case BpType
			Case "1"   '����ó 
				frm1.txtSpplCd.Value = arrRet(0) : frm1.txtSpplNm.Value = arrRet(1)
				frm1.txtSpplCd.focus
			Case "3"   '���ݰ�꼭����ó 
				frm1.txtBuildCd.Value = arrRet(0) : frm1.txtBuildNm.Value = arrRet(1) ': frm1.txtSpplRegNo.Value = arrRet(2)				
		        frm1.txtBuildCd.focus
		End Select 	
		
	End If	
			
    lblnWinEvent = False
    
    Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------

Function OpenPoNo()
	
	Dim strRet
	Dim lblnWinEvent
	Dim iCalledAspName
	Dim arrParam(2)
	
	If lblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag
	
	iCalledAspName = AskPRAspName("m3111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "m3111pa1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus	
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function

'------------------------------------------  OpenIvNo()  -------------------------------------------------
Function OpenIvNo()
	
	Dim strRet
	Dim arrParam(0)
	Dim iCalledAspName
	
		If lblnWinEvent = True Then Exit Function

		lblnWinEvent = True

		arrParam(0) = ivType
		
		iCalledAspName = AskPRAspName("m5111pa1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "m5111pa1", "X")
			lgIsOpenPop = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		lblnWinEvent = False
	
		If strRet(0) = "" Then
			frm1.txtIvNo.focus
			Exit Function
		Else
			frm1.txtIvNo.value = strRet(0)
			frm1.txtIvNo.focus
		End If	
		Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet,lgIsOpenPop
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"	
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else	
		frm1.txtPlantCd.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function	


'===========================================================================
' Function Name : OpenSoNo
' Function Desc : OpenSoNo Reference Popup
'===========================================================================
 Function OpenSoNo()

	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True
		
'	strRet = window.showModalDialog("../s3/s3111pa1.asp", "", _
'		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	iCalledAspName = AskPRAspName("S3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtSoNo.value = strRet
	End If	

End Function
<%
'===========================================================================
' Function Name : OpenTrackingNo
' Function Desc : OpenTrackingNo Reference Popup
'===========================================================================
%>

Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
    arrParam(2) = ""	'���� 
    arrParam(3) = ""	'��ǰ�� 
    arrParam(4) = ""	'���ֹ�ȣ 
    arrParam(5) = ""	'�߰� Where�� 
    
'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		lgBlnFlgChgValue = True
	End If	

End Function
 
<%
'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
%>
'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderBy()
	Dim arrRet
	
	On Error Resume Next
	
	'If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("B"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("B",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


<% '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ %>
<% '------------------------------------------  SetSorgCode()  --------------------------------------------------
'	Name : SetBPCd()
'	Description : SetSorgCode Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- %>

<%
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>

<%
'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
%>
<%
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
%>
<%
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
%>
Sub Form_Load()

    Call LoadInfTB19029													'��: Load table , B_numeric_format
'    ReDim lgPopUpR(C_MaxSelList - 1,1)
    
	'Call GetAdoFieldInf("M2111RA1_1","S","A")			              '��: spread sheet �ʵ����� query
	'
                                                                  ' 1. Program id
                                                                  ' 2. G is for Qroup , S is for Sort     
                                                                  ' 3. Spreadsheet no     
    'Html���� tag ���ڰ� 1�� 2�� �����ϴ� �κ� ����Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
    
'    Call MakePopData(gDefaultT,gFieldNM,gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,C_MaxSelList)    ' You must not this line    
    Call InitVariables											  '��: Initializes local global variables
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

Sub SetDefaultVal()
		Dim arrParam
		
		arrParam = arrParent(1)
		
		frm1.vspdData1.OperationMode = 3 
		frm1.vspdData2.OperationMode = 5
		
		frm1.txtIvTypeCd.value = arrParam(0)
		frm1.txtIvTypeNm.value = arrParam(1)
		frm1.txtGrpCd.value = arrParam(2)
		frm1.txtGrpNm.value = arrParam(3)
		frm1.txtVatCd.value = arrParam(4)
		frm1.txtVatNm.value = arrParam(5)
		frm1.txtSpplCd.value = arrParam(6)
		frm1.txtSpplNm.value = arrParam(7)
		frm1.txtBuildCd.value = arrParam(8)
		frm1.txtBuildNm.value = arrParam(9)
		frm1.txtIvNo.value = arrParam(10)
		frm1.txtCur.value = arrParam(11)
		
		
		If arrParam(2) = "" then
			frm1.txtGrpCd.value = PopupParent.gPurGrp
		End if

		If arrParam(0) <> "" then		'2002-12-04(LJT)
			ggoOper.SetReqAttr		frm1.txtIvTypeCd, "Q"
			ggoOper.SetReqAttr		frm1.txtIvTypeNm, "Q"
		End if
		
		if  arrParam(2) <> "" then
			ggoOper.SetReqAttr		frm1.txtGrpCd, "Q"
			ggoOper.SetReqAttr		frm1.txtGrpNm, "Q"
		End if

		if  arrParam(4) <> "" then
			ggoOper.SetReqAttr		frm1.txtVatCd, "Q"
			ggoOper.SetReqAttr		frm1.txtVatNm, "Q"
		End if

		if  arrParam(6) <> "" then
			ggoOper.SetReqAttr		frm1.txtSpplCd, "Q"
			ggoOper.SetReqAttr		frm1.txtSpplNm, "Q"
		End if

		if  arrParam(8) <> "" then
			ggoOper.SetReqAttr		frm1.txtBuildCd, "Q"
			ggoOper.SetReqAttr		frm1.txtBuildNm, "Q"
		End if

		
		frm1.txtFrIvDt.text 	= EndDate
		frm1.txtToIvDt.text 	= UnIDateAdd("m", +1, EndDate, PopupParent.gDateFormat)
		
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGrpCd, "Q") 
		frm1.txtGrpCd.Tag = left(frm1.txtGrpCd.Tag,1) & "4" & mid(frm1.txtGrpCd.Tag,3,len(frm1.txtGrpCd.Tag))
        frm1.txtGrpCd.value = lgPGCd
	End If
End Sub

<%
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
%>
	Sub Form_QueryUnload(Cancel, UnloadMode)
	   
	End Sub
<%
'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************
%>



<%
'==========================================================================================
'   Event Name : OCX_Keypress()
'   Event Desc : 
'==========================================================================================
%>
	Sub txtFrIvDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtToIvDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub
<%
'==========================================================================================
'   Event Name : txtFrIvDt
'   Event Desc :
'==========================================================================================
%>
Sub txtFrIvDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrIvDt.Action = 7
	End if
End Sub

<%
'==========================================================================================
'   Event Name : txtToIvDt
'   Event Desc :
'==========================================================================================
%>
Sub txtToIvDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToIvDt.Action = 7
	End if
End Sub


<%
'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
%>
	Function vspdData2_DblClick(ByVal Col, ByVal Row)
	
	 If Row = 0 Or Frm1.vspdData2.MaxRows = 0 Then 
          Exit Function
     End If
	With frm1.vspdData2 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
	End Function
'========================================================================================
' Function Name : vspdData1_Click
' Function Desc : 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	
	ggoSpread.Source = frm1.vspdData1
	gMouseClickStatus = "SPC"   
	
	'�̼��� �߰� 
	lgIntFlgMode = PopupParent.OPMD_CMODE
	
	frm1.vspdData2.MaxRows = 0
	
	Set gActiveSpdSheet = frm1.vspdData1
	Call SetPopupMenuItemInf("0000111111")

	If frm1.vspdData1.MaxRows <= 0 Then Exit Sub

	If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
	Else
 		'------ Developer Coding part (Start)
 		
		Call SetHiddenArea(row)		
		
		
		'�̼��� �߰� 
		lgPageNo = ""
			
		If DbQuery2() = False Then
			'	Call ResetToolBar(lgOldRow)
				Exit Sub 
		End If	
	 	'------ Developer Coding part (End)
 	End If
End Sub

Function SetHiddenArea(Byval row)
	Dim strIvType, strBuildCd , strPayeeCd , strSupplCd , strGrpCd
	Dim strBizAreaCd , strCurr , strVatCd , strPayTermCd
		
		frm1.vspdData1.row	= row

		frm1.vspdData1.col = C_IvType
		strIvType = frm1.vspdData1.value
		
		frm1.vspdData1.col = C_PayeeCd 
		strPayeeCd = frm1.vspdData1.value
		
		frm1.vspdData1.col = C_BuildCd 
		strBuildCd = frm1.vspdData1.value
						
		frm1.vspdData1.col = C_SupplCd
		strSupplCd = frm1.vspdData1.value	
			
		frm1.vspdData1.col = C_GrpCd 
		strGrpCd = frm1.vspdData1.value
		
		frm1.vspdData1.col = C_BizAreaCd
		strBizAreaCd = frm1.vspdData1.value
		
		frm1.vspdData1.col = C_Curr
		strCurr = frm1.vspdData1.value		
				
		frm1.vspdData1.col = C_VatCd
		strVatCd = frm1.vspdData1.value
		
		frm1.vspdData1.col = C_PayTermCd
		strPayTermCd = frm1.vspdData1.value	
		
		frm1.hdnFrDt1.value		= frm1.hdnFrDt.value
		frm1.hdnToDt1.value		= frm1.hdnToDt.value
		frm1.hdnIvTypeCd1.value = Trim(strIvType)
		frm1.hdnGrpCd1.value	= Trim(strGrpCd)
		frm1.hdnVatCd1.value	= Trim(strVatCd)
		frm1.hdnSpplCd1.value	= Trim(strSupplCd)
		frm1.hdnBuildCd1.value	= Trim(strBuildCd)
		frm1.hdnPoNo1.value		= frm1.hdnPoNo.value
		frm1.hdnIvNo1.value		= frm1.hdnIvNo.value
		frm1.hdnPayeeCd1.value	= Trim(strPayeeCd)
		frm1.hdnCurr1.value		= Trim(strCurr)
		frm1.hdnPayTermCd1.value	= Trim(strPayTermCd)
		
		frm1.vspdData1.col	=	C_BizAreaCd
		frm1.hdnBizAreaCd1.value = frm1.vspdData1.value
		
		frm1.vspdData1.col	=	C_IvTypeNm
		frm1.hdnIvTypeNm1.value = frm1.vspdData1.value
		
		frm1.vspdData1.col	=	C_BuildNm
		frm1.hdnBuildNm1.value = frm1.vspdData1.value
		
		frm1.vspdData1.col	=	C_PayeeNm
		frm1.hdnPayeeNm1.value = frm1.vspdData1.value
		
		frm1.vspdData1.col	=	C_SupplNm
		frm1.hdnSpplNm1.value = frm1.vspdData1.value
		
		frm1.vspdData1.col	=	C_GrpNm
		frm1.hdnGrpNm1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_BizAreaNm
		frm1.hdnBizAreaNm1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_VatNm
		frm1.hdnVatNm1.value = frm1.vspdData1.value	
		
		frm1.vspdData1.col	=	C_VatRt
		frm1.hdnVatRt.value = frm1.vspdData1.value	

		frm1.vspdData1.col	=	C_VatRt
		frm1.hdnVatRt.value = frm1.vspdData1.value
		
		frm1.vspdData1.col	=	C_PayTermNm
		frm1.hdnPayTermNm1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_SpplRegNo
		frm1.hdnSpplRegNo1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_SpplInvNo
		frm1.hdnSpplInvNo1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_PayDur
		frm1.hdnPayDur1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_PayTypeCd
		frm1.hdnPayTypeCd1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_PayTypeNm
		frm1.hdnPayTypeNm1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_PayTermsTxt
		frm1.hdnPayTermsTxt1.value = frm1.vspdData1.value		
		
		frm1.vspdData1.col	=	C_Remark
		frm1.hdnRemark1.value = frm1.vspdData1.value		
		
		
				

End Function


Function vspdData2_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData2.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
	Sub vspdData1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If		
		
		lgIntFlgMode = PopupParent.OPMD_UMODE
		
		If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	    '��: ������ üũ	
			'If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			'�̼��� 
			If lgPageNo_1 <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbQuery = False Then
					Exit Sub
				End if
			End If
		End If		 
	End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
	Sub vspdData2_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
		Dim lRow, i, strIvType, strBuildCd , strPayeeCd , strSupplCd , strGrpCd
		Dim strBizAreaCd , strCurr , strVatCd , strPayTermCd


		If OldLeft <> NewLeft Then
		    Exit Sub
		End If		
				
		lgIntFlgMode = PopupParent.OPMD_UMODE
		
		If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	    '��: ������ üũ	
			If lgPageNo <> "" Then                '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbQuery2() = False Then
					Exit Sub
				End if
			End If
		End If		 
	End Sub
<% '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'######################################################################################################### %>

<% '#########################################################################################################
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
'######################################################################################################### %>

<% '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* %>
<%
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
%>
Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    '�̼��� �߰� 
    lgIntFlgMode = PopupParent.OPMD_CMODE
    Err.Clear                                                               '��: Protect system from crashing
	
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� ũ�ų� ���ƾ� �Ҷ� **
	If ValidDateCheck(frm1.txtFrIvDt, frm1.txtToIvDt) = False Then Exit Function
   
    '-----------------------
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
    Call InitVariables
    
    
    ggoSpread.Source = frm1.vspdData2	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function


<%
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
%>
Function DbQuery() 

	Err.Clear														'��: Protect system from crashing
	DbQuery = False													'��: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ����	
			strVal = strVal & "&txtFrIvDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToIvDt=" & .hdnToDt.value
			strVal = strVal & "&txtIvTypeCd=" & .hdnIvTypeCd.value
			strVal = strVal & "&txtGrpCd=" & .hdnGrpCd.value		
			strVal = strVal & "&txtVatCd=" & .hdnVatCd.value
			strVal = strVal & "&txtSpplCd=" & .hdnSpplCd.value
			strVal = strVal & "&txtBuildCd=" & .hdnBuildCd.value 
			strVal = strVal & "&txtPoNo=" & .hdnPoNo.value
			strVal = strVal & "&txtIvNo=" & .hdnIvNo.value		
			strVal = strVal & "&txtCur=" & .txtCur.value		
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey   
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtFrIvDt=" & Trim(.txtFrIvDt.text)
			strVal = strVal & "&txtToIvDt=" & Trim(.txtToIvDt.text)
			strVal = strVal & "&txtIvTypeCd=" & Trim(.txtIvTypeCd.value)
			strVal = strVal & "&txtGrpCd=" & Trim(.txtGrpCd.value)
			strVal = strVal & "&txtVatCd=" & Trim(.txtVatCd.value)
			strVal = strVal & "&txtSpplCd=" & Trim(.txtSpplCd.value)
			strVal = strVal & "&txtBuildCd=" & Trim(.txtBuildCd.value )
			strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
			strVal = strVal & "&txtIvNo=" & Trim(.txtIvNo.value)
			strVal = strVal & "&txtCur=" & .txtCur.value
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If				
	    strVal = strVal & "&lgPageNo="		 & lgPageNo_1						'��: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�  
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&txtGridNum="	 & "A"

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

		Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
		
    End With
    
    DbQuery = True    

End Function

<%
'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=========================================================================================================
%>
Function DbQueryOk()	    												'��: ��ȸ ������ ������� 
	Dim lRow, i
		
	'lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Focus
		frm1.vspdData1.Row = 1	
		
		
	

		Call SetHiddenArea(1)	
									
		If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			Call DbQuery2()
			lgIntFlgMode = PopupParent.OPMD_UMODE
		End If
		
		frm1.vspdData1.SelModeSelected = True		
	Else
	'	frm1.txtDnType.focus
	End If
	
	call SetSpreadLock("A")

End Function
'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2()
	Err.Clear														'��: Protect system from crashing
	DbQuery2 = False													'��: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
	'frm1.vspdData2.MaxRows = 0

    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ����	
			strVal = strVal & "&txtFrIvDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToIvDt=" & .hdnToDt.value
			strVal = strVal & "&txtIvTypeCd=" & .hdnIvTypeCd1.value
			strVal = strVal & "&txtGrpCd=" & .hdnGrpCd1.value		
			strVal = strVal & "&txtVatCd=" & .hdnVatCd1.value
			strVal = strVal & "&txtSpplCd=" & .hdnSpplCd1.value
			strVal = strVal & "&txtBuildCd=" & .hdnBuildCd1.value 
			strVal = strVal & "&txtPoNo=" & .hdnPoNo.value
			strVal = strVal & "&txtIvNo=" & .hdnIvNo.value
			
			strVal = strVal & "&txtPayeeCd=" & .hdnPayeeCd1.value
			strVal = strVal & "&txtCurr=" & .hdnCurr1.value
			strVal = strVal & "&txtPayTermCd=" & .hdnPayTermCd1.value
			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey   
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtFrIvDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToIvDt=" & .hdnToDt.value
			strVal = strVal & "&txtIvTypeCd=" & .hdnIvTypeCd1.value
			strVal = strVal & "&txtGrpCd=" & .hdnGrpCd1.value		
			strVal = strVal & "&txtVatCd=" & .hdnVatCd1.value
			strVal = strVal & "&txtSpplCd=" & .hdnSpplCd1.value
			strVal = strVal & "&txtBuildCd=" & .hdnBuildCd1.value
			strVal = strVal & "&txtPoNo=" & .hdnPoNo.value
			strVal = strVal & "&txtIvNo=" & .hdnIvNo.value
			
			strVal = strVal & "&txtPayeeCd=" & .hdnPayeeCd1.value
			strVal = strVal & "&txtCurr=" & .hdnCurr1.value
			strVal = strVal & "&txtPayTermCd=" & .hdnPayTermCd1.value
			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If	
		
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�  
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
		strVal = strVal & "&txtGridNum="	 & "B"
		
		Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
        
    End With

    DbQuery2 = True    
End Function

Function DbQuery2Ok()
	DbQuery2Ok = False
	call SetSpreadLock("B")
	DbQuery2Ok = true
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>������</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtFrIvDt" style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME" ALT="������"></OBJECT>');</SCRIPT>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtToIvDt" style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME" ALT="������"></OBJECT>');</SCRIPT>
									</td>
								<tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>��������</TD>
						<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvTypeCd" ALT="��������" MAXLENGTH=5 style="HEIGHT: 20px; WIDTH: 80px" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" style="HEIGHT: 21px; WIDTH: 16px"  align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">
											   <INPUT CLASS = protected readonly TYPE=TEXT NAME="txtIvTypeNm" ALT="��������" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" nowrap>���ű׷�</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtGrpCd" ALT="���ű׷�" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=4 tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" style="HEIGHT: 21px; WIDTH: 16px" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGrp()" >
														   <INPUT TYPE=TEXT CLASS = protected readonly = True NAME="txtGrpNm" ALT="���ű׷�" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
						<TD CLASS="TD5" nowrap>VAT</TD>
								<TD CLASS="TD6" NOWRAP>
									<Table cellpadding=0 cellspacing=0>
										<TR>
											<TD NOWRAP><INPUT TYPE=TEXT NAME="txtVatCd" ALT="VAT" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="11NXXU"
											ONChange="vbscript:SetVatType()"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnVat" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenVat()">
													   <INPUT TYPE=TEXT NAME="txtVatNm" ALT="VAT" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" >&nbsp;
											</TD>
										</TR>
									</Table>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó" NAME="txtSpplCd" MAXLENGTH=10 SIZE=10 tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(1)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT AlT="����ó" ID="txtSpplNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>									
						<TD CLASS="TD5" NOWRAP>���ݰ�꼭����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="���ݰ�꼭����ó" NAME="txtBuildCd" MAXLENGTH=4 SIZE=10 tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(3)">
														   <INPUT TYPE=TEXT AlT="���ݰ�꼭����ó" NAME="txtBuildNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=32 MAXLENGTH=18 ALT="���ֹ�ȣ" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"><div style="Display:none"><input type="text" name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>���Թ�ȣ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=32  MAXLENGTH=18 ALT="���Թ�ȣ" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvNo()"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=60% valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=40% valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
					<IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderBy()"></IMG></TD>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>	</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" WIDTH=100% SRC="../../blank.htm" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvTypeCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGrpCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVatCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSpplCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBuildCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvNo" tag="14">

<INPUT TYPE=HIDDEN NAME="hdnFrDt1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvTypeCd1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvTypeNm1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGrpCd1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGrpNm1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVatCd1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVatNm1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSpplCd1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSpplNm1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBuildCd1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBuildNm1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoNo1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvNo1" tag="14">

<INPUT TYPE=HIDDEN NAME="hdnPayeeCd1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPayeeNm1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnCurr1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVatRt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPayTermCd1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPayTermNm1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBizAreaCd1" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBizAreaNm1" tag="14">

<INPUT TYPE=HIDDEN NAME="hdnSpplRegNo1" tag="14">  
<INPUT TYPE=HIDDEN NAME="hdnSpplInvNo1" tag="14">  
<INPUT TYPE=HIDDEN NAME="hdnPayDur1" tag="14">  
<INPUT TYPE=HIDDEN NAME="hdnPayTypeCd1" tag="14">  
<INPUT TYPE=HIDDEN NAME="hdnPayTypeNm1" tag="14">  
<INPUT TYPE=HIDDEN NAME="hdnPayTermsTxt1" tag="14">  
<INPUT TYPE=HIDDEN NAME="hdnRemark1" tag="14">    

<INPUT TYPE=HIDDEN NAME="txtCur" tag="14">
<INPUT TYPE=HIDDEN NAME="txtCurNm" tag="14">


</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     