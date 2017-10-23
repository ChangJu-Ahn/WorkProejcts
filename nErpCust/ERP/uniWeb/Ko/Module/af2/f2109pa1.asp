<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2109pa1
'*  4. Program Name         : ����󼼳������� 
'*  5. Program Desc         : Popup of Budget Detail
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.04.01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================
Response.Expires = -1
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<% '#########################################################################################################
'												1. �� �� �� 
'############################################################################################################%>
<% '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<%'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<%'=====================================  1.1.2 ���� Include   =============================================
'===========================================================================================================%>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/eventpopup.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

<%'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
Dim lgBlnFlgChgValue                                        <%'��: Variable is for Dirty flag            %>
Dim lgStrPrevKey                                            <%'��: Next Key tag                          %>
Dim lgSortKey                                               <%'��: Sort���� ���庯��                     %> 
Dim lgIsOpenPop                                             <%'��: Popup status                          %> 

Dim lgSelectList                                            <%'��: SpreadSheet�� �ʱ�  ��ġ�������� ���� %>
Dim lgSelectListDT                                          <%'��: SpreadSheet�� �ʱ�  ��ġ�������� ���� %>

Dim lgTypeCD                                                <%'��: 'G' is for group , 'S' is for Sort    %>
Dim lgFieldCD                                               <%'��: �ʵ� �ڵ尪                           %>
Dim lgFieldNM                                               <%'��: �ʵ� ����                           %>
Dim lgFieldLen                                              <%'��: �ʵ� ��(Spreadsheet����)              %>
Dim lgFieldType                                             <%'��: �ʵ� ����                           %>
Dim lgDefaultT                                              <%'��: �ʵ� �⺻��                           %>
Dim lgNextSeq                                               <%'��: �ʵ� Pair��                           %>
Dim lgKeyTag                                                <%'��: Key ����                                %>

Dim lgSortFieldNm                                           <%'��: Orderby popup�� ����Ÿ(�ʵ弳��)      %>
Dim lgSortFieldCD                                          <%'��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      %>

Dim lgKeyPos                                                <%'��: Key��ġ                               %>
Dim lgKeyPosVal                                             <%'��: Key��ġ Value                         %>

Dim lgPopUpR                                                <%'��: Orderby default ��                    %>
Dim lgMark

Dim IsOpenPop                                                  <%'��: ��ũ                                  %>
Dim lgIntFlgMode               ' Variable is for Operation Status
<%'---------------  coding part(�������,Start)-----------------------------------------------------------
   Call GetAdoFiledInf("F2109PA1","S","A")                        '��: spread sheet �ʵ����� query   -----
                                                                  ' 1. Program id
                                                                  ' 2. G is for Qroup , S is for Sort     
                                                                  ' 3. Spreadsheet no   
'--------------- ������ coding part(�������,End)-------------------------------------------------------------
%>

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f2109pb1.asp"
Const C_SHEETMAXROWS    = 50                                   '��: Spread sheet���� �������� row
Const C_SHEETMAXROWS_D  = 100                                  '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 4                                    '�١١١�: Max key value

Dim arrReturn
Dim arrParent
Dim arrParam					

	<% '------ Set Parameters from Parent ASP ------ %>
	arrParent = Window.DialogArguments
	arrParam = arrParent(0)
	
	top.document.title = "����󼼳���"

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

<% '#########################################################################################################
'												2. Function�� 
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### %>

<% '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= %>
Sub InitVariables()
    Redim arrReturn(0)
    
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgIntFlgMode     = OPMD_CMODE
    
	Self.Returnvalue = arrReturn

End Sub

<% '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= %>
Sub SetDefaultVal()
	Dim ii,kk	
	Dim iCast
	
    lgTypeCD    = Split ("<%=gTypeCD%>"   ,Chr(11))                                 <%'  �ʵ� ��          %>
    lgFieldCD   = Split ("<%=gFieldCD%>"  ,Chr(11))                                 <%'  �ʵ� �ڵ尪      %>
    lgFieldNM   = Split ("<%=gFieldNM%>"  ,Chr(11))                                 <%'  �ʵ� ����      %>
    lgFieldLen  = Split ("<%=gFieldLen%>" ,Chr(11))                                 <%'  �ʵ� ��          %>
    lgFieldType = Split ("<%=gFieldType%>",Chr(11))                                 <%'  �ʵ� ����Ÿ Ÿ�� %>
    lgDefaultT  = Split ("<%=gDefaultT%>" ,Chr(11))                                 <%'  �ʵ� �⺻��      %>
    lgNextSeq   = Split ("<%=gNextSeq%>"  ,Chr(11))                                 <%'  �ʵ� Pair��      %>
    lgKeyTag    = Split ("<%=gKeyTag%>"   ,Chr(11))                                 <%'  �ʵ� Pair��      %>
    
    lgSortFieldNm   = ""
    lgSortFieldCD  = ""

    Redim  lgMark(UBound(lgFieldNM)) 
    
    For ii = 0 To UBound(lgFieldNM) - 1                                            <%'Sort ��󸮽�Ʈ   ���� %>
        iCast = lgDefaultT(ii)
        If  IsNumeric(iCast) Or Trim(lgDefaultT(ii)) = "V" Then
            If IsNumeric(iCast) Then 
               If IsBetween(1,C_MaxSelList,CInt(iCast)) Then    <%'Sort����default�� ���� %>
                  lgPopUpR(CInt(lgDefaultT(ii)) - 1,0) = Trim(lgFieldCD(ii))
                  lgPopUpR(CInt(lgDefaultT(ii)) - 1,1) = "ASC"
               End If
            End If
            lgSortFieldNm   = lgSortFieldNm   & Trim(lgFieldNM (ii)) & Chr(11)
            lgSortFieldCD  = lgSortFieldCD  & Trim(lgFieldCD(ii)) & Chr(11)
        End If
    Next
    
    lgSortFieldNm = split (lgSortFieldNm ,Chr(11))
    lgSortFieldCD = split (lgSortFieldCD,Chr(11))

<%'--------------- ������ coding part(�������,Start)--------------------------------------------------%>
	Dim strSvrDate
	
	frm1.txtBdgCd.value = arrParam(0)
	frm1.txtBdgNm.value = arrParam(1)
	frm1.txtBdgYymm.Text = UNIDateClientFormat(arrParam(2))
	frm1.txtDeptCd.value = arrParam(3)
	frm1.txtDeptNm.value = arrParam(4)
	frm1.hOrgChangeId.value = arrParam(5)

	Call ggoOper.FormatDate(frm1.txtBdgYymm, gDateFormat, 2)
	
<%'--------------- ������ coding part(�������,End)----------------------------------------------------%>

End Sub

<%'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================%>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/ComLoadInfTB19029.asp"  -->
End Sub

<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================
%>	
Function OKClick()
		
	Dim intColCnt, intRowCnt, intInsRow
		
	if frm1.vspdData.ActiveRow > 0 Then 			
		
		intInsRow = 0

		Redim arrReturn(9)
			
		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1
			
			frm1.vspdData.Row = intRowCnt + 1
		
			If frm1.vspdData.SelModeSelected Then
				frm1.vspdData.Col = 1	'�����ڵ� 
				arrReturn(0) = frm1.vspdData.Text
				frm1.vspdData.Col = 2	'����� 
				arrReturn(1) = frm1.vspdData.Text
				frm1.vspdData.Col = 3	'������ 
				arrReturn(2) = frm1.vspdData.Text
				frm1.vspdData.Col = 4	'�μ��ڵ� 
				arrReturn(3) = frm1.vspdData.Text
				frm1.vspdData.Col = 5	'�μ��� 
				arrReturn(4) = frm1.vspdData.Text
				frm1.vspdData.Col = 6	'��������ID
				arrReturn(5) = frm1.vspdData.Text
				frm1.vspdData.Col = 9	'�����Ⱓ���� 
				arrReturn(6) = frm1.vspdData.Text
				frm1.vspdData.Col = 10	'�߰� 
				arrReturn(7) = frm1.vspdData.Text
				frm1.vspdData.Col = 11	'�̿� 
				arrReturn(8) = frm1.vspdData.Text
				frm1.vspdData.Col = 12	'���� 
				arrReturn(9) = frm1.vspdData.Text
				
				intInsRow = intInsRow + 1
			End IF
		Next
		
	End if			
		
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


<%'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================%>
Sub InitSpreadSheet()
    Dim ii,jj,kk,iSeq
    
    lgSelectList   = ""
    lgSelectListDT = ""
    iSeq           = 0 
    
    'frm1.vspdData.OperationMode = 3

    ReDim lgKeyPos(C_MaxKey)
    ReDim lgKeyPosVal(C_MaxKey)

    Redim  lgMark(UBound(lgFieldNM)) 

	With frm1.vspdData

		.MaxCols = 0
		.MaxCols = UBound(lgFieldNM)
	    .MaxRows = 0
	    ggoSpread.Source = frm1.vspdData
		.ReDraw = false
		
	    ggoSpread.Spreadinit

        For ii = 0 to C_MaxSelList - 1
            For jj = 0 to UBound(lgFieldNM) - 1
                If lgMark(jj) <> "X" Then
                   If lgPopUpR(ii,0) = lgFieldCD(jj) Then
                      iSeq = iSeq + 1
                      Call InitSpreadSheetRow(iSeq,jj)
                      If IsBetween(1,UBound(lgFieldNM),CInt(lgNextSeq(jj))) Then 
                         kk = CInt(lgNextSeq(jj)) 
                         iSeq = iSeq + 1
                         Call InitSpreadSheetRow(iSeq,kk-1)
                      End If    
                   End If 
                 End If 
            Next       
        Next      
         
        For ii = 0 to UBound(lgFieldNM) - 1
            If lgMark(ii) <> "X" Then
               If lgTypeCD(0) = "S" Or (lgTypeCD(0) = "G" And lgDefaultT(ii) = "L") Then
                  iSeq = iSeq + 1
                  Call InitSpreadSheetRow(iSeq,ii)
                  If IsBetween(1,UBound(lgFieldNM),CInt(lgNextSeq(ii))) Then 
                     kk = CInt(lgNextSeq(ii)) 
                     iSeq = iSeq + 1
                     Call InitSpreadSheetRow(iSeq,kk-1)
                  End If   
               End If   
            End If 
        Next       

	   .MaxCols = iSeq
       .ReDraw = true
	    Call SetSpreadLock 
    End With        
End Sub


<%'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheetRow
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================%>
Sub InitSpreadSheetRow(Byval iCol,ByVal iDx)
   Dim iAlign
   
   lgMark(iDx) = "X"
   
   iAlign = Trim(Mid(lgFieldType(iDx),3,1))
   
   If  iAlign = "" Then
       If Mid(lgFieldType(iDx),1,1) = "F" Then
          iAlign = "1"
       Else 
          iAlign = "0"
       End If   
   End If
   
   iAlign =  CInt(iAlign)

   Select Case  Mid(lgFieldType(iDx),1,2)
     Case "BT" 'Button
		    ggoSpread.SSSetButton iCol
     Case "CB" 'Combo
            ggoSpread.SSSetCombo  iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign
     Case "CK" 'Check
            ggoSpread.SSSetCheck  iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign, "", True, -1
     Case "DD"   '��¥ 
            ggoSpread.SSSetDate   iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign, gDateFormat
     Case "ED"   '���� 
            ggoSpread.SSSetEdit   iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign
     Case "F2"  ' �ݾ� 
            Call SetSpreadFloat  (iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign,2)
     Case "F3"  ' ���� 
            Call SetSpreadFloat  (iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign,3)
     Case "F4"  ' �ܰ� 
            Call SetSpreadFloat  (iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign,4)
     Case "F5"   ' ȯ�� 
            Call SetSpreadFloat  (iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign,5)
     Case "MK"   ' Mask
            ggoSpread.SSSetMask   iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign
     Case "ST"   ' Static
            ggoSpread.SSSetStatic iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign
     Case "TT"   ' Time
            ggoSpread.SSSetTime   iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign   ,1,1
     Case "HH"   ' Hidden
            ggoSpread.Source.Col = iCol
            ggoSpread.Source.ColHidden = true            
     Case Else
            ggoSpread.SSSetEdit   iCol , lgFieldNM(iDx), lgFieldLen(iDx), iAlign
   End Select
   
   If Len(Trim(lgSelectList)) > 0  And Len(Trim(lgFieldCD(iDx))) > 0 Then
      lgSelectList   = lgSelectList & " , " 
   End If   
   lgSelectList   = lgSelectList & lgFieldCD(iDx)         

   lgSelectListDT = lgSelectListDT & lgFieldType(iDx) & gColSep
   
   <%' Spreadsheet #2�˻��� ���� Ű ����ġ ���� %>
   If CInt(lgKeyTag(iDx)) > 0 And CInt(lgKeyTag(iDx)) <= C_MaxKey Then  
      lgKeyPos(CInt(lgKeyTag(iDx))) =  iCol
   End If

End Sub

<%'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================%>
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	 ggoSpread.SpreadLock 1 , -1
    .vspdData.ReDraw = True

    End With
End Sub


<% '**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** %>

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	Select Case iWhere

		Case "BdgCdFr", "BdgCdTo"
			arrParam(0) = "�����ڵ� �˾�"								<%' �˾� ��Ī %>
			arrParam(1) = "F_BDG_ACCT A "									<%' TABLE ��Ī %>
			arrParam(2) = strCode											<%' Code Condition%>
			arrParam(3) = ""												<%' Name Cindition%>
			arrParam(4) = ""
			arrParam(5) = "�����ڵ�"									<%' �����ʵ��� �� ��Ī %>

			arrField(0) = "A.BDG_CD"	     								<%' Field��(0)%>
			arrField(1) = "A.GP_ACCT_NM"			    					<%' Field��(1)%>
			
			arrHeader(0) = "�����ڵ�"									<%' Header��(0)%>
			arrHeader(1) = "�����"										<%' Header��(1)%>
			
	   Case "DeptCd"
			arrParam(0) = "�μ��ڵ� �˾�"								<%' �˾� ��Ī %>
			arrParam(1) = "B_ACCT_DEPT A "									<%' TABLE ��Ī %>
			arrParam(2) = strCode											<%' Code Condition%>
			arrParam(3) = ""												<%' Name Cindition%>
			arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(gChangeOrgId , "''", "S") & ""
			arrParam(5) = "�μ��ڵ�"									<%' �����ʵ��� �� ��Ī %>

			arrField(0) = "A.DEPT_CD"
			arrField(1) = "A.DEPT_NM"
			
			arrHeader(0) = "�μ��ڵ�"									<%' Header��(0)%>
			arrHeader(1) = "�μ���"										<%' Header��(1)%>
		
		Case Else
			Exit Function
	End Select	
	
	lgIsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		    
		    Case "BdgCdFr"
				.txtBdgCdFr.value = arrRet(0)
				.txtBdgNmFr.value = arrRet(1)
				
		    Case "BdgCdTo"
				.txtBdgCdTo.value = arrRet(0)
				.txtBdgNmTo.value = arrRet(1)
				
			Case "DeptCd"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
			
		End Select
    
	End With

End Function

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenOrderBy()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = "<%=gMethodText%>"    
  
	For ii = 0 to C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(lgSortFieldCD,lgSortFieldNm,arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function


<% '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ %>

<% '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== %>

<% '#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################%>
<% '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* %>
<% '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= %>
Sub Form_Load()

    Call LoadInfTB19029														'��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    ReDim lgPopUpR(C_MaxSelList - 1,1)
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	    
<%'--------------- ������ coding part(�������,Start)----------------------------------------------------%>
	Call FncQuery
<%'--------------- ������ coding part(�������,End)------------------------------------------------------%>
End Sub
<%
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
%>
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

<% '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* %>

<% '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* %>

<%
'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************

%>

<%
'==========================================================================================
'   Event Name : DblClick
'   Event Desc :
'==========================================================================================
%>

<%
'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================
%>

<%
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
%>
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ'
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DbQuery
		End If
   End if
    
End Sub

<%'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================%>
Sub vspdData_Click(ByVal Col, ByVal Row)
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
<%'--------------- ������ coding part(�������,Start)----------------------------------------------------%>
	If Row < 1 Then Exit Sub

<%'--------------- ������ coding part(�������,End)------------------------------------------------------%>
    
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        'Call OKClick()
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

<% '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
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
Function FncQuery() 
Dim IntRetCD
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery															'��: Query db data

    FncQuery = True		
End Function

<%
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
%>
Function FncPrint() 
    Call parent.FncPrint()
End Function

<%
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
%>
Function FncExcel() 
	Call parent.FncExport(C_MULTI)
End Function

<%
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
%>
Function FncFind() 
    Call parent.FncFind(C_MULTI , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function

<%
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
%>
Function FncExit()
    FncExit = True
End Function

<% '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* %>
<%
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
%>
Function DbQuery() 
	Dim strVal
	Dim ColYymm1, ii

    DbQuery = False
    
    Err.Clear           

    frm1.vspdData.MaxRows = 0                                                   '��: Protect system from crashing                                                    '��: Protect system from crashing
    
	Call LayerShowHide(1)

	For ii = LBound(lgFieldCD) To UBound(lgFieldCD)
		If lgFieldCD(ii) = "A.BDG_YYYYMM" Then 
			ColYymm1 = ii	'������ �÷� 
			Exit For
		End If
	Next
    
    With frm1
<%'--------------- ������ coding part(�������,Start)----------------------------------------------%>
	If lgIntFlgMode = OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtBdgYymm=" & Trim(.hBdgYymm.value)
		strVal = strVal & "&txtDeptCd=" & Trim(.hDeptCd.value)
		strVal = strVal & "&txtBdgCd=" & Trim(.hBdgCd.value)
		strVal = strVal & "&txtOrgChangeId=" & Trim(.hOrgChangeId.value)
		strVal = strVal & "&txtColYymm=" & ColYymm1
		strVal = strVal & "&txtDateType=" & gComDateType
	Else
		strVal = BIZ_PGM_ID & "?txtBdgYymm=" & ggoOper.RetFormat(.txtBdgYymm.Text, "yyyyMM")
		strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptCd.value)
		strVal = strVal & "&txtBdgCd=" & Trim(.txtBdgCd.value)
		strVal = strVal & "&txtOrgChangeId=" & Trim(.hOrgChangeId.value)
		strVal = strVal & "&txtColYymm=" & ColYymm1
		strVal = strVal & "&txtDateType=" & gComDateType
	End If
	
<%'--------------- ������ coding part(�������,End)------------------------------------------------%>
	strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
	strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	strVal = strVal & "&lgSelectListDT=" & lgSelectListDT
	strVal = strVal & "&lgTailList="     & MakeSql()
	strVal = strVal & "&lgSelectList="   & EnCoding(lgSelectList)

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    End With
    
    DbQuery = True

End Function

<%
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
%>
Function DbQueryOk()														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = False                                                 'Indicates that no value changed
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	lgIntFlgMode = OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	End If
	
End Function

<%
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
%>
<%
'========================================================================================
' Function Name : MakeSql()
' Function Desc : Order by ���� group by ���� �����.
'========================================================================================
%>
Function MakeSql()
    Dim iStr,jStr
    Dim ii,jj,kk
    Dim iFirst
    Dim tmpPopUpR
    
    '2001/03/30 �ڵ�, �ڵ�� ���İ��� ���� 
    Redim tmpPopUpR(C_MaxSelList - 1)
    For kk = 0 to C_MaxSelList - 1
		tmpPopUpR(kk) = lgPopUpR(kk,0)
    Next
    
    iFirst = "N"
    iStr   = ""  
    jStr   = ""      

    Redim  lgMark(0) 
    Redim  lgMark(UBound(lgFieldNM)) 
    lgMark(0) = ""
    
    For ii = 0 to C_MaxSelList - 1
        If tmpPopUpR(ii) <> "" Then	'2001/03/30 �ڵ�, �ڵ�� ���İ��� ���� 
           If lgTypeCD(0) = "G" Then
              For jj = 0 To UBound(lgFieldNM) - 1                                            <%'Sort ��󸮽�Ʈ   ���� %>
                  If lgMark(jj) <> "X" Then
                     If lgPopUpR(ii,0) = lgFieldCD(jj) Then
                        If iFirst = "Y" Then
                           iStr = iStr & " , "
                           jStr = jStr & " , " 
                        End If   
                        If CInt(Trim(lgNextSeq(jj))) >= 1 And CInt(Trim(lgNextSeq(jj))) <= UBound(lgFieldNM) Then
                           iStr = iStr & " " & lgPopUpR(ii,0) & " " & lgPopUpR(ii,1) & " , " & lgFieldCD(CInt(lgNextSeq(jj)) - 1)
                           jStr = jStr & " " & lgPopUpR(ii,0) & " " &          " , " & lgFieldCD(CInt(lgNextSeq(jj)) - 1)
                           '2001/03/30 �ڵ�, �ڵ�� ���İ��� ���� 
                           If (ii + 1) < C_MaxSelList Then
								For kk = ii + 1 to C_MaxSelList - 1
									If lgPopUpR(kk,0) = lgFieldCD(CInt(lgNextSeq(jj)) - 1) Then
										iStr = iStr & " " & lgPopUpR(kk,1)
										tmpPopUpR(kk) = ""
									End If
								Next
                           End If
                           lgMark(CInt(lgNextSeq(jj)) - 1) = "X"
                        Else
                          iStr = iStr & " " & lgPopUpR(ii,0) & " " & lgPopUpR(ii,1)
                          jStr = jStr & " " & lgPopUpR(ii,0) 
                        End If
                        iFirst = "Y"
                        lgMark(jj) = "X"
                     End If
                     
                  End If
              Next
           Else
              If iFirst = "Y" Then
                 iStr = iStr & " , "
                 jStr = jStr & " , " 
              End If   
              iStr = iStr & " " & lgPopUpR(ii,0) & " " & lgPopUpR(ii,1)
              iFirst = "Y"
           End If
              
        End If
    Next     
    
    If lgTypeCD(0) = "G" Then
       MakeSql =  "Group By " & jStr  & " Order By " & iStr 
    Else
       MakeSql = "Order By" & iStr
    End If   


End Function
'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### %>
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
						<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgCd" MAXLENGTH="18" SIZE=10  ALT ="�����ڵ�" tag="14XXXU">&nbsp;<INPUT NAME="txtBdgNm" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="�����ڵ��" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>������</TD>
						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f2109pa1_fpBdgYymmFr_txtBdgYymm.js'></script></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>�μ��ڵ�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" MAXLENGTH="10" SIZE=10 ALT ="�μ��ڵ�" tag="14XXXU">&nbsp;<INPUT NAME="txtDeptNm" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="�μ���" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
						<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<script language =javascript src='./js/f2109pa1_vspdData_vspdData.js'></script>
					</TD>
				</TR>
				<TR HEIGHT=20>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ݾ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgPlanAmt" SIZE=15  ALT ="����ݾ�"     tag="24XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>�����ѵ��ݾ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgAmt"     SIZE=15  ALT ="�����ѵ��ݾ�" tag="24XXXU"></TD>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>&nbsp;<!--<IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Query.gif',1)" ONCLICK="FncQuery()"></IMG> --></TD>
					<TD WIDTH=30% ALIGN=RIGHT><!--<IMG SRC="../../image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/OK.gif',1)" ONCLICK="OkClick()"></IMG> -->
							                  <IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CLOSE" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hBdgYymm" tag="14">
<INPUT TYPE=HIDDEN NAME="hDeptCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hBdgCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
