<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5101ra1
'*  4. Program Name         : ������ǥ��ȣ PopUp
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2000/12/09
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Ahn Hye Jin
'* 11. Comment              :
'*                            2000/12/09
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE>����¹�ȣ��ȸ �˾�</TITLE>
<!--
'############################################################################################################
'												1. �� �� �� 
'############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->


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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID        = "f3101rb1_ko441.asp"
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'Const C_SHEETMAXROWS_D  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 5					                          '��: SpreadSheet�� Ű�� ���� 

Dim lsPoNo                                                 '��: Jump�� Cookie�� ���� Grid value
Dim lgIsOpenPop
Dim lgParentsPgmID
Dim IsOpenPop   
Dim lgAuthorityFlag
Dim arrReturn
Dim arrParent
Dim arrParam
Dim lgdiffer
Dim lgcd
Dim Title

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'------ Set Parameters from Parent ASP -----------------------------------------------------------------------

	arrParent		= window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam		= arrParent(1)

	top.document.title = "���������˾�"


'========================================================================================================= 
Sub InitVariables()
    Redim arrReturn(0)
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgAuthorityFlag = arrParam(4)                          '���Ѱ��� �߰� 
    'lgdiffer          = arrParam(5)
    lgcd              = arrParam(6)
    If lgcd <> "" Then
		frm1.txtcd.value = lgcd
	End if	

	Title = "����¹�ȣ"

		
	Self.Returnvalue = arrReturn

	' ���Ѱ��� �߰� 
	If UBound(arrParam) > 7 Then
		lgAuthBizAreaCd		= arrParam(8)
		lgInternalCd		= arrParam(9)
		lgSubInternalCd		= arrParam(10)
		lgAuthUsrID			= arrParam(11)
	End If

End Sub


'========================================================================================================= 
Sub SetDefaultVal()
	lblTitle.innerHTML = Title

End Sub


'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '��: 

End Sub



'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================	
Function OKClick()

	If frm1.vspdData.ActiveRow > 0 Then 				
		Redim arrReturn(5)
		frm1.vspdData.Row	= frm1.vspdData.ActiveRow

		frm1.vspdData.Col	= GetKeyPos("A",2)		
		arrReturn(0)		= frm1.vspdData.Text
		
		frm1.vspdData.Col	= GetKeyPos("A",3)		
		arrReturn(1)		= frm1.vspdData.Text
		
		frm1.vspdData.Col	= GetKeyPos("A",1)		
		arrReturn(2)		= frm1.vspdData.Text
		
		frm1.vspdData.Col	= GetKeyPos("A",4)		
		arrReturn(3)		= frm1.vspdData.Text
		
		frm1.vspdData.Col	= GetKeyPos("A",5)		
		arrReturn(4)		= frm1.vspdData.Text
   End If		
	
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



'==========================================================================================================
Sub InitSpreadSheet()
	frm1.vspdData.OperationMode = 3

	Call SetZAdoSpreadSheet("F3101RA1_KO441", "S", "A", "V20030923", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	
	Call SetSpreadLock() 
End Sub




'=========================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
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
Function OpenOrderBy()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(PopupParent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = "<%'=PopupParent.gMethodText%>"    
  
	For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
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
	
	   For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function


 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 

Sub CookiePage(ByVal Kubun)

	Select Case Kubun		
		Case "FORM_LOAD"
			lgParentsPgmID = PopupParent.ReadCookie("PGMID")
			Call PopupParent.WriteCookie("PGMID", "")			
		Case Else			
	End Select
End Sub


'===========================================================================
Function OpenSortPopup()
   
   	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function

 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029()
	Call InitVariables()	
	Call SetDefaultVal()	
	Call InitSpreadSheet()
	Call CookiePage("FORM_LOAD")
	Call FncQuery()
End Sub


'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub



'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			DbQuery
		End If
    End if
   
End Sub



'=======================================================================================================
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
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub


'--------------- ������ coding part(�������,End)------------------------------------------------------
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)	
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then
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

Sub txtCd_Keypress(KeyAscii)
    On Error Resume Next    
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call Fncquery()
    End if    
End Sub

Sub txtNm_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call Fncquery()
    End if
End Sub


'********************************************************************************************************* 
Function FncQuery() 
	Dim IntRetCD
	
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
    Call InitVariables 														'��: Initializes local global variables
    '-----------------------
    'Query function call area
    '-----------------------

    IF DbQuery = False Then										'��: Query db data
    	Exit Function
    End If															

    FncQuery = True		
End Function



'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function



'========================================================================================

Function FncExcel() 
	Call parent.FncExport(PopupParent.C_MULTI)
End Function


'========================================================================================

Function FncFind() 
    Call parent.FncFind(PopupParent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

 
'========================================================================================

Function DbQuery() 
	Dim strVal
	Dim txtCd
	Dim txtNm

    DbQuery = False
    
    Err.Clear            
	Call LayerShowHide(1)
    
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtCd=" & Trim(.txtcd.value)
		strVal = strVal & "&txtNm=" & Trim(.txtNm.value)
	    'strVal = strVal & "&txtdiffer= " & lgdiffer 

'--------------- ������ coding part(�������,End)------------------------------------------------
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")         
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&lgAuthorityFlag="   & EnCoding(lgAuthorityFlag)            '���Ѱ��� �߰�		

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ����		

        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True

End Function


'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
	Dim IntRetCD
  
    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = True                                                 'Indicates that no value changed

   If frm1.vspdData.MaxRows = 0 Then

      IntRetCD = DisplayMsgBox("900014","X","X","X") 
      If Trim(txtCd.value) > "" Then
         txtCd.Select 
         txtCd.Focus
      Else   
         txtNm.Select 
         txtNm.Focus
     End If
   Else
	 frm1.vspdData.Focus
   End If      
End Function


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtCd" SIZE=20 MAXLENGTH=50 tag="12XXXU" ALT="�ڵ�" ID="Text1"></TD>
			</TR>		
			<TR>
				<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtNm" SIZE=30 MAXLENGTH=50 tag="12"   ALT="�ڵ��" ID="Text2"></TD>
			</TR>		
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vspdData tag="2"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"></OBJECT>');</SCRIPT>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>&nbsp;<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>		
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


