<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution                                                      *
'*  2. Function Name        :                                                                           *
'*  3. Program ID           : S5211MA1
'*  4. Program Name         : ���� B/L���                                                              *
'*  5. Program Desc         : ���� B/L���																*
'*  6. Comproxy List        : PS7G131.dll,PS7G115.dll										            *
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2002/11/15																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Ahn Tae Hee												                *
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/04/18 : ȭ�� design												*
'*							  2. 2000/04/18 : Coding Start												*
'*                            3. 2002/11/15 : UI ǥ������												*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                             '��: indicates that All variables must be declared in advance
'========================================================================================================
 Const BIZ_PGM_ID    = "s5211mb1.asp" 
 Const BIZ_PGM_SOQRY_ID   = "s5211mb2.asp" 
 Const BIZ_PGM_LCQRY_ID   = "s5211mb3.asp"
 Const BIZ_PGM_CCQRY_ID   = "s5211mb4.asp" 
 'Const EXBL_DETAIL_ENTRY_ID  = "s5212ma1"  '20120725 ����ȣ �ּ�ó�� b/l������Ͻ� ko441�� �����Ȱ� ���ǵ��� ��û: ��̿�
 Const EXBL_DETAIL_ENTRY_ID  = "s5212ma1_ko441"  
 Const EXPORT_CHARGE_ENTRY_ID = "s6111ma1"  
 Const BIZ_BillCollect_JUMP_ID = "s5115ma1"
'========================================================================================================
 Const TAB1 = 1
 Const TAB2 = 2
 Const TAB3 = 3
 
 Const PostFlag = "PostFlag"
 
 '------ Minor Code PopUp�� ���� Major Code���� ------ 
 Const gstrTransportMajor  = "B9009"
 Const gstrFreightMajor   = "S9007"
 Const gstrPackingTypeMajor  = "B9007"
 Const gstrPayTypeMajor   = "A1006"
 Const gstrOriginMajor   = "B9094" 
 Const gstrVATTypeMajor   = "B9001"
'========================================================================================================
 Dim lgBlnFlgChgValue    '��: Variable is for Dirty flag 
 Dim lgIntGrpCount     '��: Group View Size�� ������ ���� 
 Dim lgIntFlgMode     '��: Variable is for Operation Status 

 Dim gSelframeFlg     '���� TAB�� ��ġ�� ��Ÿ���� Flag %>
 Dim gblnWinEvent     '~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
              ' PopUp Window�� ��������� ���θ� ��Ÿ���� variable 
'========================================================================================================
Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

Dim IsOpenPop
'========================================================================================================
 Function InitVariables()
  lgIntFlgMode = Parent.OPMD_CMODE        '��: Indicates that current mode is Create mode
  lgBlnFlgChgValue = False        '��: Indicates that no value changed
  lgIntGrpCount = 0          '��: Initializes Group View Size
  
  '------ Coding part ------ 
  gblnWinEvent = False
  Call BtnDisabled(1)
 End Function
'=========================================================================================================
 Sub SetDefaultVal()
  With frm1
   .txtBLIssueDt.text  = EndDate
   .txtLoadingDt.text  = EndDate
   .txtDocAmt.text   = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
   .txtDocAmt1.text  = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
   .txtMoney.text   = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
   .txtXchRate.text  = UNIFormatNumber(0, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
   .txtLocAmt.text   = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
   .txtLocAmt1.text  = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
   .txtGrossWeight.text = UNIFormatNumber(0, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
   .txtGrossVolumn.value = UNIFormatNumber(0, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
   .txtLocCurrency.value = Parent.gCurrency
   .txtLocCurrency1.value = Parent.gCurrency
   .btnPosting.disabled = True
   .btnGLView.disabled = True
   .btnPreRcptView.disabled = True
  End With
  
  lgBlnFlgChgValue = False
 End Sub
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub 
'========================================================================================================

 Function ClickTab1()
  If gSelframeFlg = TAB1 Then Exit Function
  
  Call changeTabs(TAB1)
  
  gSelframeFlg = TAB1
 End Function

 Function ClickTab2()
  If gSelframeFlg = TAB2 Then Exit Function
  
  Call changeTabs(TAB2)
  
  gSelframeFlg = TAB2
 End Function
 
 Function ClickTab3()
  If gSelframeFlg = TAB3 Then Exit Function

  Call changeTabs(TAB3)
  
  gSelframeFlg = TAB3
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenEXBLNoPop()
	Dim iCalledAspName
	Dim strRet
	  
	If gblnWinEvent = True Or UCase(frm1.txtBLNo.className) = "PROTECTED" Then Exit Function
	  
	iCalledAspName = AskPRAspName("s5211pa1_KO441")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5211pa1_KO441", "x")
		gblnWinEvent = False
		exit Function
	end if

	gblnWinEvent = True
	  
	strRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	  
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetExBLNo(strRet)
	End If 
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenSORef()
	Dim iCalledAspName
	Dim strRet
	    
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 
	  
	If gblnWinEvent = True Then Exit Function
	   
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3111ra8_KO441")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111ra8_KO441", "x")
		gblnWinEvent = False
		exit Function
	end if

	gblnWinEvent = True
	    
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	    
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetSORef(strRet)
	End If 
End Function 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenLCRef()
	Dim iCalledAspName
	Dim strRet
	  
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 

	iCalledAspName = AskPRAspName("s3211ra8_KO441")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3211ra8_KO441", "x")
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
	 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	  
	If strRet(0) = "" Then
		Exit Function
	Else
		Call ggoOper.ClearField(Document, "A")        '��: Clear Condition,Contents  Field
		Call SetRadio()
		Call InitVariables             '��: Initializes local global variables
		Call SetDefaultVal

		Call SetLCRef(strRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenCCRef()
	Dim iCalledAspName
	Dim arrRet
	  
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 

	iCalledAspName = AskPRAspName("s4211ra8_KO441")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s4211ra8_KO441", "x")
		exit Function
	end if

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	  
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCCRef(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenSalesGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "���ݿ����׷�"     
	arrParam(1) = "B_SALES_GRP"       
	arrParam(2) = Trim(frm1.txtToSalesGroup.value)  
	arrParam(3) = ""         
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "      
	arrParam(5) = "���ݿ����׷�"     

	arrField(0) = "SALES_GRP"       
	arrField(1) = "SALES_GRP_NM"      

	arrHeader(0) = "���ݿ����׷�"     
	arrHeader(1) = "���ݿ����׷��"     

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesGroup(arrRet)
	End If
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenPayType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "�Ա�����"				' �˾� ��Ī 
	arrParam(1) = "B_MINOR,B_CONFIGURATION," _
				& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & ""_
    		    & "And MINOR_CD= " & FilterVar(frm1.txtPayTerms.value, "''", "S") & " And SEQ_NO>=2)C" ' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPayType.value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
				& "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("R", "''", "S") & " )" ' Where Condition
	arrParam(5) = "�Ա�����"				' TextBox ��Ī 
	  
	arrField(0) = "B_MINOR.MINOR_CD"			' Field��(0)
	arrField(1) = "B_MINOR.MINOR_NM"			' Field��(1)
		   
	arrHeader(0) = "�Ա�����"				' Header��(0)
	arrHeader(1) = "�Ա�������"				' Header��(1)
		
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPayType(arrRet)
	End If
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
	 Dim arrRet
	 Dim arrParam(5), arrField(6), arrHeader(6)

	 If gblnWinEvent = True Then Exit Function

	 gblnWinEvent = True

	 arrParam(0) = strPopPos        
	 arrParam(1) = "B_Minor"        
	 arrParam(2) = Trim(strMinorCD)      
	 arrParam(3) = ""         
	 arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""  
	 arrParam(5) = strPopPos        

	 arrField(0) = "Minor_CD"       
	 arrField(1) = "Minor_NM"       

	 arrHeader(0) = strPopPos       
	 arrHeader(1) = strPopPos & "��"     

	 arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	   "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	 gblnWinEvent = False

	 If arrRet(0) = "" Then
		Exit Function
	 Else
		Call SetMinorCd(strMajorCd, arrRet)
	 End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenCountry(strCntryCD, strPopPos)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)

  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "����"       
  arrParam(1) = "B_COUNTRY"       
  arrParam(2) = Trim(strCntryCD)      
  arrParam(3) = ""         
  arrParam(4) = ""         
  arrParam(5) = "����"       

  arrField(0) = "COUNTRY_CD"       
  arrField(1) = "COUNTRY_NM"       

  arrHeader(0) = "����"    
  arrHeader(1) = "������"    

  arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetCountry(strPopPos, arrRet)
  End If
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenBizPartner(strBizPartnerCD, strBizPartnerNM, strPopPos)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)

  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = strPopPos       
  arrParam(1) = "B_BIZ_PARTNER"     
  arrParam(2) = Trim(strBizPartnerCD)    
  arrParam(3) = ""        
  arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "     
  arrParam(5) = strPopPos       

  arrField(0) = "BP_CD"       
  arrField(1) = "BP_NM"       

  arrHeader(0) = strPopPos      
  arrHeader(1) = strPopPos & "��"    

  arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetBizPartner(strPopPos, arrRet)
  End If
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenUnit(strUnitCD, strDim, strPopPos)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)

  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = strPopPos       
  arrParam(1) = "B_UNIT_OF_MEASURE"    
  arrParam(2) = Trim(strUnitCD)     
  arrParam(3) = ""        
  arrParam(4) = "DIMENSION= " & FilterVar(strDim, "''", "S") & ""  
  arrParam(5) = strPopPos       

  arrField(0) = "UNIT"       
  arrField(1) = "UNIT_NM"       

  arrHeader(0) = strPopPos      
  arrHeader(1) = strPopPos & "��"    

  arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetUnit(strDim, arrRet)
  End If
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenTaxOffice()
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)

  OpenTaxOffice = False
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "���ݽŰ�����"		<%' �˾� ��Ī %>
  arrParam(1) = "B_TAX_Biz_Area"			<%' TABLE ��Ī %>
  arrParam(2) = Trim(frm1.txtTaxBizArea.value)  <%' Code Condition%>
  arrParam(3) = ""							<%' Name Cindition%>
  arrParam(4) = ""							<%' Where Condition%>
  arrParam(5) = "���ݽŰ�����"		<%' TextBox ��Ī %>

  arrField(0) = "TAX_BIZ_AREA_CD"			<%' Field��(0)%>
  arrField(1) = "TAX_BIZ_AREA_NM"			<%' Field��(1)%>

  arrHeader(0) = "���ݽŰ�����"		<%' Header��(0)%>
  arrHeader(1) = "���ݽŰ������"     <%' Header��(1)%>

  arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetTaxOffice(arrRet)
   OpenTaxOffice = True
  End If
 End Function
'==========================================================================================
Function txtTaxBizArea_OnChange()
 If Trim(frm1.txtTaxBizArea.value) = "" Then
  frm1.txtTaxBizAreaNm.value = ""
 Else
  If Not GetTaxBizArea("NM") Then txtTaxBizArea_OnChange = False
 End if
End Function
'====================================================================================================
Function GetTaxBizArea(Byval pvStrFlag)

 Dim iStrSelectList, iStrFromList, iStrWhereList
 Dim iStrApplicant, iStrSalesGrp, iStrTaxBizArea
 Dim iStrRs
 Dim iArrTaxBizArea(2), iArrTemp
 
 GetTaxBizArea = False
 
 <%'���ݽŰ� ����� Edting�� ��ȿ�� Check �� ����� �� Fetch %> 
 If pvStrFlag = "NM" Then
  iStrTaxBizArea = frm1.txtTaxBizArea.value
 Else
  iStrApplicant = frm1.txtApplicant.value
  iStrSalesGrp = frm1.txtSalesGroup.value
  <%'����ó�� ���� �׷��� ��� ��ϵǾ� �ִ� ��� �����ڵ忡 ������ rule�� ������ %>
  If Len(iStrApplicant) > 0 And Len(iStrSalesGrp) > 0 Then pvStrFlag = "*"
 End if
 
 iStrSelectList = " * "
 iStrFromList = " dbo.ufn_s_GetTaxBizArea ( " & FilterVar(iStrApplicant, "''", "S") & ",  " & FilterVar(iStrSalesGrp, "''", "S") & ",  " & FilterVar(iStrTaxBizArea, "''", "S") & ",  " & FilterVar(pvStrFlag, "''", "S") & ") "
 iStrWhereList = ""
 
 Err.Clear
    
 If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
  iArrTemp = Split(iStrRs, Chr(11))
  iArrTaxBizArea(0) = iArrTemp(1)
  iArrTaxBizArea(1) = iArrTemp(2)
  Call SetTaxOffice(iArrTaxBizArea)
  GetTaxBizArea = True
 Else
  ' ���� �Ű� ������� Editing�� ��� 
  If pvStrFlag = "NM" Then
   If Not OpenTaxOffice() Then
    frm1.txtTaxBizArea.value = ""
    frm1.txtTaxBizAreaNm.value = ""
   Else
    GetTaxBizArea = True
   End if
  End if
 End if
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenPort(strMinorCD, strMinorNM, strPopNm, iwhere)
 
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)

  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True
  
   arrParam(0) = strPopNm       
   arrParam(1) = "B_MINOR"       
   arrParam(2) = Trim(strMinorCD)     
   arrParam(3) = ""        
   arrParam(4) = "MAJOR_CD = " & FilterVar("B9092", "''", "S") & ""    
   arrParam(5) = strPopNm       
  
   arrField(0) = "Minor_CD"      
   arrField(1) = "Minor_NM"      
     
   arrHeader(0) = strPopNm       
   arrHeader(1) = strPopNm & "��"    

   arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetOpenPort(iwhere, arrRet)
  End If 
   
 End Function  
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetExBLNo(strRet)
  frm1.txtBLNo.value = strRet(0)
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetSORef(strRet)
  Call ggoOper.ClearField(Document, "A")        <% '��: Clear Condition,Contents  Field %>
  Call InitVariables             <% '��: Initializes local global variables %>
  Call SetDefaultVal

  frm1.txtSONo.value = strRet(0)
  frm1.txtBillType.value = strRet(1)
  frm1.txtBillTypeNm.value = strRet(2)

  Dim strVal

  If LayerShowHide(1) = False Then
   Exit Function
  End If

  strVal = BIZ_PGM_SOQRY_ID & "?txtSONo=" & Trim(frm1.txtSONo.value) <%'��: �����Ͻ� ó�� ASP�� ���� %>

  Call RunMyBizASP(MyBizASP, strVal)         <%'��: �����Ͻ� ASP �� ���� %>

  lgBlnFlgChgValue = True
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetLCRef(strRet)
  Dim strVal
  
  frm1.txtHLCNo.value = strRet(0) 
  frm1.txtSONo.value = strRet(1)
  frm1.txtBillType.value = strRet(2)
  frm1.txtBillTypeNm.value = strRet(3)
  
  If LayerShowHide(1) = False Then
   Exit Function
  End If

  strVal = BIZ_PGM_LCQRY_ID & "?txtMode=" & Parent.UID_M0001       <%'��: �����Ͻ� ó�� ASP�� ���� %>
  strVal = strVal & "&txtLCNo=" & Trim(frm1.txtHLCNo.value)
  strVal = strVal & "&txtSONo=" & Trim(frm1.txtSONo.value)

  Call RunMyBizASP(MyBizASP, strVal)           <%'��: �����Ͻ� ASP �� ���� %>

  lgBlnFlgChgValue = True
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetCCRef(strRet)
  Call ggoOper.ClearField(Document, "2")        <% '��: Clear Contents  Field %>
  Call SetRadio()
  Call InitVariables             <% '��: Initializes local global variables %>
  Call SetDefaultVal

  frm1.txtCCNo.value = strRet(0)
  frm1.txtSONo.value = strRet(1)
  frm1.txtBillType.value = strRet(2)
  frm1.txtBillTypeNm.value = strRet(3)

  Dim strVal

  If LayerShowHide(1) = False Then
   Exit Function
  End If

  strVal = BIZ_PGM_CCQRY_ID & "?txtCCNo=" & Trim(frm1.txtCCNo.value) <%'��: �����Ͻ� ó�� ASP�� ���� %>
  strVal = strVal & "&txtSONo=" & Trim(frm1.txtSONo.value)
  
  Call RunMyBizASP(MyBizASP, strVal)         <%'��: �����Ͻ� ASP �� ���� %>

  lgBlnFlgChgValue = True
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetSalesGroup(arrRet)
     frm1.txtToSalesGroup.value = arrRet(0)
     frm1.txtToSalesGroupNm.value = arrRet(1)

     lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetPayType(arrRet)
    frm1.txtPayType.Value = arrRet(0)
    frm1.txtPayTypeNm.Value = arrRet(1)
   

    lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetMinorCd(strMajorCd, arrRet)
     Select Case strMajorCd
     Case gstrTransportMajor
          frm1.txtTransport.Value = arrRet(0)
          frm1.txtTransportNm.Value = arrRet(1)
   
     Case gstrFreightMajor
          frm1.txtFreight.Value = arrRet(0)
          frm1.txtFreightNm.Value = arrRet(1)

     Case gstrPackingTypeMajor
          frm1.txtPackingType.Value = arrRet(0)
          frm1.txtPackingTypeNm.Value = arrRet(1)
          
     Case gstrOriginMajor
          frm1.txtOrigin.Value = arrRet(0)
          frm1.txtOriginNm.Value = arrRet(1) 

     Case gstrVATTypeMajor
          frm1.txtVatType.Value = arrRet(0)
          frm1.txtVatTypeNm.Value = arrRet(1) 

     Case Else
     End Select

     lgBlnFlgChgValue = True
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetCountry(strPopPos, arrRet)
  Select Case UCase(strPopPos)
   Case "VESSEL"
    frm1.txtVesselCntry.Value = arrRet(0)
    frm1.txtVesselCntryNm.Value = arrRet(1)
        
   Case "TRANSHIP"
    frm1.txtTranshipCntry.Value = arrRet(0)
    frm1.txtTranshipCntryNm.Value = arrRet(1)
    
   Case "ORIGIN"
    frm1.txtOriginCntry.Value = arrRet(0)
    frm1.txtOriginCntryNm.value = arrRet(1)
    
   Case Else
  End Select

  lgBlnFlgChgValue = True
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetBizPartner(strPopPos, arrRet)
  Select Case UCase(strPopPos)
   Case "������"
    frm1.txtAgent.Value = arrRet(0)
    frm1.txtAgentNm.Value = arrRet(1)
    
   Case "������"
    frm1.txtManufacturer.Value = arrRet(0)
    frm1.txtManufacturerNm.Value = arrRet(1)
    
   Case "����ȸ��"
    frm1.txtForwarder.Value = arrRet(0)
    frm1.txtForwarderNm.Value = arrRet(1)
    
   Case "����ó"
    frm1.txtPayer.value = arrRet(0)
    frm1.txtPayerNm.value = arrRet(1)
   Case Else
  End Select

  lgBlnFlgChgValue = True
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetUnit(strDim, arrRet)
  Select Case UCase(strDim)
   Case "WT"
    frm1.txtWeightUnit.Value = arrRet(0)
    
   Case "WD"
    frm1.txtVolumnUnit.Value = arrRet(0)
     
   Case Else
  End Select

  lgBlnFlgChgValue = True
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetTaxOffice(arrRet)
  frm1.txtTaxBizArea.value = arrRet(0)
  frm1.txtTaxBizAreaNm.value = arrRet(1)

  lgBlnFlgChgValue = True
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetOpenPort(iwhere, arrRet)
  Select Case iwhere    
   Case 0
    frm1.txtLoadingPort.Value = arrRet(0)
    frm1.txtLoadingPortNm.Value = arrRet(1) 

   Case 1
    frm1.txtDischgePort.Value = arrRet(0)
    frm1.txtDischgePortNm.Value = arrRet(1) 
  End Select   
     
  lgBlnFlgChgValue = True
 End Function
'========================================================================================================
 Function CookiePage(ByVal Kubun)

  On Error Resume Next

  Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
  Dim strTemp, arrVal

  Select Case Kubun
  'ȭ�� Open�� 
  Case 0
   strTemp = ReadCookie(CookieSplit)
    
   If strTemp = "" then Exit Function
    
   frm1.txtBLNo.value =  strTemp
   
   WriteCookie CookieSplit , ""

   Call DbQuery()
      
  '������� 
  Case 1
   WriteCookie CookieSplit , frm1.txtHBLNo.value
  '����� 
  Case 2 
   WriteCookie CookieSplit , "EB" & Parent.gRowSep & frm1.txtSalesGroup.value & Parent.gRowSep & frm1.txtSalesGroupNm.value & Parent.gRowSep & frm1.txtHBLNo.value 
  End Select 
   
 End Function
'========================================================================================================
 Function LoadExportCharge()
  Dim strDtlOpenParam

  WriteCookie "txtChargeType", "EB"
  WriteCookie "txtBasNo", UCase(Trim(frm1.txtBLNo.value))
  PgmJump(EXPORT_CHARGE_ENTRY_ID)
 End Function
'========================================================================================================
 Function SetRadio()
  Dim blnOldFlag

  blnOldFlag = lgBlnFlgChgValue

  frm1.rdoPostingflg2.checked = True

  lgBlnFlgChgValue = blnOldFlag
 End Function
'========================================================================================================
 Function PostBL()
  If Trim(frm1.txtHBLNo.value) = "" Then
   Call DisplayMsgBox("900002", "x", "x", "x") <% '��: "Will you destory previous data" %>

   Exit Function
  End If

  Dim strVal

  If LayerShowHide(1) = False Then
   Exit Function
  End If
  strVal = BIZ_PGM_ID & "?txtMode=" & PostFlag 
  strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value) <%'��: �����Ͻ� ó�� ASP�� ���� %>
  strVal = strVal & "&txtgChangeOrgId=" & Parent.gChangeOrgId
  strVal = strVal & "&txtInsrtUserId=" & Parent.gUsrID       <%'��: ��ȸ ���� ����Ÿ %>

  Call RunMyBizASP(MyBizASP, strVal)          <%'��: �����Ͻ� ASP �� ���� %>
 End Function
'========================================================================================================
 Function PostingOk()
  Dim blnOldFlag
  blnOldFlag = lgBlnFlgChgValue

  frm1.rdoPostingflg1.checked = True

  lgBlnFlgChgValue = blnOldFlag

  lgBlnFlgChgValue = False

  Call MainQuery()
 End Function
'==========================================================================================
Function ProtectBody()

    On Error Resume Next
    
 Dim elmCnt, strTagName

 For elmCnt = 1 to frm1.length - 1
  If Left(frm1.elements(elmCnt).getAttribute("tag"),1) = "2" Then
   Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "Q")
  End If

  If Err.number <> 0 Then Err.Clear
 Next

End Function
'==========================================================================================
Function ReleaseBody()

    On Error Resume Next
    
 Dim elmCnt, strTagName

 For elmCnt = 1 to frm1.length - 1
  Select Case Left(frm1.elements(elmCnt).getAttribute("tag"),2)
  Case "21","25"
   Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "D")
  Case "22","23"
   Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "N")
  End Select

  If Err.number <> 0 Then Err.Clear
 Next

End Function
'============================================================================================================
Function ProtectXchRate()
	If frm1.txtCurrency.value = Parent.gCurrency Then
		 Call ggoOper.SetReqAttr(frm1.txtXchRate, "Q")
		 frm1.txtXchRate.text = 1
	Else
		 Call ggoOper.SetReqAttr(frm1.txtXchRate, "N")
		 frm1.txtXchRate.text = 0
	End If 
End Function
'============================================================================================================
Function JumpChgCheck(Byval pvIntCookieFlag, Byval pvStrJumpFlag)
 Dim IntRetCD

 If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
  'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
  If IntRetCD = vbNo Then Exit Function
 End If

 Call CookiePage(pvIntCookieFlag)
 Call PgmJump(pvStrJumpFlag)
End Function
'============================================================================================================
Function BtnSpreadCheck()

 BtnSpreadCheck = False

 Dim Answer
 <% '������ ������ ���� ���� ���� üũ��, YES�̸� �۾����࿩�� üũ ���Ѵ� %>
 If lgBlnFlgChgValue = True Then Answer = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")  <%'����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?%>
 If Answer = VBNO Then Exit Function

 <% '������ ������ �۾����࿩�� üũ %>
 If lgBlnFlgChgValue = False Then Answer = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x") <% '�۾��� �����Ͻðڽ��ϱ�? %> 
 If Answer = VBNO Then Exit Function

 BtnSpreadCheck = True

End Function
'========================================================================================================
 Sub ProtectGIRelITag()
  With frm1
   Call ggoOper.SetReqAttr(.txtTaxBizArea, "D")
   Call ggoOper.SetReqAttr(.txtPayer, "D")
   Call ggoOper.SetReqAttr(.txtToSalesGroup, "D")
   Call ggoOper.SetReqAttr(.txtPayType, "D")
  End With
 End Sub 
'========================================================================================================
 Sub ReleaseGIRelTag()
  With frm1
   Call ggoOper.SetReqAttr(.txtPayer, "N")
   Call ggoOper.SetReqAttr(.txtToSalesGroup, "N")
  End With   
 End Sub
'========================================================================================================
 Sub SetLocCurrency()
  frm1.txtLocCurrency.value = Parent.gCurrency
  frm1.txtLocCurrency1.value = Parent.gCurrency
 End Sub
'====================================================================================================
Sub CurFormatNumericOCX()

 With frm1
  'B/L �ݾ� 
  ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
  'B/L �ݾ� 
  ggoOper.FormatFieldByObjectOfCur .txtDocAmt1, .txtCurrency1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
  '�Ѽ��ݾ� 
  ggoOper.FormatFieldByObjectOfCur .txtMoney, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
  'ȯ�� 
  ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
  
 End With

End Sub
'========================================================================================================
 Sub Form_Load()
  Call LoadInfTB19029                <% '��: Load table , B_numeric_format %>
  Call AppendNumberPlace("6", "2", "0")
  Call AppendNumberPlace("7", "10", "0")
  Call AppendNumberPlace("8", "3", "0")
  Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
 
  Call ggoOper.LockField(Document, "N")           <% '��: Lock  Suitable  Field %>
  Call SetDefaultVal
  Call InitVariables
   
  <% '----------  Coding part  ------------------------------------------------------------- %>

  Call SetToolbar("11100000000011")          <% '��: ��ư ���� ���� %>
  Call CookiePage(0) 

  Call changeTabs(TAB1)

  gSelframeFlg = TAB1
  frm1.txtBLNo.focus
  Set gActiveElement = document.activeElement 
        gIsTab     = "Y"
        gTabMaxCnt = 3  
 End Sub
'========================================================================================================
 Sub Form_QueryUnload(Cancel, UnloadMode)
 End Sub
'========================================================================================================
 Sub btnBLNoOnClick()
  frm1.txtBLNo.focus 
  Call OpenExBLNoPop()
 End Sub
'========================================================================================================
 Sub btnLoadingPortOnClick()
  If frm1.txtLoadingPort.readOnly <> True Then
	frm1.txtLoadingPort.focus 
	Call OpenPort(frm1.txtLoadingPort.value, frm1.txtLoadingPortNm.value, "������", 0)
  End If
 End Sub
'========================================================================================================
 Sub btnDischgePortOnClick()
  If frm1.txtDischgePort.readOnly <> True Then
	frm1.txtDischgePort.focus 
	Call OpenPort(frm1.txtDischgePort.value, frm1.txtDischgePortNm.value, "������", 1)
  End If
 End Sub
'========================================================================================================
 Sub btnOriginOnClick()
  If frm1.txtOrigin.readOnly <> True Then
	frm1.txtOrigin.focus 
	Call OpenMinorCd(frm1.txtOrigin.value, frm1.txtOriginNm.value, "������", gstrOriginMajor)
  End If
 End Sub
'========================================================================================================
 Sub btnToSalesGroupOnClick()
  If frm1.txtToSalesGroup.readOnly <> True Then
	frm1.txtToSalesGroup.focus 
	Call OpenSalesGroup()
  End If
 End Sub
'========================================================================================================
 Sub btnTransportOnClick()
  If frm1.txtTransport.readOnly <> True Then
	frm1.txtTransport.focus 
	Call OpenMinorCd(frm1.txtTransport.value, frm1.txtTransportNm.value, "��۹��", gstrTransportMajor)
  End If
 End Sub
'========================================================================================================
 Sub btnFreightOnClick()
  If frm1.txtFreight.readOnly <> True Then
	frm1.txtFreight.focus 
	Call OpenMinorCd(frm1.txtFreight.value, frm1.txtFreightNm.value, "�������ҹ��", gstrFreightMajor)
  End If
 End Sub
'========================================================================================================
 Sub btnPackingTypeOnClick()
  If frm1.txtPackingType.readOnly <> True Then
	frm1.txtPackingType.focus 
	Call OpenMinorCd(frm1.txtPackingType.value, frm1.txtPackingTypeNm.value, "������", gstrPackingTypeMajor)
  End If
 End Sub
'========================================================================================================
 Sub btnVATTypeOnClick()
  If frm1.txtVatType.readOnly <> True Then
	frm1.txtVatType.focus 
	Call OpenMinorCd(frm1.txtVatType.value, frm1.txtVatTypeNm.value, "VAT����", gstrVATTypeMajor)
  End If
 End Sub
'========================================================================================================
 Sub btnVesselCntryOnClick()
  If frm1.txtVesselCntry.readOnly <> True Then
	frm1.txtVesselCntry.focus 
	Call OpenCountry(frm1.txtVesselCntry.value, "VESSEL")
  End If
 End Sub
'========================================================================================================
 Sub btnTranshipCntryOnClick()
  If frm1.txtTranshipCntry.readOnly <> True Then
	frm1.txtTranshipCntry.focus 
	Call OpenCountry(frm1.txtTranshipCntry.value, "TRANSHIP")
  End If
 End Sub
'========================================================================================================
 Sub btnOriginCntryOnClick()
  If frm1.txtOriginCntry.readOnly <> True Then
	frm1.txtOriginCntry.focus 
	Call OpenCountry(frm1.txtOriginCntry.value, "ORIGIN")
  End If
 End Sub
'========================================================================================================
 Sub btnAgentOnClick()
  If frm1.txtAgent.readOnly <> True Then
	frm1.txtAgent.focus 
	Call OpenBizPartner(frm1.txtAgent.value, frm1.txtAgentNm.value, "������")
  End If
 End Sub
'========================================================================================================
 Sub btnManufacturerOnClick()
  If frm1.txtManufacturer.readOnly <> True Then
	frm1.txtManufacturer.focus 
	Call OpenBizPartner(frm1.txtManufacturer.value, frm1.txtManufacturerNm.value, "������")
  End If
 End Sub
'========================================================================================================
 Sub btnForwarderOnClick()
  If frm1.txtForwarder.readOnly <> True Then
	frm1.txtForwarder.focus 
	Call OpenBizPartner(frm1.txtForwarder.value, frm1.txtForwarderNm.value, "����ȸ��")
  End If
 End Sub
'========================================================================================================
 Sub btnWeightUnitOnClick()
  If frm1.txtWeightUnit.readOnly <> True Then
	frm1.txtWeightUnit.focus 
	Call OpenUnit(frm1.txtWeightUnit.value, "WT", "�߷�����")
  End If
 End Sub
'========================================================================================================
 Sub btnVolumnUnitOnClick()
  If frm1.txtVolumnUnit.readOnly <> True Then
	frm1.txtVolumnUnit.focus 
	Call OpenUnit(frm1.txtVolumnUnit.value, "WD", "��������")
  End If
 End Sub
'========================================================================================================
 Sub btnTaxBizAreaOnClick()
  If frm1.txtTaxBizArea.readOnly <> True Then
	frm1.txtTaxBizArea.focus 
	Call OpenTaxOffice()
  End If
 End Sub
'========================================================================================================
Sub btnPayTypeOnClick()
     If frm1.txtPayType.readOnly <> True Then
        frm1.txtPayType.focus 
		Call OpenPayType()
     End If
End Sub
'==========================================================================================
 Sub txtBLIssueDt_DblClick(Button)
  If Button = 1 Then
	frm1.txtBLIssueDt.Action = 7
	Call SetFocusToDocument("M")   
	Frm1.txtBLIssueDt.Focus
  End If
 End Sub
 Sub txtLoadingDt_DblClick(Button)
  If Button = 1 Then
	frm1.txtLoadingDt.Action = 7
	Call SetFocusToDocument("M")   
	Frm1.txtLoadingDt.Focus
  End If
 End Sub
 Sub txtDischgeDt_DblClick(Button)
  If Button = 1 Then
	frm1.txtDischgeDt.Action = 7
	Call SetFocusToDocument("M")   
	Frm1.txtDischgeDt.Focus
  End If
 End Sub
 Sub txtTranshipDt_DblClick(Button)
  If Button = 1 Then
	frm1.txtTranshipDt.Action = 7
	Call SetFocusToDocument("M")   
	Frm1.txtTranshipDt.Focus
  End If
 End Sub
 Sub txtPayDt_DblClick(Button)
  If Button = 1 Then
	frm1.txtPayDt.Action = 7
	Call SetFocusToDocument("M")   
	Frm1.txtPayDt.Focus
  End If
 End Sub
'==========================================================================================
 Sub txtBLIssueDt_Change()
  If Trim(frm1.txtBLIssueDt.Text) = "" Then Exit Sub
  If frm1.txtCreditRot.value <> "0" and Trim(frm1.txtCreditRot.value) <> "" Then
   frm1.txtPayDt.Text = UNIDateAdd("d", frm1.txtCreditRot.value, Trim(frm1.txtBLIssueDt.Text), Parent.gDateFormat)
  Else
   'frm1.txtPayDt.Text = UniConvDateAToB("2999-12-31", Parent.gServerDateFormat, Parent.gDateFormat)
  End If

  lgBlnFlgChgValue = True
 End Sub

 Sub txtLoadingDt_Change()
  lgBlnFlgChgValue = True
 End Sub

 Sub txtDischgeDt_Change()
  lgBlnFlgChgValue = True
 End Sub

 Sub txtTranshipDt_Change()
  lgBlnFlgChgValue = True
 End Sub

 Sub txtPayDt_Change()
  lgBlnFlgChgValue = True
 End Sub

 Sub txtVatRate_Change()
  lgBlnFlgChgValue = True
 End Sub

 Sub txtXchRate_Change()
  lgBlnFlgChgValue = True
 End Sub 

 Sub txtPreRcptAmt_Change()
  lgBlnFlgChgValue = True
 End Sub

 Sub txtMoney_Change()
  lgBlnFlgChgValue = True
 End Sub 
 Sub txtBLIssueCnt_Change()
  lgBlnFlgChgValue = True
 End Sub  

 Sub txtTotPackingCnt_Change()
  lgBlnFlgChgValue = True
 End Sub  

 Sub txtContainerCnt_Change()
  lgBlnFlgChgValue = True
 End Sub  

 Sub txtGrossWeight_Change()
  lgBlnFlgChgValue = True
 End Sub  

 Sub txtGrossVolumn_Change()
  lgBlnFlgChgValue = True
 End Sub  
'========================================================================================================
 Sub rdoPostingflg1_OnPropertyChange()
  lgBlnFlgChgValue = True
 End Sub

 Sub rdoPostingflg2_OnPropertyChange()
  lgBlnFlgChgValue = True
 End Sub
'========================================================================================================
 Sub btnPosting_OnClick()
  If frm1.btnPosting.disabled <> True Then
   If BtnSpreadCheck = False Then Exit Sub
   Call PostBL()
  End If
 End Sub
'==========================================================================================
Sub btnGLView_OnClick()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent
 	
	If Trim(frm1.txtGLNo.value) <> "" Then
		 arrParam(0) = Trim(frm1.txtGLNo.value) 'ȸ����ǥ��ȣ 
		 
		 if arrParam(0) = "" THEN Exit Sub
		 
		 iCalledAspName = AskPRAspName("a5120ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If

		 arrRet = window.showModalDialog(iCalledAspName , Array(window.parent,arrParam), _
		      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		      
	ElseIf Trim(frm1.txtTempGLNo.value) <> "" Then
	     arrParam(0) = Trim(frm1.txtTempGLNo.value) '������ǥ��ȣ 
	     
	     if arrParam(0) = "" THEN Exit Sub
	     
	     iCalledAspName = AskPRAspName("a5130ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If
		 
	     arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else 
	     Call DisplayMsgBox("205154", "X", "X", "X")
	End If 
	     lblnWinEvent = False
End Sub
'==========================================================================================
Sub btnPreRcptView_OnClick()
 Dim iCalledAspName
 Dim arrRet
 Dim arrParam(4)
 
 If IsOpenPop = True Then Exit Sub

 IsOpenPop = True
 arrParam(0) = Trim(frm1.txtBLIssueDt.Text)   '������ 
 arrParam(1) = Trim(frm1.txtApplicant.value)   '������ 
 arrParam(2) = Trim(frm1.txtApplicantNm.value)  '������ 
 arrParam(3) = Trim(frm1.txtCurrency.value)   'ȭ�� 
 arrParam(4) = ""         '�����ݹ�ȣ 
 
iCalledAspName = AskPRAspName("s5111ra7")	
if Trim(iCalledAspName) = "" then
	IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5111ra7", "x")
	IsOpenPop = False
	exit sub
end if

 arrRet = window.showModalDialog(iCalledAspName & "?txtFlag=BL&txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,arrParam), _
       "dialogWidth=860px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 IsOpenPop = False
End Sub
'========================================================================================================
 Sub btnToBizAreaOnClick()
  If frm1.txtPayer.readOnly <> True Then
	frm1.txtPayer.focus 
	Call OpenBizPartner(frm1.txtPayer.value, frm1.txtPayerNm.value, "����ó")
  End If
 End Sub
'========================================================================================================
 Sub rdoVatCalcflg1_OnPropertyChange()
  lgBlnFlgChgValue = True
 End Sub

 Sub rdoVatCalcflg2_OnPropertyChange()
  lgBlnFlgChgValue = True
 End Sub
'========================================================================================================
 Function FncQuery()
  Dim IntRetCD

  FncQuery = False             <% '��: Processing is NG %>

  Err.Clear               <% '��: Protect system from crashing %>

  <% '------ Check previous data area ------ %>
  If lgBlnFlgChgValue = True Then
   IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")   <% '��: "Will you destory previous data" %>

   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  <% '------ Erase contents area ------ %>
  Call ggoOper.ClearField(Document, "2")        <% '��: Clear Contents  Field %>
  Call SetDefaultVal
  Call SetRadio()
  Call InitVariables             <% '��: Initializes local global variables %>

  
  <% '------ Check condition area ------ %>
  If Not chkField(Document, "1") Then       <% '��: This function check indispensable field %>
   Exit Function
  End If
  
  <% '------ Query function call area ------ %>
  
  Call DbQuery()              <% '��: Query db data %>

  FncQuery = True              <% '��: Processing is OK %>
 End Function
'========================================================================================================
 Function FncNew()
  Dim IntRetCD 

  FncNew = False                                                          <%'��: Processing is NG%>               <% '��: Protect system from crashing %>

  <% '------ Check previous data area ------ %>
  If lgBlnFlgChgValue = True Then
   IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "x", "x")

   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  Call ggoOper.ClearField(Document, "A")         <%'��: Clear Condition,Contents Field%>
  Call ggoOper.LockField(Document, "N")         <%'��: Lock  Suitable  Field%>
  Call SetDefaultVal
  Call InitVariables              <%'��: Initializes local global variables%>
  Call SetToolbar("11100000000011")          <% '��: ��ư ���� ���� %>
  
  Call ReleaseBody()
  Call changeTabs(TAB1)
  Call SetRadio()

  frm1.txtBLNo.focus
  Set gActiveElement = document.activeElement 
    	
  FncNew = True               <%'��: Processing is OK%>
 End Function
'========================================================================================================
 Function FncDelete()
  Dim IntRetCD

  FncDelete = False            <% '��: Processing is NG %>
  
  <% '------ Precheck area ------ %>
  If lgIntFlgMode <> Parent.OPMD_UMODE Then        <% 'Check if there is retrived data %>
   Call DisplayMsgBox("900002", "x", "x", "x")

   Exit Function
  End If

  IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "x", "x")

  If IntRetCD = vbNo Then
   Exit Function
  End If

  <% '------ Delete function call area ------ %>
  Call DbDelete             <% '��: Delete db data %>

  FncDelete = True            <% '��: Processing is OK %>
 End Function
'========================================================================================================
 Function FncSave()
  Dim IntRetCD
  
  FncSave = False            <% '��: Processing is NG %>
  
  Err.Clear              <% '��: Protect system from crashing %>
  
  <% '------ Precheck area ------ %>
  If lgBlnFlgChgValue = False Then        <% 'Check if there is retrived data %>
      IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     <% '��: No data changed!! %>
'      Call MsgBox("No data changed!!", vbInformation)
      Exit Function
  End If
  
  <% '------ Check contents area ------ %>
  If Not chkField(Document, "2") Then        <% '��: Check contents area %>
   <% ' Required Field Check�� Error�߻��� Tab �̵��� �̵��� tab page ��ȣ�� 
   ' gSelframeFlg(tab page flag)���� �Ѱ��ݴϴ�. %>
      If gPageNo > 0 Then
          gSelframeFlg = gPageNo
      End If
      Exit Function
  End If   

  '** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� Ŀ�� �Ҷ� **
  If Len(Trim(frm1.txtBLIssueDt.Text)) Then
   If UniConvDateToYYYYMMDD(frm1.txtLoadingDt.Text, Parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtBLIssueDt.Text, Parent.gDateFormat, "-") Then
    Call ClickTab1()
    Call DisplayMsgBox("970023", "x", frm1.txtBLIssueDt.Alt, frm1.txtLoadingDt.Alt)
    'MsgBox "pObjToDt(��)�� pObjFromDt���� ũ�ų� ���ƾ� �մϴ�.", vbExclamation, "uniERP(Warning)"
    frm1.txtBLIssueDt.Focus
    Set gActiveElement = document.activeElement 
    Exit Function
   End If
  End If

  If Len(Trim(frm1.txtTranshipDt.Text)) And Len(Trim(frm1.txtLoadingDt.Text)) Then
   If UniConvDateToYYYYMMDD(frm1.txtLoadingDt.Text, Parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtTranshipDt.Text, Parent.gDateFormat, "-") Then
    Call ClickTab2()
    Call DisplayMsgBox("970023", "x", frm1.txtTranshipDt.Alt, frm1.txtLoadingDt.Alt)
    'MsgBox "pObjToDt(��)�� pObjFromDt���� ũ�ų� ���ƾ� �մϴ�.", vbExclamation, "uniERP(Warning)"
    frm1.txtTranshipDt.Focus
    Set gActiveElement = document.activeElement 
    Exit Function
   End If
  End If

  If Len(Trim(frm1.txtDischgeDt.Text)) And Len(Trim(frm1.txtTranshipDt.Text)) Then
   If UniConvDateToYYYYMMDD(frm1.txtTranshipDt.Text, Parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtDischgeDt.Text, Parent.gDateFormat, "-") Then
    Call ClickTab2()
    Call DisplayMsgBox("970023", "x", frm1.txtDischgeDt.Alt, frm1.txtTranshipDt.Alt)
    'MsgBox "pObjToDt(��)�� pObjFromDt���� ũ�ų� ���ƾ� �մϴ�.", vbExclamation, "uniERP(Warning)"
    frm1.txtDischgeDt.Focus
    Set gActiveElement = document.activeElement 
    Exit Function
   End If
  End If

  If Len(Trim(frm1.txtBLIssueDt.Text)) And Len(Trim(frm1.txtPayDt.Text)) Then
   If UniConvDateToYYYYMMDD(frm1.txtBLIssueDt.Text, Parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtPayDt.Text, Parent.gDateFormat, "-") Then
    Call ClickTab3()
    Call DisplayMsgBox("970023", "x", frm1.txtPayDt.Alt, frm1.txtBLIssueDt.Alt)
    'MsgBox "pObjToDt(��)�� pObjFromDt���� ũ�ų� ���ƾ� �մϴ�.", vbExclamation, "uniERP(Warning)"
    frm1.txtPayDt.Focus
    Set gActiveElement = document.activeElement 
    Exit Function
   End If
  End If

  <% '------ Save function call area ------ %>
  Call DbSave              <% '��: Save db data %>
  
  FncSave = True             <% '��: Processing is OK %>
 End Function
'========================================================================================================
 Function FncCopy()
  Dim IntRetCD

  If lgBlnFlgChgValue = True Then
   IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")   <%'��: "Will you destory previous data"%>
'   IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)

   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  lgIntFlgMode = Parent.OPMD_CMODE             <%'��: Indicates that current mode is Crate mode%>

  <% '------ ���Ǻ� �ʵ带 �����Ѵ�. ------ %>
  Call ggoOper.ClearField(Document, "1")          <%'��: Clear Condition Field%>
  Call ggoOper.LockField(Document, "N")          <%'��: This function lock the suitable field%>
  frm1.txtBLNo1.value = "" 
  lgBlnFlgChgValue = True 
 End Function
'========================================================================================================
 Function FncCancel() 
  On Error Resume Next              <%'��: Protect system from crashing%>
 End Function
'========================================================================================================
 Function FncInsertRow()
  On Error Resume Next              <%'��: Protect system from crashing%>
 End Function
'========================================================================================================
 Function FncDeleteRow()
  On Error Resume Next              <%'��: Protect system from crashing%>
 End Function
'========================================================================================================
 Function FncPrint()
  Call parent.FncPrint()
 End Function
'========================================================================================
Function FncPrev() 
    Dim strVal
 Dim IntRetCD
 
 If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")   <% '��: "Will you destory previous data" %>

  If IntRetCD = vbNo Then
   Exit Function
  End If
 End If
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "x", "x", "x")  '�� �ٲ�κ� 
        'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
        Exit Function
    End If

 If LayerShowHide(1) = False Then
  Exit Function
 End If

 frm1.txtPrevNext.value = "P"

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001       <%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo1.value)    <%'��: ��ȸ ���� ����Ÿ %>
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)  <%'��: ��ȸ ���� ����Ÿ %>
         
 Call RunMyBizASP(MyBizASP, strVal)
End Function
'========================================================================================
Function FncNext() 
    Dim strVal
 Dim IntRetCD
 
 If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")   <% '��: "Will you destory previous data" %>

  If IntRetCD = vbNo Then
   Exit Function
  End If
 End If
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "x", "x", "x")  '�� �ٲ�κ� 
        'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
        Exit Function
    End If

 If LayerShowHide(1) = False Then
  Exit Function
 End If

 frm1.txtPrevNext.value = "N"

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001       <%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo1.value)    <%'��: ��ȸ ���� ����Ÿ %>
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)  <%'��: ��ȸ ���� ����Ÿ %>
         
 Call RunMyBizASP(MyBizASP, strVal)
End Function
'========================================================================================================
 Function FncExcel() 
  Call parent.FncExport(Parent.C_SINGLE)
 End Function
'========================================================================================================
 Function FncFind() 
  Call parent.FncFind(Parent.C_SINGLE, True)
 End Function
'========================================================================================================
 Function FncExit()
  Dim IntRetCD

  FncExit = False

  If lgBlnFlgChgValue = True Then
   IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")   <%'��: "Will you destory previous data"%>

'   IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  FncExit = True
 End Function
'********************************************************************************************************
 Function DbQuery()
  Err.Clear               <%'��: Protect system from crashing%>

  DbQuery = False              <%'��: Processing is NG%>

  Dim strVal

  If LayerShowHide(1) = False Then
   Exit Function
  End If

  frm1.txtPrevNext.value = "Q"
  
  strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001     <%'��: �����Ͻ� ó�� ASP�� ���� %>
  strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)   <%'��: ��ȸ ���� ����Ÿ %>
     strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)  <%'��: ��ȸ ���� ����Ÿ %>

  Call RunMyBizASP(MyBizASP, strVal)         <%'��: �����Ͻ� ASP �� ���� %>
 
  DbQuery = True              <%'��: Processing is NG%>
 End Function
'========================================================================================================
 Function DbSave()
  Err.Clear               <%'��: Protect system from crashing%>

  DbSave = False              <%'��: Processing is NG%>

  Dim strVal

  If frm1.chkSONoFlg.checked = True Then
   frm1.txtSoNoFlg.value = "Y"
  Else 
   frm1.txtSoNoFlg.value = "N"
  End If 

  If LayerShowHide(1) = False Then
   Exit Function
  End If

  With frm1
   .txtMode.value = Parent.UID_M0002          <%'��: �����Ͻ� ó�� ASP �� ���� %>
   .txtFlgMode.value = lgIntFlgMode
   .txtUpdtUserId.value = Parent.gUsrID
   .txtInsrtUserId.value = Parent.gUsrID

   ReleaseTag(.rdoPostingflg1)
   ReleaseTag(.rdoPostingflg2)

   .rdoPostingflg1.className = "RADIO"
   .rdoPostingflg2.className = "RADIO"

   Call ExecMyBizASP(frm1, BIZ_PGM_ID)

   ProtectTag(.rdoPostingflg1)
   ProtectTag(.rdoPostingflg2)

   .rdoPostingflg1.className = "RADIO"
   .rdoPostingflg2.className = "RADIO"
  End With

  DbSave = True              <%'��: Processing is NG%>
 End Function
'========================================================================================================
 Function DbDelete()
  Err.Clear               <%'��: Protect system from crashing%>

  DbDelete = False             <%'��: Processing is NG%>

  Dim strVal

  If LayerShowHide(1) = False Then
   Exit Function
  End If

  strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003     <%'��: �����Ͻ� ó�� ASP�� ���� %>
  strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo1.value)   <%'��: ���� ���� ����Ÿ %>
  Call RunMyBizASP(MyBizASP, strVal)         <%'��: �����Ͻ� ASP �� ���� %>

  DbDelete = True              <%'��: Processing is NG%>
 End Function
'========================================================================================================
 Function DbQueryOk()             <% '��: ��ȸ ������ ������� %>
  <% '------ Reset variables area ------ %>
  lgIntFlgMode = Parent.OPMD_UMODE           <% '��: Indicates that current mode is Update mode %>
 
  Call ggoOper.LockField(Document, "Q")        <% '��: This function lock the suitable field %>
  Call SetToolbar("111110001101111")
  
  frm1.txtLocCurrency.value = Parent.gCurrency
  frm1.txtLocCurrency1.value = Parent.gCurrency

  If frm1.rdoPostingflg1.checked = True Then
   Call ProtectBody()
  ElseIf frm1.rdoPostingflg2.checked = True Then
   Call ReleaseBody()
  End If   
  
  If frm1.txtRefFlg.value = "M" Then 
   frm1.btnPosting.disabled = True 
   Call ProtectGIRelITag()
  Else  
   If CInt(frm1.txtStatusFlg.value) < 3 Then 
    frm1.btnPosting.disabled = False
   Else 
    frm1.btnPosting.disabled = True 
   End If  
  End If 
  '���ݸ����� 
  if UniConvDateToYYYYMMDD(frm1.txtPayDt.Text,Parent.gDateFormat,"-") = "2999-12-31" then
   frm1.txtPayDt.Text = ""    
  end if
  
  If Len(Trim(frm1.txtSONo.value)) then frm1.chkSONoFlg1.checked = True
  
  Call ggoOper.SetReqAttr(frm1.txtBLNo1, "Q")
  Call ggoOper.SetReqAttr(frm1.chkSONoFlg, "Q")
  
	If frm1.txtCurrency.value = Parent.gCurrency Then
		Call ggoOper.SetReqAttr(frm1.txtXchRate, "Q")
	Else
		If frm1.rdoPostingflg2.checked = True Then
			Call ggoOper.SetReqAttr(frm1.txtXchRate, "N")
		End If
	End If 

  lgBlnFlgChgValue = False
 End Function
'========================================================================================================
 Function ReferenceQueryOk()             <% '��: ��ȸ ������ ������� %>
  Call SetToolbar("11101000000011") 
  Call SetLocCurrency()
  Call ProtectXchRate()

  If frm1.txtCreditRot.value <> "0" and Trim(frm1.txtCreditRot.value) <> "" Then
   frm1.txtPayDt.Text = UNIDateAdd("d", frm1.txtCreditRot.value, Trim(frm1.txtBLIssueDt.Text), Parent.gDateFormat)
  Else
'   frm1.txtPayDt.Text = UniConvDateAToB("2999-12-31", Parent.gServerDateFormat, Parent.gDateFormat)
  End If
  
  '���ݽŰ����� Default �� ���� 
  Call GetTaxBizArea("*")
  
  If frm1.txtRefFlg.value = "M" Then 
   Call ProtectGIRelITag()
  Else
   Call ReleaseGIRelTag()
  End If 

  frm1.btnPosting.disabled = True 
 End Function
'========================================================================================================
 Function DbSaveOk()              <%'��: ���� ������ ���� ���� %>
  Call InitVariables
  Call MainQuery()
 End Function
'========================================================================================================
 Function DbDeleteOk()             <%'��: ���� ������ ���� ���� %>
  Call InitVariables
  Call MainNew()
 End Function
</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
 <FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
  <TABLE <%=LR_SPACE_TYPE_00%>>
   <TR>
    <TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' ���� ���� %></TD>
   </TR>
   <TR HEIGHT=23>
    <TD WIDTH=100%>
     <TABLE <%=LR_SPACE_TYPE_10%>>
      <TR>
       <TD WIDTH=10>&nbsp;</TD>
       <TD CLASS="CLSSTABP">
        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
         <TR>
          <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
          <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>��������</font></td>
          <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
            </TR>
        </TABLE>
       </TD>
       <TD CLASS="CLSSTABP">
        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
         <TR>
          <td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
          <td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>������Ÿ</font></td>
          <td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
            </TR>
        </TABLE>
       </TD>
       <TD CLASS="CLSSTABP">
        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
         <TR>
          <td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
          <td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>����ä������</font></td>
          <td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
            </TR>
        </TABLE>
       </TD>
       <TD WIDTH=* align=right><A href="vbscript:OpenSORef">��������</A>&nbsp;|&nbsp;<A href="vbscript:OpenLCRef">L/C����</A>&nbsp;|&nbsp;<A href="vbscript:OpenCCRef">�������</A></TD>
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
           <TD CLASS=TD5 NOWRAP>B/L ������ȣ</TD>
           <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="B/L ������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBLNo" ALIGN=top TYPE="BUTTON"ONCLICK ="vbscript:btnBLNoOnClick()"></TD>
           <TD CLASS=TDT NOWRAP></TD>
           <TD CLASS=TD6 NOWRAP></TD>
          </TR>
         </TABLE>
        </FIELDSET>
       </TD>
      </TR>
      <TR>
       <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
      </TR>
      <TR>
       <TD WIDTH=100% VALIGN=TOP>
       <!-- ù��° �� ���� -->
        <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
            <TABLE <%=LR_SPACE_TYPE_60%>>
             <TR>
              <TD CLASS=TD5 NOWRAP>B/L ������ȣ</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo1" SIZE=20 MAXLENGTH=18 TAG="25XXXU" ALT="B/L ������ȣ"></TD>
              <TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
              <TD CLASS=TD6 NOWRAP>
               <INPUT NAME="txtSONo" TYPE=TEXT SIZE=20 MAXLENGTH=18 TAG="24XXXU" ALT="���ֹ�ȣ">&nbsp;&nbsp;&nbsp;
               <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="25X" VALUE="Y" NAME="chkSONoFlg" ID="chkSONoFlg1">
               <LABEL FOR="chkSONoFlg">���ֹ�ȣ����</LABEL>
              </TD> 
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>B/L��ȣ</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLDocNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="21XXXU" ALT="B/L��ȣ"></TD>
              <TD CLASS=TD5 NOWRAP>L/C��ȣ</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="24XXXU" ALT="L/C��ȣ" >&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtBLIssueDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="������"></OBJECT>');</SCRIPT></TD>
              <TD CLASS=TD5 NOWRAP>B/L�ݾ�</TD>
              <TD CLASS=TD6 NOWRAP>
               <TABLE CELLSPACING=0 CELLPADDING=0> 
                <TR>
                 <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtDocAmt" CLASS=FPDS140 tag="24X2" ALT="B/L�ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
                 <TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��">
                </TR>
               </TABLE>
              </TD>
             </TR>
             <TR> 
              <TD CLASS=TD5 NOWRAP>ȯ��</TD>
              <TD CLASS=TD6 NOWRAP>
               <TABLE CELLSPACING=0 CELLPADDING=0>
                <TR>
                 <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchRate" CLASS=FPDS140 tag="22X5Z" ALT="ȯ��" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
                </TR>
               </TABLE>
              </TD>
              <TD CLASS=TD5 NOWRAP>B/L�ڱ��ݾ�</TD>
              <TD CLASS=TD6 NOWRAP>
               <TABLE CELLSPACING=0 CELLPADDING=0> 
                <TR>
                 <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtLocAmt" CLASS=FPDS140 tag="24X2Z" ALT="B/L�ڱ��ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
                 <TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��">
                </TR>
               </TABLE>
              </TD>
             </TR>
             <TR> 
              <TD CLASS=TD5 NOWRAP>��۹��</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="��۹��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTransportOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=25 TAG="24"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingPort" ALT="������" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnLoadingPortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>��������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="��������">&nbsp;<INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=25 TAG="24"></TD></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgePort" ALT="������" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnDischgePortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>�����׷�</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="�����׷�">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=25 TAG="24"></TD>
             </TR>
             <TR>     
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtLoadingDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="������"></OBJECT>');</SCRIPT></TD>
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=25 TAG="24"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>�������ҹ��</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFreight" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="�������ҹ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFreight" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnFreightOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtFreightNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>B/L�������</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtBLIssueCnt" CLASS=FPDS65 tag="21X6Z" ALT="B/L�������" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>B/L�������</TD>
              <TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtBLIssuePlce" ALT="B/L�������" TYPE=TEXT MAXLENGTH=50 SIZE=84 TAG="21X"></TD>
             </TR>
            <%Call SubFillRemBodyTD5656(9)%>
            </TABLE>
           </DIV> 
           <!-- �ι�° �� ���� -->
           <DIV ID="TabDiv" SCROLL=no>
            <TABLE <%=LR_SPACE_TYPE_60%>>
             <TR>
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnAgentOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnManufacturerOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>VESSEL��</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselNm" ALT="VESSEL��" TYPE=TEXT MAXLENGTH=50 SIZE=35 TAG="21X"></TD>
              <TD CLASS=TD5 NOWRAP>������ȣ</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVoyageNo"ALT="������ȣ" MAXLENGTH=20 SIZE=35 TAG="21X"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>����ȸ��</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtForwarder" SIZE=10 MAXLENGTH=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnForwarder" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnForwarderOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtForwarderNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>���ڱ���</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVesselCntry" SIZE=10 MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVesselCntry" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnVesselCntryOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtVesselCntryNm" SIZE=20 TAG="24"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>�������</TD>
              <TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtReceiptPlce" ALT="�������" TYPE=TEXT MAXLENGTH=50 SIZE=84 TAG="21X"></TD>
             </TR>
             <TR> 
              <TD CLASS=TD5 NOWRAP>�ε����</TD>
              <TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtDeliveryPlce" ALT="�ε����" TYPE=TEXT MAXLENGTH=120 SIZE=84 TAG="21X"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>����������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtFinalDest" ALT="����������" TYPE=TEXT MAXLENGTH=50 SIZE=35 TAG="21X"></TD>
              <TD CLASS=TD5 NOWRAP>����������</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 NAME="txtDischgeDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="����������"></OBJECT>');</SCRIPT></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>ȯ������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTranshipCntry" SIZE=10 MAXLENGTH=3 SIZE=10 TAG="21XXXU" ALT="ȯ������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTranshipCntry" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTranshipCntryOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtTranshipCntryNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>ȯ����</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime5 NAME="txtTranshipDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="ȯ����"></OBJECT>');</SCRIPT></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>��������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPackingType" SIZE=10 MAXLENGTH=5 TAG="21XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPackingType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnPackingTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPackingTypeNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>�����尹��</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtTotPackingCnt" CLASS=FPDS65 tag="21X7Z" ALT="�����尳��" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
             </TR>
             <TR> 
              <TD CLASS=TD5 NOWRAP>�����������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPackingTxt" ALT="�����������" TYPE=TEXT MAXLENGTH=50 SIZE=34 TAG="21X"></TD>
              <TD CLASS=TD5 NOWRAP>�����̳ʼ�</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtContainerCnt" CLASS=FPDS65 tag="21X8Z" ALT="�����̳ʼ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>���߷�</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtGrossWeight" CLASS=FPDS140 tag="21X3Z" ALT="���߷�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
              <TD CLASS=TD5 NOWRAP>�߷�����</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtWeightUnit" ALT="�߷�����" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWeightUnit" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnWeightUnitOnClick()"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>�ѿ���</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtGrossVolumn" CLASS=FPDS140 tag="21X3Z" ALT="�ѿ���" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
              <TD CLASS=TD5 NOWRAP>��������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtVolumnUnit" ALT="��������" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVolumnUnit" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnVolumnUnitOnClick()"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="������" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnOriginOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>����������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="����������" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnOriginCntryOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginCntryNm" SIZE=20 TAG="24"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>�����������</TD>
              <TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtFreightPlce" ALT="�����������" TYPE=TEXT MAXLENGTH=30 SIZE=84 TAG="21X"></TD>
             </TR>
            <%Call SubFillRemBodyTD5656(5)%>
            </TABLE>
           </DIV>
           <!-- ����° �� ���� -->
           <DIV ID="TabDiv" SCROLL=no>
            <TABLE <%=LR_SPACE_TYPE_60%>>
             <TR>
              <TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizArea" ALT="���ݽŰ�����" TYPE=TEXT MAXLENGTH=10 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBizArea" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTaxBizAreaOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>����ä������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillType" ALT="����ä������" TYPE=TEXT MAXLENGTH=20 SIZE=10 TAG="24X">&nbsp;<INPUT TYPE=TEXT NAME="txtBillTypeNm" SIZE=25 TAG="24"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>����ó</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayer" ALT="����ó" TYPE=TEXT MAXLENGTH=10 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToBizArea" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnToBizAreaOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPayerNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>����ó</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBilltoParty" ALT="����ó" TYPE=TEXT MAXLENGTH=10 SIZE=10 TAG="24X">&nbsp;<INPUT TYPE=TEXT NAME="txtBilltoPartyNm" SIZE=25 TAG="24"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>���ݿ����׷�</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtToSalesGroup" ALT="���ݿ����׷�" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnToSalesGroupOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtToSalesGroupNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>Ȯ������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="24X" VALUE="Y" ID="rdoPostingflg1"><LABEL FOR="rdoPostingflg1">Ȯ��</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" VALUE="N" TAG="24X" CHECKED ID="rdoPostingflg2"><LABEL FOR="rdoPostingflg2">��Ȯ��</LABEL></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>���ݸ�����</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime8 NAME="txtPayDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="���ݸ�����"></OBJECT>');</SCRIPT></TD>
              <TD CLASS=TD5 NOWRAP>B/L�ݾ�</TD>
              <TD CLASS=TD6 NOWRAP>
               <TABLE CELLSPACING=0 CELLPADDING=0> 
                <TR>
                 <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtDocAmt1" CLASS=FPDS140 tag="24X2Z" ALT="B/L�ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
                 <TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency1" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��"></TD>
                </TR>
               </TABLE>
              </TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>�Ա�����</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayType" ALT="�Ա�����" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnPayTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTypeNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>�Ѽ��ݾ�</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtMoney" CLASS=FPDS140 tag="24X2Z" ALT="�Ѽ��ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
             </TR> 
             <TR> 
              <TD CLASS=TD5 NOWRAP>�������</TD>
              <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="�������">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
              <TD CLASS=TD5 NOWRAP>B/L�ڱ��ݾ�</TD>
              <TD CLASS=TD6 NOWRAP>
               <TABLE CELLSPACING=0 CELLPADDING=0> 
                <TR>
                 <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtLocAmt1" CLASS=FPDS140 tag="24X2Z" ALT="B/L�ڱ��ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
                 <TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency1" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="�ڱ�ȭ��"></TD>
                </TR>
               </TABLE>
              </TD>              
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>�����Ⱓ</TD>
              <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayDur" ALT="�����Ⱓ" STYLE="TEXT-ALIGN: right" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24X7">&nbsp;��</TD>             
              <TD CLASS=TD5 NOWRAP>�Ѽ����ڱ���</TD>
              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtCollectLocAmt" CLASS=FPDS140 tag="24X2Z" ALT="�Ѽ����ڱ���" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>

             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>��ݰ�����������</TD>
              <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtPayTermstxt" ALT="��ݰ�������" TYPE=TEXT MAXLENGTH=120 SIZE=84 TAG="21X"></TD>
             </TR>
             <TR>
              <TD CLASS=TD5 NOWRAP>���</TD>
              <TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark" ALT="���" TYPE=TEXT MAXLENGTH=120 SIZE=84 TAG="21X"></TD>
             </TR>
            <%Call SubFillRemBodyTD5656(11)%>
            </TABLE>
        </DIV>
       </TD>
      </TR>
     </TABLE> 
    </TD>
   </TR>  
   <TR HEIGHT=20>
    <TD WIDTH=100%>
     <TABLE <%=LR_SPACE_TYPE_30%>>
      <TD><BUTTON NAME="btnPosting" CLASS="CLSMBTN">Ȯ��</BUTTON>&nbsp;
       <BUTTON NAME="btnGLView" CLASS="CLSMBTN">��ǥ��ȸ</BUTTON>&nbsp;
       <BUTTON NAME="btnPreRcptView" CLASS="CLSMBTN">��������Ȳ</BUTTON></TD>
      <TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(1, EXBL_DETAIL_ENTRY_ID)">B/L�������</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1, BIZ_BillCollect_JUMP_ID)">B/L���ݳ������</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(2, EXPORT_CHARGE_ENTRY_ID)">�ǸŰ����</A></TD>
      <TD WIDTH=10>&nbsp;</TD>
     </TABLE>
    </TD>
   </TR>
   <TR>
    <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
   </TR>
  </TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLCNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCCNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHBLNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtRefFlg" TAG="24"> 
<INPUT TYPE=HIDDEN NAME="txtPrevNext" TAG="24"> 
<INPUT TYPE=HIDDEN NAME="txtStatusFlg" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtCreditRotDay" TAG="24">  
<INPUT TYPE=HIDDEN NAME="txtVatIncflag" TAG="24">  
<INPUT TYPE=HIDDEN NAME="txtVatType" TAG="24">  
<INPUT TYPE=HIDDEN NAME="txtVatTypeNm" TAG="24">  
<INPUT TYPE=HIDDEN NAME="txtVatRate" TAG="24">  
<INPUT TYPE=HIDDEN NAME="txtSoNoFlg" TAG="24">  
<INPUT TYPE=HIDDEN NAME="txtGLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCreditRot" tag="24">
<INPUT TYPE=HIDDEN NAME="txtTempGLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBatchNo" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
