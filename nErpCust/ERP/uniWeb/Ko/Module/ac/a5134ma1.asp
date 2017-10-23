<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5134ma1
'*  4. Program Name         : ���� ��ǥ ��ȸ 
'*  5. Program Desc         : Query of Base Slip
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/25
'*  8. Modified date(Last)  : Kim Sang Joong
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
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
<!-- #Include file="../../inc/IncServer.asp" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">                    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Common.vbs">                 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">                 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Event.vbs">                  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Variables.vbs">              </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Operation.vbs">              </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/AdoQuery.vbs">               </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgent.vbs">          </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">				  </SCRIPT>

<Script Language="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID		= "a5134Mb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'Const C_SHEETMAXROWS    = 30										'��: Spread sheet���� �������� row
'Const C_SHEETMAXROWS_D  = 100                                       '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_SHEETMAXROWS_D  = 30                                       '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 3					                        '��: SpreadSheet�� Ű�� ���� 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          

Dim lgSelectList											'��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT											'��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 


Dim lgSortFieldNm											'��: Orderby popup�� ����Ÿ(�ʵ弳��)                                        
Dim lgSortFieldCD											'��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)                         

Dim lgMaxFieldCount

Dim lgPopUpR												'��: Orderby default ��                                                              

Dim lgKeyPos                                              
Dim lgKeyPosVal                                         
Dim lgCookValue 

Dim lgSaveRow												

Dim IsOpenPop												'��: Popup status                           


Dim lgTypeCD                                                '��: 'G' is for group , 'S' is for Sort    
Dim lgFieldCD                                               '��: �ʵ� �ڵ尪                           
Dim lgFieldNM                                               '��: �ʵ� ����                           
Dim lgFieldLen                                              '��: �ʵ� ��(Spreadsheet����)              
Dim lgFieldType                                             '��: �ʵ� ����                           
Dim lgDefaultT                                              '��: �ʵ� �⺻��                           
Dim lgNextSeq                                               '��: �ʵ� Pair��                           
Dim lgKeyTag                                                '��: Key  ����                             

Dim lgMark                                                  '��: ��ũ 
Dim strDateYr
Dim strDateMonth
Dim strDateDay        
'�ʿ���� �� ���� ���� �� 

                          
<%
'--------------- ������ coding part(�������,Start)-----------------------------------------------------------

  Call GetAdoFiledInf("A5134MA1","S", "A")						'��: spread sheet �ʵ����� query   -----
																' G is for Qroup , S is for Sort
																' A is spreadsheet No
'--------------- ������ coding part(�������,End)-------------------------------------------------------------
%>


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgPageNo	     = ""
    lgIntFlgMode     = OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
    lgSaveRow        = 0

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim ii,kk	
	Dim iCast
	
    lgTypeCD    = Split ("<%=gTypeCD%>"   ,Chr(11))                                 '  �ʵ� ��          
    lgFieldCD   = Split ("<%=gFieldCD%>"  ,Chr(11))                                 '  �ʵ� �ڵ尪      
    lgFieldNM   = Split ("<%=gFieldNM%>"  ,Chr(11))                                 '  �ʵ� ����      
    lgFieldLen  = Split ("<%=gFieldLen%>" ,Chr(11))                                 '  �ʵ� ��          
    lgFieldType = Split ("<%=gFieldType%>",Chr(11))                                 '  �ʵ� ����Ÿ Ÿ�� 
    lgDefaultT  = Split ("<%=gDefaultT%>" ,Chr(11))                                 '  �ʵ� �⺻��      
    lgNextSeq   = Split ("<%=gNextSeq%>"  ,Chr(11))                                 '  �ʵ� Pair��      
    lgKeyTag    = Split ("<%=gKeyTag%>"   ,Chr(11))                                 '  Key����          
    
    lgSortFieldNm   = ""
    lgSortFieldCD  = ""
    Redim  lgMark(UBound(lgFieldNM)) 
    
    For ii = 0 To UBound(lgFieldNM) - 1                                            'Sort ��󸮽�Ʈ   ���� 
        iCast = lgDefaultT(ii)
        If  IsNumeric(iCast) Or Trim(lgDefaultT(ii)) = "V" Then
            If IsNumeric(iCast) Then 
               If IsBetween(1,C_MaxSelList,CInt(iCast)) Then    'Sort����default�� ���� 
                  lgPopUpR(CInt(lgDefaultT(ii)) - 1,0) = Trim(lgFieldCD(ii))
                  lgPopUpR(CInt(lgDefaultT(ii)) - 1,1) = "ASC"
               End If
            End If
            lgSortFieldNm   = lgSortFieldNm   & Trim(lgFieldNM (ii)) & Chr(11)
            lgSortFieldCD  = lgSortFieldCD  & Trim(lgFieldCD(ii)) & Chr(11)
        End If
    Next
   
    lgSortFieldNm    = split (lgSortFieldNm ,Chr(11))
    lgSortFieldCD    = split (lgSortFieldCD,Chr(11))

'--------------- ������ coding part(�������,Start)--------------------------------------------------

	<%
	Dim strYear, strMonth, strDay, dtToday,  StartDate
	StartDate = GetSvrDate
	Call ExtractDateFrom(StartDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)
	StartDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)	
	%>
	
	frm1.fpDateYr.Text = "<%=StartDate %>"
	Call ggoOper.FormatDate(frm1.txtDateYr,  gDateFormat, 3)
	frm1.txtDateYr.focus

'--------------- ������ coding part(�������,End)-----------------------------------------------------

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029(gCurrency, "Q", "A") %>
End Sub

Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						 'Cookie Split String : CookiePage Function Use

	If Kubun = 1 Then								 'Jump�� ȭ���� �̵��� ��� 

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie "PoNo" , lsPoNo					 'Jump�� ȭ���� �̵��Ҷ� �ʿ��� Cookie �������� 
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							 'Jump�� ȭ���� �̵��� ������� 

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		Dim iniSep

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call FncQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Dim ii,jj,kk,iSeq

    lgSelectList   = ""
    lgSelectListDT = ""
    iSeq           = 0 

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


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheetRow
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
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
   
'	Spreadsheet #2�˻��� ���� Ű ����ġ ���� 
   If CInt(lgKeyTag(iDx)) > 0 And CInt(lgKeyTag(iDx)) <= C_MaxKey Then  
      lgKeyPos(CInt(lgKeyTag(iDx))) =  iCol
   End If

End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
	ggoSpread.SpreadLock 1 , -1
    .vspdData.ReDraw = True

    End With
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 

Sub InitComboBox()
<%	
	Dim arrData
	
'	arrData = InitCombo("F3011", "frm1.cboDpstFg")
'	arrData = InitCombo("F3014", "frm1.cboTransSts")
%>
End Sub
 
<%
Function InitCombo(ByVal strMajorCd, ByVal objCombo)

    Dim pB1a028
    Dim intMaxRow
    Dim intLoopCnt
    Dim strCodeList
    Dim strNameList
        
    Err.Clear                                                               '��: Clear error no
	On Error Resume Next

	Set pB1a028 = Server.CreateObject("B1a028.B1a028ListMinorCode")
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set pB1a028 = Nothing												'��: ComProxy Unload
		Call MessageBox(Err.description, I_INSCRIPT)						'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	pB1a028.ImportBMajorMajorCd = Trim(strMajorCd)									'��: Major Code
    pB1a028.ServerLocation = ggServerIP
    pB1a028.ComCfg = gConnectionString
    pB1a028.Execute															'��:
    
    '-----------------------
    'Com action result check area(DB,internal)
    '-----------------------
    If Not (pB1a028.OperationStatusMessage = MSG_OK_STR) Then
		Call MessageBox(pB1a028.OperationStatusMessage, I_INSCRIPT)         '��: you must release this line if you change msg into code
		Set pB1a028 = Nothing												'��: ComProxy Unload
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
    End If

	intMaxRow = pB1a028.ExportGroupCount
	
	For intLoopCnt = 1 To intMaxRow
%>
		Call SetCombo(<%=objCombo%>, "<%=pB1a028.ExportItemBMinorMinorCd(intLoopCnt)%>", "<%=pB1a028.ExportItemBMinorMinorNm(intLoopCnt)%>")		'��: InitCombo ���� �ؾ� �Ǵµ� �ӽ÷� ���� ���� 
<%
		strCodeList = strCodeList & vbtab & pB1a028.ExportItemBMinorMinorCd(intLoopCnt)
		strNameList = strNameList & vbtab & pB1a028.ExportItemBMinorMinorNm(intLoopCnt)
	Next
	
	InitCombo = Array(strCodeList, strNameList)
		
	Set pB1a028 = Nothing

End Function
%>

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

    										'��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)

    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

    ReDim lgPopUpR(C_MaxSelList - 1,1)
    
	Call InitVariables													'��: Initializes local global variables

	Call SetDefaultVal	

	Call InitSpreadSheet()

'	Call InitComboBox()
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Call FncSetToolBar("New")

'	Call CookiePage(0)
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
'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'==========================================================================================

Sub txtDateYr_DblClick(Button)
	if Button = 1 then
		frm1.txtDateYr.Action = 7
	End if
End Sub

Sub txtDateYr_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    
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

	frm1.vspdData.Row = Row
'	lsPoNo=frm1.vspdData.Text
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
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
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then
	'	If lgStrPreglno <> "" Then
           Call DisableToolBar(TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If    
    
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 	
	
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
    If Not chkField(Document, "1") Then										'��: This function check indispensable field
       Exit Function
    End If
		
	Call ExtractDateFrom(frm1.txtDateYr.Text,frm1.txtDateYr.UserDefinedFormat,gComDateType,strDateYr,strDateMonth,strDateDay)

    '-----------------------
    'Query function call area
    '-----------------------
    IF  DbQuery	= False Then														'��: Query db data
		Exit Function
	END IF
	
    FncQuery = True		
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
	Call parent.FncExport(C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
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
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
	Call FncSetToolBar("Query")
		
    With frm1

		strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode = OPMD_CMODE Then   ' This means that it is first search
        
			strVal = strVal & "?txtDateYr=" & strDateYr	
			strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.Value)
			strVal = strVal & "&txtClassType=" & Trim(.txtClassType.Value)				'��:
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        Else
			strVal = strVal & "?txtDateYr=" & strDateYr	
			strVal = strVal & "&txtBizAreaCd=" & Trim(.hBizAreaCd.Value)
			strVal = strVal & "&txtClassType=" & Trim(.hClassType.Value)				'��: 
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        End If   
        
'##        strVal = strVal & "&

    '--------- Developer Coding Part (End) ------------------------------------------------------------
		strVal = strVal & "&lgPageNo="   & lgPageNo
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)
        strVal = strVal & "&lgSelectListDT=" & lgSelectListDT
         
        strVal = strVal & "&lgTailList="     & MakeSql()

		strVal = strVal & "&lgSelectList="   & EnCoding(lgSelectList)
		
        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ����       
        	
    End With
    
    DbQuery = True


End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

	IF Trim(frm1.txtBizAreaCd.value) = "" then
		frm1.txtBizAreaNm.value = ""
	end if
	
	IF Trim(frm1.txtClassType.value) = "" then
		frm1.txtClassType.value = ""
	end if		

	Call FncSetToolBar("New")	
	'SetGridFocus

	Set gActiveElement = document.activeElement 
	
End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================
' Function Name : MakeSql()
' Function Desc : Order by ���� group by ���� �����.
'========================================================================================

Function MakeSql()
    Dim iStr,jStr
    Dim ii,jj,kk
    Dim iFirst
    Dim tmpPopUpR
    
    '2001/03/30 �ڵ�,�ڵ�� ���İ��� ���� 
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
        If tmpPopUpR(ii) <> "" Then
           If lgTypeCD(0) = "G" Then
              For jj = 0 To UBound(lgFieldNM) - 1                                            'Sort ��󸮽�Ʈ   ���� 
                  If lgMark(jj) <> "X" Then
                     If lgPopUpR(ii,0) = lgFieldCD(jj) Then
                        If iFirst = "Y" Then
                           iStr = iStr & " , "
                           jStr = jStr & " , " 
                        End If   
                        If CInt(Trim(lgNextSeq(jj))) >= 1 And CInt(Trim(lgNextSeq(jj))) <= UBound(lgFieldNM) Then
                           iStr = iStr & " " & lgPopUpR(ii,0) & " " & lgPopUpR(ii,1) & " , " & lgFieldCD(CInt(lgNextSeq(jj)) - 1)
                           jStr = jStr & " " & lgPopUpR(ii,0) & " " &          " , " & lgFieldCD(CInt(lgNextSeq(jj)) - 1)
                           '2001/03/30 �ڵ�,�ڵ�� ���İ��� ���� 
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
       MakeSql = " Order By" & iStr
    End If   
End Function


'========================================================================================
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	Case 0
		arrParam(0) = "����� �˾�"						' �˾� ��Ī 
		arrParam(1) = "B_Biz_AREA"							' TABLE ��Ī 
		arrParam(2) = strCode								' Code Condition
		arrParam(3) = ""									' Name Cindition
		arrParam(4) = ""									' Where Condition
		arrParam(5) = "������ڵ�"			
	
	    arrField(0) = "BIZ_AREA_CD"								' Field��(0)
		arrField(1) = "BIZ_AREA_NM"								' Field��(1)
    
	    arrHeader(0) = "������ڵ�"							' Header��(0)
		arrHeader(1) = "������"							' Header��(1)
    
	Case 1
		arrParam(0) = "�Է°�� �˾�"					' �˾� ��Ī 
		arrParam(1) = "B_MINOR"						' TABLE ��Ī 
		arrParam(2) = strCode									' Code Condition
		arrParam(3) = ""										' Name Cindition
		arrParam(4) = "major_cd = " & FilterVar("A1001", "''", "S") & "  and (minor_cd = " & FilterVar("BR", "''", "S") & "  or minor_cd = " & FilterVar("TR", "''", "S") & "  or minor_cd LIKE " & FilterVar("%T", "''", "S") & " )"								' Where Condition
		arrParam(5) = "�Է°��"			
	
	    arrField(0) = "MINOR_CD"								' Field��(0)
		arrField(1) = "MINOR_NM"							' Field��(1)
    
	    arrHeader(0) = "�Է°��"						' Header��(0)
		arrHeader(1) = "�Է°�θ�"							' Header��(1)
    
	Case Else
		Exit Function
	End Select
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
		Case 0
			.txtBizAreaCd.value = arrRet(0)
			.txtBizAreaNm.value = arrRet(1)
		Case 1
			.txtClassType.value   = arrRet(0)
			.txtClassTypeNm.value = arrRet(1)
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
	
    Call CopyPopupInfABT("1")

	lgIsOpenPop = True
	
    TInf(0) = "<%=gMethodText%>"    
  
	For ii = 0 to C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR_T(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR_T(ii / 2  , 1)
    Next  
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(lgSortFieldCD_T,lgSortFieldNm_T,arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to C_MaxSelList * 2 - 1 Step 2
           lgPopUpR_T(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR_T(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call CopyTBL("1")
	   Call InitSpreadSheet()													'��: Initializes Spread Sheet 1

   End If
End Function

'========================================================================================================
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'========================================================================================================

Function OpenGroupPopup()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = gMethodText
  
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



'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarBizArea, VarClassTypeFr, VarDateFr, VarDateTo, VarDate, VarDrAmt, VarCrAmt)
Dim strGlYear
Dim strGlMonth
Dim strgGlDay
Dim strFiscYr,strFiscMnth,strFiscDt
Dim strFiscEndYr,strFiscEndMnth,strFiscEndDt

	StrEbrFile = "a5134ma1.ebr"
	
	With frm1

		If Trim(.txtBizAreaCd.value) = "" Then
			VarBizArea = "%"
		Else
			VarBizArea = UCase(Trim(.txtBizAreaCd.value))
		End If	

		If Trim(.txtClassType.value) = "" Then
			VarClassTypeFr = "%"
		Else
			VarClassTypeFr = UCase(Trim(.txtClassType.value))
		End If	

		Call ExtractDateFrom(frm1.txtDateYr.Text,frm1.txtDateYr.UserDefinedFormat,gComDateType,strGlYear,strGlMonth,strgGlDay)

		Call ExtractDateFrom(gFiscStart,gDateFormat,gComDateType,strFiscYr,strFiscMnth,strFiscDt)		
		Call ExtractDateFrom(gFiscEnd,gDateFormat,gComDateType,strFiscEndYr,strFiscEndMnth,strFiscEndDt)		
		
		VarDate = strGlYear
		VarDateFr = strGlYear + strFiscMnth + strFiscDt

		VarDateTo = CStr(CInt(VarDate) + 1) + strFiscMnth + strFiscDt

		VarDateFr = UniDateAdd("D", 0, VarDateFr, gServerDateFormat)		
		VarDateTo = UniDateAdd("D", -1, VarDateTo, gServerDateFormat)		
		
		VarDrAmt = Replace(.txtDAmt.value, gComNum1000, "")
		VarCrAmt = Replace(.txtCAmt.value, gComNum1000, "")

		VarDrAmt = Replace(.txtDAmt.value, gComNum1000, "")
		VarCrAmt = Replace(.txtCAmt.value, gComNum1000, "")
		
	End With
	
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile, VarBizArea, VarClassTypeFr, VarDateFr, VarDateTo, VarDate, VarDrAmt, VarCrAmt
	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

'	If UNIConvDateToYYYYMMDD(frm1.txtDateYr.Text, gDateFormat, "") > UNIConvDateToYYYYMMDD(frm1.txtDateTo.Text, gDateFormat, "") Then
'		Call DisplayMsgBox("970025", "X", frm1.txtDateYr.Alt, frm1.txtDateTo.Alt)
'		frm1.txtDateYr.focus
'		Exit Function
'	End If
	
	Call SetPrintCond(StrEbrFile, VarBizArea, VarClassTypeFr, VarDateFr, VarDateTo, VarDate, VarDrAmt, VarCrAmt)
	
'    On Error Resume Next                                                    '��: Protect system from crashing
    
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "BizArea|" & VarBizArea
	StrUrl = StrUrl & "|ClassTypeFr|" & VarClassTypeFr
	StrUrl = StrUrl & "|DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DateY|" & VarDate
	StrUrl = StrUrl & "|DrAmt|" & VarDrAmt
	StrUrl = StrUrl & "|CrAmt|" & VarCrAmt

	Call FncEBRPrint(EBAction,StrEbrFile,StrUrl)
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '��: Protect system from crashing
    
	Dim StrUrl
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile, VarBizArea, VarClassTypeFr, VarDateFr,VarDateTo,  VarDate, VarDrAmt, VarCrAmt
    
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	
	Call SetPrintCond(StrEbrFile, VarBizArea, VarClassTypeFr, VarDateFr, VarDateTo, VarDate, VarDrAmt, VarCrAmt)
	
	StrUrl = StrUrl & "BizArea|" & VarBizArea
	StrUrl = StrUrl & "|ClassTypeFr|" & VarClassTypeFr
	StrUrl = StrUrl & "|DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DateY|" & VarDate
	StrUrl = StrUrl & "|DrAmt|" & VarDrAmt
	StrUrl = StrUrl & "|CrAmt|" & VarCrAmt	

	Call FncEBRPreview(StrEbrFile,StrUrl)
		
End Function
'==========================================================
'���ٹ�ư ���� 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1000000000011111")
	End Select
End Function
'=========================================================================================================
' Function Name : CopyPopupInfABT
' Function Desc : set popup information according to iOpt
'===========================================================================================================
Sub  CopyPopupInfABT(Byval iOpt)
    Dim ii
    Call CopyTBL(iOpt)    
    If iOpt = "1" Then
       For ii = 0 to  C_MaxSelList - 1
           lgPopUpR_T(ii,0)   =   lgPopUpR_A(ii,0)  
           lgPopUpR_T(ii,1)   =   lgPopUpR_A(ii,1)  
       Next
       
       ReDim lgSortFieldCD_T(UBound(lgSortFieldCD_A))
       ReDim lgSortFieldNM_T(UBound(lgSortFieldNM_A))

       For ii = 0 to UBound(lgSortFieldCD_A)
           lgSortFieldCD_T(ii) = lgSortFieldCD_A(ii)
           lgSortFieldNM_T(ii) = lgSortFieldNM_A(ii)
       Next
    Else
       For ii = 0 to  C_MaxSelList - 1
           lgPopUpR_T(ii,0)   =   lgPopUpR_B(ii,0)  
           lgPopUpR_T(ii,1)   =   lgPopUpR_B(ii,1)  
       Next

       ReDim lgSortFieldCD_T(UBound(lgSortFieldCD_B))
       ReDim lgSortFieldNM_T(UBound(lgSortFieldNM_B))

       For ii = 0 to UBound(lgSortFieldCD_B)
           lgSortFieldCD_T(ii) = lgSortFieldCD_B(ii)
           lgSortFieldNM_T(ii) = lgSortFieldNM_B(ii)
       Next
    End If       
End Sub

'=========================================================================================================
' Function Name : CopyPopupInfTAB
' Function Desc : set popup information according to iOpt
'===========================================================================================================
Sub  CopyPopupInfTAB(Byval iOpt)
    Dim ii
    If iOpt = "1" Then
          
       For ii = 0 to  C_MaxSelList - 1
           lgPopUpR_A(ii,0)   =   lgPopUpR_T(ii,0)      
           lgPopUpR_A(ii,1)   =   lgPopUpR_T(ii,1)      
       Next
       
       lgSelectList_A        =   lgSelectList_T  
       lgSelectListDT_A      =   lgSelectListDT_T
    Else

       For ii = 0 to  C_MaxSelList - 1
           lgPopUpR_B(ii,0)   =   lgPopUpR_T(ii,0)      
           lgPopUpR_B(ii,1)   =   lgPopUpR_T(ii,1)      
       Next
       lgSelectList_B        =   lgSelectList_T  
       lgSelectListDT_B      =   lgSelectListDT_T
    End If       
End Sub


'========================================================================================
' Function Name : CopyTBL
' Function Desc : multi + multi�� ���� temp buffer�� copy
'========================================================================================
Sub  CopyTBL(ByVal iOpt)
   Dim ii
   Dim iSz
   Select Case iOpt
      Case "1"
              iSz  = UBound(lgTypeCD_A) 
              ReDim      lgTypeCD_T   (iSz)
              ReDim      lgFieldCD_T  (iSz)
              ReDim      lgFieldNM_T  (iSz)
              ReDim      lgFieldLen_T (iSz)
              ReDim      lgFieldType_T(iSz)
              ReDim      lgDefaultT_T (iSz)
              ReDim      lgNextSeq_T  (iSz)
              ReDim      lgKeyTag_T   (iSz)
                            
              For ii = 0 to iSz
                  lgTypeCD_T   (ii) =  lgTypeCD_A   (ii)
                  lgFieldCD_T  (ii) =  lgFieldCD_A  (ii)
                  lgFieldNM_T  (ii) =  lgFieldNM_A  (ii)
                  lgFieldLen_T (ii) =  lgFieldLen_A (ii)
                  lgFieldType_T(ii) =  lgFieldType_A(ii)
                  lgDefaultT_T (ii) =  lgDefaultT_A (ii)
                  lgNextSeq_T  (ii) =  lgNextSeq_A  (ii)
                  lgKeyTag_T   (ii) =  lgKeyTag_A   (ii)
              Next     

      Case "2"
              iSz  = UBound(lgTypeCD_B) 
              ReDim      lgTypeCD_T   (iSz)
              ReDim      lgFieldCD_T  (iSz)
              ReDim      lgFieldNM_T  (iSz)
              ReDim      lgFieldLen_T (iSz)
              ReDim      lgFieldType_T(iSz)
              ReDim      lgDefaultT_T (iSz)
              ReDim      lgNextSeq_T  (iSz)
              ReDim      lgKeyTag_T   (iSz)
                            
              For ii = 0 to iSz
                  lgTypeCD_T   (ii) =  lgTypeCD_B   (ii)
                  lgFieldCD_T  (ii) =  lgFieldCD_B  (ii)
                  lgFieldNM_T  (ii) =  lgFieldNM_B  (ii)
                  lgFieldLen_T (ii) =  lgFieldLen_B (ii)
                  lgFieldType_T(ii) =  lgFieldType_B(ii)
                  lgDefaultT_T (ii) =  lgDefaultT_B (ii)
                  lgNextSeq_T  (ii) =  lgNextSeq_B  (ii)
                  lgKeyTag_T   (ii) =  lgKeyTag_B   (ii)
              Next     
    End Select              
End Sub

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../image/table/seltab_up_bg.gif" NOWRAP><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������ǥ��ȸ</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right><button name="btnAutoSel" class="clsmbtn" ONCLICK="OpenGroupPopup()">���ļ���</button></td>					
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
									<TD CLASS="TD5" NOWRAP>ȸ��⵵</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5134ma1_fpDateYr_txtDateYr.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>������ڵ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtBizAreaCd.value,0)">&nbsp;
														   <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 tag="24X" ALT="������" STYLE="TEXT-ALIGN: Left">
									</TD>										  
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�Է°��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtClassType" SIZE=11 MAXLENGTH=4 tag="11XXXU" ALT="�Է°��" STYLE="TEXT-ALIGN:left"><IMG SRC="../../image/btnPopup.gif" NAME="btnClassType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtClassType.value,1)">&nbsp;
														   <INPUT TYPE=TEXT NAME="txtClassTypeNm" SIZE=20 tag="24X" ALT="�Է°�θ�" STYLE="TEXT-ALIGN: Left">
									</TD>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/a5134ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����հ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDAmt" TYPE="Text" MAXLENGTH="20" STYLE="TEXT-ALIGN: right" tag="24X2"></TD>
								<TD CLASS=TD5 NOWRAP>�뺯�հ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCAmt" TYPE="Text" MAXLENGTH="20" STYLE="TEXT-ALIGN: right" tag="24X2"></TD>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hBizAreaCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hClassType" tag="24">
<INPUT TYPE=HIDDEN NAME="hClassCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hDateFr" tag="24">
<INPUT TYPE=HIDDEN NAME="hDateTo" tag="24">
<INPUT TYPE=HIDDEN NAME="hCommand" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>
