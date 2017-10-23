
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : a5132ma1
'*  4. Program Name         : ������ ������ ����ǥ 
'*  5. Program Desc         : Query of Account Code
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.11.15
'*  8. Modified date(Last)  : ahj
'* 10. Modifier (Last)      : ahj
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncServer.asp"  -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim lgBlnFlgChgValue                                        '��: Variable is for Dirty flag            
Dim lgStrPrevKey                                            '��: Next Key tag                          
Dim lgSortKey                                               '��: Sort���� ���庯��                      
Dim IsOpenPop                                               '��: Popup status                           

Dim lgSelectList                                            '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 

Dim lgTypeCD                                                '��: 'G' is for group , 'S' is for Sort    
Dim lgFieldCD                                               '��: �ʵ� �ڵ尪                           
Dim lgFieldNM                                               '��: �ʵ� ����                           
Dim lgFieldLen                                              '��: �ʵ� ��(Spreadsheet����)              
Dim lgFieldType                                             '��: �ʵ� ����                           
Dim lgDefaultT                                              '��: �ʵ� �⺻��                           
Dim lgNextSeq                                               '��: �ʵ� Pair��                           
Dim lgKeyTag                                                '��: Key  ����                             

Dim lgSortFieldNm                                           '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD                                          '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark                                                  '��: ��ũ                                  
<%


'--------------- ������ coding part(�������,Start)-----------------------------------------------------------

  Call GetAdoFiledInf("A5132MA1","S", "A")						'��: spread sheet �ʵ����� query   -----
																' G is for Qroup , S is for Sort
																' A is spreadsheet No
'--------------- ������ coding part(�������,End)-------------------------------------------------------------
%>

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "A5132MB1.asp"
Const C_SHEETMAXROWS    = 30                                   '��: Spread sheet���� �������� row
Const C_SHEETMAXROWS_D  = 100                                  '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
'Dim lsPoNo								                       '��: Jump�� Cookie�� ���� Grid value
Const C_MaxKey          = 2                                    '�١١١�: Max key value
'--------------- ������ coding part(��������,End)-------------------------------------------------------------

 '#########################################################################################################
'												2. Function�� 
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
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
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	
	EndDate = GetSvrDate
	
	Call ExtractDateFrom(EndDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)

	StartDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, "01")
	EndDate   = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)
%>

frm1.txtFromGlDt.Text = "<%=StartDate %>"
frm1.txtToGlDt.Text = "<%=EndDate %>"
frm1.txtFromGlDt.focus

	
'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029(gCurrency, "Q", "A") %>
End Sub


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
   
   ' Spreadsheet #2�˻��� ���� Ű ����ġ ���� 
   If CInt(lgKeyTag(iDx)) > 0 And CInt(lgKeyTag(iDx)) <= C_MaxKey Then  
      lgKeyPos(CInt(lgKeyTag(iDx))) =  iCol
   End If

End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
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

	pB1a028.ImportBMajorMajorCd = strMajorCd									'��: Major Code
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

 '**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 
'============================================================
'�μ��ڵ� �˾� 
'============================================================


'----------------------------------------  OpenAcctCd()  -------------------------------------------------
'	Name : OpenAcctCd()
'	Description : Account PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcctCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"	
	arrParam(1) = " A_ACCT A, A_CTRL_ITEM B, A_CTRL_ITEM C, A_ACCT_GP D "
	arrParam(2) = Trim(frm1.txtAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.SUBLEDGER_1 *= B.CTRL_CD AND A.SUBLEDGER_2 *= C.CTRL_CD AND A.GP_CD = D.GP_CD AND (A.SUBLEDGER_1 IS NOT NULL AND A.SUBLEDGER_1 <> '') "
	arrParam(5) = "�����ڵ�"			
	
	arrField(0) = "A.ACCT_CD"						' Field��(0)
	arrField(1) = "D.GP_CD"						' Field��(1)
	arrField(2) = "D.GP_NM+" & FilterVar(" - ", "''", "S") & " + A.ACCT_NM"							' Field��(2)
	arrField(3) = "B.CTRL_CD"			
	arrField(4) = "B.CTRL_NM"
	arrField(5) = "C.CTRL_CD"
	arrField(6) = "C.CTRL_NM"


	arrHeader(0) = "�����ڵ�"		
	arrHeader(1) = "�׷��ڵ�"
	arrHeader(2) = "������"									' Header��(2)
	arrHeader(3) = "�����׸�1"	
	arrHeader(4) = "�����׸��1"
	arrHeader(5) = "�����׸�2"
	arrHeader(6) = "�����׸��2"


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,2)
	End If	
	
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

'========================================================================================
'                       ȸ����ǥ POPUP
' ========================================================================================  
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = 9

	
		arrParam(0) = Trim(.Text)	'ȸ����ǥ��ȣ 
		arrParam(1) = ""			'Reference��ȣ 
	End With

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 
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

'--------------- ������ coding part(�������,Start)---------------------------------------------------
		 '�ڵ���ȸ�Ǵ� ���ǰ��� �˻����Ǻ� Name�� Match 
		For iniSep = 0 To UBound(arrVal) -1
			Select Case UCase(Trim(arrVal(iniSep)))
			Case UCase("��������")
				frm1.txtPoType.value =  arrVal(iniSep + 1)
			Case UCase("�������¸�")
				frm1.txtPoTypeNm.value =  arrVal(iniSep + 1)
			Case UCase("����ó")
				frm1.txtSpplCd.value =  arrVal(iniSep + 1)
			Case UCase("����ó��")
				frm1.txtSpplNm.value =  arrVal(iniSep + 1)
			Case UCase("���ű׷�")
				frm1.txtPurGrpCd.value =  arrVal(iniSep + 1)
			Case UCase("���ű׷��")
				frm1.txtPurGrpNm.value =  arrVal(iniSep + 1)
			Case UCase("ǰ��")
				frm1.txtItemCd.value =  arrVal(iniSep + 1)
			Case UCase("ǰ���")
				frm1.txtItemNm.value =  arrVal(iniSep + 1)
			Case UCase("Tracking No.")
				frm1.txtTrackNo.value =  arrVal(iniSep + 1)
			End Select
		Next
'--------------- ������ coding part(�������,End)---------------------------------------------------

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

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

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)

    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

    ReDim lgPopUpR(C_MaxSelList - 1,1)
 
	Call InitVariables													'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call InitComboBox()
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

Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtFromGlDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtFromGlDt.Focus       
    End If
End Sub

Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtToGlDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtToGlDt.Focus       
    End If
End Sub

Sub txtFromGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call FncQuery
	End If   
End Sub

Sub txtToGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call FncQuery
	End If   
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
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ'
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DbQuery
		End If
   End if
    
End Sub


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

	If CompareDateByFormat(frm1.txtFromGlDt.Text, frm1.txtToGlDt.Text, frm1.txtFromGlDt.Alt, frm1.txtToGlDt.Alt, _
						"970025", frm1.txtFromGlDt.UserDefinedFormat, gComDateType, true) = False Then
			frm1.txtFromGlDt.focus											'��: GL Date Compare Common Function
			Exit Function
	End if
   

	Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery															'��: Query db data

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

'-------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(ByVal arrRet,ByVal field_fg) 
	With frm1	
		Select case field_fg
			case 1
				.txtBizAreaCd.Value		= arrRet(0)
				.txtBizAreaNm.Value		= arrRet(1)
			case 2
				.txtAcctCd.Value		= arrRet(0)
				.txtAcctNm.Value		= arrRet(2)
				
				 'Call DbPopUpQuery()
			case 3
				'.txtSubLedger1.value	= arrRet(0)
				'.txtSubLedger3.value	= arrRet(1)
			case 4											'OpenSubledger2
				'.txtSubLedger2.value	= arrRet(0)
				'.txtSubLedger4.value	= arrRet(1)
		End select	
	End With

End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
	        
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtFromGlDt=" & Trim(.txtFromGlDt.Text)
		strVal = strVal & "&txtToGlDt=" & Trim(.txtToGlDt.Text)
		strVal = strVal & "&txtAcctCd=" & Trim(.txtAcctCd.Value)
		strVal = strVal & "&txtAcctCd_Alt=" & Trim(.txtAcctCd.Alt)
		
'--------------- ������ coding part(�������,End)------------------------------------------------

		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

	
	Call FncSetToolBar("Query")
		
	'frm1.txtBankCd.focus
	
	'SetGridFocus
	Set gActiveElement = document.activeElement 
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################


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
                           jStr = jStr & " " & lgPopUpR(ii,0) & " " &        " , " & lgFieldCD(CInt(lgNextSeq(jj)) - 1)
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
    
  '  If lgTypeCD(0) = "G" Then
  '     MakeSql =  "Group By " & jStr  & " Order By " & iStr 
  '  Else
  '     MakeSql = " Order By" & iStr
  '  End If   
End Function

'==========================================================
'���ٹ�ư ���� 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
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
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����������������ǥ</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>ȸ����</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5132ma1_fpDateTime1_txtFromGlDt.js'></script>&nbsp;~&nbsp;
												           <script language =javascript src='./js/a5132ma1_fpDateTime2_txtToGlDt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="�����ڵ�"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenAcctCd()">&nbsp;
														   <INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=25 tag="14"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" colspan=7>
								<script language =javascript src='./js/a5132ma1_vspdData_vspdData.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�հ�ݾ�</TD>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD5 NOWRAP><INPUT NAME="txtSDrAmt" TYPE="Text" MAXLENGTH="20" STYLE="TEXT-ALIGN: right" tag="24X2"></TD>
								<TD CLASS=TD5 NOWRAP>�뺯</TD>
								<TD CLASS=TD5 NOWRAP><INPUT NAME="txtSCrAmt" TYPE="Text" MAXLENGTH="20" STYLE="TEXT-ALIGN: right" tag="24X2"></TD>								
							</TR>
						</TABLE>						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HGIEHT_TYPE_01%>></td>
	</TR>
	<tr>	
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1">
</TEXTAREA><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
