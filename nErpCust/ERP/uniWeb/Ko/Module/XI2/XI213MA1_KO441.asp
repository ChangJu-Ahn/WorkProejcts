<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : MES Interface ���۰���
*  2. Function Name        : ǰ�񸶽��� ������Ȳ 
*  3. Program ID           : XI213MA1_KO441
*  4. Program Name         : XI213MA1_KO441
*  5. Program Desc         : ǰ�񸶽��� ������Ȳ 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/01/03
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : CHCHO
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "XI213MB1_KO441.asp"                               'Biz Logic ASP 
Const BIZ_PGM_JUMP_ID = "B1B01MA1"				      '��: Jump ASP�� 
Const BIZ_PGM_JUMP_ID1 = "B1B11MA1"				      '��: Jump ASP�� 

Const C_SHEETMAXROWS    = 21	                                      '�� ȭ�鿡 �������� �ִ밹��*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop          
Dim lsConcd
Dim lsConNm
Dim IsOpenPop          

Dim C_PlantCd
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_MinorNm
Dim C_ItemGroupCd
Dim C_ItemGroupNm
Dim C_StdTime
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_CreateType
Dim C_SendDt
Dim C_MesReceiveFlag
Dim C_ErrDesc
Dim C_MesReceiveDt

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
sub InitSpreadPosVariables()
	C_PlantCd				= 1
	C_ItemCd				= 2
	C_ItemNm				= 3
	C_Spec					= 4
	C_MinorNm				= 5
	C_ItemGroupCd			= 6
	C_ItemGroupNm			= 7
	C_StdTime				= 8
	C_ValidFromDt			= 9
	C_ValidToDt				= 10
	C_CreateType			= 11
	C_SendDt				= 12
	C_MesReceiveFlag		= 13
	C_ErrDesc				= 14
	C_MesReceiveDt			= 15
	
end sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtStart_dt.Focus	
	frm1.txtStart_dt.Year = strYear 		 '����� default value setting
	frm1.txtStart_dt.Month = strMonth
	frm1.txtStart_dt.Day = strDay
	frm1.txtEnd_dt.Year = strYear 		 '����� default value setting
	frm1.txtEnd_dt.Month = strMonth 
	frm1.txtEnd_dt.Day = strDay
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

Sub InitComboBox()
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboAccount, lgF0, lgF1, Chr(11))
    
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp, arrVal

	Call vspdData_Click(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow)

	If flgs = 1 Then
		'WriteCookie CookieSplit, lsConcd & parent.gRowSep & lsConnm
		WriteCookie "txtPlantCd", lsConcd
		WriteCookie "txtItemCd", lsConnm

	ElseIf flgs = 0 Then

		strTemp =  ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		If Err.number <> 0 Then
			Err.Clear
			 WriteCookie CookieSplit , ""
			Exit Function 
		End If

		 WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

	Dim changeTxtEnd_dt
	changeTxtEnd_dt = Trim(Frm1.txtEnd_dt.text)
	changeTxtEnd_dt = UniConvDateAToB(UNIDateAdd ("D", 1, changeTxtEnd_dt, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)

    lgKeyStream = Trim(Frm1.txtPlantCd.Value) & parent.gColSep              
    lgKeyStream = lgKeyStream & Trim(Frm1.txtStart_dt.Text) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtEnd_dt.text)   & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtItemCd.Value)  & parent.gColSep

	If frm1.txtMesReceivetFlg(0).checked Then 
		lgKeyStream = lgKeyStream & "A" & parent.gColSep
	Elseif frm1.txtMesReceivetFlg(1).checked Then 
		lgKeyStream = lgKeyStream & "Y" & parent.gColSep
	Elseif frm1.txtMesReceivetFlg(2).checked Then 
		lgKeyStream = lgKeyStream & "N" & parent.gColSep
	Else
		lgKeyStream = lgKeyStream & "A" & parent.gColSep		
	End if

    lgKeyStream = lgKeyStream & Trim(Frm1.cboAccount.Value) & parent.gColSep
    'msgbox(lgKeyStream)

End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  
		
	With frm1.vspdData
	
        ggoSpread.Source = Frm1.vspdData
   		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	   .ReDraw = false

       .MaxCols = C_MesReceiveDt + 1                                                      ' ��:��: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ��:��: Hide maxcols
       .ColHidden = True                                                            ' ��:��:
    
       .MaxRows = 0	
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData        
		Call GetSpreadColumnPos("A")       
        Call  AppendNumberPlace("6","7","2")

		ggoSpread.SSSetEdit		C_PlantCd,			"����",4,,,4,2
		ggoSpread.SSSetEdit		C_ItemCd,			"ǰ��",15,,,18,2
		ggoSpread.SSSetEdit		C_ItemNm,			"ǰ���",25,,,40
		ggoSpread.SSSetEdit		C_Spec,				"�԰�",15,,,15
		ggoSpread.SSSetEdit		C_MinorNm,			"ǰ�����", 8
		ggoSpread.SSSetEdit		C_ItemGroupCd,		"ǰ��׷�",10,,,18,2
		ggoSpread.SSSetEdit		C_ItemGroupNm,		"ǰ��׷��",20,,,40

		ggoSpread.SSSetFloat	C_StdTime,			"ǥ���۾��ð�", 10, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
				
		ggoSpread.SSSetDate		C_ValidFromDt,		"��ȿ������", 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate		C_ValidToDt,		"��ȿ������", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_CreateType,		"��������",8,2,,40
		'ggoSpread.SSSetDate	C_SendDt,			"ERP �����۽��Ͻ�", 15, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_SendDt,			"ERP �����۽��Ͻ�", 18
		ggoSpread.SSSetEdit		C_MesReceiveFlag,	"MES �ݿ�����",12,2,,12
		ggoSpread.SSSetEdit		C_ErrDesc,			"��������",25,,,20
		'ggoSpread.SSSetDate		C_MesReceiveDt,		"MES ���������Ͻ�", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_MesReceiveDt,		"MES ���������Ͻ�", 18


		ggoSpread.SSSetSplit2(1)										'frozen ����߰� 

        Call ggoSpread.SSSetColHidden(C_StdTime         , C_StdTime         , True) '20080306::HANC


	   .ReDraw = true    
	
       Call SetSpreadLock 
    
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
            
            C_PlantCd			= iCurColumnPos(1)
			C_ItemCd			= iCurColumnPos(2)
			C_ItemNm			= iCurColumnPos(3)
			C_Spec				= iCurColumnPos(4)
			C_MinorNm			= iCurColumnPos(5)
			C_ItemGroupCd		= iCurColumnPos(6)
			C_ItemGroupNm		= iCurColumnPos(7)
			C_StdTime			= iCurColumnPos(8)
			C_ValidFromDt		= iCurColumnPos(9)
			C_ValidToDt			= iCurColumnPos(10)
			C_CreateType		= iCurColumnPos(11)
			C_SendDt			= iCurColumnPos(12)
			C_MesReceiveFlag	= iCurColumnPos(13)
			C_ErrDesc			= iCurColumnPos(14)
			C_MesReceiveDt		= iCurColumnPos(15)

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
 	
    Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'��: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
 	Call InitComboBox

    Call InitVariables                                                              'Initializes local global variables
    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")
   
    Call SetDefaultVal
    Call SetToolbar("1100000000001111")										        '��ư ���� ���� 

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
   End If
 
	Call CookiePage (0)                                                             '��: Check Cookie
End Sub



'------------------------------------------  OpenPlantCd()  --------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "����"						
	arrParam(1) = "B_Plant"						
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)	
	arrParam(4) = ""							
	arrParam(5) = "����"						

    arrField(0) = "Plant_Cd"					
    arrField(1) = "Plant_NM"					
    
    arrHeader(0) = "����"					
    arrHeader(1) = "�����"					
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function

'------------------------------------------  OpenItemCd()  -----------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 
    
	'arrRet = window.showModalDialog("../m1/m1111pa1.asp", Array(window.parent, arrParam, arrField, arrHeader), _
	'	"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function




'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    If  ValidDateCheck(frm1.txtStart_dt, frm1.txtEnd_dt)=False Then
        Exit Function
    End If
     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '��: Initializes local global variables
    Call MakeKeyStream("X")

	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'��: Query db data
       
    FncQuery = True                                                              '��: Processing is OK
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
                                                         '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
 
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    Dim IntRetCD
    Dim imRow
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

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
     ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
             ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
     ggoSpread.Source = Frm1.vspdData	
     ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD,imRow
    
    On Error Resume Next         
    FncInsertRow = False
    
    if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End if
	With frm1
	    .vspdData.ReDraw = False
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    ggoSpread.InsertRow,imRow
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	   .vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
	If Err.number =0 Then
		FncInsertRow = True
	End if
	
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
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
  
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport( parent.C_MULTI)                                         '��: ȭ�� ���� 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( parent.C_MULTI, False)                                    '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
     ggoSpread.Source = frm1.vspdData	
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '��: Clear err status

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey                 '��: Next key tag
    End With
		
	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
    
    DbQuery = True
    
End Function


'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
 	
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'��: Lock field
    Call InitData()
    Call SetToolbar("1100000000001111")										        '��ư ���� ���� 
	Frm1.vspdData.focus	
End Function
 

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim intIndex
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
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

    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 1
    lsConcd=frm1.vspdData.Text
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 2
    lsConnm=frm1.vspdData.Text  

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
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

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub


'========================================================================================================
' Name : txtStart_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtStart_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")         
        frm1.txtStart_dt.Action = 7 
		frm1.txtStart_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtEnd_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtEnd_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")         
        frm1.txtEnd_dt.Action = 7 
        frm1.txtEnd_dt.focus
    End If
End Sub


'==========================================================================================
'   Event Name : txtStart_dt_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtStart_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================================================================
'   Event Name : txtEnd_dt_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtEnd_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			                <TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU"  ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbScript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>ERP �۽űⰣ</TD>
							    	<TD CLASS="TD6" NOWRAP>
							    		<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD  name=txtStart_dt classid=<%=gCLSIDFPDT%> ALT="������" tag="12X1X" VIEWASTEXT></OBJECT>');</SCRIPT>
	                                    &nbsp;~&nbsp;
	                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtEnd_dt classid=<%=gCLSIDFPDT%>ALT="������" tag="12X1X" VIEWASTEXT></OBJECT>');</SCRIPT>
							    	</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU"  ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbScript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14" ALT="ǰ���"></TD>
									<TD CLASS=TD5 NOWRAP>MES ���ſ���</TD>
									<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" NAME="txtMesReceivetFlg" ID="rdoValidFlg1" Value="A" CLASS="RADIO" tag="1X" CHECKED><LABEL FOR="rdoValidFlg1">��ü</LABEL>
												<INPUT TYPE="RADIO" NAME="txtMesReceivetFlg" ID="rdoValidFlg2" Value="Y" CLASS="RADIO" tag="1X"><LABEL FOR="rdoValidFlg2">����</LABEL>
												<INPUT TYPE="RADIO" NAME="txtMesReceivetFlg" ID="rdoValidFlg3" Value="N" CLASS="RADIO" tag="1X"><LABEL FOR="rdoValidFlg3">����</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAccount" ALT="����" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=500>&nbsp;[��������]A :�ű��Է� | B : ���� | C : ����</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPayCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



