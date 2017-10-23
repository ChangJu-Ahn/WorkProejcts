<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : Human Resources
'*  2. Function Name        : ��������(���ο��ݼҵ��Ѿ׽Ű�)
'*  3. Program ID           : H5406ma1.asp
'*  4. Program Name         : H5406ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/05/31
'*  7. Modified date(Last)  : 2003/06/11
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Lee SiNa
'* 10. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'======================================================================================================== 
Const BIZ_PGM_ID = "h5406mb1.asp"												'�����Ͻ� ���� ASP�� 
Const CookieSplit = 1233
Const C_SHEETMAXROWS = 30

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

Dim C_COMP_CD
Dim C_NO
Dim C_COMP_PAGE
Dim C_RES_NO
Dim C_NAME
Dim C_WORK_MONTH
Dim C_TOT_AMT
Dim C_JISA_CODE
Dim C_EDI_CD
Dim C_EMPTY

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_COMP_CD		= 1     	
    C_NO			= 2    
    C_COMP_PAGE		= 3     
    C_RES_NO		= 4     
    C_NAME			= 5
    C_WORK_MONTH	= 6 
    C_TOT_AMT		= 7  
    C_JISA_CODE		= 8  
    C_EDI_CD		= 9 
    C_EMPTY			= 10

End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'======================================================================================================
'	Name : SeTDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
		
	frm1.txtEntrDt.Year = strYear 		'��� default value setting
	frm1.txtEntrDt.Month = strMonth
	frm1.txtEntrDt.Day = strDay
	
	frm1.txtYear.Year = strYear 		'��� default value setting
	frm1.txtYear.Month = strMonth
	frm1.txtYear.Day = strDay

	Call ggoOper.FormatDate(frm1.txtYear, Parent.gDateFormat, 3)
	frm1.txtAutoCd.value = "06"

End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

	lgKeyStream   = Frm1.txtBizArea.Value & Parent.gColSep                 '0
    lgKeyStream   = lgKeyStream & Frm1.txtYear.text & Parent.gColSep       '1
    lgKeyStream   = lgKeyStream & Frm1.txtEntrDt.Text & Parent.gColSep     '2
	lgKeyStream   = lgKeyStream & Frm1.txtArea.Value & Parent.gColSep      '3
    lgKeyStream   = lgKeyStream & Frm1.txtCompCd.Value & Parent.gColSep    '4
    lgKeyStream   = lgKeyStream & Frm1.txtAutoCd.Value & Parent.gColSep    '5

End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()   'sbk 

	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	    .ReDraw = false
	
        .MaxCols = C_EMPTY + 1												<%'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>
	    .Col = .MaxCols															<%'������Ʈ�� ��� Hidden Column%>
        .ColHidden = True
    
        .MaxRows = 0
        ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos("A") 'sbk
	
	    Call AppendNumberPlace("6","9","0")

        ggoSpread.SSSetEdit C_COMP_CD,		"������ȣ", 12
        ggoSpread.SSSetEdit C_NO,			"�Ϸù�ȣ"	, 8
        ggoSpread.SSSetEdit C_COMP_PAGE,	"�����������", 14, 2    
	    ggoSpread.SSSetEdit C_RES_NO,		"�ֹι�ȣ", 13
	    ggoSpread.SSSetEdit C_NAME,			"����", 10
	    ggoSpread.SSSetEdit C_WORK_MONTH,	"�ٹ�����", 10,,,2,2
'	    ggoSpread.SSSetEdit C_TOT_AMT,		"�ҵ��Ѿ�" ,9,,,9,2
		ggoSpread.SSSetFloat C_TOT_AMT,		"�ҵ��Ѿ�" ,  13,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit C_JISA_CODE,	"�����ڵ�", 10, 2
        ggoSpread.SSSetEdit C_EDI_CD,		"����ȭ�ڵ�", 10, 2
        ggoSpread.SSSetEdit C_EMPTY,		"����", 8, 2
    
	    .ReDraw = true
	
        Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================%>
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_COMP_CD,			-1 , -1
    ggoSpread.SpreadLock C_NO,				-1 , -1
    ggoSpread.SpreadLock C_COMP_PAGE,		-1 , -1
    ggoSpread.SpreadLock C_RES_NO,			-1 , -1
    ggoSpread.SpreadLock C_NAME,			-1 , -1

	ggoSpread.SSSetRequired C_WORK_MONTH,	-1 , -1
	ggoSpread.SSSetRequired C_TOT_AMT,		-1 , -1    
		
    ggoSpread.SpreadLock C_JISA_CODE,		-1 , -1
    ggoSpread.SpreadLock C_EDI_CD,			-1 , -1
    ggoSpread.SpreadLock C_EMPTY,			-1 , -1
    ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
     
    .vspdData.ReDraw = True

    End With
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

            C_COMP_CD		= iCurColumnPos(1)
            C_NO			= iCurColumnPos(2)
            C_COMP_PAGE		= iCurColumnPos(3)
            C_RES_NO		= iCurColumnPos(4)
            C_NAME			= iCurColumnPos(5)
            C_WORK_MONTH	= iCurColumnPos(6)
            C_TOT_AMT		= iCurColumnPos(7)
            C_JISA_CODE		= iCurColumnPos(8)
            C_EDI_CD		= iCurColumnPos(9)
            C_EMPTY			= iCurColumnPos(10)            
    End Select    
End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================%>
Sub Form_Load()

    Call LoadInfTB19029                                                     'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                            'Lock  Suitable  Field%>                         
                                                                            'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
            
    Call InitSpreadSheet                                                    'Setup the Spread sheet%>
    Call InitVariables                                                      'Initializes local global variables%>
    
    Call SetDefaultVal
    Call SetToolbar("1100000000011111")										'��ư ���� ���� %>
    Call CookiePage(0)
    
    frm1.txtBizArea.focus
    
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================%>
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub
'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing%>

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
   
   Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field%>
   ggoSpread.ClearSpreadData

    Call InitVariables                                                      'Initializes local global variables%>

    If Not chkField(Document, "1") Then						         'This function check indispensable field%>
       Exit Function
    End If
    if txtArea_OnChange=false then
		exit function
	end if
	
    If Len(frm1.txtBizArea.value) <> 8 Then
		
		Call DisplayMsgBox("970029", "X", frm1.txtBizArea.alt,"X")
		Exit Function
    ElseIf Len(frm1.txtCompCd.value) <> 4 Then
		Call DisplayMsgBox("970029", "X", frm1.txtCompCd.alt,"X")
		Exit Function
	ElseIf Len(frm1.txtAutoCd.value) <> 2 Then
		Call DisplayMsgBox("970029", "X", frm1.txtAutoCd.alt,"X")
		Exit Function
	ElseIf (Len(frm1.txtYear.text) <> 4 or frm1.txtYear.text > "3000" or frm1.txtYear.text < "1900") Then
		Call DisplayMsgBox("970029", "X", frm1.txtYear.alt,"X")
		Exit Function
    End If
    
    Call MakeKeyStream("X")

    If DbQuery = False Then
        Exit Function
    End If
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete()

End Function

 '========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
  
     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '��: No data changed!!
        Exit Function
    End If
 	 
    If Not chkField(Document, "2") Then
       Exit Function
    End If
  
	ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_SAVE)
	If DBSave=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
    
    FncSave = True                                                              '��: Processing is OK
    
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
		
    If LayerShowHide(1)=False Then
		Exit Function
    End If

	With frm1
		.txtMode.value        = parent.UID_M0002                                        '��: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1
 
	With Frm1
   
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
               Case ggoSpread.InsertFlag                                      '��: insert�߰� 
                   
																		   strVal = strVal & "C" & parent.gColSep 'array(0)
																		   strVal = strVal & lRow & parent.gColSep
                                                                           strVal = strVal & Trim(.txtYear.year) &  parent.gColSep
                    .vspdData.Col = C_COMP_CD							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_NO								 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_COMP_PAGE							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                 
                    .vspdData.Col = C_RES_NO							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_NAME								 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_WORK_MONTH						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                 
                    .vspdData.Col = C_TOT_AMT							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_JISA_CODE							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_EDI_CD							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep
                    lGrpCnt = lGrpCnt + 1                                                               
               Case  ggoSpread.UpdateFlag                                      '��: Update
                                                                           strVal = strVal & "U" &  parent.gColSep
                                                                           strVal = strVal & lRow &  parent.gColSep
                                                                           strVal = strVal & Trim(.txtYear.year) & parent.gColSep
                    .vspdData.Col = C_RES_NO							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_WORK_MONTH						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_TOT_AMT							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep

                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '��: Processing is NG
End Function
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
    Call InitVariables															'��: Initializes local global variables
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
        Call RestoreToolBar()
        Exit Function
    End If
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()

End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================%>
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================%>
Function FncInsertRow() 

End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================%>
Function FncDeleteRow() 
    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================%>
Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================%>
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 <%'��: ȭ�� ���� %>
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================%>
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      <%'��:ȭ�� ����, Tab ���� %>
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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================%>
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                <%'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================%>
Function DbQuery() 
	Dim strVal
	
    DbQuery = False
    
    Err.Clear                                                               <%'��: Protect system from crashing%>

	If LayerShowHide(1) =False Then
       Exit Function
    End If
	
    With frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    	    
    End With
    
    Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
    DbQuery = True
    
End Function

'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================%>
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")									<%'This function lock the suitable field%>
    Call SetToolbar("1100100000011111")	    
	frm1.vspdData.focus	
End Function

'------------------------------------------  OpenArea()  -------------------------------------------
'	Name : OpenArea()
'	Description : �ٹ����� PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ٹ����� �˾�"				<%' �˾� ��Ī %>
	arrParam(1) = "b_minor"							<%' TABLE ��Ī %>
	arrParam(2) = frm1.txtArea.value				<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = " major_cd = " & FilterVar("H0035", "''", "S") & " "			<%' Where Condition%>
	arrParam(5) = "�ٹ�����"					<%' �����ʵ��� �� ��Ī %>
	
    arrField(0) = "minor_cd"					<%' Field��(0)%>
    arrField(1) = "minor_nm"					<%' Field��(1)%>
    
    arrHeader(0) = "�ڵ�"					<%' Header��(0)%>
    arrHeader(1) = "�ڵ��"					<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtArea.focus
		Exit Function
	Else
		Call SetArea(arrRet)
	End If	
	
End Function

'------------------------------------------  SetArea()  --------------------------------------------
'	Name : SetArea()
'	Description : �����ڵ� Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetArea(Byval arrRet)
		
	With frm1		
		.txtArea.value = arrRet(0)
		.txtAreaNm.value = arrRet(1)
		.txtArea.focus
	End With
	
End Function

'==========================================================================================
'   Event Name : btnBatch_OnClick()
'   Event Desc : ���ϻ��� 
'==========================================================================================
Function btnBatch_OnClick()
	Dim RetFlag
	Dim strVal
	Dim intRetCD

    Err.Clear                                                                   '��: Clear err status
    
    If Not chkField(Document, "1") Then                                         ' Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
       Exit Function                            
    End If
    
    If Len(frm1.txtBizArea.value) <> 8 Then
		Call DisplayMsgBox("970029", "X", frm1.txtBizArea.alt,"X")
		Exit Function
    ElseIf Len(frm1.txtCompCd.value) <> 4 Then
		Call DisplayMsgBox("970029", "X", frm1.txtCompCd.alt,"X")
		Exit Function
	ElseIf Len(frm1.txtAutoCd.value) <> 2 Then
		Call DisplayMsgBox("970029", "X", frm1.txtAutoCd.alt,"X")
		Exit Function
	ElseIf (Len(frm1.txtYear.text) <> 4 or frm1.txtYear.text > "3000" or frm1.txtYear.text < "1900") Then
		Call DisplayMsgBox("970029", "X", frm1.txtYear.alt,"X")
		Exit Function
    End If
    
    If frm1.vspdData.MaxRows <= 0 Then
		Call DisplayMsgBox("900002", "X","X","X")			 '��: Query First 
		Exit Function		
    End If

	RetFlag = DisplayMsgBox("900018", Parent.VB_YES_NO,"x","x")   '�� �ٲ�κ�	
	If RetFlag = VBNO Then
		Exit Function
	End IF

	If LayerShowHide(1) =False Then
       Exit Function
    End If
	
    With frm1
	    
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0003						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    End With    
    
    Call RunMyBizASP(MyBizASP, strVal)
End Function

Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                               '��: Protect system from crashing

	strVal = BIZ_PGM_ID & "?txtMode=" & "7"							'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtFileName=" & pFileName							'��: ��ȸ ���� ����Ÿ	
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
End Function



'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================%>
Sub vspdData_Change(ByVal Col , ByVal Row )
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================%>
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC" 

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If
        
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
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

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
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
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================%>
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
'=======================================================================================================
'   Event Name : txtEntrDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtEntrDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtEntrDt.Action = 7
        frm1.txtEntrDt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtEntrDt_Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================
Sub txtEntrDt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub



'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtYear_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYear.Action = 7
        frm1.txtYear.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYear_Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================
Sub txtYear_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYear_Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================
Sub txtYear_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub


'========================================================================================================
'   Event Name : txtCd_OnChange
'   Event Desc :
'========================================================================================================
Function txtArea_OnChange()    

    Dim IntRetCd
	dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    If frm1.txtArea.value = "" Then
        frm1.txtAreaNm.value = ""
		txtArea_OnChange = true
    ELSE    
        IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0035", "''", "S") & " AND minor_cd =  " & FilterVar(frm1.txtArea.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
            Call DisplayMsgBox("970000","X","�ٹ�����","X")
            frm1.txtAreaNm.value=""
            frm1.txtArea.focus
			Set gActiveElement = document.ActiveElement 
            Exit Function
        Else
            frm1.txtAreaNm.value=Trim(Replace(lgF0,Chr(11),""))
			txtArea_OnChange = true
        End If
    End If
End Function
'==========================================================================================
'   Event Name : btnCb_select_OnClick
'   Event Desc : ������ �������� 
'==========================================================================================
Function btnCb_select_OnClick()
	Dim RetFlag ,RetFlag2
	Dim strVal
	Dim intRetCD,strWhere, strEmp_no

    Err.Clear                                                                           '��: Clear err status
'	If gSelframeFlg = TAB1 Then      
'		If Not chkField(Document, "1") Then                                                 'Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
'		   Exit Function                            
'		End If
'	End If
		
      ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
   
   Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field%>
   ggoSpread.ClearSpreadData

    Call InitVariables                                                      'Initializes local global variables%>

    If Not chkField(Document, "1") Then						         'This function check indispensable field%>
       Exit Function
    End If
 		
	strWhere = " YEAR_YY = " & FilterVar(Frm1.txtYear.Year, "''", "S")
	 
 	IntRetCD = CommonQueryRs(" * "," HDB040T ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 	
	If IntRetCD = True  Then
		        
		IntRetCD = DisplayMsgBox("800502", 35,"X","X")	    '�̹� ������ �ڷᰡ �ֽ��ϴ�.�����Ͻðڽ��ϱ�?
		If IntRetCD = vbCancel Then
		   	Exit Function
		End If
	End If
					
'	RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                         '�� �۾��� ����Ͻðڽ��ϱ�?
'	If RetFlag = VBNO Then
'		Exit Function
'	End IF
	ggoSpread.ClearSpreadData
	
    With frm1
        Call LayerShowHide(1)					 
        lgCurrentSpd = "A"		
	    Call MakeKeyStream(lgCurrentSpd)  
        
		strVal = BIZ_PGM_ID    & "?txtMode="           & "5"						'��: �����Ͻ� ó�� ASP�� ���� 	    	    		    
		strVal = strVal         & "&lgCurrentSpd="      & lgCurrentSpd                  '��: Mulit�� ���� 
		strVal = strVal         & "&txtKeyStream="      & lgKeyStream                   '��: Query Key

		Call RunMyBizASP(MyBizASP, strVal)
 
    End With    
End Function

Sub DBAutoQueryOk()
    Dim lRow
    With Frm1

        .vspdData.ReDraw = false
        ggoSpread.Source = .vspdData
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0

            .vspdData.Text = ggoSpread.InsertFlag
        Next

'      ggoSpread.SpreadLock C_CHANG_DT, -1,C_CHANG_DT
    .vspdData.ReDraw = TRUE
    ggoSpread.ClearSpreadData "T"            
    End With    
    Call SetToolbar("1100100000011111")	    
    lgStrPrevKey = ""
    Set gActiveElement = document.ActiveElement   
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���ο��ݼҵ��Ѿ׽Ű�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
						        <TD CLASS=TD5 NOWRAP>������ȣ</TD>       
					            <TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBizArea" SIZE=20 MAXLENGTH=8 tag="12XXXU"  ALT="������ȣ"></TD>
							    <TD CLASS=TD5 NOWRAP>���س⵵</TD>
			                    <TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtYear CLASS=FPDTYYYY title=FPDATETIME tag="12X1" ALT="���س⵵"></OBJECT>');</SCRIPT></TD>
			                <TR>
								<TD CLASS=TD5 NOWRAP>�Ի������</TD>
			                    <TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtEntrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="�Ի������"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5>�ٹ�����</TD>
								<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtArea" SIZE=10 MAXLENGTH=2 tag="11XXXU"  ALT="�ٹ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnArea" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenArea()">
								<INPUT TYPE=TEXT NAME="txtAreaNm" tag="14X"></TD>					           
							</TR>
						    <TR>
								<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>       
					            <TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtCompCd" SIZE=10 MAXLENGTH=4 tag="12" STYLE="TEXT-ALIGN: center" ALT="�����ڵ�"></TD>
					            <TD CLASS=TD5 NOWRAP>����ȭ�ڵ�</TD>       
					            <TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtAutoCd" SIZE=10 MAXLENGTH=2 tag="12" STYLE="TEXT-ALIGN: center" ALT="����ȭ�ڵ�"></TD>
					        </TR>		
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
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
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnCb_select" CLASS="CLSMBTN">�����ͻ���</BUTTON>&nbsp;
						<BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag="1">���ϻ���</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hEmp_no" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<!--
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
-->
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

