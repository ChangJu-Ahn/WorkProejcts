<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>
<!--
======================================================================================================
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

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "S3101MB1_KO441.asp"                                      '�����Ͻ� ���� ASP�� 
Const C_SHEETMAXROWS    = 21	                                      '�� ȭ�鿡 �������� �ִ밹��*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd

Dim C_PROFORMA_NO
Dim C_BP_CD
Dim C_BP_POP
Dim C_BP_NM
Dim C_TEL_NO
Dim C_CUST_USER
Dim C_PROFORMA_DT
Dim C_DOCUMENT
Dim C_AMT
Dim C_USR_NM
Dim C_CONFIRM_YN
Dim C_BILL_YN
Dim C_TAX_DT
Dim C_REMARK

Dim IsOpenPop          

Dim FromDateOfDB
Dim ToDateOfDB

FromDateOfDB	= UNIConvDateAToB(UniDateAdd("m",-1,"<%=GetSvrDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
ToDateOfDB		= UNIConvDateAToB(UniDateAdd("m", 0,"<%=GetSvrDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)


'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================

Sub initSpreadPosVariables()  
	C_PROFORMA_NO = 1
	C_BP_CD 			= 2
	C_BP_POP 			= 3
	C_BP_NM 			= 4
	C_TEL_NO 			= 5
	C_CUST_USER 	= 6
	C_PROFORMA_DT = 7
	C_DOCUMENT 		= 8
	C_AMT 				= 9
	C_USR_NM 			= 10
	C_CONFIRM_YN 	= 11
	C_BILL_YN 		= 12
	C_TAX_DT 			= 13
	C_REMARK 			= 14

End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
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
	frm1.txtProFrDt.Text = FromDateOfDB
	frm1.txtProToDt.Text = ToDateOfDB	
	frm1.txtReqFrDt.Text = FromDateOfDB
	frm1.txtReqToDt.Text = ToDateOfDB
	frm1.txtBpCd.Focus
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)

End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
	Dim strFrDept, strToDept,IntRetCd
      
    lgKeyStream   = Frm1.txtBpCd.value & parent.gColSep 
    lgKeyStream   = lgKeyStream & Frm1.txtTelNo.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtProFrDt.Text & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtProToDt.Text & parent.gColSep    
    If frm1.rdoConfirm1.Checked Then
	    lgKeyStream   = lgKeyStream & parent.gColSep    
    ElseIf frm1.rdoConfirm2.Checked Then
  	  lgKeyStream   = lgKeyStream & "Y" & parent.gColSep    
  	Else
  	  lgKeyStream   = lgKeyStream & "N" & parent.gColSep    
    End If
    lgKeyStream   = lgKeyStream & Frm1.txtReqFrDt.Text & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtReqToDt.Text & parent.gColSep    
    If frm1.rdoAR1.Checked Then
	    lgKeyStream   = lgKeyStream & parent.gColSep    
    ElseIf frm1.rdoAR2.Checked Then
  	  lgKeyStream   = lgKeyStream & "Y" & parent.gColSep    
  	Else
  	  lgKeyStream   = lgKeyStream & "N" & parent.gColSep    
    End If
    lgKeyStream   = lgKeyStream & Frm1.txtDocument.value & parent.gColSep    
End Sub        

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	    .ReDraw = false
        .MaxCols = C_REMARK + 1												<%'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>
	    .Col = .MaxCols															<%'������Ʈ�� ��� Hidden Column%>
			.ColHeaderRows = 2
	    
        .ColHidden = True
        .MaxRows = 0
			Call GetSpreadColumnPos("A")  	

			ggoSpread.SSSetEdit   C_PROFORMA_NO     , "����NO", 20,,, 18, 2
			ggoSpread.SSSetEdit   C_BP_CD     			, "���ڵ�", 10,,, 10, 2
			ggoSpread.SSSetButton C_BP_POP
			ggoSpread.SSSetEdit  	C_BP_NM           , "����", 16
			ggoSpread.SSSetEdit   C_TEL_NO     			, "����ó", 10,,, 20, 1
			ggoSpread.SSSetEdit   C_CUST_USER       , "�������", 10,,, 20, 1
		  ggoSpread.SSSetDate		C_PROFORMA_DT			,	"������",		10,		2,					parent.gDateFormat
			ggoSpread.SSSetEdit   C_DOCUMENT   			, "����", 30,,, 50, 1
	  	ggoSpread.SSSetFloat	C_AMT							,	"�����ܰ�(�ݾ�)",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"			
			ggoSpread.SSSetEdit  	C_USR_NM          , "�ۼ���", 10
			ggoSpread.SSSetCheck	C_CONFIRM_YN			, "����Ȯ������",10,,,true
			ggoSpread.SSSetCheck	C_BILL_YN					, "�������뿩��",10,,,true
		  ggoSpread.SSSetDate		C_TAX_DT					,	"û�������",		10,		2,					parent.gDateFormat
			ggoSpread.SSSetEdit  	C_REMARK          , "���", 50,,, 50, 1
			        
'     call ggoSpread.SSSetColHidden(C_FLAG,C_FLAG,True)

			.Row = 1
			
			Call .AddCellSpan(0,							-1000,1,2)							'�����÷�, ���۷�, �����÷�, ������ 

			Call .AddCellSpan(C_PROFORMA_NO,	-1000,1,2)	'�����÷�, ���۷�, �����÷�, ������ 
			.Col = C_PROFORMA_NO :.Text ="����NO"
			
			Call .AddCellSpan(C_BP_CD,				-1000,5,1)	'�����÷�, ���۷�, �����÷�, ������ 
			.Col = C_BP_CD :.Text ="������"
			
			Call .AddCellSpan(C_PROFORMA_DT,	-1000,5,1)	'�����÷�, ���۷�, �����÷�, ������ 
			.Col = C_PROFORMA_DT :.Text ="��������"
			
			Call .AddCellSpan(C_BILL_YN,			-1000,2,1)	'�����÷�, ���۷�, �����÷�, ������ 			
			.Col = C_BILL_YN :.Text ="���⿬��"
			
			Call .AddCellSpan(C_REMARK,				-1000,1,2)	'�����÷�, ���۷�, �����÷�, ������ 
			
			.Row = 2
			.Col = C_BP_CD :.Text ="���ڵ�"
			.Col = C_BP_NM :.Text ="����"
			.Col = C_TEL_NO :.Text ="����ó"
			.Col = C_CUST_USER :.Text ="�������"
			.Col = C_PROFORMA_DT :.Text ="������"
			.Col = C_DOCUMENT :.Text ="����"
			.Col = C_AMT :.Text ="�����ܰ�(�ݾ�)"
			.Col = C_USR_NM :.Text ="�ۼ���"
			.Col = C_CONFIRM_YN :.Text ="����Ȯ������"
			.Col = C_BILL_YN :.Text ="�������뿩��"
			.Col = C_TAX_DT :.Text ="û�������"
			.Col = C_REMARK :.Text ="���"
					        

			.rowheight(-1000) =13	' ���� ������ 
			.rowheight(-999) = 13	' ���� ������ 

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

			C_PROFORMA_NO = iCurColumnPos(1)
			C_BP_CD 			= iCurColumnPos(2)
			C_BP_POP 			= iCurColumnPos(3)
			C_BP_NM 			= iCurColumnPos(4)
			C_TEL_NO 			= iCurColumnPos(5)
			C_CUST_USER 	= iCurColumnPos(6)
			C_PROFORMA_DT = iCurColumnPos(7)
			C_DOCUMENT 		= iCurColumnPos(8)
			C_AMT 				= iCurColumnPos(9)
			C_USR_NM 			= iCurColumnPos(10)
			C_CONFIRM_YN 	= iCurColumnPos(11)
			C_BILL_YN 		= iCurColumnPos(12)
			C_TAX_DT 			= iCurColumnPos(13)
			C_REMARK 			= iCurColumnPos(14)

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock    C_PROFORMA_NO, -1, C_PROFORMA_NO
        ggoSpread.SSSetRequired	C_BP_CD, -1, -1
        ggoSpread.SpreadLock    C_BP_NM, -1, C_BP_NM
        ggoSpread.SpreadLock    C_USR_NM, -1, C_USR_NM        
   	    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1 
        .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
       .vspdData.ReDraw = False
         ggoSpread.SSSetRequired		C_PROFORMA_NO, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_BP_CD, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_BP_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_USR_NM, pvStartRow, pvEndRow
         
       .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
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
  Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock Field
            
  Call InitSpreadSheet                                                            'Setup the Spread sheet
  Call InitVariables                                                              'Initializes local global variables
        
  Call SetDefaultVal
    
  Call SetToolbar("1100110100101111")										        '��ư ���� ���� 
       
End Sub

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
    Dim strFrDept, strToDept
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If   
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      															
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    Call InitVariables                                                        '��: Initializes local global variables
    Call MakeKeyStream("X")

    Call DisableToolBar(parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True                                                              '��: Processing is OK

End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    FncDelete = True                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim strReturn_value, strSQL
    Dim HFlag,MFlag,Rowcnt
    Dim strVdate
    Dim strWhere
    Dim strDay_time
    
    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
       Exit Function
    End If

    FncSave = True                                            
    
		Call DisableToolBar(parent.TBC_SAVE)
		If DbSave = False Then                                    '��: Save db data     Processing is OK
			Call RestoreToolBar()
      Exit Function
    End If
    
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    FncCopy = False           
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
            .Col = C_PROFORMA_NO
            .Text = ""
		    .Focus
		    .Action = 0 ' go to 
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
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

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
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
    Call parent.FncExport(parent.C_MULTI)                                         '��: ȭ�� ���� 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '��:ȭ�� ����, Tab ���� 
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

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
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
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '��: Clear err status

	 If LayerShowHide(1) = False then
    		Exit Function 
    	End if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    End With
		
    If lgIntFlgMode = parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal, strDel
	
    DbSave = False                                                          
    
     If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        
               Case ggoSpread.InsertFlag                                      '��: Insert
                    strVal = strVal & "C" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_PROFORMA_NO,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_BP_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_TEL_NO,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CUST_USER,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_PROFORMA_DT,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_DOCUMENT,lRow,"X","X") & parent.gColSep
										strVal = strVal & UniConvNum(GetSpreadText(frm1.vspdData,C_AMT,lRow,"X","X"),0) & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CONFIRM_YN,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_BILL_YN,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_TAX_DT,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_REMARK,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '��: Update
                    strVal = strVal & "U" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_PROFORMA_NO,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_BP_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_TEL_NO,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CUST_USER,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_PROFORMA_DT,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_DOCUMENT,lRow,"X","X") & parent.gColSep
										strVal = strVal & UniConvNum(GetSpreadText(frm1.vspdData,C_AMT,lRow,"X","X"),0) & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CONFIRM_YN,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_BILL_YN,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_TAX_DT,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_REMARK,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '��: Delete

                    strDel = strDel & "D" & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_PROFORMA_NO,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_BP_CD,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_TEL_NO,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_CUST_USER,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_PROFORMA_DT,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_DOCUMENT,lRow,"X","X") & parent.gColSep
										strDel = strDel & UniConvNum(GetSpreadText(frm1.vspdData,C_AMT,lRow,"X","X"),0) & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_CONFIRM_YN,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_BILL_YN,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_TAX_DT,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_REMARK,lRow,"X","X") & parent.gColSep
                    strDel = strDel & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

       .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '��: Processing is NG

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '��:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If    
    
    Call DisableToolBar(parent.TBC_DELETE)
	If DbDelete = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncDelete = True                                                        '��: Processing is OK

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
	Call SetToolbar("110011110011111")									
	frm1.vspdData.focus	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
    Call InitVariables															'��: Initializes local global variables
	Call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'==========================================================================================================
Function OpenBp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "B_BIZ_PARTNER"							
	arrParam(2) = Trim(frm1.txtBpCd.value)		
	'arrParam(3) = Trim(frm1.txtBpCd.value)	
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag = " & FilterVar("Y", "''", "S") & " "							
	arrParam(5) = "����"							
	arrParam(0) = arrParam(5)								
		
	arrField(0) = "BP_CD"									
	arrField(1) = "BP_NM"									
	arrField(2) = "BP_RGST_NO"
	
	arrHeader(0) = "����"							
	arrHeader(1) = "�����"						
	arrHeader(2) = "����ڵ�Ϲ�ȣ"
	    
	If frm1.txtBpCd.readOnly = True Then
		IsOpenPop = False
		Exit Function
	End If
												
	If UCase(frm1.txtBpCd.className) = parent.UCN_PROTECTED Then 
		IsOpenPop = False			
		Exit Function
	End IF
	
	frm1.txtBpCd.focus 
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtBpCd.value = arrRet(0)
		frm1.txtBpNm.value = arrRet(1)
	End If	
End Function

'==========================================================================================================
Function OpenBp2(pRow)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "B_BIZ_PARTNER"							
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_BP_CD,pRow,"X","X"))
	'arrParam(3) = Trim(frm1.txtBpCd.value)	
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag = " & FilterVar("Y", "''", "S") & " "							
	arrParam(5) = "����"							
	arrParam(0) = arrParam(5)								
		
	arrField(0) = "BP_CD"									
	arrField(1) = "BP_NM"									
	arrField(2) = "BP_RGST_NO"
	
	arrHeader(0) = "����"							
	arrHeader(1) = "�����"						
	arrHeader(2) = "����ڵ�Ϲ�ȣ"	    												
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		call frm1.vspdData.SetText(C_BP_CD,pRow,arrRet(0))
		call frm1.vspdData.SetText(C_BP_NM,pRow,arrRet(1))
	End If	
End Function
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	ggoSpread.Source = frm1.vspdData
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
    
    If Row > 0 Then
		Select Case Col
			Case C_BP_POP
				call OpenBp2(Row)
		End Select    
	End If
            
End Sub
'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    'Select Case Col
    '     Case  C_EMPNO
    'End Select    
             	
	ggoSpread.Source = frm1.vspdData
  ggoSpread.UpdateRow Row
End Function

'========================================================================================================
'   Event Name : txtBpCd_OnChange
'   Event Desc :
'========================================================================================================
Function txtBpCd_OnChange()    
    Dim IntRetCd

    If frm1.txtBpCd.value = "" Then
        frm1.txtBpNm.value = ""
    ELSE    
        IntRetCd = CommonQueryRs(" bp_nm "," b_biz_partner "," bp_cd =  " & FilterVar(frm1.txtBpCd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
            Call DisplayMsgBox("17A003","X","�ŷ�ó�ڵ�","X")                         '�� : �ش����� �������� �ʽ��ϴ�.            
            frm1.txtBpNm.value=""
            frm1.txtBpCd.focus
						txtBpCd_OnChange = true
						Exit Function
        Else
            frm1.txtBpNm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
   
End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")       

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
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================

Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
        End If
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

'=======================================================================================================
'   Event Name : txtProFrDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtProFrDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtProFrDt.Action = 7
        frm1.txtProFrDt.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtProToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtProToDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtProToDt.Action = 7
        frm1.txtProToDt.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtReqFrDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReqFrDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtReqFrDt.Action = 7
        frm1.txtReqFrDt.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtReqToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReqToDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtReqToDt.Action = 7
        frm1.txtReqToDt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtProFrDt_Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================
Sub txtProFrDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub
'=======================================================================================================
'   Event Name : txtProToDt_Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================
Sub txtProToDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub
'=======================================================================================================
'   Event Name : txtReqFrDt_Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================
Sub txtReqFrDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub
'=======================================================================================================
'   Event Name : txtReqToDt_Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================
Sub txtReqToDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- space Area-->

	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>����������</font></td>
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
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBpCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU" class=required STYLE="text-transform:uppercase" ALT="��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp()" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtBpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14" class = protected readonly = true TABINDEX="-1"></TD>
								<TD CLASS="TD5" NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTelNo" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="11XXXU" class=required STYLE="text-transform:uppercase" ALT="��"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtProFrDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="������"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtProToDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="������"></OBJECT>');</SCRIPT>
									</TD>
								<TD CLASS="TD5" NOWRAP>����Ȯ������</TD>
								<TD CLASS="TD6" NOWRAP>
									
										<input type=radio CLASS="RADIO" name="rdoConfirm" id="rdoConfirm1" value="" tag = "11" checked>
											<label for="rdoConfirm1">��ü</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoConfirm" id="rdoConfirm2" value="Y" tag = "11">
											<label for="rdoConfirm2">Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoConfirm" id="rdoConfirm3" value="N" tag = "11">
											<label for="rdoConfirm3">��Ȯ��</label></TD>

									</TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>û����</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtReqFrDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="û����"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtReqToDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="û����"></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS="TD5" NOWRAP>�������뿩��</TD>
								<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoAR" id="rdoAR1" value="" tag = "11" checked>
											<label for="rdoAR1">��ü</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoAR" id="rdoAR2" value="Y" tag = "11">
											<label for="rdoAR2">Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoAR" id="rdoAR3" value="N" tag = "11">
											<label for="rdoAR3">��Ȯ��</label></TD>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocument" TYPE="Text" MAXLENGTH="50" SIZE=30 tag="11XXXU" class=required STYLE="text-transform:uppercase" ALT="��"></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
								<TD HEIGHT=100% WIDTH=100% >
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

