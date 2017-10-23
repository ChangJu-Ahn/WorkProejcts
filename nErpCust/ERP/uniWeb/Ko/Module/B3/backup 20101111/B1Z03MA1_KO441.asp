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
Const BIZ_PGM_ID = "B1Z03MB1_KO441.asp"                                      '비지니스 로직 ASP명 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd

Dim C_SEQ
Dim C_TODO_DOC
Dim C_COMBO_YN
Dim C_UD_MAJOR_CD
Dim C_UD_MAJOR_POP
Dim C_UD_MAJOR_NM
Dim C_UD_MINOR_CD
Dim C_UD_MINOR_POP
Dim C_SAMPLE_DATA
Dim C_PROCESS_TYPE
Dim C_MES_USE_YN
Dim C_CDN_BIZ
Dim C_CDN_BMP
Dim C_CDN_PKG
Dim C_CDN_PRD
Dim C_CDN_TQC
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

	C_SEQ						= 1
	C_TODO_DOC			= 2
	C_COMBO_YN			= 3
	C_UD_MAJOR_CD		= 4
	C_UD_MAJOR_POP	= 5
	C_UD_MAJOR_NM		= 6
	C_UD_MINOR_CD		= 7
	C_UD_MINOR_POP	= 8
	C_SAMPLE_DATA		= 9
	C_PROCESS_TYPE	= 10
	C_MES_USE_YN		= 11
	C_CDN_BIZ				= 12
	C_CDN_BMP				= 13
	C_CDN_PKG				= 14
	C_CDN_PRD				= 15
	C_CDN_TQC				= 16
	C_REMARK				= 17

End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
  lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
  lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================	
Sub SetDefaultVal()	
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
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)

End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
End Sub        

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
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
			.MaxCols = C_REMARK + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
			.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
			.ColHidden = True
			.MaxRows = 0
			Call GetSpreadColumnPos("A")  	

			ggoSpread.SSSetEdit   C_SEQ     		, "순번", 10,,, 3, 2
			ggoSpread.SSSetEdit   C_TODO_DOC    , "항목", 20,,, 50, 1
			ggoSpread.SSSetCheck	C_COMBO_YN		, "선택",	10,,,true
			ggoSpread.SSSetEdit   C_UD_MAJOR_CD	, "코드그룹", 10,,, 10, 2
			ggoSpread.SSSetButton C_UD_MAJOR_POP
			ggoSpread.SSSetEdit   C_UD_MAJOR_NM	, "코드그룹명", 20
			ggoSpread.SSSetEdit   C_UD_MINOR_CD	, "공통코드", 10,,, 10, 2
			ggoSpread.SSSetButton C_UD_MINOR_POP
			ggoSpread.SSSetEdit   C_SAMPLE_DATA	, "Sample Data", 20,,, 50, 1
			ggoSpread.SSSetEdit   C_PROCESS_TYPE, "Process Type", 20,,, 20, 1
			ggoSpread.SSSetCheck	C_MES_USE_YN	, "Mes Code운영",10,,,true
			ggoSpread.SSSetCheck	C_CDN_BIZ			, "영업등록",10,,,true
			ggoSpread.SSSetCheck	C_CDN_BMP			, "기술등록(Bump)",10,,,true
			ggoSpread.SSSetCheck	C_CDN_PKG			, "기술등록(Pkg)",10,,,true
			ggoSpread.SSSetCheck	C_CDN_PRD			, "품질등록",10,,,true
			ggoSpread.SSSetCheck	C_CDN_TQC			, "생산등록",10,,,true			
			ggoSpread.SSSetEdit  	C_REMARK      , "비고", 50,,, 50, 1
 
        call ggoSpread.SSSetColHidden(C_SEQ,C_SEQ,True)
		        
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


						C_SEQ						= iCurColumnPos(1)
						C_TODO_DOC			= iCurColumnPos(2)
						C_COMBO_YN			= iCurColumnPos(3)
						C_UD_MAJOR_CD		= iCurColumnPos(4)
						C_UD_MAJOR_POP	= iCurColumnPos(5)
						C_UD_MAJOR_NM		= iCurColumnPos(6)
						C_UD_MINOR_CD		= iCurColumnPos(7)
						C_UD_MINOR_POP	= iCurColumnPos(8)
						C_SAMPLE_DATA		= iCurColumnPos(9)
						C_PROCESS_TYPE	= iCurColumnPos(10)
						C_MES_USE_YN		= iCurColumnPos(11)
						C_CDN_BIZ				= iCurColumnPos(12)
						C_CDN_BMP				= iCurColumnPos(13)
						C_CDN_PKG				= iCurColumnPos(14)
						C_CDN_PRD				= iCurColumnPos(15)
						C_CDN_TQC				= iCurColumnPos(16)
						C_REMARK				= iCurColumnPos(17)

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False

				ggoSpread.SSSetRequired		C_TODO_DOC, -1, -1
        ggoSpread.SpreadLock    	C_UD_MAJOR_NM, -1, C_UD_MAJOR_NM
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
       
         ggoSpread.SSSetRequired		C_TODO_DOC, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_UD_MAJOR_CD, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_UD_MAJOR_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_UD_MINOR_CD, pvStartRow, pvEndRow
         
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
    
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
  Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
  Call InitSpreadSheet                                                            'Setup the Spread sheet
  Call InitVariables                                                              'Initializes local global variables
        
  Call SetDefaultVal
    
  Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
       
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
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If   
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                        '⊙: Initializes local global variables
    Call MakeKeyStream("X")

    Call DisableToolBar(parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True                                                              '☜: Processing is OK

End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                            '☜: Processing is OK
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
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

    FncSave = True                                            
    
		Call DisableToolBar(parent.TBC_SAVE)
		If DbSave = False Then                                    '☜: Save db data     Processing is OK
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
            .Col = C_SEQ
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
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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
    
    Err.Clear                                                                        '☜: Clear err status

	 If LayerShowHide(1) = False then
    		Exit Function 
    	End if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
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
        
               Case ggoSpread.InsertFlag                                      '☜: Insert
                    strVal = strVal & "C" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_SEQ,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_TODO_DOC,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_COMBO_YN,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_UD_MINOR_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_SAMPLE_DATA,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_PROCESS_TYPE,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_MES_USE_YN,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_BIZ,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_BMP,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_PKG,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_PRD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_TQC,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_REMARK,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '☜: Update
                    strVal = strVal & "U" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_SEQ,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_TODO_DOC,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_COMBO_YN,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_UD_MINOR_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_SAMPLE_DATA,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_PROCESS_TYPE,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_MES_USE_YN,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_BIZ,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_BMP,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_PKG,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_PRD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_CDN_TQC,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_REMARK,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                    strDel = strDel & "D" & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_SEQ,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_TODO_DOC,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_COMBO_YN,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_UD_MINOR_CD,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_SAMPLE_DATA,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_PROCESS_TYPE,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_MES_USE_YN,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_CDN_BIZ,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_CDN_BMP,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_CDN_PKG,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_CDN_PRD,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_CDN_TQC,lRow,"X","X") & parent.gColSep
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
    
    FncDelete = False                                                      '⊙: Processing is NG

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If    
    
    Call DisableToolBar(parent.TBC_DELETE)
	If DbDelete = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncDelete = True                                                        '⊙: Processing is OK

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    
    Dim iCnt
    For iCnt=1 To frm1.vspdData.MaxRows
	    Call SetSpreadColorAfterQuery(C_COMBO_YN, iCnt)
    Next
    
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
    Call InitVariables															'⊙: Initializes local global variables
	Call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'==========================================================================================================
Function OpenMajor(pRow)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사용자Major코드 팝업"			' 팝업 명칭 
	arrParam(1) = "B_User_Defined_MAJOR"		 		' TABLE 명칭 
	arrParam(2) = GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,pRow,"X","X")				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = " ud_major_cd like 'ZZ9%'"									' Where Condition
	arrParam(5) = "사용자Major코드"			
	
    arrField(0) = "ud_major_cd"							' Field명(0)
    arrField(1) = "ud_major_nm"							' Field명(1)
    
    arrHeader(0) = "사용자Major코드"				' Header명(0)
    arrHeader(1) = "사용자Major코드명"				' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_UD_MAJOR_CD,pRow,arrRet(0))
		Call frm1.vspdData.SetText(C_UD_MAJOR_NM,pRow,arrRet(1))
		call vspdData_Change(C_UD_MAJOR_CD , pRow)
	End If	

End Function

'==========================================================================================================
Function OpenMinor(pRow)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,pRow,"X","X") = "" Then
		Call DisplayMsgBox("971012", "X", "코드그룹", "X")
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사용자Minor코드 팝업"			' 팝업 명칭 
	arrParam(1) = "B_User_Defined_MINOR"		 		' TABLE 명칭 
	arrParam(2) = GetSpreadText(frm1.vspdData,C_UD_MINOR_CD,pRow,"X","X")				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "UD_MAJOR_CD=" & FilterVar(GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,pRow,"X","X"),"''","S")									' Where Condition
	arrParam(5) = "사용자Minor코드"			
	
    arrField(0) = "UD_MINOR_CD"							' Field명(0)
    arrField(1) = "UD_MINOR_NM"							' Field명(1)
    arrField(2) = "UD_REFERENCE"						' Field명(2)
    
    arrHeader(0) = "사용자Major코드"				' Header명(0)
    arrHeader(1) = "사용자Major코드명"			' Header명(1)
    arrHeader(2) = "Reference"							' Header명(2)
    


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_UD_MINOR_CD,pRow,arrRet(0))
		Call frm1.vspdData.SetText(C_SAMPLE_DATA,pRow,arrRet(2))
		call vspdData_Change(C_UD_MINOR_CD , pRow)
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
			Case C_UD_MAJOR_POP
				call OpenMajor(Row)
			Case C_UD_MINOR_POP
				call OpenMinor(Row)
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

    Call SetSpreadColorAfterQuery(Col,Row)
    
	ggoSpread.Source = frm1.vspdData
  ggoSpread.UpdateRow Row
End Function

Function SetSpreadColorAfterQuery(Col, Row)
    With frm1
    
       .vspdData.ReDraw = False

    Select Case Col
         Case  C_COMBO_YN
         	If GetSpreadText(frm1.vspdData,C_COMBO_YN,Row,"X","X")="1" Then
		        ggoSpread.SpreadUnLock    	C_UD_MAJOR_CD, Row, C_UD_MAJOR_CD, Row
						ggoSpread.SSSetRequired			C_UD_MAJOR_CD, Row, Row
		        ggoSpread.SpreadUnLock    	C_UD_MINOR_CD, Row, C_UD_MINOR_CD, Row
						ggoSpread.SSSetRequired			C_UD_MINOR_CD, Row, Row
         	Else
         		Call frm1.vspdData.SetText(C_UD_MAJOR_CD,Row,"")
         		Call frm1.vspdData.SetText(C_UD_MAJOR_NM,Row,"")
         		Call frm1.vspdData.SetText(C_UD_MINOR_CD,Row,"")
		        ggoSpread.SSSetProtected    	C_UD_MAJOR_CD, Row, Row
		        ggoSpread.SSSetProtected    	C_UD_MINOR_CD, Row, Row
         	End If
    End Select    
       .vspdData.ReDraw = True
    
    End With
End Function


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>CDN 부서별 TODO LIST등록</font></td>
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
<!--				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>고객</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBpCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU" class=required STYLE="text-transform:uppercase" ALT="고객"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp()" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtBpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14" class = protected readonly = true TABINDEX="-1"></TD>
								<TD CLASS="TD5" NOWRAP>연락처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTelNo" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="11XXXU" class=required STYLE="text-transform:uppercase" ALT="고객"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>견적일</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtProFrDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="견적일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtProToDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="견적일"></OBJECT>');</SCRIPT>
									</TD>
								<TD CLASS="TD5" NOWRAP>견적확정여부</TD>
								<TD CLASS="TD6" NOWRAP>
									
										<input type=radio CLASS="RADIO" name="rdoConfirm" id="rdoConfirm1" value="" tag = "11" checked>
											<label for="rdoConfirm1">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoConfirm" id="rdoConfirm2" value="Y" tag = "11">
											<label for="rdoConfirm2">확정</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoConfirm" id="rdoConfirm3" value="N" tag = "11">
											<label for="rdoConfirm3">미확정</label></TD>

									</TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>청구일</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtReqFrDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="청구일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtReqToDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="청구일"></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS="TD5" NOWRAP>매출적용여부</TD>
								<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoAR" id="rdoAR1" value="" tag = "11" checked>
											<label for="rdoAR1">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoAR" id="rdoAR2" value="Y" tag = "11">
											<label for="rdoAR2">확정</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoAR" id="rdoAR3" value="N" tag = "11">
											<label for="rdoAR3">미확정</label></TD>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>내용</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocument" TYPE="Text" MAXLENGTH="50" SIZE=30 tag="11XXXU" class=required STYLE="text-transform:uppercase" ALT="고객"></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>	-->
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

