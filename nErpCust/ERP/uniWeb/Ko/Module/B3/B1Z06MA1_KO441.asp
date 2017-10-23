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
Const BIZ_PGM_ID = "B1Z06MB1_KO441.asp"                                      '비지니스 로직 ASP명 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd

Dim C_MES_SND_YN
Dim C_USR_NM
Dim C_NOTE_DT
Dim C_DP_GUBUN
Dim C_ITEM_CD
Dim C_ITEM_NM
Dim C_CBM_DESCRIPTION
Dim C_CDN_BIZ
Dim C_CDN_BMP
Dim C_CDN_PKG
Dim C_CDN_PRD
Dim C_CDN_TQC
Dim C_CONFIRM_FLG


Dim IsOpenPop          

Dim FromDateOfDB
Dim ToDateOfDB

FromDateOfDB	= UNIConvDateAToB(UniDateAdd("d",-15,"<%=GetSvrDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
ToDateOfDB		= UNIConvDateAToB(UniDateAdd("d", 15,"<%=GetSvrDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)


'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================

Sub initSpreadPosVariables()  


	C_MES_SND_YN			= 1
	C_USR_NM					= 2
	C_NOTE_DT					= 3
	C_DP_GUBUN				= 4
	C_ITEM_CD					= 5
	C_ITEM_NM					= 6
	C_CBM_DESCRIPTION	= 7
	C_CDN_BIZ					= 8
	C_CDN_BMP					= 9
	C_CDN_PKG					= 10
	C_CDN_PRD					= 11
	C_CDN_TQC					= 12
	C_CONFIRM_FLG			= 13

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
	frm1.txtFrDt.Text = FromDateOfDB
	frm1.txtToDt.Text = ToDateOfDB
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
			ggoSpread.Spreadinit "V20081125",,parent.gAllowDragDropSpread    
			.ReDraw = false
			.MaxCols = C_CONFIRM_FLG + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
			.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
			.ColHidden = True
			.MaxRows = 0
			Call GetSpreadColumnPos("A")  	

			ggoSpread.SSSetCheck	C_MES_SND_YN			, "전송구분",10,,,true
			ggoSpread.SSSetEdit   C_USR_NM					, "작성자", 10
   	  ggoSpread.SSSetDate 	C_NOTE_DT					, "작성일", 10, 2, Parent.gDateFormat
			ggoSpread.SSSetEdit   C_DP_GUBUN				, "D/P", 10
			ggoSpread.SSSetEdit   C_ITEM_CD					, "품목코드", 18
			ggoSpread.SSSetEdit   C_ITEM_NM					, "품목명", 20
			ggoSpread.SSSetEdit   C_CBM_DESCRIPTION	, "Mes Device", 20
			ggoSpread.SSSetCheck	C_CDN_BIZ					, "승인(영업)",10,,,true
			ggoSpread.SSSetCheck	C_CDN_BMP					, "승인(기술Bump)",10,,,true
			ggoSpread.SSSetCheck	C_CDN_PKG					, "승인(기술Pkg)",10,,,true
			ggoSpread.SSSetCheck	C_CDN_PRD					, "승인(품질)",10,,,true
			ggoSpread.SSSetCheck	C_CDN_TQC					, "승인(생산)",10,,,true			
			ggoSpread.SSSetCheck	C_CONFIRM_FLG			, "확정여부",10,,,true			
 
      Call ggoSpread.SSSetColHidden(C_CONFIRM_FLG,C_CONFIRM_FLG,True)
		        
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

						C_MES_SND_YN			= iCurColumnPos(1)
						C_USR_NM					= iCurColumnPos(2)
						C_NOTE_DT					= iCurColumnPos(3)
						C_DP_GUBUN				= iCurColumnPos(4)
						C_ITEM_CD					= iCurColumnPos(5)
						C_ITEM_NM					= iCurColumnPos(6)
						C_CBM_DESCRIPTION	= iCurColumnPos(7)
						C_CDN_BIZ					= iCurColumnPos(8)
						C_CDN_BMP					= iCurColumnPos(9)
						C_CDN_PKG					= iCurColumnPos(10)
						C_CDN_PRD					= iCurColumnPos(11)
						C_CDN_TQC					= iCurColumnPos(12)
						C_CONFIRM_FLG			= iCurColumnPos(13)

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock    	-1, -1
   	    
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
    
  Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
       
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
        strVal = strVal     & "&txtItemCd="       & frm1.txtItemCd.value
        strVal = strVal     & "&txtFrDt="       & frm1.txtFrDt.Text
        strVal = strVal     & "&txtToDt="       & frm1.txtToDt.Text
        strVal = strVal     & "&txtDevice="       & frm1.txtDevice.value
        
        If frm1.rdoDp1.Checked = True Then
	        strVal = strVal     & "&rdoDp="
        ElseIf frm1.rdoDp2.Checked = True Then
	        strVal = strVal     & "&rdoDp=Y"
        Else
	        strVal = strVal     & "&rdoDp=N"
        End If
        
        If frm1.rdoTrans1.Checked = True Then
	        strVal = strVal     & "&rdoTrans=Y"
        ElseIf frm1.rdoTrans2.Checked = True Then
	        strVal = strVal     & "&rdoTrans=N"
        End If

        If frm1.rdoCfmDept1.Checked = True Then
	        strVal = strVal     & "&rdoCfmDept="
        ElseIf frm1.rdoCfmDept2.Checked = True Then
	        strVal = strVal     & "&rdoCfmDept=SO"
        ElseIf frm1.rdoCfmDept3.Checked = True Then
	        strVal = strVal     & "&rdoCfmDept=TB"
        ElseIf frm1.rdoCfmDept4.Checked = True Then
	        strVal = strVal     & "&rdoCfmDept=TP"
        ElseIf frm1.rdoCfmDept5.Checked = True Then
	        strVal = strVal     & "&rdoCfmDept=QA"
        ElseIf frm1.rdoCfmDept6.Checked = True Then
	        strVal = strVal     & "&rdoCfmDept=PP"
        End If

        If frm1.rdoCfm1.Checked = True Then
	        strVal = strVal     & "&rdoCfm="
        ElseIf frm1.rdoCfm2.Checked = True Then
	        strVal = strVal     & "&rdoCfm=Y"
        Else
	        strVal = strVal     & "&rdoCfm=N"
        End If

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
                            
               Case ggoSpread.UpdateFlag                                      '☜: Update
                    strVal = strVal & "U" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_ITEM_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_MES_SND_YN,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
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
	    Call SetSpreadColorAfterQuery(C_CONFIRM_FLG, iCnt)
    Next
    
	Call SetToolbar("110010010001111")									
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
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목코드 팝업"			' 팝업 명칭 
	arrParam(1) = "B_CDN_REQ_HDR_KO441"		 		' TABLE 명칭 
	arrParam(2) = frm1.txtItemCd.Value				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "품목코드"			
	
    arrField(0) = "ITEM_CD"							' Field명(0)
    arrField(1) = "ITEM_NM"							' Field명(1)
    
    arrHeader(0) = "품목코드"				' Header명(0)
    arrHeader(1) = "품목명"			' Header명(1)    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1)
		frm1.txtItemCd.focus
	End If	

End Function

'==========================================================================================================
Function OpenDevice()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Device 팝업"			' 팝업 명칭 
	arrParam(1) = "B_CDN_REQ_HDR_KO441"		 		' TABLE 명칭 
	arrParam(2) = frm1.txtDevice.Value				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "Device코드"			
	
    arrField(0) = "CBM_DESCRIPTION"							' Field명(0)
    arrField(1) = "ITEM_NM"							' Field명(1)
    
    arrHeader(0) = "Device"				' Header명(0)
    arrHeader(1) = "품목명"			' Header명(1)    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtDevice.value = arrRet(0)
		frm1.txtDevice.focus
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

Function SetSpreadColorAfterQuery(Col, Row)
    With frm1
    
       .vspdData.ReDraw = False

    Select Case Col
         Case  C_CONFIRM_FLG
         	If GetSpreadText(frm1.vspdData,C_CONFIRM_FLG,Row,"X","X")="1" Then
		        ggoSpread.SpreadUnLock    		C_MES_SND_YN, Row, C_MES_SND_YN, Row
         	Else
		        ggoSpread.SSSetProtected    	C_MES_SND_YN, Row, Row
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

'=======================================================================================================
'   Event Name : txtFrDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFrDt.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>CDN 승인코드 MES전송</font></td>
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
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 CLASS=required STYLE="text-transform:uppercase" tag="11XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" CLASS=protected READONLY=true TABINDEX="-1" SIZE=30 tag="14"></TD>
							    <TD CLASS="TD5" NOWRAP>작성일자</TD>
									<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtFrDt name=txtFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="작성일자"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtToDt name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="작성일자"></OBJECT>');</SCRIPT>
									</TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>Mes Device</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDevice" SIZE=20 MAXLENGTH=18 CLASS=required STYLE="text-transform:uppercase" tag="11XXXU"  ALT="Mes Device"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDevice()"></TD>
							    <TD CLASS="TD5" NOWRAP>D/P</TD>
									<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoDp" id="rdoDp1" value="" tag = "11" checked>
											<label for="rdoDp1">All</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoDp" id="rdoDp2" value="Y" tag = "11">
											<label for="rdoDp2">Development</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoDp" id="rdoDp3" value="N" tag = "11">
											<label for="rdoDp3">Production</label></TD>
									</TD>
								</TR>	
								<TR>
							    <TD CLASS="TD5" NOWRAP>전송구분</TD>
									<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoTrans" id="rdoTrans1" value="Y" tag = "11">
											<label for="rdoTrans1">전송</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoTrans" id="rdoTrans2" value="N" tag = "11" checked>
											<label for="rdoTrans2">미전송</label></TD>
									</TD>
							    <TD CLASS="TD5" NOWRAP>승인부서</TD>
									<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoCfmDept" id="rdoCfmDept1" value="" tag = "11" checked>
											<label for="rdoCfmDept1">전체</label>
										<input type=radio CLASS="RADIO" name="rdoCfmDept" id="rdoCfmDept2" value="S" tag = "11">
											<label for="rdoCfmDept2">영업</label>
										<input type=radio CLASS="RADIO" name="rdoCfmDept" id="rdoCfmDept3" value="B" tag = "11">
											<label for="rdoCfmDept3">기술(BUMP)</label>
										<input type=radio CLASS="RADIO" name="rdoCfmDept" id="rdoCfmDept4" value="P" tag = "11">
											<label for="rdoCfmDept4">기술(PKG)</label>
										<input type=radio CLASS="RADIO" name="rdoCfmDept" id="rdoCfmDept5" value="Q" tag = "11">
											<label for="rdoCfmDept5">품질</label>
										<input type=radio CLASS="RADIO" name="rdoCfmDept" id="rdoCfmDept6" value="M" tag = "11">
											<label for="rdoCfmDept6">생산관리</label>
											</TD>
									</TD>
							<TR>
								<TD CLASS="TD5" NOWRAP>승인구분</TD>
								<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoCfm" id="rdoCfm1" value="" tag = "11" style="display:none">
											<label for="rdoCfm1" style="display:none">전체</label>
										<input type=radio CLASS="RADIO" name="rdoCfm" id="rdoCfm2" value="Y" tag = "11">
											<label for="rdoCfm2">승인</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoCfm" id="rdoCfm3" value="N" tag = "11" checked>
											<label for="rdoCfm3">미승인</label>&nbsp;									
									</TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

