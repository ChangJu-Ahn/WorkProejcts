<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m1511ma1
'*  4. Program Name         : 입출고형태등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/26
'*  8. Modified date(Last)  : 2003/08/08
'*  9. Modifier (First)     : Shin jin hyun
'* 10. Modifier (Last)      : Kim Duk Hyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2003-05-20
'*        
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit              
'=======================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'=======================================================================================================================
Const BIZ_PGM_ID = "m1511mb1.asp" 

Dim C_IotypeCd
Dim C_IotypeNm
Dim C_GmCond		'입고여부 
Dim C_RetCond		'반품여부 
Dim C_SubCond		'사급여부 
Dim C_SubCond2		'외주가공여부 
Dim	C_ChildCond		'자품목 처리여부 
Dim C_ImportCond	'수입여부 
Dim C_UseCond		'사용여부 
Dim C_MovTypeCd
Dim C_MovTypePop
Dim C_MovTypeNm

'KSJ추가 2007.06.14 매입일괄등록 
Dim C_ExptIvTypeCd
Dim C_ExptIvTypePop
Dim C_ExptIvTypeNm

Dim lgQuery
Dim lgCopyRow
Dim IsOpenPop          

'=======================================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                 
    lgBlnFlgChgValue = False                  
    lgIntGrpCount = 0                         
    lgStrPrevKey = ""                         
    lgLngCurRows = 0                          
End Sub
'=======================================================================================================================
Sub initSpreadPosVariables()  
	C_IotypeCd   = 1
	C_IotypeNm   = 2
	C_GmCond     = 3
	C_RetCond    = 4
	C_SubCond    = 5
	C_ChildCond	 = 6    '자품목처리여부 
	C_SubCond2	 = 7	'자품목정산여부	
	C_ImportCond = 8    
	C_UseCond    = 9    
	C_MovTypeCd  = 10   
	C_MovTypePop = 11   
	C_MovTypeNm  = 12
	
	'KSJ추가 2007.06.14 매입일괄등록 
	C_ExptIvTypeCd  = 13
    C_ExptIvTypePop = 14
    C_ExptIvTypeNm  = 15

End Sub
'=======================================================================================================================
Sub SetDefaultVal()
	frm1.rdoUseflg(0).Checked = true
    Call SetToolbar("1110110100101111")
    frm1.txtGmTypeCd.focus 
	Set gActiveElement = document.activeElement
End Sub
'=======================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub
'=======================================================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()  
	ggoSpread.Source = frm1.vspdData  
	ggoSpread.Spreadinit "V20051201",,parent.gAllowDragDropSpread    
	
	With frm1.vspdData
	 
		.ReDraw = False
		 .MaxCols = C_ExptIvTypeNm+1      
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit  C_IotypeCd, "입출고형태", 10,,,5,2
		ggoSpread.SSSetEdit  C_IotypeNm, "입출고형태명", 20,,,50
		ggoSpread.SSSetCheck  C_GmCond,"입고여부",15,,,true
		ggoSpread.SSSetCheck  C_RetCond,"반품여부",15,,,true
		ggoSpread.SSSetCheck  C_SubCond,"사급여부",15,,,true
		ggoSpread.SSSetCheck  C_ChildCond,"자품목처리여부",15,,,true
		ggoSpread.SSSetCheck  C_SubCond2,"자품목정산여부",15,,,true
		ggoSpread.SSSetCheck  C_ImportCond,"수입여부",15,,,true
		ggoSpread.SSSetCheck  C_UseCond,"사용여부",15,,,true
		ggoSpread.SSSetEdit  C_MovTypeCd, "재고처리형태", 15,,,3,2
		ggoSpread.SSSetButton  C_MovTypePop
		ggoSpread.SSSetEdit  C_MovTypeNm, "재고처리형태명", 20
		
		'KSJ추가 2007.06.14 매입일괄등록 
		ggoSpread.SSSetEdit  C_ExptIvTypeCd , "일괄매입형태", 15,,,3,2
		ggoSpread.SSSetButton  C_ExptIvTypePop
		ggoSpread.SSSetEdit  C_ExptIvTypeNm, "일괄매입형태명", 20
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.MakePairsColumn(C_MovTypeCd, C_MovTypePop)
		Call ggoSpread.MakePairsColumn(C_ExptIvTypeCd, C_ExptIvTypePop)
		Call ggoSpread.SSSetSplit2(2)
		
		Call SetSpreadLock 
		    
		.ReDraw = true
	 
	End With
    
End Sub
'=======================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_IotypeCd   = iCurColumnPos(1)
			C_IotypeNm   = iCurColumnPos(2)
			C_GmCond     = iCurColumnPos(3)
			C_RetCond    = iCurColumnPos(4)
			C_SubCond    = iCurColumnPos(5)
			C_ChildCond  = iCurColumnPos(6)	'자품목 처리여부 
			C_SubCond2   = iCurColumnPos(7)
			C_ImportCond = iCurColumnPos(8)
			C_UseCond    = iCurColumnPos(9)
			C_MovTypeCd  = iCurColumnPos(10)
			C_MovTypePop = iCurColumnPos(11)
			C_MovTypeNm  = iCurColumnPos(12)
			
			'KSJ추가 2007.06.14 매입일괄등록 
			C_ExptIvTypeCd   = iCurColumnPos(13)
			C_ExptIvTypePop  = iCurColumnPos(14)
			C_ExptIvTypeNm   = iCurColumnPos(15)
		End Select    
End Sub
'=======================================================================================================================
Sub SetSpreadLock()
    With ggoSpread
    
		.spreadunlock  C_IotypeCd, -1
		.sssetrequired C_IotypeCd, -1
		.sssetrequired C_IotypeNm, -1
		.spreadunlock  C_GmCond,   -1
		.spreadunlock  C_RetCond,  -1
		.spreadunlock  C_SubCond,  -1
		.spreadunlock  C_SubCond2,  -1
		.spreadunlock  C_ChildCond, -1	'자품목 처리여부 
		.spreadunlock  C_ImportCond, -1
		.spreadunlock  C_UseCond,  -1
		.spreadunlock  C_MovTypecd, -1
		.sssetrequired C_MovTypeCd, -1
		.spreadlock  C_MovTypeNm, -1
		
		'KSJ추가 2007.06.14 매입일괄등록 
		.spreadunlock  C_ExptIvTypeCd, -1
		.spreadlock  C_ExptIvTypeNm, -1
		
		.SSSetProtected frm1.vspdData.MaxCols, -1
    End With
End Sub
'=======================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, Byval pvEndRow)
	frm1.vspdData.ReDraw = False
    With ggoSpread
		.spreadunlock  C_IotypeCd, pvStartRow, C_IotypeNm,	pvEndRow
		.sssetrequired C_IotypeCd, pvStartRow,  pvEndRow
		.sssetrequired C_IotypeNm, pvStartRow,  pvEndRow
'		.SSSetProtected C_SubCond, pvStartRow, pvEndRow
'		.SSSetProtected C_SubCond2, pvStartRow, pvEndRow
		.SSSetProtected C_MovTypeNm, pvStartRow, pvEndRow
		.spreadunlock  C_MovTypeCd, pvStartRow, C_MovTypeCd,		pvEndRow
		.sssetrequired C_MovTypeCd, pvStartRow,  pvEndRow
		.spreadlock  C_MovTypeNm, pvStartRow, C_MovTypeNm,		pvEndRow
		
		'KSJ추가 2007.06.14 매입일괄등록 
		.spreadunlock  C_ExptIvTypeCd, pvStartRow, C_ExptIvTypeCd,		pvEndRow
		.spreadlock  C_ExptIvTypeNm, pvStartRow, C_ExptIvTypeNm,		pvEndRow
		
		.SSSetProtected frm1.vspdData.MaxCols,	pvStartRow, pvEndRow
		
    End With
    frm1.vspdData.ReDraw = True
End Sub
'=======================================================================================================================
Function OpenIotype()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	 
	arrParam(0) = "입출고형태" 
	arrParam(1) = "M_MVMT_TYPE"    
	arrParam(2) = UCase(Trim(frm1.txtGmTypeCd.Value))
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "입출고형태" 
	 
	arrField(0) = "io_type_cd" 
	arrField(1) = "io_type_Nm" 
	    
	arrHeader(0) = "입출고형태" 
	arrHeader(1) = "입출고형태명" 
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtGmTypeCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtGmTypeCd.Value = arrRet(0)
		frm1.txtGmTypeNm.Value = arrRet(1)
		frm1.txtGmTypeCd.focus	
		Set gActiveElement = document.activeElement
	End If 
 
End Function
'=======================================================================================================================
Function OpenMovType()
 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCurRow = frm1.vspdData.ActiveRow
	 
	arrParam(0) = "재고처리형태" 
	arrParam(1) = "I_MOVETYPE_CONFIGURATION,b_minor"    
	'arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_MovTypeCd,iCurRow,"X","X")))
	arrParam(3) = ""
	arrParam(4) = "I_MOVETYPE_CONFIGURATION.MOV_TYPE=b_minor.minor_cd And b_minor.major_cd=" & FilterVar("I0001", "''", "S") & " and " & _
				  "(I_MOVETYPE_CONFIGURATION.TRNS_TYPE=" & FilterVar("PR", "''", "S") & " or I_MOVETYPE_CONFIGURATION.TRNS_TYPE=" & FilterVar("ST", "''", "S") & ")"
	arrParam(5) = "재고처리형태" 
	 
	arrField(0) = "I_MOVETYPE_CONFIGURATION.MOV_TYPE" 
	arrField(1) = "b_minor.minor_Nm" 
	    
	arrHeader(0) = "재고처리형태" 
	arrHeader(1) = "재고처리형태명" 
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	  
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_MovTypeCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_MovTypeNM,	iCurRow, arrRet(1))
		Call vspdData_Change(0, iCurRow)
	End If
 
End Function

'=======================================================================================================================
Function OpenExptIvtype()

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)
 Dim iCurRow
 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCurRow = frm1.vspdData.ActiveRow
	
	arrParam(0) = "매입형태" 
	arrParam(1) = "m_iv_type"    
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_ExptIvTypeCd,iCurRow,"X","X")))
	arrParam(3) = ""
	arrParam(4) = ""
	 
	arrParam(5) = "매입형태" 
	 
	arrField(0) = "iv_type_cd" 
	arrField(1) = "iv_type_nm" 
	    
	arrHeader(0) = "매입형태" 
	arrHeader(1) = "매입형태명" 
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_ExptIvTypeCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_ExptIvTypeNm,	iCurRow, arrRet(1))
		Call vspdData_Change(0, iCurRow)
	End If 
 
End Function

'=======================================================================================================================
Sub changeGm(ByVal curRow)
	
	if Trim(GetSpreadText(frm1.vspdData,C_GmCond,iCurRow,"X","X")) <> "1" then
		ggoSpread.spreadunlock  C_RetCond, curRow, C_RetCond, curRow
	else
		ggoSpread.SSSetProtected C_RetCond, curRow, curRow
		Call frm1.vspdData.SetText(C_RetCond,	curRow, "0")
	end if
	
End Sub
'=======================================================================================================================
Sub changeRet(ByVal curRow)
	
	if Trim(GetSpreadText(frm1.vspdData,C_RetCond,iCurRow,"X","X")) <> "1" then
		ggoSpread.spreadunlock  C_GmCond, curRow, C_GmCond, curRow
		ggoSpread.spreadunlock  C_SubCond, curRow, C_SubCond, curRow
	else
		ggoSpread.SSSetProtected C_GmCond, curRow, curRow
		ggoSpread.SSSetProtected C_SubCond, curRow, curRow
			
		Call frm1.vspdData.SetText(C_GmCond,	curRow, "0")
		Call frm1.vspdData.SetText(C_SubCond,	curRow, "0")
	end if
	
End Sub
'=======================================================================================================================
Sub changeSub(ByVal curRow)
	
	if Trim(GetSpreadText(frm1.vspdData,C_SubCond,iCurRow,"X","X")) <> "1" then
		ggoSpread.spreadunlock  C_RetCond, curRow, C_RetCond, curRow
	else
		ggoSpread.SSSetProtected C_RetCond, curRow, curRow
		Call frm1.vspdData.SetText(C_RetCond,	curRow, "0")
	end if
	
End Sub

'=======================================================================================================================
'	Name : changeExptIvType()
'	Description : 일괄매입여부 Check Event
'   KSJ추가 2007.06.14 매입일괄등록 
'=======================================================================================================================
sub changeExptIvType(ByVal curRow, ByVal Actionflg)
	With frm1
		If Trim(GetSpreadText(frm1.vspdData,C_ImportCond,curRow,"X","X")) = "1" OR Trim(GetSpreadText(frm1.vspdData,C_SubCond,curRow,"X","X")) = "1" Then
			
			.vspdData.Row = curRow
			.vspdData.Col = C_ExptIvTypeCd
			.vspdData.Text = ""
			
			.vspdData.Col = C_ExptIvTypeNm
			.vspdData.Text = ""
			
			ggoSpread.SSSetProtected C_ExptIvTypeCd, curRow, curRow
			ggoSpread.SSSetProtected C_ExptIvTypePop, curRow, curRow
			
			If Actionflg = True then
				Call .vspdData.SetText(C_ExptIvTypeCd,	curRow, "")
				Call .vspdData.SetText(C_ExptIvTypeNm,	curRow, "")
			End If
		else
			ggoSpread.spreadunlock  C_ExptIvTypeCd, curRow, C_ExptIvTypePop, curRow
		End if
		
		
	End With
End Sub
'=======================================================================================================================
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")           
    
    Call InitSpreadSheet                            
    Call SetDefaultVal
    Call InitVariables                              
End Sub
'=======================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	IF lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
	
	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
	   Exit Sub
	End If
	   	    
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If

	frm1.vspdData.Row = Row   
End Sub
'=======================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'=======================================================================================================================
Sub vspdData_MouseDown(ByVal Button , ByVal Shift , ByVal x , ByVal y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'=======================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'=======================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'=======================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
	ggoSpread.SSSetProtected C_IotypeCd , -1
	ggoSpread.SSSetProtected C_MovTypeNm , -1
End Sub
'=======================================================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub
'=======================================================================================================================
 Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.source= frm1.vspdData     
    ggoSpread.UpdateRow Row
 
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	Call changeExptIvType(Row,false)
	frm1.vspdData.ReDraw = False
	 
	Select Case Col
		Case C_GmCond
		Case C_RetCond
		Case C_SubCond
	End Select 
	 
	frm1.vspdData.ReDraw = True
End Sub
'=======================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
	If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End if
End Sub
'=======================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
 
	if lgQuery = true then Exit Sub
	if lgCopyRow = true then Exit Sub
	if Col = C_MovTypePop then
		Call OpenMovType() 
	elseif Col = C_ExptIvTypePop Then
	    Call OpenExptIvtype()
	elseif Col = C_ImportCond or Col = C_SubCond Then  'KSJ추가 2007.06.14 매입일괄등록 
	    Call changeExptIvType(Row, false)
	End if
End Sub
'=======================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
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
    End if
    
End Sub
'=======================================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    Err.Clear                                                 
 
	ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")     
    Call InitVariables
                  
    If Not ChkField(Document, "1") Then      
       Exit Function
    End If
        
    If DbQuery = False Then Exit Function
       
    FncQuery = True           
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    Err.Clear                                               
    On Error Resume Next                                   
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
   			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call InitVariables                                      
    Call SetDefaultVal
    
    FncNew = True                                           
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                         
    
    Err.Clear                                               
    On Error Resume Next                                   
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If
    
    If Not ChkField(Document, "2")  OR Not ggoSpread.SSDefaultCheck Then                            
       Exit Function
    End If
    
	If DbSave = False Then Exit Function
    
    FncSave = True                                    
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncCopy() 
	If frm1.vspdData.Maxrows < 1 then exit function
	lgCopyRow = true
	frm1.vspdData.ReDraw = false
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	Call frm1.vspdData.SetText(C_IotypeCd,	frm1.vspdData.ActiveRow, "")
	frm1.vspdData.ReDraw = True 
	lgCopyRow = false
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncCancel() 
	if frm1.vspdData.Maxrows < 1 then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo 
    Set gActiveElement = document.ActiveElement                                  
End Function
'=======================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
	Dim IntRetCD
    Dim imRow, iRow
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		
		If imRow = "" Then
			Exit Function
		End if
    End If
    
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow -1
			Call .vspdData.SetText(C_UseCond,	iRow, "1")
		Next
		.vspdData.ReDraw = True
	End With
    
    Set gActiveElement = document.ActiveElement   
    
    FncInsertRow = True                                                          '☜: Processing is OK
    
End Function
'=======================================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    if frm1.vspdData.Maxrows < 1 then exit function
    
    frm1.vspdData.focus
    ggoSpread.Source = frm1.vspdData 
        
	lDelRows = ggoSpread.DeleteRow
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncPrev() 
 ggoSpread.Source = frm1.vspdData
    On Error Resume Next                              
End Function
'=======================================================================================================================
Function FncExcel() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(parent.C_MULTI)     
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncFind() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(parent.C_MULTI , False) 
    Set gActiveElement = document.ActiveElement              
End Function
'=======================================================================================================================
Function FncExit()
 
	Dim IntRetCD
	 
	FncExit = False
	 
	ggoSpread.Source = frm1.vspdData
	    
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X") 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	FncExit = True
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim strVal
	Dim strText
	 

    DbQuery = False
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If 
	    
	Err.Clear                                                 
	With frm1
  		If lgIntFlgMode = parent.OPMD_UMODE Then
     
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtGmTypeCd=" & .hdnGmType.value
			strVal = strVal & "&txtUseflg=" & .hdnUseflg.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtGmTypeCd=" & Trim(.txtGmTypeCd.value)
		    
		    if .rdoUseflg(0).checked = True then
		    	strVal = strVal & "&txtUseflg=" & ""
		    elseif .rdoUseflg(1).checked = True then
		    	strVal = strVal & "&txtUseflg=" & "Y"
		    else
		     	strVal = strVal & "&txtUseflg=" & "N"
		    end if
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		end if 
    
	Call RunMyBizASP(MyBizASP, strVal)
	End With
    
    DbQuery = True
End Function
'=======================================================================================================================
Function DbQueryOk()          
 
	Dim index 
    
    lgIntFlgMode = parent.OPMD_UMODE        
    
    Call ggoOper.LockField(Document, "Q")     
    Call SetToolbar("1110111100111111")
    
    frm1.vspdData.ReDraw = False
    
    ggoSpread.spreadlock  C_IotypeCd, 1,C_IotypeCd, frm1.vspdData.MaxRows
	ggoSpread.sssetrequired C_IotypeNm, 1
	
	For index = 1 To frm1.vspdData.MaxRows
		Call changeExptIvType(index, false)  '일괄매입형태 
	Next
	
	frm1.vspdData.ReDraw = True
      
End Function
'=======================================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal
	Dim saveflg
	Dim index 
	Dim PvArr
	Dim iColSep
	
	DbSave = False    
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If
    
    iColSep = Parent.gColSep
    
	With frm1
		.txtMode.value = parent.UID_M0002
		
		lGrpCnt = 0
		strVal = ""
		ReDim PvArr(0)
		
		For lRow = 1 To .vspdData.MaxRows
			
			.vspdData.Row = lRow
			Call changeExptIvType(lRow,false)
			
			.vspdData.Col = 0
			
			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag ,ggoSpread.DeleteFlag
					
					If  .vspdData.Text = ggoSpread.InsertFlag then
						strVal = "C" & iColSep    '☜: C=Create
					ElseIf  .vspdData.Text = ggoSpread.UpdateFlag then
						strVal = "U" & iColSep    '☜: U=Update
					Else
						strVal = "D" & iColSep    '☜: D=Delete
					End if
					
					.vspdData.Col = C_IotypeCd:		strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
					.vspdData.Col = C_IotypeNm:		strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_GmCond:		strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
					.vspdData.Col = C_RetCond:		strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
					.vspdData.Col = C_SubCond:		strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
					.vspdData.Col = C_ImportCond:	strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
					.vspdData.Col = C_UseCond:		strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
					.vspdData.Col = C_MovTypeCd:	strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
													strVal = strVal & "" & iColSep
													strVal = strVal & "" & iColSep
													strVal = strVal & "" & iColSep
													strVal = strVal & "" & iColSep
													strVal = strVal & "" & iColSep
					.vspdData.Col = C_ChildCond:	strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep '자품목 처리여부 
					.vspdData.Col = C_SubCond2:		strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep '외주가공여부 
					.vspdData.Col = C_ExptIvTypeCd:	strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep 'KSJ추가 2007.06.14 매입일괄등록 
				    strVal = strVal & lRow & parent.gRowSep

				    ReDim Preserve PvArr(lGrpCnt)
				    PvArr(lGrpCnt) = strVal
					lGrpCnt = lGrpCnt + 1
			End Select
        Next
 
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = Join(PvArr, "")

		Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
 
	End With
 
    DbSave = True           
    
End Function
'=======================================================================================================================
Function DbSaveOk()   
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
End Function
'=======================================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<!-- '#########################################################################################################
'            6. Tag부 
'######################################################################################################### -->
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 border="0">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>입출고형태</font></td>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>입출고형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGmTypeCd" ALT="입출고형태" SIZE=10 MAXLENGTH=5  tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIotype()">
										<INPUT TYPE=TEXT ID="txtGmTypeNm" ALT="입출고형태" NAME="arrCond" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>사용여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="사용여부" NAME="rdoUseflg" id = "rdoUseflg1" Value="A" checked tag="1X"><label for="rdoUseflg1">&nbsp;전체&nbsp;</label>
										<INPUT TYPE=radio Class="Radio" ALT="사용여부" NAME="rdoUseflg" id = "rdoUseflg2" Value="Y" tag="1X"><label for="rdoUseflg2">&nbsp;사용&nbsp;</label>
										<INPUT TYPE=radio Class="Radio" ALT="사용여부" NAME="rdoUseflg" id = "rdoUseflg3" Value="N" tag="1X"><label for="rdoUseflg3">&nbsp;미사용&nbsp;</label></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</tr>
   <TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGmType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnUseflg" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
