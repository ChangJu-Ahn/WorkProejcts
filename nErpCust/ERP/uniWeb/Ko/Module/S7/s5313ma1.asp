<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5313MA1
'*  4. Program Name         : 세금계산서번호등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G338.cSListTaxDocNoSvr,PS7G331.cSTaxDocNoSvr
'*  7. Modified date(First) : 2002/11/14
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : Ahn tae hee
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2001/06/26 : 6차 화면 layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*							  -2002/11/14 : UI성능 적용	
'**********************************************************************************************
%>

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                            

Const BIZ_PGM_ID = "s5313mb1.asp"            

' Constant variables 
'========================================

Dim C_TaxBillNo    '세금계산서번호 
Dim C_BookKun      '책번호(권)
Dim C_BookHo       '책번호(호)
Dim C_UseFlag      '사용여부 
Dim C_Udate        '유효일 
Dim C_CreatedMeth  '생성방법 
Dim C_Used         '사용실적 
Dim C_TaxBillDocNo '세금계산서관리번호 
Dim C_Date         '발행일 

Const PostFlag = "PostFlag"

' Common variables 
'========================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey

' User-defind Variables
'========================================
Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

Dim IsOpenPop      ' Popup

'========================================
Sub initSpreadPosVariables()  

	C_TaxBillNo = 1    '세금계산서번호 
	C_BookKun   = 2    '책번호(권)
	C_BookHo    = 3    '책번호(호)
	C_UseFlag   = 4    '사용여부 
	C_Udate     = 5    '유효일 
	C_CreatedMeth = 6  '생성방법 
	C_Used    = 7      '사용실적 
	C_TaxBillDocNo = 8 '세금계산서관리번호 
	C_Date   = 9       '발행일 

End Sub

'========================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE            
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    lgStrPrevKey = ""
    lgLngCurRows = 0
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtSoNo1.className = "TD6"
	frm1.rdoUseFlagA.checked = True
	frm1.txtTaxDocBillNo.focus
	Set gActiveElement = document.activeElement 
	
	lgBlnFlgChgValue = False
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()   

		With frm1.vspdData

			ggoSpread.Source = frm1.vspdData
			
			ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  
			.ReDraw = False
				  
			.MaxRows = 0 : .MaxCols = 0
			.MaxCols = C_Date+1             '☜: 최대 Columns의 항상 1개 증가시킴				  
				      
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit C_TaxBillNo, "세금계산서번호", 30,,,30,2
			Call AppendNumberPlace("7","15","0")
			ggoSpread.SSSetFloat C_BookKun,"책번호(권)" ,15,"7",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat C_BookHo,"책번호(호)" ,15,"7",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetCombo C_UseFlag, "사용여부", 10,2,False

			ggoSpread.SSSetDate C_Udate, "유효일",10,2,Parent.gDateFormat
				  
			ggoSpread.SSSetEdit C_CreatedMeth, "생성방법", 15

			ggoSpread.SetCombo "Y" & vbTab & "N" ,C_UseFlag
			ggoSpread.SSSetEdit C_Used, "사용실적", 15
			ggoSpread.SSSetEdit C_TaxBillDocNo, "세금계산서관리번호", 30
			ggoSpread.SSSetDate C_Date, "발행일",10,2,Parent.gDateFormat
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)				'☜: 공통콘트롤 사용 Hidden Column
			.ReDraw = True
			   
		End With
	    
	End Sub

'===========================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
  
    With frm1
		.vspdData.ReDraw = False
         ggoSpread.Source = frm1.vspdData

         ggoSpread.SSSetRequired C_TaxBillNo, pvStartRow, pvEndRow   
         ggoSpread.SSSetRequired C_UseFlag, pvStartRow, pvEndRow
 
         ggoSpread.SSSetProtected C_CreatedMeth, pvStartRow, pvEndRow

         ggoSpread.SSSetProtected C_Used, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected C_TaxBillDocNo, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected C_Date, pvStartRow, pvEndRow 
        .vspdData.ReDraw = True
    End With

End Sub

'========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_TaxBillNo		= iCurColumnPos(1)
			C_BookKun       = iCurColumnPos(2)
			C_BookHo	    = iCurColumnPos(3)    
			C_UseFlag       = iCurColumnPos(4)
			C_Udate			= iCurColumnPos(5)
			C_CreatedMeth	= iCurColumnPos(6)
			C_Used			= iCurColumnPos(7)
			C_TaxBillDocNo	= iCurColumnPos(8)
			C_Date			= iCurColumnPos(9)
    End Select    
End Sub

'========================================
Function OpenTaxDocBillNo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "세금계산서번호"    
	arrParam(1) = "S_TAX_DOC_NO"
	arrParam(2) = Trim(frm1.txtTaxDocBillNo.value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "세금계산서번호"

	arrField(0) = "ED30" & Parent.gColSep & "TAX_DOC_NO"   
	arrField(1) = "ED15" & Parent.gColSep & "CASE used_flag WHEN " & FilterVar("C", "''", "S") & "  THEN " & FilterVar("Created", "''", "S") & " WHEN " & FilterVar("R", "''", "S") & "  THEN " & FilterVar("Referenced", "''", "S") & " WHEN " & FilterVar("I", "''", "S") & "  THEN " & FilterVar("Issued", "''", "S") & " WHEN " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("Deleted", "''", "S") & " ELSE '' END Used_flag"  
	        
	arrHeader(0) = "세금계산서번호"
	arrHeader(1) = "사용실적"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
									"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtTaxDocBillNo.focus
	
	If arrRet(0) <> "" Then	frm1.txtTaxDocBillNo.value = arrRet(0)

End Function

'========================================
Sub SetQuerySpreadColor(ByVal lRow)
 
 Dim Index
 
    With frm1

    .vspdData.ReDraw = False

     ggoSpread.SSSetProtected C_TaxBillNo, lRow, .vspdData.MaxRows    
     ggoSpread.SSSetRequired C_UseFlag, lRow, .vspdData.MaxRows
  
     ggoSpread.SSSetProtected C_CreatedMeth, lRow, .vspdData.MaxRows

     ggoSpread.SSSetProtected C_Used, lRow, .vspdData.MaxRows
     ggoSpread.SSSetProtected C_TaxBillDocNo, lRow, .vspdData.MaxRows
     ggoSpread.SSSetProtected C_Date, lRow, .vspdData.MaxRows      
  
  For Index = 1 to .vspdData.MaxRows 
     .vspdData.Row = Index
        .vspdData.Col = 0

   If .vspdData.Text = ggoSpread.InsertFlag then
    Call SetSpreadColor(Index,Index)
   End if
  Next

  .vspdData.ReDraw = True
  
    End With

End Sub

'========================================
Sub Form_Load()
	Call SetDefaultVal
	Call InitVariables
	Call LoadInfTB19029
	Call InitComboBox
	Call AppendNumberPlace("6","15","0")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N") 
	    
	Call InitSpreadSheet
	Call SetToolBar("11001111001011")
End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================
Sub txtBookNo_KeyDown(KeyCode, Shift)
 If KeyCode = 13 Then Call MainQuery()
End Sub

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("1101111111")

     gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		Exit Sub
	End If 
End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	If Row < 0 Then Exit Sub

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
End Sub

'==========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess Then Exit Sub
		    
		Call DisableToolBar(Parent.TBC_QUERY)
		Call DBQuery
	End If
End Sub

Sub InitComboBox()
<%
 Dim intLoopCnt
 Dim Arrvalue(3),Arrname(3)
 
 Arrvalue(0) = "C"
 Arrvalue(1) = "R"
 Arrvalue(2) = "I"
 Arrvalue(3) = "D"

 
 Arrname(0)="Created"
 Arrname(1)="Referenced"
 Arrname(2)="Issued"
 Arrname(3)="Deleted"
%>

With frm1
<% 
  For intLoopCnt = 0 To 3
%>   
  Call SetCombo(.cboConfFg, "<%=Arrvalue(intLoopCnt)%>", "<%=Arrname(intLoopCnt)%>")
<%
 Next  
%>  

End With 

End Sub

'========================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If

    Call ggoOper.ClearField(Document, "2")          
    Call InitVariables               

	If frm1.rdoUseFlagA.checked = True Then
		frm1.HUseFlag.value = frm1.rdoUseFlagA.value
	ElseIf frm1.rdoUseFlagY.checked = True Then
		frm1.HUseFlag.value = frm1.rdoUseFlagY.value
	Else
		frm1.HUseFlag.value = frm1.rdoUseFlagN.value
	End If

    Call DbQuery                

    FncQuery = True                
        
End Function

'========================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then Exit Function
    End If
    
    Call ggoOper.ClearField(Document, "A")                                      
    Call ggoOper.LockField(Document, "N")                                       
    Call SetToolBar("11001111001011")          
    Call SetDefaultVal
    Call InitVariables               

    FncNew = True                

End Function

'========================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    
    ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If

    If ggoSpread.SSDefaultCheck = False Then Exit Function
    CAll DbSave                                                    
    
    FncSave = True                                                          
    
End Function

'========================================
Function FncCopy() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	With frm1
		.vspdData.ReDraw = False
		 
		ggoSpread.Source = frm1.vspdData 
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
		  
		.vspdData.Col  = C_TaxBillNo
		.vspdData.text = ""
		  
		.vspdData.ReDraw = True
	End With
	    
End Function

'========================================
Function FncCancel() 
 If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData 
    ggoSpread.EditUndo  
End Function

'========================================
Function FncInsertRow(ByVal pvRowCnt) 

    Dim imRow
    On Error Resume Next                                                          
    Err.Clear                                                                     
    
    FncInsertRow = False                                                         

   If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
				
		lgBlnFlgChgValue = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True
    End If   

End Function

'========================================
Function FncDeleteRow() 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
     ggoSpread.Source = .vspdData 
    
	 lDelRows = ggoSpread.DeleteRow
 
    lgBlnFlgChgValue = True
    
    End With
    
End Function

'========================================
Function FncPrint() 
 Call parent.FncPrint()
End Function

'========================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function

'========================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                                         
End Function

'========================================
Sub FncSplitColumn()    
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If

	FncExit = True
End Function

'========================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetQuerySpreadColor(1)
End Sub

'========================================
Function DbQuery() 

    Err.Clear                                                               
    
    DbQuery = False                                                         

   
  If   LayerShowHide(1) = False Then
             Exit Function 
        End If

    
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then    
  strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         
  strVal = strVal & "&txtTaxDocBillNo=" & Trim(frm1.HTaxBillDocNo.value)   
  strVal = strVal & "&txtBookNo=" & Trim(frm1.HBookNo.value)
  
  strVal = strVal & "&HUsed=" & Trim(frm1.HUsed.value)
  strVal = strVal & "&HUseFlag=" & Trim(frm1.HUseFlag.value)
  strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 Else
  strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         
  strVal = strVal & "&txtTaxDocBillNo=" & Trim(frm1.txtTaxDocBillNo.value)  
  strVal = strVal & "&txtBookNo=" & Trim(frm1.txtBookNo.Text)
  
  strVal = strVal & "&HUsed=" & Trim(frm1.cboConfFg.value)
  strVal = strVal & "&HUseFlag=" & Trim(frm1.HUseFlag.value)
  strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 End If 
 
 strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
 
 Call RunMyBizASP(MyBizASP, strVal)            
 
    DbQuery = True                 

End Function

'========================================
Function DbQueryOk()
 
    lgIntFlgMode = Parent.OPMD_UMODE
	lgBlnFlgChgValue = False
    lgIntGrpCount = 0              

    Call SetToolBar("11101111001111")
	Call SetQuerySpreadColor(1)
 
 If frm1.vspdData.MaxRows > 0 Then 
  frm1.vspdData.Focus  
 Else
  frm1.txtTaxDocBillNo.focus
    End If     

End Function

'========================================
Function DbSave()

    Err.Clear                
 
    Dim lRow        
    Dim lGrpCnt     
	 Dim strVal, strDel
 
    DbSave = False
    
	If LayerShowHide(1) = False Then Exit Function 


 With frm1
  .txtMode.value = Parent.UID_M0002
  .txtUpdtUserId.value = Parent.gUsrID
  .txtInsrtUserId.value = Parent.gUsrID
    
  lGrpCnt = 0    
  strVal = ""
  strDel = ""
    
  For lRow = 1 To .vspdData.MaxRows
    
      .vspdData.Row = lRow
      .vspdData.Col = 0
          
   Dim Udate 
   Dim iRet
   
      Select Case .vspdData.Text
          Case ggoSpread.InsertFlag       '☜: 신규 
     strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep'☜: C=Create
          Case ggoSpread.UpdateFlag       '☜: 수정 
     strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep'☜: U=Update
          Case ggoSpread.DeleteFlag       '☜: 삭제 
     strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep'☜: D=Delete
     '--- 세금계산서번호 
              .vspdData.Col = C_TaxBillNo 
              strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep

              lGrpCnt = lGrpCnt + 1 
   End Select

   Select Case .vspdData.Text
    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

     '--- 세금계산서번호 
              .vspdData.Col = C_TaxBillNo
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '--- 책번호(권)
              .vspdData.Col = C_BookKun
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
     '--- 책번호(호)
              .vspdData.Col = C_BookHo
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
     '--- 사용여부 
              .vspdData.Col = C_UseFlag
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     
     .vspdData.Col = C_Used
     if .vspdData.Text = "Created" then
       '--- 유효일 
       .vspdData.Col = C_Udate
       If UniConvDateToYYYYMMDD(.vspdData.Text,Parent.gDateFormat,"") <> "" then
        If UniConvDateToYYYYMMDD(.vspdData.Text,Parent.gDateFormat,"") <  UniConvDateToYYYYMMDD("<%=EndDate%>",Parent.gDateFormat,"") Then  
         iRet = DisplayMsgBox("205824", "X", lRow&"행:" , "X")
         LayerShowHide(0)
         Call SetToolBar("11001111001011") 
         
         Exit Function
      
        End If
       
       End if
     else
       .vspdData.Col = C_Udate
       If UniConvDateToYYYYMMDD(.vspdData.Text,Parent.gDateFormat,"") <> "" then
       
        Udate = UniConvDateToYYYYMMDD(.vspdData.Text,Parent.gDateFormat,"") 

        .vspdData.Col = C_Date '발행일 
        If  Udate < UniConvDateToYYYYMMDD(.vspdData.Text,Parent.gDateFormat,"") Then  
         
         iRet = DisplayMsgBox("205823", "X", lRow&"행:" , "X")
       
         LayerShowHide(0)
         Call SetToolBar("11001111001011") 
         
         Exit Function
      
        End If
       
       End if
          
     end if
                    
     .vspdData.Col = C_Udate
              strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & Parent.gRowSep
               
     lGrpCnt = lGrpCnt + 1 
      End Select       
  Next
 
  .txtMaxRows.value = lGrpCnt
  .txtSpread.value = strDel & strVal
  
  Call ExecMyBizASP(frm1, BIZ_PGM_ID)
 
 End With
 
    DbSave = True
    
End Function

'========================================
Function DbSaveOk()
	Call InitVariables
	Call ggoOper.ClearField(Document, "2")
	Call MainQuery()
End Function

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
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>세금계산서번호</font></td>
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
         <TD CLASS="TD5" NOWRAP>세금계산서번호</TD>
         <TD CLASS="TD6" NOWRAP><INPUT NAME="txtTaxDocBillNo" ALT="세금계산서번호" TYPE="Text" MAXLENGTH=30 SiZE=30 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBillNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTaxDocBillNo()"><div STYLE="DISPLAY: none"><INPUT NAME="txtSoNo1" STYLE="BORDER-RIGHT: 0px solid;BORDER-TOP: 0px solid;BORDER-LEFT: 0px solid;BORDER-BOTTOM: 0px solid" TYPE="Text" SIZE=1 DISABLED=TRUE Tag="11"></div></TD>
         <TD CLASS=TD5 NOWRAP>사용실적</TD>
         <TD CLASS=TD6 NOWRAP>&nbsp;
         <SELECT NAME="cboConfFg" tag="11X" STYLE="WIDTH:82px:" style="width:100px"><OPTION VALUE="" selected></OPTION></SELECT>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>책번호(권)</TD>
         <TD CLASS="TD6"><script language =javascript src='./js/s5313ma1_fpDoubleSingle1_txtBookNo.js'></script></TD>
         <TD CLASS=TD5 NOWRAP>사용여부</TD>
         <TD CLASS=TD6 NOWRAP>
          <INPUT TYPE=radio CLASS="RADIO" NAME="rdoUseFlag" id="rdoUseFlagA" VALUE="" tag = "11" CHECKED>
           <LABEL FOR="rdoUseFlagA">전체</LABEL>&nbsp;&nbsp;
          <INPUT TYPE=radio CLASS="RADIO" NAME="rdoUseFlag" id="rdoUseFlagY" VALUE="Y" tag = "11">
           <LABEL FOR="rdoUseFlagY">예</LABEL>&nbsp;&nbsp;
          <INPUT TYPE=radio CLASS = "RADIO" NAME="rdoUseFlag" id="rdoUseFlagN" VALUE="N" tag = "11">
           <LABEL FOR="rdoUseFlagN">아니오</LABEL></TD>
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
          <script language =javascript src='./js/s5313ma1_vaSpread1_vspdData.js'></script>
         </TD>
        </TR>
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="HTaxBillDocNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HTaxBillNo" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HUsed" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HUseFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HBookNo" tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
