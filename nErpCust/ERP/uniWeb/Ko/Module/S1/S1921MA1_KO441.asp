<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : S1921MA1_KO441
'*  4. Program Name         : 업체별적용환율등록(KO441)
'*  5. Program Desc         : 업체별적용환율등록(KO441)
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2008/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : ajc
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/11/20 : Grid성능 적용, Kang Jun Gu
'*				            : 2002/12/10 : INCLUDE 다시 성능 적용, Kang Jun Gu
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                        '☜: Turn on the Option Explicit option.

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim prDBSYSDate

Dim EndDate ,StartDate

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim IsOpenPop

prDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
StartDate = UniDateAdd("m", 0, EndDate,parent.gDateFormat)

Const BIZ_PGM_ID = "S1921MB1_KO441.asp"            '☆: 비지니스 로직 ASP명 

Dim C_ConDt
Dim C_DocCur
Dim C_DocCurPopup
Dim C_ExchRate_Fir
Dim C_ExchRate_Las
Dim C_Remark
Dim C_ConDt_H

Dim C_BpCd
Dim C_BpCdPopup
Dim C_BpNm
Dim C_ExchApplyCd
Dim C_ExchApply
Dim C_Remark2

Dim C_ConDt2
Dim C_DocCur2

'========================================================================================================= 
Dim gblnWinEvent
Dim CheckDTL
Dim gSpreadFlg
'========================================================================================================
Sub initSpreadPosVariables() 

	Dim i
'Grid1 -----------------------------------------------------------------	
	i = 1
	C_ConDt			= i	: i=i+1					' 년월
	C_DocCur		= i	: i=i+1					' 통화	
	C_DocCurPopup	= i	: i=i+1					' 통화 -------	팝		
	C_ExchRate_Fir	= i	: i=i+1					' 최초고시환율
	C_ExchRate_Las	= i	: i=i+1					' 최종고시환율
	C_Remark		= i	: i=i+1					' 비고
	C_ConDt_H		= i	: i=i+1					' 년월	Key
	
'Grid2 -----------------------------------------------------------------	
	i = 1
	C_BpCd			= i	: i=i+1					' 업체코드
	C_BpCdPopup		= i	: i=i+1					' 업체 -------	팝
	C_BpNm			= i	: i=i+1					' 업체명
	C_ExchApplyCd	= i	: i=i+1					' 적용고시코드
	C_ExchApply		= i	: i=i+1					' 적용고시
	C_Remark2		= i	: i=i+1					' 비고2

	C_ConDt2		= i	: i=i+1					' 년월
	C_DocCur2		= i	: i=i+1					' 통화

End Sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	
	gSpreadFlg = 1
	frm1.txtConDt.Text = StartDate
	frm1.txtConDt.focus
	lgBlnFlgChgValue = False

End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	
 
	With frm1.vspdData
	   ggoSpread.Source = frm1.vspdData
       ggoSpread.Spreadinit "V20080623",,parent.gAllowDragDropSpread    

       .MaxCols   = C_ConDt_H + 1							    				  ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True
       .MaxRows = 0															          ' ☜: Clear spreadsheet data 
        ggoSpread.Source = frm1.vspdData
       Call GetSpreadColumnPos("A")
	   .ReDraw = false

		ggoSpread.SSSetEdit		C_ConDt,			"년월"			,10,2,,7,2						' 년월
		ggoSpread.SSSetEdit		C_DocCur,			"통화"			,10,0,,8,2						' 통화
		ggoSpread.SSSetButton   C_DocCurPopup														' 통화 -------	팝	
		ggoSpread.SSSetFloat    C_ExchRate_Fir,		"최초고시환율"	,15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec ' 최초고시환율
		ggoSpread.SSSetFloat	C_ExchRate_Las,		"최종고시환율"	,15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec ' 최종고시환율
		ggoSpread.SSSetEdit		C_Remark,			"비고"			,40,0,,50,2						' 비고
		ggoSpread.SSSetEdit		C_ConDt_H,			"년월"			,10,2,,7,2						' 년월

		.ReDraw = true 
	   
	   Call ggoSpread.SSSetColHidden(C_ConDt_H,	C_ConDt_H,	True)
       Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)
    
    End With

	With frm1.vspdData2
	   ggoSpread.Source = frm1.vspdData2
       ggoSpread.Spreadinit "V20080623",,parent.gAllowDragDropSpread    

       .MaxCols   = C_DocCur2 + 1							    				  ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True
       .MaxRows = 0															          ' ☜: Clear spreadsheet data 
        ggoSpread.Source = frm1.vspdData2

       Call GetSpreadColumnPos("B")
	   .ReDraw = false		
		
		ggoSpread.SSSetEdit		C_BpCd,				"업체"			,15,0,,10,2						' 업체
		ggoSpread.SSSetButton   C_BpCdPopup															' 업체 -------	팝	
		ggoSpread.SSSetEdit     C_BpNm,				"업체명"		,30,0							' 업체명	
		ggoSpread.SSSetCombo	C_ExchApplyCd,		"적용고시코드"	,10, 0							' 적용고시코드
		ggoSpread.SSSetCombo	C_ExchApply,		"적용고시"		,15, 0							' 적용고시
		ggoSpread.SSSetEdit		C_Remark2,			"비고사항"		,30,0,,50,2						' 비고사항
				
		ggoSpread.SSSetEdit		C_ConDt2,			"년월"			,10,0,,8,2						' 년월
		ggoSpread.SSSetEdit		C_DocCur2,			"통화"			,10,0,,8,2						' 통화
		
		.ReDraw = true 
    
	   Call ggoSpread.SSSetColHidden(C_ExchApplyCd,	C_ExchApplyCd,	True)
	   Call ggoSpread.SSSetColHidden(C_ConDt2,	C_ConDt2,	True)
	   Call ggoSpread.SSSetColHidden(C_DocCur2,	C_DocCur2,	True)
       Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)
    
    End With
	
	Call SetSpreadLock 
    
End Sub

'===========================================================================================================
Sub SetSpreadLock()
    
    With frm1
    
    .vspdData.ReDraw = False
	 
		Call SetSpreadColor(-1,-1)
	
    .vspdData.ReDraw = True

    End With
    
    With frm1
    
    .vspdData2.ReDraw = False
	 
		Call SetSpreadColor2(-1,-1)
	
    .vspdData2.ReDraw = True

    End With

End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
            
    With frm1
    
    .vspdData.ReDraw = False
		ggoSpread.Source = .vspdData
	
		ggoSpread.SSSetRequired    C_ConDt,				  pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_DocCur,			  pvStartRow, pvEndRow
		
    .vspdData.ReDraw = True
    
    End With

End Sub
'========================================================================================
Sub SetSpreadColor2(ByVal pvStartRow,ByVal pvEndRow)
            
    With frm1
    
    .vspdData2.ReDraw = False
		ggoSpread.Source = .vspdData2
	
		ggoSpread.SSSetRequired    C_BpCd,				  pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_BpNm,				  pvStartRow, pvEndRow
'		ggoSpread.SSSetRequired    C_ExchApplyCd,		  pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_ExchApply,			  pvStartRow, pvEndRow
		
    .vspdData2.ReDraw = True
    
    End With

End Sub
'========================================================================================
Sub SetSpreadLockAfterQuery(ByVal pvStartRow,ByVal pvEndRow)
  	          
    With frm1
    
    .vspdData.ReDraw = False
		ggoSpread.Source = .vspdData		
		
		'ggoSpread.SSSetRequired     C_ConDt,			  pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_ConDt,			  pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_DocCur,			  pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_DocCurPopup,		  pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected    C_ConDt_H,			  pvStartRow, pvEndRow
		
    .vspdData.ReDraw = True
    
    End With

End Sub
'========================================================================================
Sub SetSpreadLockAfterQuery2(ByVal pvStartRow,ByVal pvEndRow)
  	          
    With frm1
    
    .vspdData2.ReDraw = False
		ggoSpread.Source = .vspdData2		
		
		ggoSpread.SSSetProtected    C_BpCd,				  pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_BpCdPopup,		  pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected    C_ConDt2,			  pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_DocCur2,			  pvStartRow, pvEndRow
		
    .vspdData2.ReDraw = True
    
    End With

End Sub
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    Dim i
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            i = 1
            C_ConDt				= iCurColumnPos(i)	: i=i+1					' 년월
			C_DocCur			= iCurColumnPos(i)	: i=i+1					' 통화
			C_DocCurPopup		= iCurColumnPos(i)	: i=i+1					' 통화 -------	팝
			C_ExchRate_Fir		= iCurColumnPos(i)	: i=i+1					' 최초고시환율
			C_ExchRate_Las		= iCurColumnPos(i)	: i=i+1					' 최종고시환율
			C_Remark			= iCurColumnPos(i)	: i=i+1					' 비고
			C_ConDt_H			= iCurColumnPos(i)	: i=i+1					' 년월 Key
		
		Case "B"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			i = 1
            C_BpCd				= iCurColumnPos(i)	: i=i+1					' 업체코드
			C_BpCdPopup			= iCurColumnPos(i)	: i=i+1					' 업체 -------	팝
			C_BpNm				= iCurColumnPos(i)	: i=i+1					' 업체명 
			C_ExchApplyCd		= iCurColumnPos(i)	: i=i+1					' 적용고시코드
			C_ExchApply			= iCurColumnPos(i)	: i=i+1					' 적용고시
			C_Remark2			= iCurColumnPos(i)	: i=i+1					' 비고2
			C_ConDt2			= iCurColumnPos(i)	: i=i+1					' 년월
			C_DocCur2			= iCurColumnPos(i)	: i=i+1					' 통화
	
    End Select    
End Sub

'========================================================================================================= 
Sub Form_Load()
 
	Call LoadInfTB19029()
 
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) '⊙: Format Contents  Field
	Call ggoOper.LockField(Document, "N")                                     '⊙: Lock  Suitable  Field
	Call ggoOper.FormatDate(frm1.txtConDt, parent.gDateFormat, 2)
	Call InitSpreadSheet
	Call SetDefaultVal
	Call InitVariables   
	Call SetGridCombo  
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("1110110100001111")          '⊙: 버튼 툴바 제어 
	
End Sub
'=======================================================================================================
Function SetGridCombo()
	Dim strCombo, strVal
	
	ggoSpread.Source = frm1.vspdData2
	
	strCombo = "F" & vbTab & "L" 
	strVal = "최초고시환율" & vbTab & "최종고시환율"	
   
    ggoSpread.SetCombo Replace(strCombo ,Chr(11),vbTab), C_ExchApplyCd
    ggoSpread.SetCombo Replace(strVal,Chr(11),vbTab), C_ExchApply
	
End Function
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	ggoSpread.Source = frm1.vspdData
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
    
    With frm1.vspdData		
		If Row > 0 Then
			Select Case Col
				Case C_DocCurPopup
					.Col = C_DocCur
					.Row = Row
					Call OpenPopUp(.Text, 2)
			End Select    
		End If
	End With

End Sub
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	ggoSpread.Source = frm1.vspdData2
   	frm1.vspdData2.Row = Row
    frm1.vspdData2.Col = Col
    
    With frm1.vspdData2		
		If Row > 0 Then
			Select Case Col
				Case C_BpCdPopup
					.Col = C_BpCd
					.Row = Row
					Call OpenPopUp(.Text, 3)
			End Select    
		End If
	End With

End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData
	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
    
    gSpreadFlg = 1
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If CheckDTL <> "" Then
   		Call SetActiveCell(frm1.vspdData, 1, CheckDTL,"M","X","X")
   		Exit Sub
   	End If
   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData2
	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
    
    gSpreadFlg = 2
    
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If CheckDTL <> "" Then
   		Call SetActiveCell(frm1.vspdData2, 1, CheckDTL,"M","X","X")
   		Exit Sub
   	End If
   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData2
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData2.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	
	Dim IntRetCD
	
	CheckDTL = ""
	if frm1.vspddata.row = 0 then exit sub
		If Row <> NewRow And NewRow > 0 Then
			If CheckRunningBizProcess = True Then
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = 1 
				frm1.vspdData.Action = 0
				Exit Sub
			End If
			
			ggoSpread.Source = frm1.vspdData2
			If ggoSpread.SSCheckChange = True Then
				IntRetCD = DisplayMsgBox("700001", parent.VB_YES_NO,Row & "행 :","x")   <% '⊙: "Will you destory previous data" %>
				If IntRetCD = vbNo Then
					CheckDTL = Row
				    Exit Sub
				End If
			End If 
			
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = NewRow					
			If frm1.vspdData.Text = ggoSpread.InsertFlag Then 			    	
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.ClearSpreadData
				Exit Sub	
			End If
	 
		    If DbQuery2(Trim(GetSpreadText(frm1.vspdData,C_ConDt_H, NewRow,"X","X")), Trim(GetSpreadText(frm1.vspdData,C_DocCur, NewRow,"X","X"))) = False Then	Exit Sub
		
		End If
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	
	Dim dblAllow, dblSubAmt, dblAmt
	Dim sYear,sMon,sDay
	Dim tempVal
	
	With frm1.vspdData
	.Row = Row
	.Col = Col
	 
		ggoSpread.Source = frm1.vspdData
		.Col = Col	: .Row = Row
	
	    Select Case Col			
				
			Case C_ConDt
			
				tempVal = .Text
					If Len(tempVal) = 6 Then 
						tempVal = left(tempVal,4) & "-" & Right(tempVal,2)
						Call frm1.vspdData.SetText(C_ConDt, Row , tempVal)
					End If
					
					If Len(tempVal) < 6 Then 
						Call DisplayMsgBox("700003","x","x","x")
						Call frm1.vspdData.SetText(C_ConDt, Row , "")
						Exit Sub
					End If
					If IsNumeric(Left(tempVal,4)) = False Or UNICDbl(Left(tempVal,4)) < 1900 Or UNICDbl(Left(tempVal,4)) > 3000 Then 						
						Call DisplayMsgBox("700003","x","x","x")
						Call frm1.vspdData.SetText(C_ConDt, Row , "")
						Exit Sub
					End If					
					If IsNumeric(Right(tempVal,2)) = False Or UNICDbl(Right(tempVal,2)) <= 0 Or UNICDbl(Right(tempVal,2)) > 12 Then 						
						Call DisplayMsgBox("700003","x","x","x")
						Call frm1.vspdData.SetText(C_ConDt, Row , "")
						Exit Sub
					End If
					If Mid(tempVal,5,1) <> "-" Then
						Call DisplayMsgBox("700003","x","x","x")
						Call frm1.vspdData.SetText(C_ConDt, Row , "")
						Exit Sub
					End If
				
				.Col = C_DocCur
				.Row = Row
				If Len(.Text) Then
					If UCase(Trim(.Text)) = parent.gCurrency Then
						.Col = C_ExchRate_Fir : .Text = 1
						.Col = C_ExchRate_Las : .Text = 1
					Else
						Call FindExchRate(UniConvDateToYYYYMMDD(Trim(GetSpreadText(frm1.vspdData,C_ConDt, Row,"X","X")) & "-01",parent.gDateFormat,""), UCase(Trim(GetSpreadText(frm1.vspdData,C_DocCur, Row,"X","X"))),Row)
					End If
						
					Call DocCur_OnChange(Row,Row)
				End If
					
            Case  C_DocCur
				
				.Col = C_ConDt
				If .Text = "" Then 
					Call DisplayMsgBox("700003","x","x","x")
					Call frm1.vspdData.SetText(C_DocCur, Row , "")
					Call SetActiveCell(frm1.vspdData, C_ConDt, Row,"M","X","X")
					Exit Sub
				End If
				
				.Col = C_DocCur
				If UCase(Trim(.Text)) = parent.gCurrency Then
					.Col = C_ExchRate_Fir : .Text = 1
					.Col = C_ExchRate_Las : .Text = 1
				Else
				
					Call FindExchRate(UniConvDateToYYYYMMDD(Trim(GetSpreadText(frm1.vspdData,C_ConDt, Row,"X","X")) & "-01",parent.gDateFormat,""), UCase(Trim(GetSpreadText(frm1.vspdData,C_DocCur, Row,"X","X"))),Row)
				End If
					
				Call DocCur_OnChange(Row,Row)
			
        End Select  			
			
	End With

	ggoSpread.UpdateRow Row
End Sub
'==========================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )

	Dim sYear,sMon,sDay
	With frm1.vspdData2
	.Row = Row
	.Col = Col

		ggoSpread.Source = frm1.vspdData2
		.Col = Col	: .Row = Row
	
	    Select Case Col
			
			Case C_BpCd
				Call CommonQueryRs(" BP_NM ", " B_BIZ_PARTNER ", " BP_TYPE IN('C','CS') AND BP_CD = " & FilterVar(.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If lgF0 <> "" Then
					Call frm1.vspdData2.SetText(C_BpNm, Row, RePlace(lgF0, Chr(11), ""))
				Else
					Call frm1.vspdData2.SetText(C_BpCd, Row, "")
					Call frm1.vspdData2.SetText(C_BpNm, Row, "")
					Call OpenPopup(Trim(GetSpreadText(frm1.vspdData2,C_BpCd, Row,"X","X")),3)
				End If
			Case C_ExchApply
	
				If .Text = "최초고시환율" Then
					Call frm1.vspdData2.SetText(C_ExchApplyCd, Row, "F")
				Else
					Call frm1.vspdData2.SetText(C_ExchApplyCd, Row, "L")
				End If
				
			Case C_ExchApplyCd
	
				If .Text = "F" Then
					Call frm1.vspdData2.SetText(C_ExchApply, Row, "최초고시환율")
				Else
					Call frm1.vspdData2.SetText(C_ExchApply, Row, "최종고시환율")
				End If
        
        End Select  			
			
	End With

	ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)

 '   Select Case Col
 '       Case C_Item_Price
 '           Call EditModeCheck(frm1.vspdData, Row, C_Cur, C_Item_Price, "C" ,"I", Mode, "X", "X")        
 '   End Select
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess Then	Exit Sub
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery()
	End if     

End Sub

'============================================================================================================
Function OpenPopUp(ByVal strCode, ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 2
		
				frm1.vspdData.Col = C_ConDt
				If frm1.vspdData.Text = "" Then 
					Call DisplayMsgBox("700003","x","x","x")
					Call frm1.vspdData.SetText(C_DocCur, frm1.vspdData.ActiveRow , "")
					Call SetActiveCell(frm1.vspdData, C_ConDt, frm1.vspdData.ActiveRow,"M","X","X")
					IsOpenPop = False
					Exit Function
				End If
				
			arrParam(0) = "통화코드 팝업"	
			arrParam(1) = "B_Currency"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "통화코드"

			arrField(0) = "Currency"
			arrField(1) = "Currency_desc"

			arrHeader(0) = "통화코드"	
			arrHeader(1) = "통화코드명"
			
'			arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=B_CURRENCY_00", Array(Array(Trim(strCode))), _
'			           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	    	    			
			
		Case 3
			arrParam(0) = "업체코드팝업"
			arrParam(1) = "B_BIZ_PARTNER"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "BP_TYPE IN('C','CS')"
			arrParam(5) = "업체코드"

			arrField(0) = "BP_CD"
			arrField(1) = "BP_NM"

			arrHeader(0) = "업체코드"
			arrHeader(1) = "업체코드명"

'			arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=A_ACCT_00", Array(Array(Trim(strCode))), _
'			           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	    
	End Select

    If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	IsOpenPop = False

	If arrRet(0) <> "" Then     
		Call SetPopup(arrRet, iWhere)	
	End If

	Call FocusAfterPopup(iWhere)
End Function
'========================================================================================================= 
Function SetPopUp(ByRef arrRet, ByVal iWhere)
	With frm1
		Select Case iWhere	
			Case 2
				.vspdData.Row = .vspdData.ActiveRow 
				ggoSpread.Source = .vspdData
				ggoSpread.UpdateRow .vspdData.ActiveRow 
			
				.vspdData.Col  = C_DocCur
				.vspdData.Text = UCase(Trim(arrRet(0)))
				If Trim(.vspdData.Text) = parent.gCurrency Then
					.vspdData.Col  = C_ExchRate_Fir
					.vspdData.Text = 1
					.vspdData.Col  = C_ExchRate_Las
					.vspdData.Text = 1
				Else
					Call FindExchRate(UniConvDateToYYYYMMDD(Trim(GetSpreadText(.vspdData,C_ConDt, .vspdData.ActiveRow,"X","X")) & "-01",parent.gDateFormat,""), UCase(Trim(arrRet(0))),.vspdData.ActiveRow)
				End If
				
				Call DocCur_OnChange(.vspdData.ActiveRow,.vspdData.ActiveRow)
			Case 3
				Call frm1.vspdData2.SetText(C_BpCd, .vspdData2.ActiveRow , UCase(Trim(arrRet(0))))
				Call frm1.vspdData2.SetText(C_BpNm, .vspdData2.ActiveRow , UCase(Trim(arrRet(1))))
		End Select
	End With
End Function
'=================================================================================================================
Function FindExchRate(Byval strDate, Byval FromCurrency,Byval Row )

	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strExchFg
	Dim strExchRate
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6	

	strSelect	= "b.minor_cd"
	strFrom		= "b_company a, b_minor b"
	strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
	If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
		arrTemp = Split(lgF0, chr(11))
		strExchFg =  arrTemp(0)
	End If

	If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
		strDate = Mid(strDate, 1, 6)
		strSelect	= "std_rate"
		strFrom		= "b_monthly_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData.row  = Row
			frm1.vspdData.Col  = C_ExchRate_Fir
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
			frm1.vspdData.Col  = C_ExchRate_Las
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			Call DisplayMsgBox("121600", "X", "X", "X")
		End If
	Else					' Floating Exchange Rate

		strSelect	= "top 1 std_rate"
		strFrom		= "b_daily_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt desc"

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData.row  = Row
			frm1.vspdData.Col  = C_ExchRate_Fir
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
			frm1.vspdData.Col  = C_ExchRate_Las
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			Call DisplayMsgBox("121500", "X", "X", "X")
			Call frm1.vspdData.SetText(C_ExchRate_Fir, Row, "")
			Call frm1.vspdData.SetText(C_ExchRate_Las, Row, "")
		End If
	End If
	
End Function    
'==========================================================================================
Sub DocCur_OnChange(FromRow, ToRow)
	Dim ii
	
    lgBlnFlgChgValue = True
    
	For ii = FromRow	To	ToRow
		frm1.vspdData.Row	= ii
		frm1.vspdData.Col	= C_DocCur
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.vspdData.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			Call CurFormatNumSprSheet(ii)
		End If
	Next  
End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet(Row)
	With frm1
		ggoSpread.Source = frm1.vspdData
		.vspdData.Row	= Row
		.vspdData.Col	= C_DocCur
		'Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_DocCur,C_ItemAmt, "A" ,"I","X","X")         
        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_DocCur,C_ExchRate_Fir,"D" ,"I","X","X")		
        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_DocCur,C_ExchRate_Las,"D" ,"I","X","X")		
	End With
End Sub
'=======================================================================================================
Function FocusAfterPopup(ByVal iWhere)
	With frm1
		Select Case iWhere
			Case 2 
				Call SetActiveCell(.vspdData,C_DocCur,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				Call SetActiveCell(.vspdData,C_BpCd,.vspdData.ActiveRow ,"M","X","X")
		End Select    
	End With
End Function
'========================================================================================================
Function FncQuery()

	 Dim IntRetCD

	 FncQuery = False             <% '⊙: Processing is NG %>

	 Err.Clear               <% '☜: Protect system from crashing %>
         
	 <% '------ Check previous data area ------ %>
	 ggoSpread.Source = frm1.vspdData
	 If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")   <% '⊙: "Will you destory previous data" %>
		If IntRetCD = vbNo Then
		    Exit Function
		End If
	 End If 
	 
	 ggoSpread.Source = frm1.vspdData2
	 If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")   <% '⊙: "Will you destory previous data" %>
		If IntRetCD = vbNo Then
		    Exit Function
		End If
	 End If 
	
	 ggoSpread.Source = frm1.vspdData
	 ggoSpread.ClearSpreadData
	 ggoSpread.Source = frm1.vspdData2
	 ggoSpread.ClearSpreadData
		
	 <% '------ Erase contents area ------ %>
	 Call ggoOper.ClearField(Document, "2")        <% '⊙: Clear Contents  Field %>
	 Call InitVariables             <% '⊙: Initializes local global variables %>

	 <% '------ Check condition area ------ %>
	 If Not chkField(Document, "1") Then       <% '⊙: This function check indispensable field %>
	  Exit Function
	 End If
	  
	    
	 <% '------ Query function call area ------ %>
	 Call DbQuery()              <% '☜: Query db data %>
	        
	 FncQuery = True              <% '⊙: Processing is OK %>

End Function
 
'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False              <% '☜: Protect system from crashing %>

	<% '------ Check previous data area ------ %>
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	<% '------ Erase condition area ----- %>
	<% '------ Erase contents area ------ %>
	Call ggoOper.ClearField(Document, "A")        <%'⊙: Clear Condition Field%>
	Call ggoOper.LockField(Document, "N")        <%'⊙: Lock  Suitable  Field%>
	Call SetDefaultVal
	Call SetToolBar("1110110100101111")          '⊙: 버튼 툴바 제어 
	Call InitVariables             <%'⊙: Initializes local global variables%>

	FncNew = True              <%'⊙: Processing is OK%>

End Function
 
'========================================================================================================
Function FncSave()
	Dim IntRetCD
	Dim ChFlag  
	FncSave = False                  <% '⊙: Processing is NG %>
	  
	Err.Clear                   <% '☜: Protect system from crashing %>
	
	With frm1  
		If Not chkField(Document, "2") Then  <% '⊙: Check contents area %>
			Exit Function
		End If
	End With
	<% '------ Precheck area ------ %> 
	
	If lgIntFlgMode = parent.OPMD_CMODE Then	
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = False Then        <% 'Check if there is retrived data %>
			IntRetCD = DisplayMsgBox("900001","x","x","x")     <% '⊙: No data changed!! %>
			Exit Function
		End If
	Else
		ChFlag = False
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = True Then ChFlag = True 
		
		ggoSpread.Source = frm1.vspdData2
		If ggoSpread.SSCheckChange = True Then ChFlag = True 
		
		If ChFlag = False Then
			IntRetCD = DisplayMsgBox("900001","x","x","x")     <% '⊙: No data changed!! %>
			Exit Function
		End If
	End If
	  
	<% '------ Check contents area ------ %>
	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck Then  <% '⊙: Check contents area %>
		Exit Function
	End If
	
	ggoSpread.Source = frm1.vspdData2
	If Not ggoSpread.SSDefaultCheck Then  <% '⊙: Check contents area %>
		Exit Function
	End If
	  
	'-----------------------------------------------------------------------------
	Dim iRow, temp1, temp2
	 
	For iRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Col = 0
		frm1.vspdData.Row = iRow
			
		If frm1.vspdData.Text = ggoSpread.InsertFlag Or frm1.vspdData.Text = ggoSpread.UpdateFlag Then			
			temp1 = UNICDbl(Trim(GetSpreadText(frm1.vspdData,C_ExchRate_Fir, iRow,"X","X")))
			temp2 = UNICDbl(Trim(GetSpreadText(frm1.vspdData,C_ExchRate_Las, iRow,"X","X")))
			If temp1 + temp2 = 0 Then
				Call DisplayMsgBox("700002","x",iRow & "행 :","x")
'				Msgbox iRow & "행 : 환율을 입력 하십시오.(최소한가지)"
				Exit Function
			End If
		End If
	Next
	  
	<% '------ Save function call area ------ %>
	Call DbSave                   <% '☜: Save db data %>
	  
	FncSave = True                  <% '⊙: Processing is OK %>
End Function

'========================================================================================================
Function FncCopy()
	Dim IntRetCD
	
	If gSpreadFlg = 1 Then
        If Frm1.vspdData.MaxRows < 1 Then
           Exit Function
        End If
        
            ggoSpread.Source = frm1.vspdData2
			If ggoSpread.SSCheckChange = True Then
				IntRetCD = DisplayMsgBox("700001", parent.VB_YES_NO,frm1.vspdData.ActiveRow & "행 :","x")   <% '⊙: "Will you destory previous data" %>
				If IntRetCD = vbNo Then
					CheckDTL = frm1.vspdData.ActiveRow
				    Exit Function		
				End If
			End If 			
		
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData
			
        With frm1.vspdData        				
			If .ActiveRow > 0 Then
				.focus
				.ReDraw = False
		
				ggoSpread.Source = frm1.vspdData	
				ggoSpread.CopyRow
                SetSpreadColor .ActiveRow, .ActiveRow
    
				.ReDraw = True
    		    .Focus
			End If
		End With
	Else
        If Frm1.vspdData2.MaxRows < 1 Then
           Exit Function
        End If
    
        With frm1.vspdData2
			If .ActiveRow > 0 Then
				.focus
				.ReDraw = False
		
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.CopyRow
                SetSpreadColor2 .ActiveRow, .ActiveRow
    
				.ReDraw = True
    		    .Focus
			End If
		End With
	End If
	
    Set gActiveElement = document.ActiveElement   
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
		
    If gSpreadFlg = 1 Then
		frm1.vspdData.Redraw = False 
		if frm1.vspdData.maxrows < 1 then exit function
			
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.EditUndo
		
		Call DbQuery2(Trim(GetSpreadText(frm1.vspdData,C_ConDt_H, frm1.vspdData.ActiveRow,"X","X")), Trim(GetSpreadText(frm1.vspdData,C_DocCur, frm1.vspdData.ActiveRow,"X","X")))
		frm1.vspdData.Redraw = True 
	Else
		frm1.vspdData2.Redraw = False 
		if frm1.vspdData2.maxrows < 1 then exit function
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.EditUndo
		frm1.vspdData2.Redraw = True
	End If
	
End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    
    Dim Spread
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status    
    
    FncInsertRow = False                                                         '☜: Processing is NG
	
	frm1.vspdData.Col = 0
	frm1.vspdData.Row = frm1.vspdData.ActiveRow					
	If frm1.vspdData.Text = ggoSpread.InsertFlag And gSpreadFlg = 2 Then Exit Function
					
    If gSpreadFlg = 1 Then
		Set Spread = frm1.vspdData
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData

			ggoSpread.Source = frm1.vspdData2
			If ggoSpread.SSCheckChange = True Then
				IntRetCD = DisplayMsgBox("700001", parent.VB_YES_NO,frm1.vspdData.ActiveRow & "행 :","x")   <% '⊙: "Will you destory previous data" %>
				If IntRetCD = vbNo Then
					CheckDTL = frm1.vspdData.ActiveRow
				    Exit Function		
				End If
			End If 	
	
	Else
		Set Spread = frm1.vspdData2
	End If
							
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If
		
	With frm1
        Spread.ReDraw = False
        Spread.focus
        ggoSpread.Source = Spread
        ggoSpread.InsertRow ,imRow
        
        If gSpreadFlg = 1 Then
			SetSpreadColor Spread.ActiveRow, Spread.ActiveRow + imRow - 1
		Else
			SetSpreadColor2 Spread.ActiveRow, Spread.ActiveRow + imRow - 1
		End If	
			
  		lgBlnFlgChgValue = True
        Spread.ReDraw = True       
        
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If gSpreadFlg = 2 Then	
		Call Spread.SetText(C_ConDt2, Spread.ActiveRow, Replace(Trim(GetSpreadText(frm1.vspdData,C_ConDt_H, frm1.vspdData.ActiveRow,"X","X")), "-", ""))
		Call Spread.SetText(C_DocCur2, Spread.ActiveRow, Replace(Trim(GetSpreadText(frm1.vspdData,C_DocCur, frm1.vspdData.ActiveRow,"X","X")), "-", ""))
	Else
		Call Spread.SetText(C_ConDt, Spread.ActiveRow, Left(prDBSYSDate,7))
	End If
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
 End Function
'========================================================================================================
Function FncDeleteRow()
	Dim lDelRows
	Dim iDelRowCnt, i
	  
	'-------------------    
    If gSpreadFlg = 1 Then
		If Frm1.vspdData.MaxRows < 1 then
			Exit function
		End if	
		
		With Frm1.vspdData 
    		.focus
    		ggoSpread.Source = frm1.vspdData 
    		lDelRows = ggoSpread.DeleteRow
    		
		End With
	Else
		If Frm1.vspdData2.MaxRows < 1 then
			Exit function
		End if	
		
		With Frm1.vspdData2
    		.focus
    		ggoSpread.Source = frm1.vspdData2
    		lDelRows = ggoSpread.DeleteRow
		End With
	End If
	
	lgBlnFlgChgValue = True
    Set gActiveElement = document.ActiveElement 
    '-----------------------------------------
End Function

'========================================================================================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData
	Call parent.FncPrint()             <%'☜: Protect system from crashing%>
End Function

'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_MULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
'    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
'	Call SetSpreadColor1(-1)
'    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Cur,C_Item_Price,"C","I","X","X")
End Sub

'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	ggoSpread.Source = frm1.vspdData
	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")   <%'⊙: "Will you destory previous data"%>

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
Function DbQuery()
	Err.Clear													<%'☜: Protect system from crashing%>

	DbQuery = False												<%'⊙: Processing is NG%>

	Dim strVal
	  
	If   LayerShowHide(1) = False Then
		Exit Function 
	End If	  
	
		If lgIntFlgMode = parent.OPMD_UMODE Then
	
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtConDt=" & Trim(frm1.HtxtConDt.value)		<%'☆: 조회 조건 데이타 %>
			
			strVal = strVal & "&lgStrPrevKey=" & Trim(frm1.HlgStrPrevKey.value)
			strVal = strVal & "&lgStrPrevKey2=" & Trim(frm1.HlgStrPrevKey2.value)

		Else

			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtConDt=" & Trim(frm1.txtConDt.Text)		<%'☆: 조회 조건 데이타 %>		
			
			strVal = strVal & "&lgStrPrevKey=" & Trim(frm1.HlgStrPrevKey.value)
			strVal = strVal & "&lgStrPrevKey2=" & Trim(frm1.HlgStrPrevKey2.value)
		End If
		
'		msgbox strVal
		
  Call RunMyBizASP(MyBizASP, strVal)
				 
	DbQuery = True												<%'⊙: Processing is NG%>
End Function
'========================================================================================================
Function DbQuery2(Key1, Key2)
	Err.Clear													<%'☜: Protect system from crashing%>

	DbQuery2 = False												<%'⊙: Processing is NG%>

	Dim strVal
	  		
	If   LayerShowHide(1) = False Then
		Exit Function 
	End If	  
	
		strVal = BIZ_PGM_ID & "?txtMode=" & "Grid2"					<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&Key1=" & Left(Replace(Key1, "-", ""),6)							<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&Key2=" & Key2	
		
		strVal = strVal & "&lgStrPrevKey=" & Trim(frm1.HlgStrPrevKey3.value)
  
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		
  Call RunMyBizASP(MyBizASP, strVal)
				 
	DbQuery2 = True												<%'⊙: Processing is NG%>
End Function 
'========================================================================================================
 Function DbSave() 
    Dim lRow
    Dim lRow2
	Dim lGrpCnt
	Dim strVal
	Dim strDel
    Dim parentRow
	Dim Zsep
	Dim iColSep
	Dim iRowSep

	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규]
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규]
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제]
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size

	Call LayerShowHide(1)
	
	'-----------------------
	'Data manipulate area
	'-----------------------

	strVal = ""

    Zsep = "@"
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	'-----------------------
	'Data manipulate area
	'-----------------------

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID

		lGrpCnt = 1

		strVal = ""
		strDel = ""

		For lRow = 1 To .vspdData.MaxRows
		 .vspdData.Row = lRow
		 .vspdData.Col = 0
		 
		 Select Case .vspdData.Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag         <% '☜: 업데이트 %>
					  
					If .vspdData.Text = ggoSpread.InsertFlag Then
						strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		<% '☜: C=Create, Row위치 정보 %>
					Else
						strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		<% '☜: U=Update, Row위치 정보 %>
					End If					

				 .vspdData.Col = C_ConDt													<% '2 년월 %>
				 strVal = strVal & Trim(Left(Replace(.vspdData.Text, "-", ""),6)) & parent.gColSep
				
				 .vspdData.Col = C_DocCur													<% '3 통화 %>
				 strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				 
				 .vspdData.Col = C_ExchRate_Fir												<% '4 최초고시환율 %>
				 strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
				      
				 .vspdData.Col = C_ExchRate_Las												<% '5 최종고시환율 %>
				 strVal = strVal & UNICDbl(.vspdData.Text) & parent.gColSep
				                        
				 .vspdData.Col = C_Remark											        <% '6 비고 %>
				 strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				  
				 .vspdData.Col = C_ConDt_H													<% '7 년월 Key %>
				 strVal = strVal & Trim(Left(Replace(.vspdData.Text, "-", ""),6)) & parent.gRowSep 

				Case ggoSpread.DeleteFlag													<% '☜: 삭제 %>
					strVal = strVal & "D" & parent.gColSep & lRow & parent.gColSep									<% '☜: D=Update, Row위치 정보 %>
					
					.vspdData.Col = C_ConDt_H												<% '2 년월 %>
					strVal = strVal & Trim(Left(Replace(.vspdData.Text, "-", ""),6)) & parent.gColSep
				    
				    .vspdData.Col = C_DocCur												<% '3 통화 %>
				    strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				 					
					strVal = strVal & parent.gRowSep

			End Select
			
				If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 

				   Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
				   objTEXTAREA.name = "txtCUSpread"
				   objTEXTAREA.value = Join(iTmpCUBuffer,"")
				   divTextArea.appendChild(objTEXTAREA)

				   iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
				   ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
				   iTmpCUBufferCount = -1
				   strCUTotalvalLen  = 0
				End If

				iTmpCUBufferCount = iTmpCUBufferCount + 1

				If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
				   iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
				   ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
				End If				         
				iTmpCUBuffer(iTmpCUBufferCount) =  strVal
				strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
				
				strVal  = ""
				
		Next
	
		If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread" 
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")		  
		   divTextArea.appendChild(objTEXTAREA)		   
		End If
		
		strVal = ""
		For lRow = 1 To .vspdData2.MaxRows
			 .vspdData2.Row = lRow
			 .vspdData2.Col = 0
			 
			 Select Case .vspdData2.Text
					Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag         <% '☜: 업데이트 %>
						  
						If .vspdData2.Text = ggoSpread.InsertFlag Then
							strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		<% '☜: C=Create, Row위치 정보 %>
						Else
							strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		<% '☜: U=Update, Row위치 정보 %>
						End If
						
						.vspdData2.Col = C_BpCd													<% '2 업체 %>
						strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
					
						.vspdData2.Col = C_ExchApplyCd											<% '3 적용고시 %>
						strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
						     
						.vspdData2.Col = C_Remark2												<% '4 비고사항 %>
						strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
						                       
						.vspdData2.Col = C_ConDt2											    <% '5 년월 %>
						strVal = strVal & Trim(Left(Replace(.vspdData2.Text, "-", ""),6)) & parent.gColSep
						
						.vspdData2.Col = C_DocCur2											    <% '6 통화 %>
						strVal = strVal & Trim(.vspdData2.Text) & parent.gRowSep

					Case ggoSpread.DeleteFlag													<% '☜: 삭제 %>
						strVal = strVal & "D" & parent.gColSep & lRow & parent.gColSep									<% '☜: D=Update, Row위치 정보 %>
						
					    .vspdData2.Col = C_ConDt2												<% '2 년월 %>
					    strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
					 					
					 	.vspdData2.Col = C_DocCur2											    <% '3 통화 %>
						strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
						
						.vspdData2.Col = C_BpCd													<% '4 업체 %>
						strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
						
						'---삭제 no
						.vspdData2.Col = C_ConDt2												<% '5 년월 %>
					    strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
					 					
					 	.vspdData2.Col = C_DocCur2											    <% '6 통화 %>
						strVal = strVal & Trim(.vspdData2.Text) & parent.gRowSep
					    
			End Select	
		Next	
						
   .txtSpread.value =  strVal
   .txtFlgMode.value = lgIntFlgMode
   
'   msgbox .txtSpread.value
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)													<% '☜: 비지니스 ASP 를 가동 %>

	End With

	DbSave = True              <% '⊙: Processing is NG %>
 End Function
 
'========================================================================================================
 Function DbQueryOk()             <% '☆: 조회 성공후 실행로직 %>

  <% '------ Reset variables area ------ %>
  lgIntFlgMode = parent.OPMD_UMODE           <% '⊙: Indicates that current mode is Update mode %>
  lgBlnFlgChgValue = False
  Call ggoOper.LockField(Document, "Q")        <% '⊙: This function lock the suitable field %>
  Call SetToolBar("1110111100111111")         <% '⊙: 버튼 툴바 제어 %>
  If frm1.vspdData.MaxRows > 0 Then
	Call DbQuery2(Trim(GetSpreadText(frm1.vspdData,C_ConDt_H, 1,"X","X")), Trim(GetSpreadText(frm1.vspdData,C_DocCur, 1,"X","X")))
	frm1.vspdData.Focus	
  Else
	frm1.txtConDt.focus
  End If
  
  
  
 End Function
'========================================================================================================
 Function DbSaveOk()              <%'☆: 저장 성공후 실행 로직 %>
  Call ggoOper.ClearField(Document, "2")
  Call InitVariables
  Call MainQuery()
 End Function

'=========================================================================================================
Sub txtConDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtConDt.Focus
	End If
End Sub

'=========================================================================================================
Sub txtConDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

 '========================================================
Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>마감-업체별적용환율등록</font></td>
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
						<TD CLASS="TD5" NOWRAP>년월</TD>
						<TD CLASS="TD6" NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtConDt" CLASS=FPDTYYYYMM tag="11XXXU" Title="FPDATETIME" ALT=시작지급일 id=txtConDt></OBJECT>');</SCRIPT>			   
						</TD>
						<TD CLASS="TD6" NOWRAP></TD>
						<TD CLASS="TD6"></TD>
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
	 	<DIV ID="TabDiv"  STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
          <TABLE <%=LR_SPACE_TYPE_60%>>	
			<TR>	 
				<TD HEIGHT=220 WIDTH=100% valign=top COLSPAN=4>
					<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
				</TD>
			</TR>	
			<TR>	 
				<TD HEIGHT=100% WIDTH=100% valign=top COLSPAN=4>
					<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
				</TD>
			</TR>
		   </TABLE>
		  </DIV>
	   </TD>	
	</TR>	
    <TR>
     <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>    
   </TABLE>
  </TD>
 </TR>
 <TR>
  <TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
 </TR>  
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%>  FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>

<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HlgStrPrevKey" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HlgStrPrevKey2" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HlgStrPrevKey3" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HtxtConDt" tag="24" TABINDEX="-1">     

</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
