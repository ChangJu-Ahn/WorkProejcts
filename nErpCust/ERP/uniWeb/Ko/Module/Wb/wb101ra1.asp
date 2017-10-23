<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incServer.asp" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/commonPopup.vbs">     </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js">         </SCRIPT>

<Script Language="VBScript">
Option Explicit                            

Const POPUP_TITLE = 0                                                           '--- Index of POP-UP Title
Const TABLE_NAME  = 1                                                           '--- Index of DB table name to query
Const CODE_CON    = 2                                                           '--- Index of Code Condition value
Const NAME_CON    = 3                                                           '--- Index of Name Condition value
Const WHERE_CON   = 4                                                           '--- Index of Where Clause
Const TEXT_NAME   = 5                                                           '--- Index of Textbox Name

Dim lgSortKey

Dim lgStrNextCodeKey                                                            '--- Next code
Dim lgStrNextNameKey                                                            '--- Next name
Dim lgNameDuplication

Dim arrParent
Dim arrParam                                                                    '--- First Parameter Group
Dim arrTblField                                                                 '--- Second Parameter Group(DB Table Field Name)
Dim arrGridHdr                                                                  '--- Third Parameter Group(Column Captions of the SpreadSheet)
Dim arrReturn                                                                   '--- Return Parameter Group
Dim gintDataCnt                                                                 '--- Data Counts to Query
Dim arrFieldType

    arrParent   = window.dialogArguments    
    
	arrParam    = arrParent(0)
	arrTblField = arrParent(1)
	arrGridHdr  = arrParent(2)	
		
	top.document.title = arrParam(POPUP_TITLE)

'========================================================================================================
Sub InitSpreadSheet()
    Dim i
    Dim iArr
    Dim iLen
	
    ReDim arrFieldType(gintDataCnt - 1)    
	
    vspdData.ReDraw = False
		    
    ggoSpread.Source = vspdData
    vspdData.OperationMode = 3
	
    vspdData.MaxCols = gintDataCnt
    vspdData.MaxRows = 0
	    
    ggoSpread.Spreadinit
		
    ggoSpread.SSSetEdit 1, arrGridHdr(0), 18	' 코드 
		
    For i = 1 To gintDataCnt - 1
        If InStr(1, UCase(arrTblField(i)), "CONVERT") > 0 And InStr(1, UCase(arrTblField(i)), "CHAR") > 0 Then 
           ggoSpread.SSSetEdit i + 1, arrGridHdr(i), 25, 1,,999
        Else
           ggoSpread.SSSetEdit i + 1, arrGridHdr(i), 30,  ,,999
        End If
    Next
	    
    For i = 0 To gintDataCnt - 1
        If InStr(1, UCase(arrTblField(i)), gColSep) > 0 Then
           iArr = Split(UCase(arrTblField(i)),gColSep)
              
           iLen = 0
               
           If Len(Trim(iArr(0))) > 2 Then
                  iLen = Cint(Mid(iArr(0),3))
           End If

           arrFieldType(i) = ""
               
           Select Case Mid(iArr(0),1,2)
                    Case "ED"   '일반문자 
                           If iLen > 0 Then
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), iLen,,,999
                           Else    
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), 30  ,,,999
                           End If   
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)
                    Case "DD"   '날짜 
                           If iLen > 0 Then
                              ggoSpread.SSSetDate   i + 1,arrGridHdr(i),iLen,2,gDateFormat
                           Else    
                              ggoSpread.SSSetDate   i + 1,arrGridHdr(i),  12,2,gDateFormat
                           End If   
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)
                    Case "F2","F3","F4","F5"
                           If iLen > 0 Then
                              ggoSpread.SSSetFloat  i + 1,arrGridHdr(i),iLen,Mid(iArr(0),2,1),ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
                           Else
                              ggoSpread.SSSetFloat  i + 1,arrGridHdr(i),  17,Mid(iArr(0),2,1),ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
                           End If
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)
                    Case "TT"   ' Time
                           If iLen > 0 Then
                              ggoSpread.SSSetTime   i + 1,arrGridHdr(i),iLen,,1,1
                           Else
                              ggoSpread.SSSetTime   i + 1,arrGridHdr(i),  12,,1,1
                           End If
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)
                    Case "HH"  
                             vspdData.Col = i + 1
                             vspdData.ColHidden = True
                           If iLen > 0 Then
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), iLen
                           Else    
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), 30
                           End If   
                           arrTblField(i) = iArr(1) 
                           arrFieldType(i) = Mid(iArr(0),1,2)
                           
           End Select             
        End If    
    Next

    ggoSpread.Source = vspdData                         '2003/06/26 leejinsoo
    ggoSpread.SpreadLockWithOddEvenRowColor()

    vspdData.ReDraw = True
End Sub

'========================================================================================================
Sub Form_Load()
    
	Dim intLoopCnt

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	
	<% Call loadInfTB19029A("Q", "*","NOCOOKIE","COMMONPOPUP") %>
 	
	lgSortKey        = 1	
	lgStrNextCodeKey = ""
	lgStrNextNameKey = ""
	lgNameDuplication = "F"

	gintDataCnt      = 0
	
	For intLoopCnt = 0 To Ubound(arrTblField)
		If arrTblField(intLoopCnt) <> "" Then
           gintDataCnt = gintDataCnt + 1	
		Else
           Exit For
		End If
	Next
	
	If gintDataCnt < 2 Then 
	    txtNm.classname = UCN_PROTECTED 
	    txtNm.ReadOnly = True
	    txtNm.TabIndex = "-1"
	End If

	lblTitle.innerHTML = arrParam(TEXT_NAME)
	txtCd.value        = arrParam(CODE_CON)
	txtNm.value        = arrParam(NAME_CON)
	Self.Returnvalue = Array("")

	Call InitSpreadSheet()
	Call FncQuery()

End Sub

'========================================================================================================
Sub OKClick()

	Dim intColCnt

	If vspdData.MaxRows < 1 Then
	   Call CancelClick()
	   Exit Sub		
	End If
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
		
		vspdData.Row = vspdData.ActiveRow
					
		For intColCnt = 0 To vspdData.MaxCols - 1
			vspdData.Col = intColCnt + 1
			arrReturn(intColCnt) = vspdData.Text
		Next
							
		Self.Returnvalue = arrReturn
	End If	
	
    Call CancelClick()
	
End Sub

'========================================================================================================
Sub CancelClick()
	Self.Close()
End Sub

'========================================================================================================
Function FncQuery()

    vspdData.MaxRows = 0

	lgStrNextCodeKey = Trim(txtCd.value)
	lgStrNextNameKey = Trim(txtNm.value)
	lgNameDuplication = "F"

	Call DbQuery()
	
End Function

'========================================================================================================
Sub vspdData_Click(Col, Row)
    
    If Row = 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey
            lgSortKey = 1
        End If
    End If
    
End Sub

'========================================================================================================
Function vspdData_DblClick( Col,  Row)

	If Row = 0 Then
	   Exit Function
	End If
	
	If vspdData.MaxRows = 0 Then
	   Exit Function
	End If

    If Row > 0 Then
       Call OKClick()
    End If

End Function

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	
	If vspdData.MaxRows = 0 Then
	   Exit Function
	End If
	
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function
	
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop)  Then
       If lgStrNextCodeKey <> "" Or lgStrNextNameKey <> "" Then
          Call DbQuery
       End If
    End if
End Sub
	
'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim arrStrDT
    Dim iLoop

	Call LayerShowHide(1)

    arrStrDT = ""
        
    For iLoop = 0 To gintDataCnt - 1
        arrStrDT = arrStrDT & Trim(arrFieldType(iLoop)) & gColSep
    Next
    
    strVal = "wb101rb1.asp" & "?txtTable="    & Trim(arrParam(TABLE_NAME)) 
    strVal =     strVal & "&txtNextCode=" & lgStrNextCodeKey
    strVal =     strVal & "&txtNextName=" & lgStrNextNameKey
    strVal =     strVal & "&NameDuplication=" & lgNameDuplication    
    strVal =     strVal & "&txtWhere="    & Trim(arrParam(WHERE_CON)) 
    strVal =     strVal & "&arrField1="   & Trim(arrTblField(0)) 
    strVal =     strVal & "&arrField2="   & Trim(arrTblField(1)) 
    strVal =     strVal & "&arrField3="   & Trim(arrTblField(2)) 
    strVal =     strVal & "&arrField4="   & Trim(arrTblField(3)) 
    strVal =     strVal & "&arrField5="   & Trim(arrTblField(4)) 
    strVal =     strVal & "&arrField6="   & Trim(arrTblField(5)) 
    strVal =     strVal & "&arrField7="   & Trim(arrTblField(gintDataCnt - 1))
    strVal =     strVal & "&arrStrDT="    & arrStrDT
    strVal =     strVal & "&gintDataCnt=" & gintDataCnt
		
    Call RunMyBizASP(MyBizASP, strVal)                                      '☜: 비지니스 ASP 를 가동 
		
End Function

'========================================================================================================
Function DbQueryOk()
   Dim IntRetCD

   If vspdData.MaxRows = 0 Then

      IntRetCD = DisplayMsgBoxA("900014") 
      If Trim(txtCd.value) > "" Then
         txtCd.Select 
         txtCd.Focus
      Else   
         txtNm.Select 
         txtNm.Focus
      End If
   Else
	  vspdData.Focus
   End If      

End Function	
</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	


</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtCd" SIZE=20 MAXLENGTH=50 tag="12XXXU" ALT="코드" ID="Text1"></TD>
			</TR>		
			<TR>
				<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtNm" SIZE=30 MAXLENGTH=50 tag="12"   ALT="코드명" ID="Text2"></TD>
			</TR>		
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/wb101ra1_vaSpread1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()"    onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   ONCLICK="OkClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=10><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=10 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

