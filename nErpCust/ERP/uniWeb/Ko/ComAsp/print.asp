<HTML>
<HEAD><TITLE>Print Preview </TITLE>
<% '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################%>
<% '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* %>
<!-- #Include file="../inc/IncServer.asp" -->
<%'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">

<%'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================%>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
	Option Explicit

	'Title은 첫 페이지를 위한 것입니다.
	'표지와 같은 의미로 큰 Font size를 사용합니다.
	Const FONT_TITLE1 = "/c/fn""MS UI Gothic""/fz""25""/fb1/fi0/fu1/fk0/fs1"
	Const FONT_TITLE2 = "/c/fn""MS UI Gothic""/fz""20""/fb1/fi0/fu1/fk0/fs1"
	Const FONT_TITLE3 = "/fn""MS UI Gothic""/fz""12""/fb0/fu0/fs0"

	'Header는 두번째 페이지 이후에 
	'실제 spread내용의 header부분입니다.
	Const FONT_HEADER1 = "/c/fn""MS UI Gothic""/fz""25""/fb1/fi0/fu1/fk0/fs1"
	Const FONT_HEADER2 = "/c/fn""MS UI Gothic""/fz""20""/fb1/fi0/fu1/fk0/fs1"
	Const FONT_HEADER3 = "/fn""MS UI Gothic""/fz""12""/fb0/fu0/fs0"

	'DD를 위해 처리해야 합니다.
	Const TAG_USER = "사용자 :"
	Const TAG_DATE = "인쇄일자 :"
	Const TAG_GRID = "그리드"

	Dim lgIntSelSpd
	Dim lgIntSpd
	Dim lgObjSpd()

	Dim lgStrTitle
	Dim lgStrHeader

	Dim lgObjEl
	Dim lgObjCond
	Dim lgObjComDtl

	Dim lgObjDoc
	Set lgObjDoc = window.dialogArguments

    lgIntSpd = 0
    lgIntSelSpd = 0

'기본 logic은 Excel과 유사합니다.
'단 spread는 따로 작업하지 않고 preview control에 연결만 해 줍니다.
'
'spread 이외의 부분은 spdData 라는 안보이는 스프레드에 넣습니다.
'첫 페이지는 spdData의 내용을 보여주고 
'다음 페이지부터는 combo박스에서 선택된 spread를 보여줍니다.

	Sub ShowPrint(objDoc)
	    Dim strDate
	    Dim strTitle
	    Dim strTabName
	    Dim objOption
	    Dim i
	    '0:탭하나    0:탭2이상     2:탭2이상 공통조건    4:탭2이상 조건그리드 b1b11ma2    8:탭2개이상 공통디테일 a3111ma1
	    Dim nIsDiv '  by Shin hyoung jae, 2001/4/3

	    Dim nTemp

	    On Error Resume Next
		
		If gLang = "JA" Then
	           spdData.Font.CharSet = 128   		 
		End If

	    strDate = "<%=UNIDateClientFormat(GetSvrDate)%>"

	    i = 0
	    i = objDoc.All.MyTab.Length

	    Err.Clear
	    On Error GoTo 0

	    strTitle = objDoc.Title

	    If i <= 1 Then 'Tab이 하나인 경우 

	        strTabName = objDoc.All.MyTab.Rows(0).cells(1).innerText

			nIsDiv = 0 '  by Shin hyoung jae, 2001/4/3

	        For i = 0 To objDoc.All.frm1.All.Length - 1
	            If UCase(objDoc.All.frm1.All(i).TagName) = "TD" Then
	                If Left(UCase(objDoc.All.frm1.All(i).className), 3) = "TAB" Then
	                    Set lgObjEl = objDoc.All.frm1.All(i)
	                    Exit For
	                End If
	            End If
	        Next
        	Call SearchSpread(lgObjEl)

	    Else 'Tab이 하나이상 

	        For i = 0 To objDoc.All.MyTab.Length - 1
	            If objDoc.All.TabDiv(i).Style.display = "" Then
	                strTabName = objDoc.All.MyTab(i).Rows(0).cells(1).innerText
	                Set lgObjEl = objDoc.All.TabDiv(i)
	                Exit For
	            End If
	        Next

			'  by Shin hyoung jae, 2001/4/3
			nIsDiv = 0
			nTemp = nIsDiv

			Call SearchSpread(lgObjEl)

			' 공통 조건은 반드시 filedset에 쌓여져 있어야 한다. DIV가 나오기 전에 
	        For i = 0 To objDoc.All.frm1.All.Length - 1
				If UCase(objDoc.All.frm1.All(i).TagName) = "DIV" Then
					Exit For
				End If

	            If UCase(objDoc.All.frm1.All(i).TagName) = "FIELDSET" Then
					Set lgObjCond = objDoc.All.frm1.All(i)
					nIsDiv = nIsDiv + 2
					nTemp = nIsDiv
					Exit For
	            End If
	        Next

			' 공통 조건 Spread Sheet. DIV가 나오기전에 
			' b1b11ma2 같은 엽기적 화면땜시 
			' 공통 그리드가 있을경우, div 뒤에 있음. h6014ma1
			For i = 0 To objDoc.All.frm1.All.Length - 1
				If UCase(objDoc.All.frm1.All(i).TagName) = "DIV" Then
					If lgIntSpd > 0 Then
						Exit For
					End If
				End If

	            If UCase(objDoc.All.frm1.All(i).TagName) = "OBJECT" And UCase(objDoc.All.frm1.All(i).Title) = "SPREAD" Then
					If lgIntSpd = 0 Then
						lgIntSpd = lgIntSpd + 1
						ReDim Preserve lgObjSpd(lgIntSpd)
						Set lgObjSpd(lgIntSpd) = objDoc.All.frm1.All(i)
						nIsDiv = nIsDiv + 4
						nTemp = nIsDiv
						Exit For
					End If
	            End If

	        Next

			' 공통 싱글 detail a3111ma1 땜시 
	        For i = 0 To objDoc.All.frm1.All.Length - 1
				If UCase(objDoc.All.frm1.All(i).TagName) = "DIV" Then
					Set lgObjComDtl = Nothing
					i = i + objDoc.All.frm1.All(i).All.Length
					nIsDiv = -999
				End If

	            If UCase(objDoc.All.frm1.All(i).TagName) = "TABLE" And nIsDiv = -999 Then
					Set lgObjComDtl = objDoc.All.frm1.All(i)
					nIsDiv = nTemp
					nIsDiv = nIsDiv + 8
					Exit For
	            End If
	        Next

			If nIsDiv = -999 Then
				nIsDiv = nTemp
			End If

	    End If

	    lgStrTitle = GetTitleStr(strTitle, strTabName, gUsrNm, strDate,nIsDiv)
	    lgStrHeader = GetHeaderStr(strTitle, strTabName, gUsrNm, strDate,nIsDiv)

	    cboZoom.value = 6

	    spdData.PrintGrid = False
	    spdData.PrintBorder = False

	    '조건 Field와 Single Data로 구성된 첫페이지를 만든다.
		'  by Shin hyoung jae, 2001/4/3
		Call SetPrvwDataSheet
		'0:탭하나    0:탭2이상     2:탭2이상 공통조건    4:탭2이상 조건그리드 b1b11ma2    8:탭2개이상 공통디테일 a3111ma1
		'msgbox "nIsDiv => " & nIsDiv
		select case nIsDiv
			Case 2
				Call MakeTitlePage(lgObjCond)
			Case 4
			Case 6
				Call MakeTitlePage(lgObjCond)
			Case 8
				Call MakeTitlePage(lgObjComDtl)
			Case 10
				Call MakeTitlePage(lgObjCond)
				Call MakeTitlePage(lgObjComDtl)
			Case 12
				Call MakeTitlePage(lgObjComDtl)
			Case 14
				Call MakeTitlePage(lgObjCond)
				Call MakeTitlePage(lgObjComDtl)
		End Select

	    Call MakeTitlePage(lgObjEl)

	    If lgIntSpd = 1 Then
	        lgIntSelSpd = 1
	    ElseIf lgIntSpd > 0 Then
	        cboGrid.disabled = False

	        '스프레드수만큼 콤보 아이템 추가 
	        For i = 1 To lgIntSpd
				Set objOption = Document.CreateElement("OPTION")
				objOption.Text = TAG_GRID & i
				objOption.Value = i

				cboGrid.add(objOption)
				Set objOption = Nothing
	        Next

	        lgIntSelSpd = 1
	        cboGrid.Value = 1
	    End If

            spdData.PrintOrientation = 2

  	    If lgIntSpd > 0 Then
	        lgObjSpd(lgIntSelSpd).PrintOrientation = 2
	    End If

	    Call ResetPreview

	End Sub

	'첫페이지의 헤더 
	Function GetTitleStr(strTitle, strTab, strUser, strDate, nIsDiv)
	    Dim strString

		If nIsDiv = 0 Then
		strString = "/n/n" & FONT_TITLE1 & Replace(StrTitle, "/", "//")
	    strString = strString & "/n/r" & FONT_TITLE3 & TAG_USER & " " & strUser
	    strString = strString & "/n/r" & FONT_TITLE3 & TAG_DATE & " " & strDate & "/n"
		Else
		strString = "/n/n" & FONT_TITLE1 & Replace(StrTitle, "/", "//")
	    strString = strString & "/n/n" & FONT_TITLE2 & Replace(strTab, "/", "//") & "%%Grid%%"
	    strString = strString & "/n/r" & FONT_TITLE3 & TAG_USER & " " & strUser
	    strString = strString & "/n/r" & FONT_TITLE3 & TAG_DATE & " " & strDate & "/n"
		End If

	    GetTitleStr = strString
	End Function	

	'두번째 페이지 이후의 헤더 
	Function GetHeaderStr(strTitle, strTab, strUser, strDate, nIsDiv)
	    Dim strString

		If nIsDiv = 0 Then
	    strString = FONT_HEADER1 & Replace(StrTitle, "/", "//")
	    strString = strString & "/n/r" & FONT_HEADER3 & TAG_USER & strUser
	    strString = strString & "/n/r" & FONT_HEADER3 & TAG_DATE & strDate & "/n"
	    Else
	    strString = FONT_HEADER1 & Replace(StrTitle, "/", "//")
	    strString = strString & "/n/n" & FONT_HEADER2 & Replace(strTab, "/", "//") & "%%Grid%%"
	    strString = strString & "/n/r" & FONT_HEADER3 & TAG_USER & strUser
	    strString = strString & "/n/r" & FONT_HEADER3 & TAG_DATE & strDate & "/n"
	    End If
	    GetHeaderStr = strString
	End Function

	'Grid Combo가 바뀔때 선택된 spread를 preview control에 연결하는 함수입니다.
	Sub ResetPreview()
	    Dim Index
	    Dim strHeader
	    Dim strTitle
 	    Dim IntRetCD

	    Call SetMargin

	    If cboGrid.value = "" Then
			strTitle = Replace(lgStrTitle, "%%Grid%%", "")
			strHeader = Replace(lgStrHeader, "%%Grid%%", "")
	    Else
			strTitle = Replace(lgStrTitle, "%%Grid%%", " " & TAG_GRID & cboGrid.value)
			strHeader = Replace(lgStrHeader, "%%Grid%%", " " & TAG_GRID & cboGrid.value)
	    End If

	    spdData.PrintHeader = strTitle

	    spdData.PrintFooter = "/c/p"
	    spdData.PrintFooter = "/c/p // " & CStr(spdData.PrintPageCount)

	    spdData.StartingRowNumber = 1
	    spdData.PrintFirstPageNumber = 1

	    If lgIntSpd > 0 Then
	        lgObjSpd(lgIntSelSpd).PrintFooter = "/c/p"
	        lgObjSpd(lgIntSelSpd).PrintHeader = strHeader
	        lgObjSpd(lgIntSelSpd).StartingRowNumber = 1
	        lgObjSpd(lgIntSelSpd).PrintFirstPageNumber = spdData.PrintPageCount + 1
            lgObjSpd(lgIntSelSpd).PrintRowHeaders = False 	        
			'재고이동 화면 관련 수정.
			if (lgObjSpd(lgIntSelSpd).PrintPageCount) = -709 then
			    spdData.PrintFooter = "/c/p // " & CStr(CInt(spdData.PrintPageCount)+1)
            else
                spdData.PrintFooter = "/c/p // " & CStr(CInt(spdData.PrintPageCount) + CInt(lgObjSpd(lgIntSelSpd).PrintPageCount))
	        end if
	        lgObjSpd(lgIntSelSpd).PrintFooter = "/c/p // " & CStr(CInt(spdData.PrintPageCount) + CInt(lgObjSpd(lgIntSelSpd).PrintPageCount))
	    End If

	    spvwData.hWndSpread = spdData.hWnd
	    spvwData.PageCurrent = 1

		'  by hurjun 2002/7/31  *********************************************
		If spvwData.hWndSpread = 0 Then
			IntRetCD = DisplayMsgBox("900033","X","X","X") 		
		Close
		End If
		
	'  *******************************************************************

	    Call spvwData_PageChange(1)

	End Sub

	Sub SetMargin()
	    If spdData.PrintMarginLeft = 0 _
	    And spdData.PrintMarginRight = 0 _
	    And spdData.PrintMarginTop = 0 _
	    And spdData.PrintMarginBottom = 0 Then
	        spdData.PrintMarginLeft = 0.5 * 1440
	        spdData.PrintMarginRight = 0.5 * 1440
	        spdData.PrintMarginTop = 0.75 * 1440
	        spdData.PrintMarginBottom = 0.75 * 1440
	    End If

	    If lgIntSpd > 0 Then
	    
			lgObjSpd(lgIntSelSpd).FontSize = 10
			lgObjSpd(lgIntSelSpd).RowHeight(0) = 20
			
	        lgObjSpd(lgIntSelSpd).PrintMarginLeft = spdData.PrintMarginLeft
	        lgObjSpd(lgIntSelSpd).PrintMarginRight = spdData.PrintMarginRight
	        lgObjSpd(lgIntSelSpd).PrintMarginTop = spdData.PrintMarginTop
	        lgObjSpd(lgIntSelSpd).PrintMarginBottom = spdData.PrintMarginBottom
	        tempdata.PrintMarginLeft = spdData.PrintMarginLeft
	        tempdata.PrintMarginRight = spdData.PrintMarginRight
	        
	        If optVorH(0).checked = True Then
				tempdata.PrintMarginTop = 600 '세로 
	        else
				tempdata.PrintMarginTop = 900 '가로 
	        end if
	        
	        tempdata.PrintMarginBottom = spdData.PrintMarginBottom
	    End If
	End Sub


	'document를 검색하여 spread를 찾아 lgObjSpd에 저장합니다.
	Sub SearchSpread(objEl)
	    Dim i

	    For i = 0 To objEl.All.Length - 1
	        If UCase(objEl.All(i).TagName) = "OBJECT" Then
	            If UCase(objEl.All(i).Title) = "SPREAD" And objEl.All(i).Style.display <> "none" Then
	                lgIntSpd = lgIntSpd + 1
	                ReDim Preserve lgObjSpd(lgIntSpd)
	                Set lgObjSpd(lgIntSpd) = objEl.All(i)
	            End If
	        End If
	    Next
	End Sub

	' by Shin hyoung jae, 2001/4/3
	' by Shin hyoung jae, 2001/7/11  spdData.MaxCols = 8
	Sub SetPrvwDataSheet()
		Dim i

	    spdData.MaxRows = 0
	    spdData.MaxCols = 10 ' by Shin hyoung jae, 2002/3/11 org -> spdData.MaxCols = 8

	    For i = 1 To 10 ' by Shin hyoung jae, 2002/3/11 org -> For i = 1 To 8
			spdData.ColWidth(i) = 2
	    Next
	End Sub

	'조건과 Single 데이터로 구성된 첫페이지를 만든다.
	'logic은 Excel의 경우과 동일하며 spread내용 설정부분만이 없다.
	Sub MakeTitlePage(objEl)

	    On Error Resume Next

	    Dim blnFirstTD
	    Dim i, j, k
		Dim bIsSameLine ' by Shin hyoung jae, 2001/4/2
		Dim bIsDisplay  ' by Shin hyoung jae, 2001/6/28

		bIsSameLine = False ' by Shin hyoung jae, 2001/4/2
		bIsDisplay = True   ' by Shin hyoung jae, 2001/6/28

		If gLang = "JA" Then
           spdData.Font.CharSet = 128   		 
		End If

		spdData.MaxRows = spdData.MaxRows + 1
		spdData.Row     = spdData.MaxRows
    
	    For i = 0 To objEl.All.Length - 1

            Select Case UCase(objEl.All(i).TagName)
	            Case "TR"
					' by Shin hyoung jae, 2001/6/28
					' q1411ma1 에서 앞탭에서 선택에 따라 뒤탭의 TR 이 display가 none 이됨.
					If UCase(objEl.All(i).Style.Display) = "NONE" Then
						bIsDisplay = False
					Else
						bIsDisplay = True
					End If

					' by Shin hyoung jae, 2001/4/2
					If bIsSameLine = False Then
						spdData.MaxRows = spdData.MaxRows + 1
						spdData.Row = spdData.MaxRows
						spdData.Col = 1
						blnFirstTD = True
					End If

	            Case "TD"
					If bIsDisplay = True Then ' by Shin hyoung jae, 2001/6/28
						' by Shin hyoung jae, 2001/4/2 add TD18 TD19 추가 
						If UCase(objEl.All(i).className) = "TD5" Or UCase(objEl.All(i).className) = "TD18" Or UCase(objEl.All(i).className) = "TDT" Then  ' 라벨명 
						    ' by Shin hyoung jae, 2002/3/11, a5114ma1
							If Not blnFirstTD Then
								spdData.Col = 5
							End If

							Call SetText(objEl.All(i).innerText)
							spdData.Col = spdData.Col + 1
							bIsSameLine = False ' by Shin hyoung jae, 2001/4/2
						End If

						If UCase(objEl.All(i).className) = "TD6" Then
						' by Shin hyoung jae, 2001/4/2
						' 프린트시 위치 잘못나오는것 땜시 

							If bIsSameLine = False Then
								bIsSameLine = True
							Else
								bIsSameLine = False
							End If

							' 없는건 반드시 <TD CLASS=TD6></TD>  사이가 붙어야 한다. 반드시 
							If Trim(objEl.All(i).innerText) = "" Then
								bIsSameLine = False
							Else
                                ' by Shin hyoung jae, 2002/05/30 c2010ba1 에서 에러 잡힘.
                                'If objEl.All.Length - 1 > i Then
                                    If UCase(objEl.All(i+1).TagName) = "INPUT" Then
                                    ElseIf UCase(objEl.All(i+1).TagName) = "SELECT" Then
                                    ElseIf UCase(objEl.All(i+1).TagName) = "OBJECT" Then
                                    ElseIf UCase(objEl.All(i+1).TagName) <> "TABLE" Then ' by Shin hyoung jae, 2001/7/11
                                        Call SetText(objEl.All(i).innerText)
                                        spdData.Col = spdData.Col + 1
                                    End If
                                'Else
                                '    Call SetText(objEl.All(i).innerText)
                                '    spdData.Col = spdData.Col + 1
                                'End If
							End If
					If UCase(objEl.All(i+1).TagName) = "TABLE" Then
						bIsSameLine = True
					End if

						End If

						blnFirstTD = False
					End If

	            Case "LEGEND"
	                spdData.MaxRows = spdData.MaxRows + 1
	                spdData.Row = spdData.MaxRows

	                spdData.Col = 1

	                Call SetText(objEl.All(i).innerText)
	                spdData.Col = spdData.Col + 1

	            Case "INPUT"
	                If UCase(objEl.All(i).Type) = "TEXT" Then
						' by Shin hyoung jae, 2001/4/2
						If UCase(objEl.All(i).Style.textTransform) = "UPPERCASE" Then
							Call SetText(UCase(objEl.All(i).Value))
						ElseIf UCase(objEl.All(i).Style.textTransform) = "LOWERCASE" Then
							Call SetText(LCase(objEl.All(i).Value))
						Else
							Call SetText(objEl.All(i).Value)
						End If
	                    spdData.Col = spdData.Col + 1

						' by Shin hyoung jae, 2001/7/27
						' text ~ text 이런거 찍히게 p1401ma3
						If UCase(objEl.All(i-1).TagName) = "TD" Then
							If UCase(objEl.All(i-1).className) <> "TD5" _		 
							   And UCase(objEl.All(i-1).className) <> "TD18" _  
							   And UCase(objEl.All(i-1).className) <> "TDT" _  
							   And UCase(objEl.All(i-1).valign) <> "TOP" Then   
								If objEl.All(i-1).innerText <> "" Then
									Call SetText(objEl.All(i-1).innerText)
									spdData.Col = spdData.Col + 1
								End If
							End If
						End If
						
						If (i + 1) < objEl.All.Length Then
							If UCase(objEl.All(i+1).TagName) = "TD" Then
								If UCase(objEl.All(i+1).className) <> "TD5" _	
								   And UCase(objEl.All(i+1).className) <> "TD18" _ 
								   And UCase(objEl.All(i+1).className) <> "TDT" _ 
								   And UCase(objEl.All(i+1).valign) <> "TOP" _
								   And UCase(objEl.All(i+2).TagName) <> "FIELDSET" Then  'p4611ma1
									If objEl.All(i+1).innerText <> "" Then
										Call SetText(objEl.All(i+1).innerText)
										spdData.Col = spdData.Col + 1
									End If
								End If
							End If
						End If

	                End If

					' by Shin hyoung jae, 2001/3/29
	                If UCase(objEl.All(i).Type) = "CHECKBOX" Or UCase(objEl.All(i).Type) = "RADIO" Then
						If objEl.All(i).checked = True Then
							Call SetText(objEl.All(i+1).innerText)
		                    spdData.Col = spdData.Col + 1
						End If
	                End If

					bIsSameLine = False ' by Shin hyoung jae, 2001/4/2

	            Case "TEXTAREA"
	                Call SetText(objEl.All(i).Value)
	                spdData.Col = spdData.Col + 1

					bIsSameLine = False ' by Shin hyoung jae, 2001/4/3

	            Case "SELECT"
					' by Shin hyoung jae, 2001/3/31
					If objEl.All(i).selectedindex >= 0 Then
						Call SetText(objEl.All(i).Options(objEl.All(i).selectedindex).Text)
						spdData.Col = spdData.Col + 1
					End If

					bIsSameLine = False ' by Shin hyoung jae, 2001/4/3

				Case "OBJECT"
	                If UCase(objEl.All(i).Title) = "FPDOUBLESINGLE" Or _
	                   UCase(objEl.All(i).Title) = "FPDATETIME" Then
	                    Call SetText(objEl.All(i).Text)
	                    spdData.Col = spdData.Col + 1

						' by Shin hyoung jae, 2001/6/28
						' OCX바로뒤의 일, day 이런거 찍히게 
						If UCase(objEl.All(i-1).TagName) = "TD" Then
							If UCase(objEl.All(i-1).className) <> "TD5" _		 
							   And UCase(objEl.All(i-1).className) <> "TD18" _  
							   And UCase(objEl.All(i-1).valign) <> "TOP" Then   
								If objEl.All(i-1).innerText <> "" Then
									Call SetText(objEl.All(i-1).innerText)
									spdData.Col = spdData.Col + 1
								End If
							End If
						End If
						
						If (i + 1) < objEl.All.Length Then
							' by Shin hyoung jae, 2001/6/28
							' OCX바로뒤의 일, day 이런거 찍히게 
							If UCase(objEl.All(i+1).TagName) = "TD" Then
								If UCase(objEl.All(i+1).className) <> "TD5" _	
								   And UCase(objEl.All(i+1).className) <> "TD18" _ 
								   And UCase(objEl.All(i+1).valign) <> "TOP" Then  
									If objEl.All(i+1).innerText <> "" Then
										Call SetText(objEl.All(i+1).innerText)
										spdData.Col = spdData.Col + 1
									End If
								End If
							End If
						End If

					End If

					bIsSameLine = False ' by Shin hyoung jae, 2001/4/2

	        End Select
	    Next

	End Sub

	'각 object의 Text길이에 맞게 컬럼의 길이 결정 
	Sub SetText(strText)
	    spdData.TypeMaxEditLen = 200
	    spdData.Text = strText

	    If spdData.ColWidth(spdData.Col) < Len(spdData.Text) + 3 And Len(strText) < 100 Then
	        spdData.ColWidth(spdData.Col) = Len(spdData.Text) + 3
	    End If
	End Sub

'//////// Event 처리 
	 Sub window_onload()
	    Dim ii
	    Dim iTotal 
	   
	    lgIntSpd = 0
	    lgIntSelSpd = 0

		Call GetGlobalVar()
		Call ShowPrint(lgObjDoc)
		
		spdData.Col  = -1
		spdData.Row  = -1
		
		spdData.Col2 = -1
		spdData.Row2 = -1
		spdData.FontName = "돋움체"

		For ii = 0 To spdData.maxcols
		    iTotal = iTotal + spdData.ColWidth(ii)
        Next    

		
		'For ii = 0 To spdData.maxcols
        '    spdData.ColWidth(ii) = spdData.ColWidth(ii) * 44 / iTotal + spdData.ColWidth(ii)
        'Next    


		
	End Sub

	Sub window_onUnload()
	    Dim i
	    If lgIntSpd > 0 Then
	        For i = 1 To lgIntSpd
	            Set lgObjSpd(i) = Nothing
	        Next
	    End If

	    Set lgObjEl = Nothing
	    Set lgObjDoc = Nothing
            Set lgObjCond = Nothing
	    Set lgObjComDtl = Nothing
	End Sub

	'page가 변경될 경우 Prev, Next 버튼 Enable/Disable

	Sub spvwData_PageChange(ByVal Page)

		If spvwData.hWndSpread = 0 Then	
			btnPrev.disabled = True
			btnNext.disabled = True
		ElseIf spvwData.hWndSpread = spdData.hWnd Then
	        	If spvwData.PageCurrent > 1 Then
	        	    btnPrev.disabled = False
	        	Else
	        	    btnPrev.disabled = True
	        	End If

		        If lgIntSpd > 0 Or spvwData.PageCurrent =< spdData.PrintPageCount Then
				btnNext.disabled = False
	        	Else
	            		btnNext.disabled = True
	        	End If
	    	Else
            		btnPrev.disabled = False

	        	If spvwData.PageCurrent < lgObjSpd(lgIntSelSpd).PrintPageCount Then
	            		btnNext.disabled = False
	        	Else
	            		btnNext.disabled = True
	        	End If
		End If

            	if lgIntSpd <= 0 and spvwData.PageCurrent = CInt(spdData.PrintPageCount) then
                   	btnNext.disabled = True
            	end if

	End Sub

	Sub cmdPrint_Click()
		Dim arrRet
		Dim nConSP
		Dim nConEP
		Dim nSpdSP
		Dim nSpdEP
	        Dim strDate
	        Dim strTitle
	        Dim strTabName
                Dim strHeader

        ' by Shin hyoung jae 2002/3/19
		On Error Resume Next
		nTotalPageCount.value = CInt(spdData.PrintPageCount) + CInt(lgObjSpd(lgIntSelSpd).PrintPageCount)

		arrRet = window.showModalDialog("PrintRange.asp", document, _
			"dialogWidth=240px; dialogHeight=170px; center=yes; help: No; resizable=True; scroll=No;status:No;")
		
		If arrRet(0) = "CLOSE" Then
			Exit Sub
		End If

		If arrRet(0) = "ALL" Then
			spdData.PrintType = 0 'PrintTypeAll
			spdData.Action = 13 'ActionPrint
		Else
                   
			If CInt(spdData.PrintPageCount) >= CInt(arrRet(1)) Then
				nConSP = arrRet(1)   ' 인쇄 시작 페이지 
			Else
				nConSP = 0
				nConEP = 0
			End If

			If CInt(spdData.PrintPageCount) >= CInt(arrRet(2)) Then
				nConEP = arrRet(2) '인쇄 끝페이지 
				nSpdSP = 0
				nSpdEP = 0
			Else
				If nConSP <> 0 Then
					nConEP = spdData.PrintPageCount
					nSpdSP = 1
					'nSpdEP = arrRet(2) - spdData.PrintPageCount
					nSpdEP = arrRet(2)   '2003-01-28 김승진 수정 
				Else
					nSpdSP = arrRet(1)
					nSpdEP = arrRet(2)
				End If
			End If

			If nConSP <> 0 Then
				spdData.PrintPageStart = nConSP
				spdData.PrintPageEnd = nConEP
				spdData.PrintType = 3 'PrintTypePageRange
				spdData.Action = 13 'ActionPrint
			End If
		End If

	    If lgIntSpd > 0 Then
			If arrRet(0) = "ALL" Then
				lgObjSpd(lgIntSelSpd).PrintType = 0 'PrintTypeAll
				lgObjSpd(lgIntSelSpd).Action = 13 'ActionPrint
			Else
				If nSpdSP <> 0 Then
					tempData.ColHeaderDisplay = 0  '2003-01-28 김승진 수정 
					tempData.col=1
					tempData.row=1
					tempData.col2=lgObjSpd(lgIntSelSpd).maxcols
					tempData.Row2=lgObjSpd(lgIntSelSpd).maxrows

					lgObjSpd(lgIntSelSpd).col = 1
					lgObjSpd(lgIntSelSpd).row = 1
					lgObjSpd(lgIntSelSpd).col2=lgObjSpd(lgIntSelSpd).maxcols
					lgObjSpd(lgIntSelSpd).Row2=lgObjSpd(lgIntSelSpd).maxrows

					tempData.maxcols=lgObjSpd(lgIntSelSpd).maxcols
					tempData.maxrows=lgObjSpd(lgIntSelSpd).maxrows
	
					tempData.clip=lgObjSpd(lgIntSelSpd).clip
					tempdata.PrintOrientation=lgObjSpd(lgIntSelSpd).PrintOrientation
					tempData.Col = -1
					tempData.Row = -1
'					tempData.FontSize = 10
'					tempData.RowHeight(0) = 14
'					tempData.CellBorderType = 16
'					tempData.CellBorderStyle = 1

					Dim iii
					for iii=1 to tempData.maxcols
						tempData.ColWidth(iii)=lgObjSpd(lgIntSelSpd).ColWidth(iii)
						tempData.Row=0
						tempData.col=iii
						lgObjSpd(lgIntSelSpd).Row=0
						lgObjSpd(lgIntSelSpd).Col=iii
						tempData.text=lgObjSpd(lgIntSelSpd).text
	
						lgObjSpd(lgIntSelSpd).Row = -1
						tempData.Row = -1
						tempdata.TypeHAlign =lgObjSpd(lgIntSelSpd).TypeHAlign 
					next


					tempdata.PrintPageStart = nSpdSP - 1 
					tempdata.PrintPageEnd = nSpdEP - 1
					tempdata.PrintType = 3 'PrintTypePageRange

					strDate = "<%=UNIDateClientFormat(GetSvrDate)%>"
					strTitle = objDoc.Title
					strTabName = objDoc.All.MyTab.Rows(0).cells(1).innerText
                    
                    If cboGrid.value = "" Then
						strTitle = Replace(lgStrTitle, "%%Grid%%", "")
						strHeader = Replace(lgStrHeader, "%%Grid%%", "")
					Else
						strTitle = Replace(lgStrTitle, "%%Grid%%", " " & TAG_GRID & cboGrid.value)
						strHeader = Replace(lgStrHeader, "%%Grid%%", " " & TAG_GRID & cboGrid.value)
					End If
				        tempdata.PrintFooter = "/c/p"
				        tempdata.PrintHeader = strHeader
	        			tempdata.StartingRowNumber = 1
				        tempdata.PrintFirstPageNumber = spdData.PrintPageCount + 1
       					tempdata.PrintRowHeaders = False 	        
				        tempdata.PrintFooter = "/c/p // " & CStr(nTotalPageCount.value)

					tempdata.Action = 13 'ActionPrint
				End If
			End If
	    End If

	End Sub

    ' by Shin hyoung jae 2001/2/28
    Sub optVorH_Click()
        If optVorH(0).checked = True Then
            spdData.PrintOrientation = 1
			If lgIntSpd > 0 Then
	            lgObjSpd(lgIntSelSpd).PrintOrientation = 1
			End If
        Else
            spdData.PrintOrientation = 2
			If lgIntSpd > 0 Then
	            lgObjSpd(lgIntSelSpd).PrintOrientation = 2
			End If
        End If

        spvwData.hWndSpread = spdData.hWnd
        Call ResetPreview()
    End Sub


    Sub CheckVorH()
        If optVorH(0).checked = True Then
            spdData.PrintOrientation = 1
			If lgIntSpd > 0 Then
	            lgObjSpd(lgIntSelSpd).PrintOrientation = 1
			End If
        Else
            spdData.PrintOrientation = 2
			If lgIntSpd > 0 Then
	            lgObjSpd(lgIntSelSpd).PrintOrientation = 2
			End If
        End If

    End Sub

	'첫페이지인 경우는 숨겨진 spdData의 내용을 
	'두번째 페이지 이후인 경우는 실제 spread의 내용을 보여줍니다.
	Sub cmdPrev_Click()
        Call CheckVorH
	    If spvwData.hWndSpread = spdData.hWnd Then
	        If spvwData.PageCurrent = 1 Then
	        Else
	            spvwData.PageCurrent = spvwData.PageCurrent - 1
	        End If
	    Else
	        If spvwData.PageCurrent = 1 Then
	            spvwData.hWndSpread = spdData.hWnd
	            spvwData.PageCurrent = spdData.PrintPageCount
	            Call spvwData_PageChange(1)
	        Else
	            spvwData.PageCurrent = spvwData.PageCurrent - 1
	        End If
	    End If

	End Sub

	'첫페이지인 경우는 숨겨진 spdData의 내용을 
	'두번째 페이지 이후인 경우는 실제 spread의 내용을 보여줍니다.

	Sub cmdNext_Click()
        Call CheckVorH
	    If spvwData.hWndSpread = spdData.hWnd Then
	        If spvwData.PageCurrent = spdData.PrintPageCount Then
	            If lgIntSpd > 0 Then
	                spvwData.hWndSpread = lgObjSpd(lgIntSelSpd).hWnd
	                spvwData.PageCurrent = 1
	                Call spvwData_PageChange(1)
	            End If
	        Else
	            spvwData.PageCurrent = spvwData.PageCurrent + 1
	        End If
	    Else
	        If spvwData.PageCurrent = lgObjSpd(lgIntSelSpd).PrintPageCount Then
	        Else
	            spvwData.PageCurrent = spvwData.PageCurrent + 1
	        End If
	    End If
	End Sub

	Sub cmdExit_Click()
	    window.self.close
	End Sub

	Sub spvwData_Zoom()
		If spvwData.PageViewType = 0 Then
			cboZoom.value = 6
		ElseIf spvwData.PageViewType = 1 Then
			cboZoom.value = 2
		End If
	End Sub

	Sub cboZoom_onChange()
	    Select Case cboZoom.value
	        Case 0 'per200
	            spvwData.PageViewType = 2 'PageViewTypePercentage
	            spvwData.PageViewPercentage = 200

	        Case 1 'per150
	            spvwData.PageViewType = 2 'PageViewTypePercentage
	            spvwData.PageViewPercentage = 150

	        Case 2 'per100
	            spvwData.PageViewType = 2 'PageViewTypePercentage
	            spvwData.PageViewPercentage = 100

	        Case 3 'per50
	            spvwData.PageViewType = 2 'PageViewTypePercentage
	            spvwData.PageViewPercentage = 50

	        Case 4
	            spvwData.PageViewType = 3 'PageViewTypePageWidth

	        Case 5
	            spvwData.PageViewType = 4 'PageViewTypePageHeight

	        Case 6
	            spvwData.PageViewType = 0 'PageViewTypeWholePage

	    End Select
	End Sub

	'Grid를 선택하는 콤보에서 다른 spread를 선택하는 경우 
	Sub cboGrid_onChange()
	    If cboGrid.value > 0 And lgIntSelSpd <> cboGrid.value Then
	        lgIntSelSpd = cboGrid.value
	        Call ResetPreview
	    End If
	End Sub

</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<BODY SCROLL=no>
<input type="hidden" name="nTotalPageCount">
	<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
		<TR>
			<TD WIDTH=100% HEIGHT=30px>
                <% 'by Shin hyoung jae 2001/2/28 %>
				<TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0 STYLE="BACKGROUND-COLOR:ButtonFace;">
					<TR>
						<TD WIDTH=100px ALIGN="Center"><INPUT type="Button" value="Print" name=btnPrint onclick="cmdPrint_Click" CLASS="NormalButton"></TD>
						<TD WIDTH=100px ALIGN="Center"><INPUT type="Button" value="Prev" name=btnPrev onclick="cmdPrev_Click" CLASS="NormalButton"></TD>
						<TD WIDTH=100px ALIGN="Center"><INPUT type="Button" value="Next" name=btnNext onclick="cmdNext_Click" CLASS="NormalButton"></TD>
						<TD WIDTH=100px ALIGN="Center">
                            <SELECT name=cboZoom Style="width=80">
  						    <OPTION Value="0">200%</OPTION>
                            <OPTION Value="1">150%</OPTION>
                            <OPTION Value="2">100%</OPTION>
                            <OPTION Value="3">50%</OPTION>
                            <OPTION Value="4">Width</OPTION>
                            <OPTION Value="5">Height</OPTION>
                            <OPTION Value="6">Whole</OPTION>
                            </SELECT>
						</TD>
						<TD WIDTH=200px ALIGN="Center">
                        <label for="V"><input id="V" OnClick="optVorH_Click" class="radio" type="radio" name="optVorH" value="V">세로</label> &nbsp;&nbsp;&nbsp;
                        <label for="H"><input id="H" OnClick="optVorH_Click" class="radio" type="radio" name="optVorH" value="H" checked>가로</label>
                        </TD>
						<TD WIDTH=100px ALIGN="Center"><SELECT name=cboGrid Style="width=80" disabled></SELECT></TD>
						<TD WIDTH=100px ALIGN="Center"><INPUT type="Button" value="Exit" name=btnExit onclick="cmdExit_Click" CLASS="NormalButton"></TD>
						<TD WIDTH=*>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=*>
			    <script language =javascript src='./js/print_spvwData_N764971220.js'></script>			
				<script language =javascript src='./js/print_tempData_tempData.js'></script>			
				<script language =javascript src='./js/print_I430784767_spdData.js'></script>
			</TD>
		</TR>
	</TABLE>
</BODY>
</HTML>
