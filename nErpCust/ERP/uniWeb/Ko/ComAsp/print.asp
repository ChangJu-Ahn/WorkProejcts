<HTML>
<HEAD><TITLE>Print Preview </TITLE>
<% '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################%>
<% '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* %>
<!-- #Include file="../inc/IncServer.asp" -->
<%'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">

<%'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================%>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
	Option Explicit

	'Title�� ù �������� ���� ���Դϴ�.
	'ǥ���� ���� �ǹ̷� ū Font size�� ����մϴ�.
	Const FONT_TITLE1 = "/c/fn""MS UI Gothic""/fz""25""/fb1/fi0/fu1/fk0/fs1"
	Const FONT_TITLE2 = "/c/fn""MS UI Gothic""/fz""20""/fb1/fi0/fu1/fk0/fs1"
	Const FONT_TITLE3 = "/fn""MS UI Gothic""/fz""12""/fb0/fu0/fs0"

	'Header�� �ι�° ������ ���Ŀ� 
	'���� spread������ header�κ��Դϴ�.
	Const FONT_HEADER1 = "/c/fn""MS UI Gothic""/fz""25""/fb1/fi0/fu1/fk0/fs1"
	Const FONT_HEADER2 = "/c/fn""MS UI Gothic""/fz""20""/fb1/fi0/fu1/fk0/fs1"
	Const FONT_HEADER3 = "/fn""MS UI Gothic""/fz""12""/fb0/fu0/fs0"

	'DD�� ���� ó���ؾ� �մϴ�.
	Const TAG_USER = "����� :"
	Const TAG_DATE = "�μ����� :"
	Const TAG_GRID = "�׸���"

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

'�⺻ logic�� Excel�� �����մϴ�.
'�� spread�� ���� �۾����� �ʰ� preview control�� ���Ḹ �� �ݴϴ�.
'
'spread �̿��� �κ��� spdData ��� �Ⱥ��̴� �������忡 �ֽ��ϴ�.
'ù �������� spdData�� ������ �����ְ� 
'���� ���������ʹ� combo�ڽ����� ���õ� spread�� �����ݴϴ�.

	Sub ShowPrint(objDoc)
	    Dim strDate
	    Dim strTitle
	    Dim strTabName
	    Dim objOption
	    Dim i
	    '0:���ϳ�    0:��2�̻�     2:��2�̻� ��������    4:��2�̻� ���Ǳ׸��� b1b11ma2    8:��2���̻� ��������� a3111ma1
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

	    If i <= 1 Then 'Tab�� �ϳ��� ��� 

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

	    Else 'Tab�� �ϳ��̻� 

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

			' ���� ������ �ݵ�� filedset�� �׿��� �־�� �Ѵ�. DIV�� ������ ���� 
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

			' ���� ���� Spread Sheet. DIV�� ���������� 
			' b1b11ma2 ���� ������ ȭ�鶫�� 
			' ���� �׸��尡 �������, div �ڿ� ����. h6014ma1
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

			' ���� �̱� detail a3111ma1 ���� 
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

	    '���� Field�� Single Data�� ������ ù�������� �����.
		'  by Shin hyoung jae, 2001/4/3
		Call SetPrvwDataSheet
		'0:���ϳ�    0:��2�̻�     2:��2�̻� ��������    4:��2�̻� ���Ǳ׸��� b1b11ma2    8:��2���̻� ��������� a3111ma1
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

	        '�����������ŭ �޺� ������ �߰� 
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

	'ù�������� ��� 
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

	'�ι�° ������ ������ ��� 
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

	'Grid Combo�� �ٲ� ���õ� spread�� preview control�� �����ϴ� �Լ��Դϴ�.
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
			'����̵� ȭ�� ���� ����.
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
				tempdata.PrintMarginTop = 600 '���� 
	        else
				tempdata.PrintMarginTop = 900 '���� 
	        end if
	        
	        tempdata.PrintMarginBottom = spdData.PrintMarginBottom
	    End If
	End Sub


	'document�� �˻��Ͽ� spread�� ã�� lgObjSpd�� �����մϴ�.
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

	'���ǰ� Single �����ͷ� ������ ù�������� �����.
	'logic�� Excel�� ���� �����ϸ� spread���� �����κи��� ����.
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
					' q1411ma1 ���� ���ǿ��� ���ÿ� ���� ������ TR �� display�� none �̵�.
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
						' by Shin hyoung jae, 2001/4/2 add TD18 TD19 �߰� 
						If UCase(objEl.All(i).className) = "TD5" Or UCase(objEl.All(i).className) = "TD18" Or UCase(objEl.All(i).className) = "TDT" Then  ' �󺧸� 
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
						' ����Ʈ�� ��ġ �߸������°� ���� 

							If bIsSameLine = False Then
								bIsSameLine = True
							Else
								bIsSameLine = False
							End If

							' ���°� �ݵ�� <TD CLASS=TD6></TD>  ���̰� �پ�� �Ѵ�. �ݵ�� 
							If Trim(objEl.All(i).innerText) = "" Then
								bIsSameLine = False
							Else
                                ' by Shin hyoung jae, 2002/05/30 c2010ba1 ���� ���� ����.
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
						' text ~ text �̷��� ������ p1401ma3
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
						' OCX�ٷε��� ��, day �̷��� ������ 
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
							' OCX�ٷε��� ��, day �̷��� ������ 
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

	'�� object�� Text���̿� �°� �÷��� ���� ���� 
	Sub SetText(strText)
	    spdData.TypeMaxEditLen = 200
	    spdData.Text = strText

	    If spdData.ColWidth(spdData.Col) < Len(spdData.Text) + 3 And Len(strText) < 100 Then
	        spdData.ColWidth(spdData.Col) = Len(spdData.Text) + 3
	    End If
	End Sub

'//////// Event ó�� 
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
		spdData.FontName = "����ü"

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

	'page�� ����� ��� Prev, Next ��ư Enable/Disable

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
				nConSP = arrRet(1)   ' �μ� ���� ������ 
			Else
				nConSP = 0
				nConEP = 0
			End If

			If CInt(spdData.PrintPageCount) >= CInt(arrRet(2)) Then
				nConEP = arrRet(2) '�μ� �������� 
				nSpdSP = 0
				nSpdEP = 0
			Else
				If nConSP <> 0 Then
					nConEP = spdData.PrintPageCount
					nSpdSP = 1
					'nSpdEP = arrRet(2) - spdData.PrintPageCount
					nSpdEP = arrRet(2)   '2003-01-28 ����� ���� 
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
					tempData.ColHeaderDisplay = 0  '2003-01-28 ����� ���� 
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

	'ù�������� ���� ������ spdData�� ������ 
	'�ι�° ������ ������ ���� ���� spread�� ������ �����ݴϴ�.
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

	'ù�������� ���� ������ spdData�� ������ 
	'�ι�° ������ ������ ���� ���� spread�� ������ �����ݴϴ�.

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

	'Grid�� �����ϴ� �޺����� �ٸ� spread�� �����ϴ� ��� 
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
                        <label for="V"><input id="V" OnClick="optVorH_Click" class="radio" type="radio" name="optVorH" value="V">����</label> &nbsp;&nbsp;&nbsp;
                        <label for="H"><input id="H" OnClick="optVorH_Click" class="radio" type="radio" name="optVorH" value="H" checked>����</label>
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
