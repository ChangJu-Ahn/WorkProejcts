<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : �ǸŰ�ȹ���� 
'*  3. Program ID           : S2214QB2
'*  4. Program Name         : �ǸŰ�ȹ�������ȸ(ǰ��׷�)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/16
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Yong Sik
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                         '�� : DBAgent Parameter ���� 
    Dim lgstrData                                                              '�� : data for spreadsheet data
    Dim lgStrPrevKey                                                           '�� : ���� �� 
    Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgDataExist
    Dim lgPageNo
    Dim lgStrColorFlag, lgStrDisplayType, lgStrGrpNm
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgPageNo		= ""
    lgSelectList   = Request("txtSelectLIst")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtTailList")                                 '�� : Orderby value
	lgStrDisplayType = Request("cboDisplayType")

    lgDataExist    = "No"

    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 100     
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    iLoopCount = 0
    lgStrColorFlag = ""
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
		If rs0(0) > 0 Then
			lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & rs0(0) & gRowSep
		End If
		
		lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        rs0.MoveNext
	Loop

	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim iStrFromDt, iStrToDt, iStrLocExpFlag, iStrItemGroupCd, iStrBaseCur, iStrCur
	Dim iIntGrpLvl
	
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(2,15)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	iStrFromDt		= UNIGetFirstDay(Request("txtConFromDt"),gDateFormatYYYYMM)
	iStrToDt		= UNIGetLastDay(Request("txtConToDt"), gDateFormatYYYYMM)
	iStrLocExpFlag	= Request("cboConLocExpFlag")
	iIntGrpLvl		= Request("cboConGrpLvl")
	iStrItemGroupCd = Trim(Request("txtConItemGroupCd"))
	iStrBaseCur		= Request("cboConBaseCur")
	iStrCur			= Trim(Request("txtConCur"))
	
	If lgStrDisplayType = "H" Then
	    UNISqlId(0) = "S2214QA201"
	    lgSelectList = Replace(lgSelectList, "1?", "�Ѱ�")
	    lgSelectList = Replace(lgSelectList, "2?", "�׷�Ұ�")
	Else
	    UNISqlId(0) = "S2214QA202"
		lgSelectList = Replace(lgSelectList, "1?", "�Ѱ�")
		lgSelectList = Replace(lgSelectList, "2?", "��Ұ�")
		lgSelectList = Replace(lgSelectList, "3?", "���Ұ�")
		lgSelectList = Replace(lgSelectList, "4?", "�׷�Ұ�")
	End If

	UNIValue(0,0) = lgSelectList

	UNIValue(0,1) = " " & FilterVar(UNIConvDate(iStrFromDt), "''", "S") & ""			' ������ 
	UNIValue(0,2) = " " & FilterVar(UNIConvDate(iStrToDt), "''", "S") & ""			' ������ 
	UNIValue(0,3) = "" & FilterVar("IG", "''", "S") & ""										' ǰ��׷캰 ��ȸ 
	UNIValue(0,4) = "NULL"										' ������������ 
	UNIValue(0,5) = "NULL"										' ǰ��׷� 
	UNIValue(0,6) = "NULL"										' ��뿩�� 
	
	UNIValue(0,7)  = iIntGrpLvl									' ǰ�񷹺� 
	If iStrItemGroupCd = "" Then								' ǰ��׷� 
		UNIValue(0,8) = "NULL"
	Else
		UNIValue(0,8) = " " & FilterVar(iStrItemGroupCd, "''", "S") & ""
		UNISqlId(1) = "I224QA1A5"								' ǰ��׷�� Fetch
		UNIValue(1,0) = UNIValue(0,8) & " AND ITEM_GROUP_LEVEL = " & iIntGrpLvl
	End If
	
	UNIValue(0,9) = "" & FilterVar("%", "''", "S") & ""										' ��뿩�� 

	UNIValue(0,10) = " " & FilterVar(Request("cboConSpType"), "''", "S") & ""		' �ǸŰ�ȹ���� 

	If iStrLocExpFlag = "" Then									' ����/���⿩�� 
		UNIValue(0,11) = "" & FilterVar("%", "''", "S") & ""
	Else
		UNIValue(0,11) = " " & FilterVar(iStrLocExpFlag, "''", "S") & ""
	End If

	UNIValue(0,12) = " " & FilterVar(iStrBaseCur, "''", "S") & ""					' ȭ����� 
	
	If iStrCur <> "" Then										' ȭ����� 
		UNIValue(0,13) = " " & FilterVar(iStrCur, "''", "S") & ""
		If iStrBaseCur = "D" Then
			If lgStrDisplayType = "H" Then
				UNIValue(0,14) = "AND GROUPING_FLAG <> 1 "		' �׷캰 �Ұ�� ���ܽ�Ŵ 
			Else
				UNIValue(0,14) = "WHERE GROUPING_FLAG <> 1 "
			End If
		Else
			UNIValue(0,14) = ""
		End If
		
		' ȭ�����翩�� Check
		UNISqlId(2) = "s0000qa014"
		UNIValue(2,0) = FilterVar(iStrCur, "''", "S")
	Else
		UNIValue(0,13) = "" & FilterVar("%", "''", "S") & ""
		UNIValue(0,14) = ""
	End If
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If

	' Write the Script Tag(<Script language=vbscript>)
	Call BeginScriptTag()

    If  UNIValue(1,0) <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing
			Call ConNotFound("txtConItemGroupCd")
			Exit Sub
		Else
			Call WriteConDesc("txtConItemGroupNm", Rs1(0))
		End If
	Else
		Call WriteConDesc("txtConItemGroupNm", "")
    End If
    
    ' ȭ�� ���翩��    
    If  UNIValue(2,0) <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing
			Call ConNotFound("txtConCur")
			Exit Sub
		End If
    End If

    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("cboConSpType")
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
        Call WriteResult()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Write the Result
' ���Html�� �ۼ��Ѵ�.
'----------------------------------------------------------------------------------------------------------
Sub BeginScriptTag()
	Response.Write "<Script language=VBScript> " & VbCr
End Sub

Sub EndScriptTag()
	Response.Write "</Script> " & VbCr
End Sub

' �����Ͱ� �������� �ʴ� ��� ó�� Script �ۼ�(��ȸ���� ����)
Sub ConNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1." & pvStrField & ".alt, ""X"") " & VbCr
	Response.Write "Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ���ǿ� �ش��ϴ� ���� Display�ϴ� Script �ۼ� 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write "Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' �����Ͱ� �������� �ʴ� ��� ó�� Script �ۼ� 
Sub DataNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write "Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ ����� Display�ϴ� Script �ۼ� 
Sub WriteResult()
	If lgStrDisplayType = "H" Then
		Response.Write "Parent.ggoSpread.Source  = Parent.frm1.vspdData " & VbCr
		Response.Write  "Parent.frm1.vspdData.Redraw = False  "      & vbCr      	
		Response.Write  "Parent.ggoSpread.SSShowDataByClip   """ & lgstrData & """ ,""F""" & vbCr
		Response.Write "parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
		Response.Write "Parent.DbQueryOk " & VbCr	
		Response.Write "Parent.frm1.vspdData.Redraw = True  "       & vbCr      
	Else
		Response.Write "Parent.ggoSpread.Source  = Parent.frm1.vspdData2 " & VbCr
		Response.Write  "Parent.frm1.vspdData2.Redraw = False  "      & vbCr      	
		Response.Write  "Parent.ggoSpread.SSShowDataByClip   """ & lgstrData & """ ,""F""" & vbCr
		Response.Write "parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
		Response.Write "Parent.DbQueryOk " & VbCr
		Response.Write "Parent.frm1.vspdData2.Redraw = True  "       & vbCr      
	End If	
	Call EndScriptTag()
End Sub
%>
