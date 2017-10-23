<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : �ǸŰ�ȹ���� 
'*  3. Program ID           : S2214QB3
'*  4. Program Name         : �ǸŰ�ȹ�������ȸ(��������/ǰ��׷�)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/15
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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, Rs1, Rs2, Rs3               '�� : DBAgent Parameter ���� 
    Dim lgstrData                                                              '�� : data for spreadsheet data
    Dim lgStrPrevKey                                                           '�� : ���� �� 
    Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgDataExist
    Dim lgPageNo
    Dim lgStrColorFlag, lgStrDisplayType
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    'lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgPageNo		= ""
    'lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
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
	Dim iStrFromDt, iStrToDt, iStrLocExpFlag, iStrBaseCur, iStrCur
	Dim iStrSalesOrg, iStrGrpFlag, iStrItemGroupCd
	Dim iIntOrgLvl, iIntGrpLvl
	
    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(3,15)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	iStrFromDt		= UNIGetFirstDay(Request("txtConFromDt"),gDateFormatYYYYMM)
	iStrToDt		= UNIGetLastDay(Request("txtConToDt"), gDateFormatYYYYMM)
	iStrLocExpFlag	= Request("cboConLocExpFlag")
	iIntOrgLvl		= Request("cboConOrgLvl")
	iIntGrpLvl		= Request("cboConGrpLvl")
	iStrItemGroupCd = Request("txtConItemGroupCd")
	iStrGrpFlag		= Request("cboConGrpFlag")
	
	If iStrGrpFlag = "G" Then
		iIntOrgLvl		= 0
		iStrSalesOrg	= Request("txtConSalesGrp")
	Else
		iIntOrgLvl		= Request("cboConOrgLvl")
		iStrSalesOrg	= Request("txtConSalesOrg")
	End If
	
	iStrBaseCur		= Request("cboConBaseCur")
	iStrCur			= Request("txtConCur")
	
	If lgStrDisplayType = "H" Then
	    lgSelectList = Replace(lgSelectList, "1?", "�Ѱ�")
	    If iStrGrpFlag = "G" Then
		    UNISqlId(0) = "S2214QA301"
			lgSelectList = Replace(lgSelectList, "2?", "�׷�Ұ�")
			lgSelectList = Replace(lgSelectList, "SALES_ORG_NM", "SALES_GRP_NM")
		Else
		    UNISqlId(0) = "S2214QA302"
			lgSelectList = Replace(lgSelectList, "2?", "�����Ұ�")
		End If
		lgSelectList = Replace(lgSelectList, "3?", "ǰ��׷�Ұ�")
	Else
		lgSelectList = Replace(lgSelectList, "1?", "�Ѱ�")
		lgSelectList = Replace(lgSelectList, "2?", "��Ұ�")
		lgSelectList = Replace(lgSelectList, "3?", "���Ұ�")
	    If iStrGrpFlag = "G" Then
		    UNISqlId(0) = "S2214QA303"
			lgSelectList = Replace(lgSelectList, "4?", "�׷�Ұ�")
			lgSelectList = Replace(lgSelectList, "SALES_ORG_NM", "SALES_GRP_NM")
		Else
		    UNISqlId(0) = "S2214QA304"
			lgSelectList = Replace(lgSelectList, "4?", "�����Ұ�")
		End If
		lgSelectList = Replace(lgSelectList, "5?", "ǰ��׷�Ұ�")
	End If

	UNIValue(0,0) = lgSelectList

	UNIValue(0,1) = " " & FilterVar(UNIConvDate(iStrFromDt), "''", "S") & ""			' ������ 
	UNIValue(0,2) = " " & FilterVar(UNIConvDate(iStrToDt), "''", "S") & ""			' ������ 


	If iStrGrpFlag = "G" Then
		UNIValue(0,3) = "" & FilterVar("GG", "''", "S") & ""									' �����׷캰��ȸ 
		UNIValue(0,4) = "NULL"									' ������������ 

		If iStrSalesOrg = "" Then
			UNIValue(0,5) = "NULL"
		Else
			UNIValue(0,5) = " " & FilterVar(iStrSalesOrg, "''", "S") & ""
			
			UNISqlId(1) = "B1254MA802"								' �����׷�� Fetch
			UNIValue(1,0) = UNIValue(0,5)
		End If
	Else
		UNIValue(0,3) = "" & FilterVar("OG", "''", "S") & ""									' ���������� ��ȸ 
		UNIValue(0,4) = iIntOrgLvl								' ������������ 

		If iStrSalesOrg = "" Then
			UNIValue(0,5) = "NULL"
		Else
			UNIValue(0,5) = " " & FilterVar(iStrSalesOrg, "''", "S") & ""
			
			UNISqlId(1) = "B1254MA803"								' ���������� Fetch
			UNIValue(1,0) = UNIValue(0,5) & " AND LVL = " & iIntOrgLvl
		End If
	End If

	UNIValue(0,6) = "" & FilterVar("%", "''", "S") & ""										' ��뿩�� 
	
	UNIValue(0,7)  = iIntGrpLvl									' ǰ�񷹺� 
	If iStrItemGroupCd = "" Then								' ǰ��׷� 
		UNIValue(0,8) = "NULL"
	Else
		UNIValue(0,8) = " " & FilterVar(iStrItemGroupCd, "''", "S") & ""
		UNISqlId(2) = "I224QA1A5"								' ǰ��׷�� Fetch
		UNIValue(2,0) = UNIValue(0,8) & " AND ITEM_GROUP_LEVEL = " & iIntGrpLvl
	End If
	
	UNIValue(0,9) = "" & FilterVar("%", "''", "S") & ""										' ��뿩�� 

	UNIValue(0,10) = " " & FilterVar(Request("cboConSpType"), "''", "S") & ""		' �ǸŰ�ȹ���� 

	If iStrLocExpFlag = "" Then									' ����/���⿩�� 
		UNIValue(0,11) = "" & FilterVar("%", "''", "S") & ""
	Else
		UNIValue(0,11) = " " & FilterVar(iStrLocExpFlag, "''", "S") & ""
	End If

	UNIValue(0,12) = " " & FilterVar(iStrBaseCur, "''", "S") & ""						' ȭ����� 
	
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
		UNISqlId(3) = "s0000qa014"
		UNIValue(3,0) = FilterVar(iStrCur, "''", "S")
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, Rs1, Rs2, Rs3)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If

	' Write the Script Tag(<Script language=vbscript>)
	Call BeginScriptTag()
	
	' �����׷�/���� ���翩�� 
    If  UNIValue(1,0) <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing
			If UNIValue(0,3) = "" & FilterVar("GG", "''", "S") & "" Then
				Call ConNotFound("txtConSalesGrp")
			Else
				Call ConNotFound("txtConSalesOrg")
			End If
			Exit Sub
		Else
			If UNIValue(0,3) = "" & FilterVar("GG", "''", "S") & "" Then
				Call WriteConDesc("txtConSalesGrpNm", Rs1(1))
			Else
				Call WriteConDesc("txtConSalesOrgNm", Rs1(1))
			End If
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")
		Call WriteConDesc("txtConSalesOrgNm", "")
    End If
    
    ' ǰ��׷� ���翩�� Check
    If  UNIValue(2,0) <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing
			Call ConNotFound("txtConItemGroupCd")
			Exit Sub
		Else
			Call WriteConDesc("txtConItemGroupNm", Rs2(0))
		End If
	Else
		Call WriteConDesc("txtConItemGroupNm", "")
    End If

    ' ȭ�� ���翩��    
    If  UNIValue(3,0) <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing
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
        Call MakeSpreadSheetData()
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
