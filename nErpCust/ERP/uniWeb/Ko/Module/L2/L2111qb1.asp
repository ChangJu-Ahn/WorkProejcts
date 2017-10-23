<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
    Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1 '�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgPageNo
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

    Dim iPrevEndRow
    
    Dim lgFromDt			'��ȸ�Ⱓ���� 
    Dim lgToDt				'��ȸ�Ⱓ�� 
    Dim lgSoldToParty		'�ֹ�ó 
    Dim lgPoNo				'�ֹ���ȣ 
    Dim lgRcptFromDt		'������ 
    Dim lgRcptToDt			'������ 
    
    lgFromDt		= Trim(Request("txtHConFromDt"))    
    lgToDt			= Trim(Request("txtHConToDt"))
    lgSoldToParty	= Replace(Trim(Request("txtHConSoldToPartyCd")),"'","''")
    lgPoNo			= Replace(Trim(Request("txtHConPoNo")),"'","''")
    lgRcptFromDt	= Trim(Request("txtHConRcptFromDt"))    
    lgRcptToDt		= Trim(Request("txtHConRcptToDt"))
    
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("txtHlgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtHlgTailList")                                 '�� : Orderby value
    iPrevEndRow = 0

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
    
    Const C_SHEETMAXROWS_D = 20     

    lgstrData      = ""
    iPrevEndRow = 0
    
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow    

    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim iStrVal    
    
    Redim UNISqlId(1)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(1,2)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
	' ������(����)
	iStrVal = iStrVal & " WHERE ISH.DOC_ISSUE_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
	
	' �ֹ���(��)
	If Len(lgToDt) Then
		iStrVal = iStrVal & " AND ISH.DOC_ISSUE_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
	End If
	
    '�ֹ�ó 
    If Len(lgSoldToParty) Then		
		UNISqlId(1)		= "s0000qa002"	
		UNIValue(1,0)	= lgSoldToParty
		iStrVal = iStrVal & " AND ISH.SOLD_TO_PARTY =  " & FilterVar(lgSoldToParty , "''", "S") & ""				
	End If	

    '�ֹ���ȣ 
    If Len(lgPoNo) Then		
		iStrVal = iStrVal & " AND ISH.DOC_NO =  " & FilterVar(lgPoNo , "''", "S") & ""				
	End If	

	' ������ 
	If Len(lgRcptFromDt) Then
		iStrVal = iStrVal & " AND ISH.RCPT_DT >=  " & FilterVar(UNIConvDate(lgRcptFromDt), "''", "S") & ""		
	End If

	' ������(��)
	If Len(lgRcptToDt) Then
		iStrVal = iStrVal & " AND ISH.RCPT_DT <=  " & FilterVar(UNIConvDate(lgRcptToDt), "''", "S") & ""		
	End If

	lgSelectList = Replace(lgSelectList, "?", _
					"CASE WHEN ISD.STS = 0 THEN " & FilterVar("�ֹ�����", "''", "S") & " " & _
						" WHEN ISD.STS = 1 THEN " & FilterVar("�ֹ�����", "''", "S") & " " & _
						" WHEN ISD.STS = 2 THEN " & FilterVar("���ֵ��", "''", "S") & " " & _
						" ELSE CASE " & _
								"	WHEN SD.CLOSE_FLAG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("���ָ���", "''", "S") & " " & _
								"	WHEN SD.REQ_QTY = 0 THEN '����Ȯ��' " & _
								"	WHEN SD.BILL_QTY > 0 AND SD.SO_QTY <= SD.BILL_QTY THEN " & FilterVar("����Ϸ�", "''", "S") & " " & _
								"	WHEN SD.BILL_QTY > 0 AND SD.SO_QTY > SD.BILL_QTY THEN '��������' " & _
								"	WHEN SD.GI_QTY > 0 AND SD.SO_QTY <= SD.GI_QTY THEN " & FilterVar("���Ϸ�", "''", "S") & " " & _
								"	WHEN SD.GI_QTY > 0 AND SD.SO_QTY > SD.GI_QTY THEN '�������' " & _
								"	WHEN SD.SO_QTY <= SD.REQ_QTY THEN " & FilterVar("����û�Ϸ�", "''", "S") & " " & _
								"	ELSE '����û����' " & _
							  " END  " & _
					    " END ")

	UNISqlId(0) = "L2111QA101"					
    UNIValue(0,0) = lgSelectList                                      
	UNIValue(0,1) = iStrVal	         
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '��: set ADO read mode
 
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                     '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                           '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	
	'�ֹ�ó ���翩�� 
	If lgSoldToParty <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConSoldToPartyCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSoldToPartyNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConSoldToPartyNm", "")
	End If
	
    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConFromDt")	
        Exit Sub
    Else    
        Call MakeSpreadSheetData()
        Call WriteResult()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Write the Result
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
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ���ǿ� �ش��ϴ� ���� Display�ϴ� Script �ۼ� 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write " Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' �����Ͱ� �������� �ʴ� ��� ó�� Script �ۼ� 
Sub DataNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ ����� Display�ϴ� Script �ۼ� 
Sub WriteResult()
	Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData " & vbCr
	Response.Write " Parent.frm1.vspdData.Redraw = False " & vbCr      	
	Response.Write " Parent.ggoSpread.SSShowData  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " Parent.lgPageNo = """ & lgPageNo & """" & vbCr
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


