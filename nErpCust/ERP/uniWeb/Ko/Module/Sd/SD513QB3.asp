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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4 '�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT        
    Dim lgStrColorFlag
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

    Dim lgBizAreaCd, lgSalesGrpCd, lgPayerCd, lgBpLevelTypeCd	
    
    lgBizAreaCd			= Trim(Request("BizAreaCd"))
    lgSalesGrpCd		= Trim(Request("SalesGrpCd"))
    lgPayerCd		= Trim(Request("PayerCd"))
    lgBpLevelTypeCd		= Trim(Request("BpLevelTypeCd"))
    
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd
    
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

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
    Dim	iInt
    
	lgstrData      = ""
    
    iLoopCount = 0
    lgStrColorFlag = ""
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			' �����Ͽ� ���� ��¥ Froamt ���� 
			' ����ó �Ұ踦 Display�ϱ� ���� Text �������� ����, ���� ��¥ �������� Conversion �ʿ� 
			If ColCnt = 5 And rs0(0) = 0 Then
				iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
			Else
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			End If
		Next
		
		If rs0(0) > 0 Then	'����Row ���� üũ 
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

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim iStrVal    
    
    Redim UNISqlId(4)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(4,1)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
    iStrVal = ""	

	'������=============================================================================================    	
	If Len(lgBizAreaCd) Then
		UNISqlId(1)		= "s0000qa013"	
		UNIValue(1,0)	= FilterVar(lgBizAreaCd , "''", "S")
		
		iStrVal	= iStrVal & " AND AR.BIZ_AREA_CD =  " & FilterVar(lgBizAreaCd , "''", "S") & ""				
	End If

    '����ó=============================================================================================
    If Len(lgPayerCd) Then		
		UNISqlId(2)		= "s0000qa002"	
		UNIValue(2,0)	= FilterVar(lgPayerCd , "''", "S")
		
		iStrVal	= iStrVal & " AND AR.PAY_BP_CD =  " & FilterVar(lgPayerCd , "''", "S") & ""				
	End If	

	'�����׷��===========================================================================================	
    If Len(lgSalesGrpCd) Then
		UNISqlId(3)		= "s0000qa005"	
		UNIValue(3,0)	= FilterVar(lgSalesGrpCd , "''", "S")
		
		iStrVal	= iStrVal & " AND AR.SALES_GRP =  " & FilterVar(lgSalesGrpCd , "''", "S") & ""				
	End If		
	
	'�ŷ�ó���������=====================================================================================	
	If Len(lgBpLevelTypeCd) Then		
		UNISqlId(4)		= "s0000qa000"	
		UNIValue(4,0)	= FilterVar("B9010" , "''", "S")
		UNIValue(4,1)	= FilterVar(lgBpLevelTypeCd , "''", "S")
		
		iStrVal	= iStrVal & " AND BP.BP_GRADE =  " & FilterVar(lgBpLevelTypeCd , "''", "S") & ""				
	End If	
	
	UNISqlId(0) = "SD513QA301"					
    UNIValue(0,0) = lgSelectList
    UNIValue(0,1) = iStrVal
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
'    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	
	'����� ���翩�� 
	If lgBizAreaCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call WriteConDesc("txtConBizAreaNm", "")		
			Call ConNotFound("txtConBizAreaCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBizAreaNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConBizAreaNm", "")		
	End If

	'����ó ���翩�� 
	If lgPayerCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call WriteConDesc("txtConPayerNm", "")		
			Call ConNotFound("txtConPayerCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConPayerNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConPayerNm", "")
	End If

	'�����׷� ���翩�� 
	If lgSalesGrpCd <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing			
			Call WriteConDesc("txtConSalesGrpNm", "")		
			Call ConNotFound("txtConSalesGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm", rs3(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")
	End If

	'�ŷ�ó������� ���翩�� 
	If lgBpLevelTypeCd <> "" Then
		If rs4.EOF And rs4.BOF Then
			rs4.Close
			Set rs4 = Nothing			
			Call ConNotFound("txtConBpLevelTypeCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBpLevelTypeNm", rs4(1))		
		End If
	Else
		Call WriteConDesc("txtConBpLevelTypeNm", "")
	End If
		 
    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConBizAreaCd")	
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
	Response.Write " parent.lgStrColorFlag = """ & lgStrColorFlag & """" & vbCr	
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


