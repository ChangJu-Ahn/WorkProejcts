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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2					'�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT        
    Dim lgStrColorFlag
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

	Dim lgConFromDt, lgConToDt, lgConSalesGrpCd, lgConBizAreaCd, lgConPostFlag, lgConExceptFlag

	lgConFromDt		= uniConvDate(Trim(Request("ConFromDt")))
	lgConToDt		= uniConvDate(Trim(Request("ConToDt")))
	lgConSalesGrpCd	= Trim(Request("SalesGrpCd"))
    lgConBizAreaCd		= Trim(Request("BizAreaCd"))
	lgConPostFlag	= Trim(Request("PostFlag"))
	lgConExceptFlag	= Trim(Request("ExceptFlag"))
	
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
    Dim	 iBlnDisplayText
    
    lgstrData      = ""
    
    iLoopCount = 0
    iBlnDisplayText = False
    lgStrColorFlag = ""
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		If rs0(0) = 0 Or iBlnDisplayText Then
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(0),rs0(0))
		    iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(1),rs0(1))
		Else
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(0),rs0(0))
		    iRowStr = iRowStr & Chr(11) & "���"
		    iBlnDisplayText = True
		End If
		
		For ColCnt = 2 To UBound(lgSelectListDT) - 1
		    iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
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
    
    Redim UNISqlId(2)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(2,1)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 

	iStrVal = " SELECT BH.SALES_GRP, " & _
					 " YEAR(BH.BILL_DT) YR, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 1 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_JAN, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 2 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_FEB, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 3 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_MAR, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 4 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_APR, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 5 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_MAY, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 6 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_JUN, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 7 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_JUL, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 8 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_AUG, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 9 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_SEP, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 10 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_OCT, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 11 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_NOV, " & _
					 " SUM(CASE WHEN MONTH(BH.BILL_DT) = 12 THEN BH.BILL_AMT_LOC + BH.VAT_AMT_LOC ELSE 0 END) AMT_DEC " & _
			  " FROM   S_BILL_HDR BH " & _
			  " WHERE BH.BILL_DT >=  " & FilterVar(lgConFromDt , "''", "S") & "" & _
			  " AND   BH.BILL_DT <=  " & FilterVar(lgConToDt , "''", "S") & ""			' ������ 

 	'�����׷��===========================================================================================	
    If Len(lgConSalesGrpCd) Then
		UNISqlId(1)		= "s0000qa005"	
		UNIValue(1,0)	= FilterVar(lgConSalesGrpCd, "''", "S")

		iStrVal = iStrVal & " AND BH.SALES_GRP =  " & FilterVar(lgConSalesGrpCd , "''", "S") & ""
	End If		

	'������=============================================================================================    	
	If Len(lgConBizAreaCd) Then
		UNISqlId(2)		= "s0000qa013"	
		UNIValue(2,0)	= FilterVar(lgConBizAreaCd, "''", "S")

		iStrVal = iStrVal & " AND BH.BIZ_AREA =  " & FilterVar(lgConBizAreaCd , "''", "S") & ""
	End If

	If lgConPostFlag <> "" Then								' Ȯ������ 
		iStrVal = iStrVal  & "AND BH.POST_FLAG =  " & FilterVar(lgConPostFlag , "''", "S") & ""
	End If

	If lgConExceptFlag <> "" Then
		iStrVal = iStrVal  & "AND BH.EXCEPT_FLAG =  " & FilterVar(lgConExceptFlag , "''", "S") & ""	'���ܿ��� 
	End If
	' Group by
	iStrVal = iStrVal  & " GROUP BY BH.SALES_GRP, YEAR(BH.BILL_DT) "

	UNISqlId(0) = "SD512QA401"					
    UNIValue(0,0) = Replace(lgSelectList, "?1", lgConFromDt)
	UNIValue(0,1)	= iStrVal
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	
	'�����׷� ���翩�� 
	If lgConSalesGrpCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConSalesGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")
	End If
	
	'����� ���翩�� 
	If lgConBizAreaCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConBizAreaCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBizAreaNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConBizAreaNm", "")		
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
	Response.Write " Call parent.SetFocusToDocument(""M"") " & vbCr	
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


