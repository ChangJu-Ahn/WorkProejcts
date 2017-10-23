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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9  '�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT        
    Dim lgStrColorFlag
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

	Dim lgConFromDt, lgConToDt, lgConSalesOrgCd, lgConPostFlag, lgConExceptFlag

	lgConFromDt		= uniConvDate(Trim(Request("ConFromDt")))
	lgConToDt		= uniConvDate(Trim(Request("ConToDt")))
	lgConSalesOrgCd	= Trim(Request("SalesOrgCd"))
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
    
    Const C_SHEETMAXROWS_D = 20     

    lgstrData      = ""
    
    iLoopCount = 0
    lgStrColorFlag = ""
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
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
    
    Redim UNISqlId(1)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(1,1)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 

	iStrVal = "WHERE BH.BILL_DT >=  " & FilterVar(lgConFromDt , "''", "S") & ""			' ������ 
	
	If lgConToDt <> "" Then
		iStrVal = iStrVal  & "AND BH.BILL_DT <=  " & FilterVar(lgConToDt , "''", "S") & ""			' ������ 
	End If
	
	If lgConPostFlag <> "" Then								' Ȯ������ 
		iStrVal = iStrVal  & "AND BH.POST_FLAG =  " & FilterVar(lgConPostFlag , "''", "S") & ""
	End If

	If lgConExceptFlag <> "" Then
		iStrVal = iStrVal  & "AND BH.EXCEPT_FLAG =  " & FilterVar(lgConExceptFlag , "''", "S") & ""	'���ܿ��� 
	End If

	'����������===========================================================================================
    If lgConSalesOrgCd <> "" Then		
		UNISqlId(1)		= "s0000qa006"	
		UNIValue(1,0)	= FilterVar(lgConSalesOrgCd, "''", "S")

		iStrVal = iStrVal  & "AND BH.SALES_ORG =  " & FilterVar(lgConSalesOrgCd , "''", "S") & ""	' �������� 
	End If	

	UNISqlId(0) = "SD513QA401"					
    UNIValue(0,0) = lgSelectList                                      
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
	
	'�������� ���翩�� 
	If lgConSalesOrgCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConSalesOrgCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesOrgNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesOrgNm", "")
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


