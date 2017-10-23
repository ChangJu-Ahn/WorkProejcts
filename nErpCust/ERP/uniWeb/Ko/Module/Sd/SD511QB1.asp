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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1  '�� : DBAgent Parameter ���� 
    Dim lgstrData										'�� : data for spreadsheet data
    Dim lgTailList                                      '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgStrColorFlag
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
   
    Dim lgFromDt			'��ȸ���۳⵵    
    Dim lgToDt				'��ȸ����⵵    
    Dim lgBillToPartyCd		'����ó 
    Dim lgBillConfFlag		'����ä��Ȯ������ 
    Dim lgBillConfFlag1		'����ä�ǿ��ܿ��� 

    lgFromDt		= Trim(Request("txtHConFromDt"))
    lgToDt			= Trim(Request("txtHConToDt"))                            
    lgBillToPartyCd = Trim(Request("txtHConBillToPartyCd"))
    lgBillConfFlag	= Trim(Request("txtHConRdoBillConfFlag"))
    lgBillConfFlag1	= Trim(Request("txtHConRdoBillConfFlag1"))    
    
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgSelectList   = Request("txtHlgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtHlgTailList")                                 '�� : Orderby value

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
 'COLOR
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

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim iStrVal    
    
    Redim UNISqlId(1)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(1,12)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
    'iStrVal = "WHERE"
	
	'��ȸ�Ⱓ(��)����==================================================================================
	If Len(lgFromDt) Then
		UNIValue(0,1) = " " & FilterVar(lgFromDt, "''", "S") & ""	
		UNIValue(0,8) = " " & FilterVar(lgFromDt, "''", "S") & ""	
	End If		
	
	'��ȸ�Ⱓ(��)��====================================================================================
	If Len(lgToDt) Then
		UNIValue(0,2) = " " & FilterVar(lgToDt, "''", "S") & ""	
		UNIValue(0,9) = " " & FilterVar(lgToDt, "''", "S") & ""	
	End If		

	'Ȯ������=============================================================================================
	If lgBillConfFlag <> "%" Then
		UNIValue(0,3)	= " " & FilterVar(lgBillConfFlag, "''", "S") & ""
		UNIValue(0,10)	= " " & FilterVar(lgBillConfFlag, "''", "S") & ""
	Else
		UNIValue(0,3)	= "NULL"
		UNIValue(0,10)	= "NULL"		
	End If

	'���ܿ���=============================================================================================
	If lgBillConfFlag1 <> "%" Then
		UNIValue(0,4)	= " " & FilterVar(lgBillConfFlag1, "''", "S") & ""
		UNIValue(0,11)	= " " & FilterVar(lgBillConfFlag1, "''", "S") & ""
	Else
		UNIValue(0,4)	= "NULL"
		UNIValue(0,11)	= "NULL"		
	End If

	'����ó��=============================================================================================
    If Len(lgBillToPartyCd) Then		
		UNISqlId(1)		= "s0000qa002"	
		UNIValue(1,0)	=  FilterVar(lgBillToPartyCd, "''", "S")
		UNIValue(0,5)	=  " " & FilterVar(lgBillToPartyCd, "''", "S") & ""
		UNIValue(0,12)	=  " " & FilterVar(lgBillToPartyCd, "''", "S") & ""
	Else
		UNIValue(0,5) = "NULL"
		UNIValue(0,12) = "NULL"
	End If

	UNIValue(0,6)	= "�հ�"

	iStrVal	= "	YEAR(BILL_DT) YR, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 1 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) JAN, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 2 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) FEB, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 3 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) MAR, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 4 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) APR, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 5 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) MAY, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 6 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) JUN, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 7 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) JUL, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 8 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) AUG, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 9 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) SEP, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 10 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) OCT, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 11 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) NOV, " & _
			"SUM(CASE WHEN MONTH(BILL_DT) = 12 THEN ISNULL((BILL_AMT + VAT_AMT),0) ELSE 0 END) DEC	"
	UNIValue(0,7) = iStrVal

	UNISqlId(0) = "SD511QA101"					
    UNIValue(0,0) = lgSelectList
        
	'UNIValue(0,1) = iStrVal	         
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
	
	'����ó ���翩�� 
	If lgBillToPartyCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConBillToPartyCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBillToPartyNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConBillToPartyNm", "")
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
	Response.Write " Parent.ggoSpread.SSShowData  """ & left(lgstrData,instr(1,lgstrData,"�հ�")) & replace(mid(lgstrData,instr(1,lgstrData,"�հ�") + 1,len(lgstrData)),"�հ�","") & """ ,""F""" & vbCr
	Response.Write " Parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


