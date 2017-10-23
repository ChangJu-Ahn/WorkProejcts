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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6			'�� : DBAgent Parameter ���� 
    Dim lgstrData																		'�� : data for spreadsheet data
    Dim lgTailList																		'�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgStrColorFlag
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

    Dim lgFromDt			'��ȸ�Ⱓ���� 
    Dim lgToDt				'��ȸ�Ⱓ�� 
	Dim lgBpCd				'�ŷ�ó 
	Dim lgBillTypeCd		'����ä������ 
    Dim lgRdoFlag			'����ä��Ȯ������ 
   
        
    lgFromDt		= Trim(Request("txtHConFromDt"))    
    lgToDt			= Trim(Request("txtHConToDt"))
    lgBpCd			= Trim(Request("txtHdnConBpCd"))
    lgBillTypeCd	= Trim(Request("txtHConBillType"))
    lgRdoFlag		= Trim(Request("txtHConRdoFlag"))
            
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
    Dim  iMaxRsCnt
            
    Const C_SHEETMAXROWS_D = 50     

    lgstrData      = ""

    iLoopCount = 0
    lgStrColorFlag = ""    

    iMaxRsCnt = rs0.recordcount
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			'If iLoopCount = iMaxRsCnt and isnumeric(rs0(0)) and ColCnt = 2 Then
			'	iRowStr = iRowStr & Chr(11) & "�հ�"
			'Else				
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			'End If
		Next
 'COLOR
 		If isnumeric(rs0(0)) Then
 			If rs0(0) > 0 Then
				lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & rs0(0) & gRowSep
			End If
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

    Redim UNIValue(2,4)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
    iStrVal = "WHERE"
	
	'��ȸ�Ⱓ����=========================================================================================
	If Len(lgFromDt) Then
		iStrVal = iStrVal & " BILL_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
	End If		
	
	'��ȸ�Ⱓ��===========================================================================================
	If Len(lgToDt) Then
		iStrVal = iStrVal & " AND BILL_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
	End If
	
	'�ŷ�ó=========================================================================================
	If Len(lgBpCd) Then
		UNISqlId(1)		= "s0000qa002"	
		UNIValue(1,0)	= FilterVar(lgBpCd, "''", "S")
		UNIValue(0,2)	=  " " & FilterVar(lgBpCd, "''", "S") & ""
	Else
		UNIValue(0,2)	= "NULL"
	End If		
	
	'����ä�����¸�=============================================================================================    	
	If Len(lgBillTypeCd) Then
		UNISqlId(2)		= "s0000qa011"	
		UNIValue(2,0)	= FilterVar(lgBillTypeCd, "''", "S")
		UNIValue(0,3)	=  " " & FilterVar(lgBillTypeCd, "''", "S") & ""
	Else
		UNIValue(0,3)	= "NULL"
	End If

	'Ȯ������===========================================================================================	
	If lgRdoFlag <> "%" Then
		UNIValue(0,4)	= " " & FilterVar(lgRdoFlag, "''", "S") & ""
	Else
		UNIValue(0,4)	= "NULL"
	End If

	UNISqlId(0) = "S5115QA301"					
    UNIValue(0,0) = lgSelectList                                      
	UNIValue(0,1) = iStrVal	         

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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	

	'�ŷ�ó ���翩�� 
	If lgBpCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConBpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBpNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConBpNm", "")		
	End If

	'����ä������ ���翩�� 
	If lgBillTypeCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConBillTypeCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBillTypeNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConBillTypeNm", "")
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
	Response.Write " Parent.ggoSpread.SSShowDataByClip  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " Parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


