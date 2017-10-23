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

    Dim lgBizAreaCd			'����� 
    Dim lgSalesGrpCd		'�����׷� 
    Dim lgSalesTypeCd		'�Ǹ�����    
    Dim lgSoldToParty		'�ֹ�ó 
    Dim lgBillToPartyCd		'����ó 
    Dim lgPayerCd			'����ó 
    Dim lgBillConfFlag		'����ä��Ȯ������ 
    Dim lgBillConfFlag1		'����ä�ǿ��ܿ��� 
        
    lgFromDt		= Trim(Request("txtHConFromDt"))
    lgToDt			= Trim(Request("txtHConToDt"))
    lgBizAreaCd		= Trim(Request("txtHConBizAreaCd"))
    lgSalesGrpCd	= Trim(Request("txtHConSalesGrpCd"))
    lgSalesTypeCd	= Trim(Request("txtHConSalesTypeCd"))      
    lgSoldToParty	= Trim(Request("txtHConSoldToPartyCd"))
    lgBillToPartyCd = Trim(Request("txtHConBillToPartyCd"))
    lgPayerCd		= Trim(Request("txtHConPayerCd"))
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
    Dim  iMaxRsCnt
            
    Const C_SHEETMAXROWS_D = 20     

    lgstrData      = ""

    iLoopCount = 0
    lgStrColorFlag = ""    

    iMaxRsCnt = rs0.recordcount
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			If iLoopCount = iMaxRsCnt and isnumeric(rs0(0)) and ColCnt = 2 Then
				iRowStr = iRowStr & Chr(11) & "�հ�"
			Else				
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			End If
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
   
    Redim UNISqlId(6)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(6,20)										'��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
	'������=============================================================================================    	
	If Len(lgBizAreaCd) Then
		UNISqlId(1)		= "s0000qa013"	
		UNIValue(1,0)	= FilterVar(lgBizAreaCd, "''", "S")
		UNIValue(0,1)	=  " " & FilterVar(lgBizAreaCd, "''", "S") & ""
		UNIValue(0,11)	=  " " & FilterVar(lgBizAreaCd, "''", "S") & ""
	Else
		UNIValue(0,1)	= "NULL"
		UNIValue(0,11)	= "NULL"		
	End If

	'�����׷��===========================================================================================	
    If Len(lgSalesGrpCd) Then
		UNISqlId(2)		= "s0000qa005"	
		UNIValue(2,0)	= FilterVar(lgSalesGrpCd, "''", "S")
		UNIValue(0,2)	= " " & FilterVar(lgSalesGrpCd, "''", "S") & ""
		UNIValue(0,12)	= " " & FilterVar(lgSalesGrpCd, "''", "S") & ""
	Else
		UNIValue(0,2)	= "NULL"
		UNIValue(0,12)	= "NULL"
	End If

	'�Ǹ�������===========================================================================================	
    If Len(lgSalesTypeCd) Then		
		UNISqlId(3)		= "s0000qa000"	
		UNIValue(3,0)	= FilterVar("S0001", "''", "S")
		UNIValue(3,1)	= FilterVar(lgSalesTypeCd, "''", "S")
		UNIValue(0,3)	= " " & FilterVar(lgSalesTypeCd, "''", "S") & ""
		UNIValue(0,13)	= " " & FilterVar(lgSalesTypeCd, "''", "S") & ""
	Else
		UNIValue(0,3)	= "NULL"
		UNIValue(0,13)	= "NULL"
	End If

    '�ֹ�ó��=============================================================================================
    If Len(lgSoldToParty) Then		
		UNISqlId(4)		= "s0000qa002"	
		UNIValue(4,0)	= FilterVar(lgSoldToParty, "''", "S")
		UNIValue(0,4)	= " " & FilterVar(lgSoldToParty, "''", "S") & ""
		UNIValue(0,14)	= " " & FilterVar(lgSoldToParty, "''", "S") & ""		
	Else
		UNIValue(0,4)	= "NULL"
		UNIValue(0,14)	= "NULL"		
	End If	

	'����ó��=============================================================================================	
    If Len(lgPayerCd) Then		
		UNISqlId(5)		= "s0000qa002"	
		UNIValue(5,0)	= FilterVar(lgPayerCd, "''", "S")
		UNIValue(0,5)	= " " & FilterVar(lgPayerCd, "''", "S") & ""
		UNIValue(0,15)	= " " & FilterVar(lgPayerCd, "''", "S") & ""		
	Else
		UNIValue(0,5)	= "NULL"
		UNIValue(0,15)	= "NULL"		
	End If
	
	'����ó��=============================================================================================
    If Len(lgBillToPartyCd) Then		
		UNISqlId(6)		= "s0000qa002"	
		UNIValue(6,0)	= FilterVar(lgBillToPartyCd, "''", "S")
		UNIValue(0,6)	= " " & FilterVar(lgBillToPartyCd, "''", "S") & ""
		UNIValue(0,16)	= " " & FilterVar(lgBillToPartyCd, "''", "S") & ""
	Else
		UNIValue(0,6)	= "NULL"
		UNIValue(0,16)	= "NULL"		
	End If

	'��ȸ�Ⱓ(��,��)����==================================================================================
	If Len(lgFromDt) Then
		UNIValue(0,7) = " " & FilterVar(lgFromDt, "''", "S") & ""                           
		UNIValue(0,17) = " " & FilterVar(lgFromDt, "''", "S") & ""                           
	End If		
	
	'��ȸ�Ⱓ(��,��)��====================================================================================
	If Len(lgToDt) Then
		UNIValue(0,8) = " " & FilterVar(lgToDt, "''", "S") & ""                           	
		UNIValue(0,18) = " " & FilterVar(lgToDt, "''", "S") & ""                           	
	End If

	If lgBillConfFlag <> "%" Then
		UNIValue(0,9)	= " " & FilterVar(lgBillConfFlag, "''", "S") & ""
		UNIValue(0,19)	= " " & FilterVar(lgBillConfFlag, "''", "S") & ""
	Else
		UNIValue(0,9)	= "NULL"
		UNIValue(0,19)	= "NULL"		
	End If

	If lgBillConfFlag1 <> "%" Then
		UNIValue(0,10)	= " " & FilterVar(lgBillConfFlag1, "''", "S") & ""
		UNIValue(0,20)	= " " & FilterVar(lgBillConfFlag1, "''", "S") & ""
	Else
		UNIValue(0,10)	= "NULL"
		UNIValue(0,20)	= "NULL"		
	End If


	UNISqlId(0) = "SD513QA101"					
    UNIValue(0,0) = lgSelectList
        
   
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)
    
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
			Call ConNotFound("txtConBizAreaCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBizAreaNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConBizAreaNm", "")		
	End If
	
	'�����׷� ���翩�� 
	If lgSalesGrpCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConSalesGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")
	End If

	'�Ǹ����� ���翩�� 
	If lgSalesTypeCd <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing			
			Call ConNotFound("txtConSalesTypeCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesTypeNm", rs3(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesTypeNm", "")
	End If

	'�ֹ�ó ���翩�� 
	If lgSoldToParty <> "" Then
		If rs4.EOF And rs4.BOF Then
			rs4.Close
			Set rs4 = Nothing			
			Call ConNotFound("txtConSoldToPartyCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSoldToPartyNm", rs4(1))		
		End If
	Else
		Call WriteConDesc("txtConSoldToPartyNm", "")
	End If

	'����ó ���翩�� 
	If lgPayerCd <> "" Then
		If rs5.EOF And rs5.BOF Then
			rs5.Close
			Set rs5 = Nothing			
			Call ConNotFound("txtConPayerCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConPayerNm", rs5(1))		
		End If
	Else
		Call WriteConDesc("txtConPayerNm", "")
	End If
	
	'����ó ���翩�� 
	If lgBillToPartyCd <> "" Then
		If rs6.EOF And rs6.BOF Then
			rs6.Close
			Set rs6 = Nothing			
			Call ConNotFound("txtConBillToPartyCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBillToPartyNm", rs6(1))		
		End If
	Else
		Call WriteConDesc("txtConBillToPartyNm", "")
	End If

    If  rs0.RecordCount <= 1 OR (rs0.EOF And rs0.BOF) Then	
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
	Response.Write " Parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


