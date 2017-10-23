<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S5115QB4
'*  4. Program Name         : ���Ⱑ�����ܾ���Ȳ 
'*  5. Program Desc         : ���Ⱑ�����ܾ���Ȳ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/06/18
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kwakeunkyoung
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
    Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3			'�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgPageNo
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

    Dim iPrevEndRow
    
    Dim lgFromDt			'��ȸ�Ⱓ����(�����)
    Dim lgToDt				'��ȸ�Ⱓ��(�����)
    Dim lgFromDt2			'��ȸ�Ⱓ����(������)
    Dim lgToDt2				'��ȸ�Ⱓ��(������)
    Dim lgBizAreaCd			'����� 
    Dim lgDnTypeCd			'�������� 
    
    Dim lgRdoFlag			'��ȸ���� (Y:���̺�,N:��ü)    
    
    lgFromDt		= Trim(Request("txtHConFromDt"))        
    lgToDt			= Trim(Request("txtHConToDt"))
    lgFromDt2		= Trim(Request("txtHConFromDt2"))        
    lgToDt2			= Trim(Request("txtHConToDt2"))
    lgBizAreaCd		= Replace(Trim(Request("txtHConBizArea")),"'","''")
    lgDnTypeCd		= Replace(Trim(Request("txtHConDnType")),"'","''")   

    lgRdoFlag		= Trim(Request("txtHConRdoFlag"))

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
Sub MakeSpreadSheetData()

    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 100     

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
Sub FixUNISQLData()

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim iStrVal    
    Dim iStrVal2    
    
    Redim UNISqlId(3)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(3,4)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
    iStrVal = "WHERE"
    	
	'�������===========================================================================================
	If Len(lgFromDt) Then
		iStrVal = iStrVal & " DH.ACTUAL_GI_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
	End If
	If Len(lgToDt) Then
		iStrVal = iStrVal & " AND DH.ACTUAL_GI_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
	End If
	
	'������=============================================================================================    	
	If Len(lgBizAreaCd) Then
		UNISqlId(1)		= "s0000qa013"	
		UNIValue(1,0)	= FilterVar(lgBizAreaCd, "''", "S")
		iStrVal = iStrVal & " AND DH.BIZ_AREA =  " & FilterVar(lgBizAreaCd , "''", "S") & ""				
	End If
	
	'�������¸�===========================================================================================	
    If Len(lgDnTypeCd) Then		    
		UNISqlId(2)		= "s0000qa000"	
		UNIValue(2,0)	= FilterVar("I0001", "''", "S")
		UNIValue(2,1)	= FilterVar(lgDnTypeCd, "''", "S")
		iStrVal = iStrVal & " AND DH.MOV_TYPE =  " & FilterVar(lgDnTypeCd , "''", "S") & ""				
	End If

	'��������===========================================================================================
	If Len(lgFromDt2) Then
		iStrVal2 = " AND BH.BILL_DT >=  " & FilterVar(UNIConvDate(lgFromDt2), "''", "S") & ""				
	End If
	If Len(lgToDt2) Then
		iStrVal2 = iStrVal2 & " AND BH.BILL_DT <=  " & FilterVar(UNIConvDate(lgToDt2), "''", "S") & ""				
	End If

	'Y:���̺�===========================================================================================	
	If lgRdoFlag = "Y" Then
		iStrVal3 = " WHERE	TT.GI_AMT - TT.BILL_AMT <> 0 "
	End If		

	'===================================================================================================	

	UNISqlId(0) = "S5115QA401"					

    UNIValue(0,0) = lgSelectList                                      
	UNIValue(0,1) = iStrVal	         
	UNIValue(0,2) = iStrVal2	         
	UNIValue(0,3) = iStrVal3	         
		
	' for summary 
	UNISqlId(3) = "S5115QA402"					
	UNIValue(3,0) = " SUM(ISNULL(TT.GI_AMT,0)) AS TOTAL_GI_AMT, SUM(ISNULL(TT.BILL_AMT,0)) AS TOTAL_BILL_AMT, ( SUM(ISNULL(TT.GI_AMT,0)) - SUM(ISNULL(TT.BILL_AMT,0)) ) AS TOTAL_DIFF_AMT "
	UNIValue(3,1) = iStrVal	         
	UNIValue(3,2) = iStrVal2	         
	UNIValue(3,3) = iStrVal3	         
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                     '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                           '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
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
			Call ConNotFound("txtConBizArea")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBizAreaNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConBizAreaNm", "")		
	End If
	
	'�������� ���翩�� 
	If lgDnTypeCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConDnType")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConDnTypeNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConDnTypeNm", "")
	End If
	 
	 
    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConToDt")	
        Exit Sub
    Else    
        Call MakeSpreadSheetData()
        Call WriteResult()
    End If
End Sub
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
	Response.Write " Parent.ggoSpread.SSShowDataByClip  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " Parent.lgPageNo = """ & lgPageNo & """" & vbCr
	Response.Write " Parent.frm1.txt_TOTAL_Gi_AMT.text = """ & rs3(0) & """" & vbCr
	Response.Write " Parent.frm1.txt_TOTAL_Bill_AMT.text = """ & rs3(1) & """" & vbCr
	Response.Write " Parent.frm1.txt_TOTAL_Balance_AMT.text = """ & rs3(2) & """" & vbCr
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


