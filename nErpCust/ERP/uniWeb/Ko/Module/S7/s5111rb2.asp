<%'======================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5111RA2
'*  4. Program Name         : ��������ä������ 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Dateǥ������ 
'=======================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgPageNo                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPoType	                                                           '�� : �������� 
Dim strPoFrDt	                                                           '�� : ������ 
Dim strPoToDt	                                                           '�� :
Dim strSpplCd	                                                           '�� : ����ó 
Dim strPurGrpCd	                                                           '�� : ���ű׷� 
Dim strItemCd	                                                           '�� : ǰ�� 
Dim strTrackNo	                                                           '�� : Tracking No
Dim BlankchkFlg

Dim iFrPoint
iFrPoint=0
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

    lgPageNo   = Request("lgPageNo")                               '�� : Next key flag
    lgMaxCount     = 30							                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
	on error resume next 
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt

    lgstrData = ""

	If IsNumeric(lgPageNo) Then 
		If CLng(lgPageNo) > 0 Then
		   rs0.Move = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		End If
	Else
		lgPageNo = 0
	End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

       If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
 
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                 '��: Check if next data exists
       lgPageNo = ""
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim iStrVal
    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(3,2)

    UNISqlId(0) = "S5111ra201"									'* : ������ ��ȸ�� ���� SQL�� 
 
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	If Len(Request("txtSoldtoParty")) Then
	    UNISqlId(1) = "s0000qa002"
	    UNIValue(1,0) = FilterVar(Trim(Request("txtSoldtoParty")), "''", "S")
		iStrVal = " AND A.SOLD_TO_PARTY = " & FilterVar(Trim(Request("txtSoldtoParty")), "''", "S") & ""
	Else
		iStrVal = ""
	End If

	If Len(Request("txtBillToParty")) Then
	    UNISqlId(2) = "s0000qa002"
	    UNIValue(2,0) = FilterVar(Trim(Request("txtBillToParty")), "''", "S")
		iStrVal =  iStrVal & " AND A.BILL_TO_PARTY = " & FilterVar(Trim(Request("txtBillToParty")), "''", "S") & ""
	End If

	If Len(Request("txtSalesGrp")) Then
	    UNISqlId(3) = "s0000qa005"
	    UNIValue(3,0) = FilterVar(Trim(Request("txtSalesGrp")), "''", "S")
		iStrVal =  iStrVal & " AND A.SALES_GRP = " & FilterVar(Trim(Request("txtSalesGrp")), "''", "S") & ""
	End If

    If Len(Trim(Request("txtBillFrDt"))) Then
		iStrVal = iStrVal & " AND A.BILL_DT >= " & FilterVar(UNIConvDate(Request("txtBillFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtBillToDt"))) Then
		iStrVal = iStrVal & " AND A.BILL_DT <= " & FilterVar(UNIConvDate(Request("txtBillToDt")), "''", "S") & ""		
	End If

    UNIValue(0,1) = iStrVal
   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3) '* : Record Set �� ���� ���� 
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    

	Call BeginScriptTag()

	'�ֹ�ó ���翩�� 
	If Trim(Request("txtSoldtoParty")) <> "" Then
		If rs1.EOF And rs1.BOF Then
			Call CloseAdoObject(rs1)
			Call WriteConDesc("txtSoldtoPartyNm", "")
			Call ConNotFound("txtSoldtoParty")			
			Exit Sub
		Else
			Call WriteConDesc("txtSoldtoPartyNm", rs1(1))		
			Call CloseAdoObject(rs1)
		End If
	Else
		Call WriteConDesc("txtSoldtoPartyNm", "")
	End If

	' ����ó ���翩�� 
	If Trim(Request("txtBillToParty")) <> "" Then
		If rs2.EOF And rs2.BOF Then
			Call CloseAdoObject(rs2)
			Call WriteConDesc("txtBillToPartyNm", "")
			Call ConNotFound("txtBillToParty")			
			Exit Sub
		Else	
			Call WriteConDesc("txtBillToPartyNm", rs2(1))		
			Call CloseAdoObject(rs2)
		End If
	Else
		Call WriteConDesc("txtBillToPartyNm", "")
	End If

	' �����׷� ���翩�� 
	If Trim(Request("txtSalesGrp")) <> "" Then
		If rs3.EOF And rs3.BOF Then
			Call CloseAdoObject(rs3)
			Call WriteConDesc("txtSalesGrpNm", "")
			Call ConNotFound("txtSalesGrp")			
			Exit Sub
		Else	
			Call WriteConDesc("txtSalesGrpNm", rs3(1))		
			Call CloseAdoObject(rs3)
		End If
	Else
		Call WriteConDesc("txtSalesGrpNm", "")
	End If

    If  rs0.EOF And rs0.BOF Then	
		Call CloseAdoObject(rs0)
        Call DataNotFound("txtSoldToParty")	
        Exit Sub
    Else    
        Call MakeSpreadSheetData()
		Call CloseAdoObject(rs0)
		If lgPageNo = "1" Then Call SetConditionData()
        Call WriteResult()
    End If

End Sub

' Recordset ��ü Release
Sub CloseAdoObject(ByRef prObjRs)
	If VarType(prObjRs) <> vbObject Then Exit Sub
	
    If Not (prObjRs Is Nothing) Then
       If prObjRs.State = 1 Then		' adStateOpen
          prObjRs.Close
       End If
       Set prObjRs = Nothing
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
	Response.Write "With parent.frm1" & vbCr
	Response.Write ".txtHSoldToParty.value	= """ & ConvSPChars(Request("txtSoldToParty")) & """" & vbCr
	Response.Write ".txtHBillToParty.value	= """ & ConvSPChars(Request("txtBillToParty")) & """" & vbCr
	Response.Write ".txtHBillFrDt.value = """ & Request("txtBillFrDt") & """" & vbCr
	Response.Write ".txtHBillToDt.value	= """ & Request("txtBillToDt") & """" & vbCr
	Response.Write ".txtHSalesGrp.value	= """ & ConvSPChars(Request("txtSalesGrp")) & """" & vbCr
	Response.Write "End with" & vbCr
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
	Response.Write "With parent.frm1" & vbCr
	Response.Write "Parent.ggoSpread.Source	= .vspdData" & vbCr
 	Response.Write ".vspdData.Redraw = False " & vbCr      
	Response.Write "parent.ggoSpread.SSShowDataByClip """ & lgstrData  & """ ,""F""" & vbCr
	Response.Write "parent.lgPageNo	= """ & lgPageNo & """" & vbCr
	Response.Write "parent.DbQueryOk" & vbCr
 	Response.Write ".vspdData.Redraw = True " & vbCr      
	Response.Write "End with" & vbCr
	Call EndScriptTag()
End Sub
%>
