<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2313QB1
'*  4. Program Name         : �ҷ�������ȸ 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPlantCd                                               '   ���� 
Dim strDtFr       	                                     '   �Ⱓ(From)
Dim strDtTo		  				'   �Ⱓ(From)
Dim strInspReqNo                                        '   �˻��Ƿڹ�ȣ 
Dim strLotNo						'  ��Ʈ��ȣ 
Dim strItemCd                                             '   ǰ�� 
Dim strInspItemCd					'�˻��׸� 

Dim FilterPlantCd
Dim FilterDtFr
Dim FilterDtTo
Dim FilterInspReqNo
Dim FilterLotNo
Dim FilterItemCd
Dim FilterInspItemCd

Dim strFlag

'Header�� Name�κп� ���� ���� 
Dim strPlantNm
Dim strItemNm
Dim strInspItemNm
Dim iOpt
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
Call HideStatusWnd 

lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgTailList     = Request("lgTailList")                                 '�� : Orderby value
iOpt		= Request("iOpt")

Call TrimData()
Call  HeaderData()                                                '�� : Header�� Name�κ� �ҷ����� 
Call FixUNISQLData()
Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '��¥ 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' �ݾ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '���� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '�ܰ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   'ȯ�� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt)) 
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '��: Check if next data exists
        lgStrPrevKey = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub HeaderData()
	Dim iStr
	
	Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(0,0)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
	
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
	
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNISqlId(0) = "Q2313QA121"
	UNIValue(0,0) = FilterPlantCd		'---���� 
	
    	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  	iStr = Split(lgstrRetMsg,gColSep)
    
    	If iStr(0) <> "0" Then
        		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    	End If    
        
    	If  rs0.EOF And rs0.BOF Then
        		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        		rs0.Close
        		Set rs0 = Nothing
        		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    	Else    
        		strPlantNm=rs0(0)
        		rs0.Close
        		Set rs0 = Nothing
    	End If
    	
	'ǰ��� 
	If strItemCd <> "" Then
		UNISqlId(0) = "Q2313QA122"
		UNIValue(0,0) = FilterItemCd		'---ǰ�� 
		
    		lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  		iStr = Split(lgstrRetMsg,gColSep)
    
    		If iStr(0) <> "0" Then
        			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    		End If    
        
    		If  rs0.EOF And rs0.BOF Then
        			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        			rs0.Close
        			Set rs0 = Nothing
        			Response.End													'��: �����Ͻ� ���� ó���� ������ 
    		Else    
        			strItemNm=rs0(0)
        			rs0.Close
        			Set rs0 = Nothing
    		End If
	End If
	
		
	'�˻��׸� 
	If strInspItemCd <> "" Then
		Redim UNIValue(0,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
		
		UNISqlId(0) = "Q2313QA123"
		UNIValue(0,0) = FilterPlantCd		'---���� 
		UNIValue(0,1) = FilterItemCd		'---ǰ�� 
		UNIValue(0,2) = FilterInspItemCd	'---�˻��׸� 
		
		lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  		iStr = Split(lgstrRetMsg,gColSep)
    
    		If iStr(0) <> "0" Then
        			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    		End If    
        
    		If  rs0.EOF And rs0.BOF Then
        			Call DisplayMsgBox("220201", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        			rs0.Close
        			Set rs0 = Nothing
        			Response.End													'��: �����Ͻ� ���� ó���� ������ 
    		Else    
        			strInspItemNm=rs0(0)
        			rs0.Close
        			Set rs0 = Nothing
    		End If
	End If
	
	'--------------- ������ coding part(�������,End)----------------------------------------------------	
     	
End Sub

Sub FixUNISQLData()
	Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Select Case strFlag
		Case "N"
			Redim UNIValue(0,6)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
	                                                                      			'    parameter�� ���� ���� ������ 
			UNISqlId(0) = "Q2313QA101"
		Case "I"
			Redim UNIValue(0,7)    
			UNISqlId(0) = "Q2313QA102"
		Case "A"
			Redim UNIValue(0,9)             
			UNISqlId(0) = "Q2313QA103"
	End Select
	
	'--------------- ������ coding part(�������,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1) = FilterPlantCd		'---���� 
    UNIValue(0,2) = FilterDtFr			'---�Ⱓ 
    UNIValue(0,3) = FilterDtTo
	UNIValue(0,4) = FilterInspReqNo		'---�˻��Ƿڹ�ȣ 
    UNIValue(0,5) = FilterLotNo		'---��Ʈ��ȣ 
	
	Select Case strFlag
	    	Case "N"
	
	    	Case "I"
	    		 UNIValue(0,6) = FilterItemCd					'---ǰ�� 
	    	Case "A"
	    		UNIValue(0,6) = FilterItemCd					'---ǰ�� 
	    		UNIValue(0,7) = FilterInspItemCd	    			'---�˻��׸� 
	    End Select
	    
'--------------- ������ coding part(�������,End)----------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)		'---	Sort By ���� 

     	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    strPlantCd = Request("txtPlantCd")
    strDtFr = UNIConvDate(Request("txtDtFr"))
	strDtTo = UNIConvDate(Request("txtDtTo"))
	strInspReqNo = Request("txtInspReqNo")
	strLotNo = Request("txtLotNo")
	strItemCd = Request("txtItemCd")
	strInspItemCd = Request("txtInspItemCd")
	
    FilterPlantCd  = FilterVar(strPlantCd, "''", "S")
    FilterDtFr =FilterVar(strDtFr, "''", "S")
    FilterDtTo =FilterVar(strDtTo, "''", "S")
	FilterInspReqNo = FilterVar(strInspReqNo, "''", "S")
	FilterLotNo = FilterVar(strLotNo, "''", "S")
	FilterItemCd = FilterVar(strItemCd, "''", "S")
	FilterInspItemCd = FilterVar(strInspItemCd, "''", "S")
		
	If strItemCd = "" And strInspItemCd = "" Then
		strFlag = "N"
	ElseIf strItemCd <> "" And strInspItemCd = "" Then
		strFlag = "I"
	ElseIf strItemCd <> "" And strInspItemCd <> "" Then
		strFlag = "A"
	End If	
	'--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub
%>
<Script Language=vbscript>
    With parent
		'�������Ÿ Display
		.frm1.txtPlantNm.Value = "<%=ConvSPChars(strPlantNm)%>"
		.frm1.txtItemNm.Value = "<%=ConvSPChars(strItemNm)%>"
		.frm1.txtInspItemNm.Value = "<%=ConvSPChars(strInspItemNm)%>"
			
		'Detail Data Display
         .ggoSpread.Source = .frm1.vspdData 
         .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                          '��: Display data 
         .lgStrPrevKey_A = "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
         .DbQueryOk("<%=iOpt%>")
	End with
</Script>	
<%
Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>