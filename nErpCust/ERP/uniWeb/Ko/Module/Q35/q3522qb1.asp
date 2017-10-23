<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'======================================================================================================
'*  1. Module Name          : Quality
'*  2. Function Name        : ADO  (QUERY)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/27
'*  7. Modified date(Last)  : 2001/01/27
'*  8. Modifier (First)     : Koh Jae Woo
'*  9. Modifier (Last)      : Koh Jae Woo
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgStrPrevKey                                            '�� : ���� �� 
Dim lgMaxCount                                              '�� : Spread sheet �� visible row �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

'--------------- ������ coding part(��������,Start)----------------------------------------------------
Dim strPlantCd                                               '   ���� 
Dim strDtFr       	                                     '   �Ⱓ(From)
Dim strDtTo		  				'   �Ⱓ(From)
Dim strItemCd                                             '   ǰ�� 
Dim strBpCd						'�ŷ�ó 
Dim strInspItemCd					'�˻��׸� 
Dim strDefectTypeCd					'�ҷ����� 

Dim FilterPlantCd
Dim FilterDtFr
Dim FilterDtTo
Dim FilterItemCd
Dim FilterBpCd
Dim FilterInspItemCd
Dim FilterDefectTypeCd

Dim strFlag

'Header�� Name�κп� ���� ���� 
Dim strPlantNm
Dim strItemNm
Dim strBpNm
Dim strInspItemNm
Dim strDefectTypeNm
Dim strDefectRatioUnit
Dim strLotRejUnit
'--------------- ������ coding part(��������,End)------------------------------------------------------

     Call HideStatusWnd 
     lgStrPrevKey     = Request("lgStrPrevKey")                           '�� : Next key flag
     lgMaxCount       = CInt(Request("lgMaxCount"))                       '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
     
     Call  TrimData()                                                     '�� : Parent�� ������ ����Ÿ ���� 
     Call  HeaderData()                                                '�� : Header�� Name�κ� �ҷ����� 
     Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
     Call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iCnt
    Dim iRCnt                                                                     
    Dim strTmpBuffer                                                              
    Dim iStr
    Dim ColCnt
     
    iCnt = 0
    lgstrData = ""
   
    If Len(Trim(lgStrPrevKey)) Then                                              '�� : Chnage str into int
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                         '�� : Discard previous data
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
            lgstrData   = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                     '��: Check if next data exists
        lgStrPrevKey = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF
    Set lgADF = Nothing                                             '��: Release ADF

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
	UNISqlId(0) = "Q3522QA121"
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
		UNISqlId(0) = "Q3522QA122"
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
	
	'�ŷ�ó 
	If strBpCd <> "" Then
		UNISqlId(0) = "Q3522QA123"
		UNIValue(0,0) = FilterBpCd		'---�ŷ�ó 
		
		lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  		iStr = Split(lgstrRetMsg,gColSep)
    
    		If iStr(0) <> "0" Then
        			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    		End If    
        
    		If  rs0.EOF And rs0.BOF Then
        			Call DisplayMsgBox("126200", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        			rs0.Close
        			Set rs0 = Nothing
        			Response.End													'��: �����Ͻ� ���� ó���� ������ 
    		Else    
        			strBpNm=rs0(0)
        			rs0.Close
        			Set rs0 = Nothing
    		End If
	End If
	
	'�˻��׸� 
	If strInspItemCd <> "" Then
		Redim UNIValue(0,2)  
		
		UNISqlId(0) = "Q3522QA124"
		
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
	
	'�ҷ����� 
	If strDefectTypeCd <> "" Then
		Redim UNIValue(0,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
		
		UNISqlId(0) = "Q3522QA125"
		UNIValue(0,0) = FilterPlantCd		'---���� 
		UNIValue(0,1) = FilterDefectTypeCd	'---�ҷ����� 
				
		lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  		iStr = Split(lgstrRetMsg,gColSep)
    
    		If iStr(0) <> "0" Then
        			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    		End If    
        
    		If  rs0.EOF And rs0.BOF Then
        			Call DisplayMsgBox("221101", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        			rs0.Close
        			Set rs0 = Nothing
        			Response.End													'��: �����Ͻ� ���� ó���� ������ 
    		Else    
        			strDefectTypeNm=rs0(0)
        			rs0.Close
        			Set rs0 = Nothing
    		End If
	End If
	
	'�ҷ��� 
	Redim UNIValue(0,0)        
	UNISqlId(0) = "Q3522QA126"
	UNIValue(0,0) = FilterPlantCd		'---���� 

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
	
	If  rs0.EOF And rs0.BOF Then
			Call DisplayMsgBox("220401", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
'			Call ServerMesgBox("900014 : " & 900014, vbCritical, I_MKSCRIPT)    	
			rs0.Close
			Set rs0 = Nothing
			Response.End													'��: �����Ͻ� ���� ó���� ������ 
	Else    
			strDefectRatioUnit=rs0(0)
			rs0.Close
			Set rs0 = Nothing
	End If
	
	'LOT���հݷ� ���� 
	strLotRejUnit = "%"
	
	'--------------- ������ coding part(�������,End)----------------------------------------------------	
     	
End Sub

Sub FixUNISQLData()

	Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Select Case strFlag
		Case "N"
			Redim UNIValue(0,5)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
			UNISqlId(0) = "Q3522QA101"
		Case "I"
			Redim UNIValue(0,6)    
			UNISqlId(0) = "Q3522QA102"
		Case "B"
			Redim UNIValue(0,6)    
			UNISqlId(0) = "Q3522QA103"
		Case "D"
			Redim UNIValue(0,6)    
			UNISqlId(0) = "Q3522QA104"
		Case "IB"
			Redim UNIValue(0,7)    
			UNISqlId(0) = "Q3522QA105"
		Case "IS"
			Redim UNIValue(0,7)    
			UNISqlId(0) = "Q3522QA106"
		Case "ID"
			Redim UNIValue(0,7)    
			UNISqlId(0) = "Q3522QA107"
		Case "BD"
			Redim UNIValue(0,7)    
			UNISqlId(0) = "Q3522QA108"
		Case "IBS"
			Redim UNIValue(0,8)    
			UNISqlId(0) = "Q3522QA109"
		Case "IBD"
			Redim UNIValue(0,8)    
			UNISqlId(0) = "Q3522QA110"
		Case "ISD"
			Redim UNIValue(0,8)    
			UNISqlId(0) = "Q3522QA111"
		Case "A"
			Redim UNIValue(0,9)             
			UNISqlId(0) = "Q3522QA112"
	End Select
	
	'--------------- ������ coding part(�������,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1) = FilterPlantCd		'---���� 
    UNIValue(0,2) = FilterDtFr			'---�Ⱓ 
    UNIValue(0,3) = FilterDtTo
	
	Select Case strFlag
		Case "N"
	
		Case "I"
			 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
		Case "B"
			 UNIValue(0,4) = FilterBpCd	    				'---�ŷ�ó 
		Case "D"
			 UNIValue(0,4) = FilterDefectTypeCd	    			'---�ҷ����� 
		Case "IB"
	  		 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
	    	 UNIValue(0,5) = FilterBpCd	    				'---�ŷ�ó 
	  	Case "IS"
	  		 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
	    	 UNIValue(0,5) = FilterInspItemCd	    			'---�˻��׸� 
	  	Case "ID"
	  		 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
	    	 UNIValue(0,5) = FilterDefectTypeCd	    			'---�ҷ����� 
	    Case "BD"
	  		 UNIValue(0,4) = FilterBpCd	    				'---�ŷ�ó 
	    	 UNIValue(0,5) = FilterDefectTypeCd	    			'---�ҷ����� 
	  	Case "IBS"
			 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
	    	 UNIValue(0,5) = FilterBpCd	    				'---�ŷ�ó 
	    	 UNIValue(0,6) = FilterInspItemCd	    			'---�˻��׸� 
		Case "IBD"
			 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
    		 UNIValue(0,5) = FilterBpCd	    				'---�ŷ�ó 
    		 UNIValue(0,6) = FilterDefectTypeCd	    			'---�ҷ����� 
		Case "ISD"
			 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
	   		 UNIValue(0,5) = FilterInspItemCd	    			'---�˻��׸� 
	   		 UNIValue(0,6) = FilterDefectTypeCd	    			'---�ҷ����� 
		Case "A"
			 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
	    	 UNIValue(0,5) = FilterBpCd	    				'---�ŷ�ó 
	    	 UNIValue(0,6) = FilterInspItemCd	    			'---�˻��׸� 
	    	 UNIValue(0,7) = FilterDefectTypeCd	    			'---�ҷ����� 
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
    'Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
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
    strDtFr = Request("txtDtFr")
	strDtTo = Request("txtDtTo")
	strItemCd = Request("txtItemCd")
	strBpCd = Request("txtBpCd")
	strInspItemCd = Request("txtInspItemCd")
	strDefectTypeCd = Request("txtDefectTypeCd")
	
    FilterPlantCd  = FilterVar(strPlantCd, "''", "S")
    FilterDtFr =FilterVar(strDtFr, "''", "S")
    FilterDtTo =FilterVar(strDtTo, "''", "S")
	FilterItemCd = FilterVar(strItemCd, "''", "S")
	FilterBpCd = FilterVar(strBpCd, "''", "S")
	FilterInspItemCd = FilterVar(strInspItemCd, "''", "S")
	FilterDefectTypeCd = FilterVar(strDefectTypeCd, "''", "S")
	
	If strItemCd = "" And strBpCd = "" And strInspItemCd = "" And strDefectTypeCd = "" Then
		strFlag = "N"
	ElseIf strItemCd <> "" And strBpCd = "" And strInspItemCd = "" And strDefectTypeCd = "" Then
		strFlag = "I"
	ElseIf strItemCd = "" And strBpCd <> "" And strInspItemCd = "" And strDefectTypeCd = "" Then
		strFlag = "B"
	ElseIf strItemCd = "" And strBpCd = "" And strInspItemCd = "" And strDefectTypeCd <> "" Then
		strFlag = "D"
	ElseIf strItemCd <> "" And strBpCd <> "" And strInspItemCd = ""  And strDefectTypeCd = "" Then
		strFlag = "IB"
	ElseIf strItemCd <> "" And strBpCd = "" And strInspItemCd <> "" And strDefectTypeCd = "" Then
		strFlag = "IS"
	ElseIf strItemCd <> "" And strBpCd = "" And strInspItemCd = "" And strDefectTypeCd <> "" Then
		strFlag = "ID"
	ElseIf strItemCd = "" And strBpCd <> "" And strInspItemCd = "" And strDefectTypeCd <> "" Then
		strFlag = "BD"
	ElseIf strItemCd <> "" And strBpCd <> "" And strInspItemCd <> "" And strDefectTypeCd = "" Then
		strFlag = "IBS"
	ElseIf strItemCd <> "" And strBpCd <> "" And strInspItemCd = "" And strDefectTypeCd <> "" Then
		strFlag = "IBD"
	ElseIf strItemCd <> "" And strBpCd = "" And strInspItemCd <> "" And strDefectTypeCd <> "" Then
		strFlag = "ISD"
	ElseIf strItemCd <> "" And strBpCd <> "" And strInspItemCd <> "" And strDefectTypeCd <> "" Then
		strFlag = "A"
	End If	
	
'--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub
%>

<Script Language=vbscript>
    
    With Parent
		'�������Ÿ Display
		.frm1.txtPlantNm.Value = "<%=ConvSPChars(strPlantNm)%>"
		.frm1.txtItemNm.Value = "<%=ConvSPChars(strItemNm)%>"
		.frm1.txtBpNm.Value = "<%=ConvSPChars(strBpNm)%>"
		.frm1.txtInspItemNm.Value = "<%=ConvSPChars(strInspItemNm)%>"
		.frm1.txtDefectTypeNm.Value = "<%=ConvSPChars(strDefectTypeNm)%>"
		.frm1.txtDefectRatioUnit.Value = "<%=ConvSPChars(strDefectRatioUnit)%>"
		.frm1.txtLotRejUnit.Value = "<%=ConvSPChars(strLotRejUnit)%>"
	
		'Detail Data Display
        .ggoSpread.Source  = .frm1.vspdData
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                  '�� : Display data
        .lgStrPrevKey =  "<%=ConvSPChars(lgStrPrevKey)%>"               '�� : Next next data tag
        .DbQueryOk
	End with
</Script>	
<%
Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
