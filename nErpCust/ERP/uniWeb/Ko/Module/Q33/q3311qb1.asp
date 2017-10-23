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
'*  7. Modified date(Last)  : 2004/08/05
'*  8. Modifier (First)     : Koh Jae Woo
'*  9. Modifier (Last)      : Lee Seung Wook
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
Dim strPlantCd
Dim strDtFr
Dim strDtTo
Dim strItemCd
Dim strRoutNo
Dim strOprNo

Dim FilterPlantCd
Dim FilterDtFr
Dim FilterDtTo
Dim FilterItemCd
Dim FilterRoutNo
Dim FilterOprNo

Dim strFlag

'Header�� Name�κп� ���� ���� 
Dim strPlantNm
Dim strItemNm
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
     Call  HeaderData()                                                   '�� : Header�� Name�κ� �ҷ����� 
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
               Case "F6"   '�ҷ���, ���հݷ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), 2, 0)
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
	UNISqlId(0) = "Q3311QA121"
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
		UNISqlId(0) = "Q3311QA122"
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
	
		
	'�ҷ��� 
	UNISqlId(0) = "Q3311QA124"
	UNIValue(0,0) = FilterPlantCd		'---���� 

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
	
	If  rs0.EOF And rs0.BOF Then
			Call DisplayMsgBox("220401", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
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
	Dim strQryCnd
	Redim UNISqlId(0) 
	Redim UNIValue(0,5)
	
	UNISqlId(0) = "Q3311QA101"
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	'Select Case strFlag
	'	Case "N"
	'		Redim UNIValue(0,4)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
	                                                                      			'    parameter�� ���� ���� ������ 
	'		UNISqlId(0) = "Q3311QA101"
	'	Case "I"
	'		Redim UNIValue(0,5)    
	'		UNISqlId(0) = "Q3311QA102"
	'	Case "W"
	'		Redim UNIValue(0,5)    
	'		UNISqlId(0) = "Q3311QA103"
	'	Case "A"
	'		Redim UNIValue(0,6)    
	'		UNISqlId(0) = "Q3311QA104"
	'End Select
	'--------------- ������ coding part(�������,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1) = FilterPlantCd		'---���� 
    UNIValue(0,2) = FilterDtFr			'---�Ⱓ 
    UNIValue(0,3) = FilterDtTo
    
    If Trim(strItemCd) <> "" Then
		strQryCnd = "AND A.ITEM_CD = " & FilterItemCd
	End If
	
	If Trim(strRoutNo) <> "" Then
		strQryCnd = strQryCnd & "AND A.ROUT_NO = " & FilterRoutNo
	End If
	
	If Trim(strOprNo) <> "" Then
		strQryCnd = strQryCnd & "AND A.OPR_NO = " & FilterOprNo
	End If
	
	UNIValue(0,4) = (Trim(strQryCnd))
	
	'Select Case strFlag
	'    Case "N"
	
	'    Case "I"
	'    	 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
	'    Case "W"
	'    	 UNIValue(0,4) = FilterWcCd	    					'---�۾��� 
	'    Case "A"
	'    	 UNIValue(0,4) = FilterItemCd					'---ǰ�� 
	'    	 UNIValue(0,5) = FilterWcCd	    					'---�۾��� 
	'End Select
     
'--------------- ������ coding part(�������,End)----------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)	'---Group By ���� 

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
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
	strRoutNo = Request("txtRoutNo")
	strOprNo = Request("txtOprNo")
	
	FilterPlantCd  = FilterVar(strPlantCd, "''", "S")
    FilterDtFr =FilterVar(strDtFr, "''", "S")
	FilterDtTo = FilterVar(strDtTo, "''", "S")
	FilterItemCd = FilterVar(strItemCd, "''", "S")
	FilterRoutNo = FilterVar(strRoutNo, "''", "S")
	FilterOprNo = FilterVar(strOprNo, "''", "S")
	
		
	'If strItemCd = "" And strWcCd = "" Then
	'	strFlag = "N"
	'ElseIf strItemCd <> "" And strWcCd = "" Then
	'	strFlag = "I"
	'ElseIf strItemCd = "" And strWcCd <> "" Then
	'	strFlag = "W"
	'ElseIf strItemCd <> "" And strWcCd <> "" Then
	'	strFlag = "A"
	'End If	
'--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub
%>

<Script Language=vbscript>
    
    With Parent
    	 '�������Ÿ Display
         .frm1.txtPlantNm.Value = "<%=ConvSPChars(strPlantNm)%>"
		.frm1.txtItemNm.Value = "<%=ConvSPChars(strItemNm)%>"
		.frm1.txtDefectRatioUnit.Value = "<%=ConvSPChars(strDefectRatioUnit)%>"
		.frm1.txtLotRejUnit.Value = "<%=ConvSPChars(strLotRejUnit)%>"
		'Detail Data Display
         .ggoSpread.Source  = .frm1.vspdData
         .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                  '�� : Display data
         .lgStrPrevKey      =  "<%=ConvSPChars(lgStrPrevKey)%>"               '�� : Next next data tag
         .DbQueryOk
	End with
</Script>	
<%
Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
