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
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         : Query Routing Header
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2002/12/05
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "QB")

Dim lgADF, ADF                                                             '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                   '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPlantCd																'�� : ���� 
Dim strFromDt																'�� : ������ 
Dim strToDt																	'�� : ������ 
Dim strItemCd																'�� : ǰ�� 
Dim strRoutNo
Dim iOpt																'�� : ����� 
Dim TmpBuffer
Dim iTotalStr
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
'======================================================================================================
'	ǰ���̸� ó�����ִ� �κ� 
'======================================================================================================
Redim UNISqlId(2)
Redim UNIValue(2, 0)
	
UNISqlId(0) = "122600sac"
UNISqlId(1) = "122700sab"
UNISqlId(2) = "181300sac"

IF Request("txtPlantCd") = "" Then
   strPlantCd = "|"
ELSE
   strPlantCd = Request("txtPlantCd")
END IF
	
IF Request("txtItemCd") = "" Then
   strItemCd = "|"
ELSE
   strItemCd = Request("txtItemCd")
END IF
	
IF Request("txtRoutNo") = "" Then
   strRoutNo = "|"
ELSE
   strRoutNo = Request("txtRoutNo")
END IF
	
UNIValue(0, 0) = FilterVar(strItemCd,"''","S")
UNIValue(1, 0) = FilterVar(strPlantCd,"''","S")
UNIValue(2, 0) = FilterVar(strRoutNo,"''","S")
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
If rs1.EOF And rs1.BOF Then
	Response.Write "<Script Language= VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End
Else
	Response.Write "<Script Language= VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
End If

If rs0.EOF And rs0.BOF Then
	Response.Write "<Script Language= VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language= VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
End If
	
If rs2.EOF And rs2.BOF Then
	Response.Write "<Script Language= VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtRoutNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language= VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtRoutNm.value = """ & ConvSPChars(rs2(0)) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
End If

rs0.Close
rs1.Close
rs2.Close
		
Set rs0 = Nothing
Set rs1 = Nothing
Set rs2 = Nothing

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
lgMaxCount     = 30							                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgTailList     = Request("lgTailList")                                 '�� : Orderby value

Call TrimData()
Call FixUNISQLData()
Call QueryData()

Response.Write "<Script Language = VBScript>" & vbCrLf
    Response.Write "With parent" & vbCrLf
         Response.Write ".ggoSpread.Source = .frm1.vspdData1" & vbCrLf
         Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf
         
         Response.Write ".lgStrPrevKey_A = """ & ConvSPChars(lgStrPrevKey) & """" & vbCrLf
         Response.Write ".DbQueryOk(""" & iOpt & """)" & vbCrLf
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
    
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

    For iRCnt = 1 To iCnt * lgMaxCount                                   '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    ReDim TmpBuffer(0)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '��¥ 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' �ݾ� 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
               Case "F3"  '���� 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
               Case "F4"  '�ܰ� 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
               Case "F5"   'ȯ�� 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit, 0)
               Case Else
                    iStr = iStr & Chr(11) & rs0(ColCnt) 
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
			ReDim Preserve TmpBuffer(iRCnt)
            TmpBuffer(iRCnt) = iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotalStr = Join(TmpBuffer, "")
	
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
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,6)

    UNISqlId(0) = "181300saa"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = UCase(Trim(strPlantCd))
    UNIValue(0,2) = UCase(Trim(strItemCd))
    UNIValue(0,3) = UCase(Trim(strRoutNo))
    UNIValue(0,4) = UCase(Trim(strFromDt))
    UNIValue(0,5) = UCase(Trim(strToDt))
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg, gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("181300", vbInformation, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

	Dim StartDt
	Dim EndDt

	strPlantCd  = FilterVar(Request("txtPlantCd"), "''", "S")
	strItemCd   = FilterVar(Request("txtItemCd"), "''", "S")
	strRoutNo   = FilterVar(Request("txtRoutNo"), "''", "S")
	StartDt		= FilterVar("1900-01-01","''","S")
	strFromDt   = FilterVar(UniConvDate(Request("txtFromDt")), StartDt, "S")
	EndDt		= FilterVar("2999-12-31","''","S")
	strToDt		= FilterVar(UniConvDate(Request("txtToDt")), EndDt, "S")
	iOpt		= Request("iOpt")

End Sub
%>
