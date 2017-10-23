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
'*  4. Program Name         : Query Routing Detail
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2000/12/05
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

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT

Dim strPlantCd																'�� : ���� 
Dim strFromDt																'�� : ������ 
Dim strToDt																	'�� : ������ 
Dim strItemCd																'�� : ǰ�� 
Dim strRoutNo
Dim PosRoutOrder															'�� : ����� 
Dim iOpt
Dim iFrPoint
Dim TmpBuffer
Dim iTotalStr

iFrPoint=0

lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
lgMaxCount     = 30							                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    
Call TrimData()
Call FixUNISQLData()
Call QueryData()

%>
<Script Language=vbscript>
    
    With Parent
		.ggoSpread.Source  = .frm1.vspdData2
		Parent.frm1.vspdData2.Redraw = False
		.ggoSpread.SSShowDataByClip "<%=iTotalStr%>","F" 
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData2,"<%=iFrPoint+1%>",.frm1.vspdData2.maxrows,.GetKeyPos("B",4),.GetKeyPos("B",3),"C","Q","X","X")
		Parent.frm1.vspdData2.Redraw = True		'by RSW 2003-08-22

		.lgStrPrevKey_B = "<%=ConvSPChars(lgStrPrevKey)%>"
		.DbQueryOk("<%=iOpt%>")
	End With
</Script>
<%

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
        iFrPoint = iCnt  *  lgMaxCount
    Next

    iRCnt = -1
    ReDim TmpBuffer(0)
    
    Do While Not (rs0.EOF Or rs0.BOF)
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
				Case "TT"   'Time�� 
				    iStr = iStr & Chr(11) & ConvNumTimeFormat(rs0(ColCnt),"00:00:00")
				Case "FA"
					iStr = iStr & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0(ColCnt), 0)  '�ݾ� 2003/03/06
				Case "FB"
					iStr = iStr & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0(ColCnt), 0)  '���� 2003/03/06
				Case "FC"
					iStr = iStr & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0(ColCnt), 0)  '�ܰ� 2003/03/06
				Case "FD"
					iStr = iStr & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0(ColCnt), 0)  'ȯ�� 2003/03/06				    
				Case Else
					If ColCnt = 4 Then
						If Trim (rs0(ColCnt)) ="N" Then
							iStr = iStr & Chr(11) & "����"
						Else  
							iStr = iStr & Chr(11) & "�系"
						End If
						
					Else
						iStr = iStr & Chr(11) & rs0(ColCnt)
					End If
					
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

    Redim UNIValue(0,4)

    UNISqlId(0) = "181300sab"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = UCase(Trim(strPlantCd))
    UNIValue(0,2) = UCase(Trim(strItemCd))
    UNIValue(0,3) = UCase(Trim(strRoutNo))
    
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
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("181200", vbInformation, "", "", I_MKSCRIPT)
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

	strPlantCd    = FilterVar(Request("txtPlantCd"), "''", "S")
	strItemCd     = FilterVar(Request("txtItemCd"), "''", "S")
	strRoutNo     = FilterVar(Request("txtRoutNo"), "''", "S")
	StartDt	      = "1900-01-01"
	EndDt		  = "2999-12-31"
	strFromDt     = FilterVar(Request("txtFromDt"), StartDt, "S")
	strToDt	      = FilterVar(Request("txtToDt"), EndDt, "S")
	iOpt		  = Request("iOpt")
End Sub

'----------------------------------------------------------------------------------------------------------
' Integer Data�� Time ������ Conversion
'----------------------------------------------------------------------------------------------------------
Function ConvNumTimeFormat(ByVal IVal, ByVal tTimeDefault)
	Dim iTime
	Dim iMin
	Dim iSec
on error resume next
err.Clear
				
	If IVal = 0 or Len(Trim(iVal)) = 0 Then
		ConvNumTimeFormat = tTimeDefault
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvNumTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
	End If
End Function
%>
