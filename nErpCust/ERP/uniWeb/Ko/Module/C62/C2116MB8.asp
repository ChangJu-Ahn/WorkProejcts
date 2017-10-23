<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim lgPlantCd
Dim lgPlantNm
Dim lgItemAcctCd
Dim lgItemAcctNm
Dim lgCItemCd
Dim lgCItemNm
Dim lgErrorStatus
Dim lgPrevKey
Dim lgSelectListDT
Dim LngRow
 
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	

    lgPageNo       = UniCInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    lgMaxCount     = Trim(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgDataExist    = "No"
    lgSelectListDT = Split(Request("lgSelectListDT"),  gColSep)             '☜ : 각 필드의 데이타 타입 
    
    lgPlantCd	   = Trim(Request("txtPlantCd"))
    lgItemAcctCd   = Trim(Request("txtItemAccntCd"))
    lgCItemCd	   = Trim(Request("txtCItemCd"))
	lgPrevKey	   = Trim(Request("lgPrevKey"))
	LngRow			= UniCInt(Trim(Request("MaxRow")),0)
	
	
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount,iLoopCount2
    
    Dim	DMI_CO		'직접재료비(내부)
	Dim DMO_CO		'직접재료비(외부)
	Dim IMI_CO		'간접재료비(내부)
	Dim IMO_CO		'간접재료비(외부)
	Dim DLI_CO		'직접노무비(내부)
	Dim DLO_CO		'직접노무비(외부)
	Dim ILI_CO		'간접노무비(내부)
	Dim ILO_CO		'간접노무비(외부)
	Dim DEI_CO		'직접경비(내부)
	Dim DEO_CO		'직접경비(외부)
	Dim IEI_CO		'간접경비(내부)
	Dim IEO_CO		'간접경비(외부)
	
	Dim iItemCd			'0
	Dim iItemNm			'1
	Dim iItemAcctNm		'2
	Dim	iBasicUnit		'3
	
	
    lgstrData = ""

    lgDataExist    = "Yes"
	
	
    IF lgPrevKey <> "" Then
		Do while Not (rs0.EOF Or rs0.BOF)
			 IF Trim(rs0(0)) = lgPrevKey  Then
				Exit Do	
			 END IF
		     rs0.MoveNext
		Loop
	END IF
    
    iLoopCount = 0
    iLoopCount2 = 0
    iItemCd = ""
    iItemNm = ""
    iItemAcctNm = ""
    iBasicUnit = ""

	DMI_CO = 0 : DMO_CO = 0 : IMI_CO = 0 : IMO_CO = 0 : DLI_CO = 0 : DLO_CO = 0
	ILI_CO = 0 : ILO_CO = 0 : DEI_CO = 0 : DEO_CO = 0 : IEI_CO = 0 : IEO_CO = 0
    
    lgstrData = ""
    
    Do while Not (rs0.EOF Or rs0.BOF)
 		IF iItemCd <> Trim(rs0(0)) Then
			IF iitemCd <> "" Then
				If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0)  Then
					lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(0),iItemCd))			'품목코드 
					lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(1),iItemAcctNm))		'품목계정 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(2),DMI_CO+IMI_CO)	'재료비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(3),DLI_CO+ILI_CO)	'노무비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(4),DEI_CO+IEI_CO)	'경비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(5),DMI_CO+IMI_CO+DLI_CO+ILI_CO+DEI_CO+IEI_CO)	'계(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(6),DMI_CO)			'직접재료비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(7),DLI_CO)			'직접노무비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(8),DEI_CO)			'직접경비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(9),DMI_CO+DLI_CO+DEI_CO) '직접비계(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(10),IMI_CO)			'간접재료비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(11),ILI_CO)			'간접노무비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(12),IEI_CO)			'간접경비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(13),IMI_CO+ILI_CO+IEI_CO) '간접비계(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(14),iItemCd) '품목코드(Hidden)
					lgstrData = lgstrData & Chr(11) & iLoopCount2 + LngRow + 1
					lgstrData = lgstrData & Chr(11) & Chr(12)
					iLoopCount2 =  iLoopCount2 + 1

					lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(0),iItemNm))			'품목명 
					lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(1),iBasicUnit))		'품목계정 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(2),DMO_CO+IMO_CO)	'재료비(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(3),DLO_CO+ILO_CO)	'노무비(내부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(4),DEO_CO+IEO_CO)	'경비(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(5),DMO_CO+IMO_CO+DLO_CO+ILO_CO+DEO_CO+IEO_CO)	'계(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(6),DMO_CO)			'직접재료비(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(7),DLO_CO)			'직접노무비(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(8),DEO_CO)			'직접경비(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(9),DMO_CO+DLO_CO+DEO_CO) '직접비계(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(10),IMO_CO)			'간접재료비(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(11),ILO_CO)			'간접노무비(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(12),IEO_CO)			'간접경비(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(13),IMO_CO+ILO_CO+IEO_CO) '간접비계(외부)
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(14),iItemCd) '품목코드(Hidden)
					lgstrData = lgstrData & Chr(11) & iLoopCount2 + LngRow + 1
					lgstrData = lgstrData & Chr(11) & Chr(12)
 					iLoopCount2 =  iLoopCount2 + 1

					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(0),"")			
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(1),"SUM")		
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(2),DMI_CO+IMI_CO+DMO_CO+IMO_CO)	'재료비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(3),DLI_CO+ILI_CO+DLO_CO+ILO_CO)	'노무비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(4),DEI_CO+IEI_CO+DEO_CO+IEO_CO)	'경비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(5),DMI_CO+IMI_CO+DLI_CO+ILI_CO+DEI_CO+IEI_CO+DMO_CO+IMO_CO+DLO_CO+ILO_CO+DEO_CO+IEO_CO)	'계 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(6),DMI_CO+DMO_CO)			'직접재료비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(7),DLI_CO+DLO_CO)			'직접노무비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(8),DEI_CO+DEO_CO)			'직접경비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(9),DMI_CO+DLI_CO+DEI_CO+DMO_CO+DLO_CO+DEO_CO) '직접비계 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(10),IMI_CO+IMO_CO)			'간접재료비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(11),ILI_CO+ILO_CO)			'간접노무비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(12),IEI_CO+IEO_CO)			'간접경비 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(13),IMI_CO+ILI_CO+IEI_CO+IMO_CO+ILO_CO+IEO_CO) '간접비계 
					lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(14),iItemCd) '품목코드(Hidden)
					lgstrData = lgstrData & Chr(11) & iLoopCount2 + LngRow + 1
					lgstrData = lgstrData & Chr(11) & Chr(12)
				
					DMI_CO = 0 : DMO_CO = 0 : IMI_CO = 0 : IMO_CO = 0 : DLI_CO = 0 : DLO_CO = 0
					ILI_CO = 0 : ILO_CO = 0 : DEI_CO = 0 : DEO_CO = 0 : IEI_CO = 0 : IEO_CO = 0

					iLoopCount2 =  iLoopCount2 + 1				
					iLoopCount =  iLoopCount + 1
				Else
				    lgPageNo = lgPageNo + 1
				    lgPrevKey = iItemCd
				    Exit Do
				End If
			END IF

			iItemCd = Trim(rs0(0))
			iItemNm = Trim(rs0(1))
			iItemAcctNm = Trim(rs0(2))
			iBasicUnit = Trim(rs0(3))
			
			
		END IF	
		
		'rs0(4) : DI_FALG rs0(5) : COST_ELMT_TYPE
		IF UCase(Trim(rs0(4))) = "D" AND UCase(Trim(rs0(5))) = "M" Then
			DMI_CO = DMI_CO + Cdbl(rs0(6))
			DMO_CO = DMO_CO + Cdbl(rs0(7))
		ELSEIF UCase(Trim(rs0(4))) = "I" AND UCase(Trim(rs0(5))) = "M" Then
			IMI_CO = IMI_CO + CDbl(rs0(6))
			IMO_CO = IMO_CO + Cdbl(rs0(7))
		ELSEIF UCase(Trim(rs0(4))) = "D" AND UCase(Trim(rs0(5))) = "L" Then
			DLI_CO = DLI_CO + Cdbl(rs0(6))
			DLO_CO = DLO_CO + CDbl(rs0(7))
		ELSEIF UCase(Trim(rs0(4))) = "I" AND UCase(Trim(rs0(5))) = "L" Then
			ILI_CO = ILI_CO + Cdbl(rs0(6))
			ILO_CO = ILO_CO + Cdbl(rs0(7))
		ELSEIF UCase(Trim(rs0(4))) = "D" AND UCase(Trim(rs0(5))) = "E" Then
			DEI_CO = DEI_CO + Cdbl(rs0(6))
			DEO_CO = DEO_CO + CDbl(rs0(7))
		ELSE
			IEI_CO = IEI_CO + Cdbl(rs0(6))
			IEO_CO = IEO_CO + Cdbl(rs0(7))
		END IF
			
        rs0.MoveNext
	Loop

	IF iItemCd <> "" and iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then
		lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(0),iItemCd))			'품목코드 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(1),iItemAcctNm))		'품목계정 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(2),DMI_CO+IMI_CO)	'재료비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(3),DLI_CO+ILI_CO)	'노무비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(4),DEI_CO+IEI_CO)	'경비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(5),DMI_CO+IMI_CO+DLI_CO+ILI_CO+DEI_CO+IEI_CO)	'계(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(6),DMI_CO)			'직접재료비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(7),DLI_CO)			'직접노무비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(8),DEI_CO)			'직접경비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(9),DMI_CO+DLI_CO+DEI_CO) '직접비계(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(10),IMI_CO)			'간접재료비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(11),ILI_CO)			'간접노무비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(12),IEI_CO)			'간접경비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(13),IMI_CO+ILI_CO+IEI_CO) '간접비계(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(14),iItemCd) '품목코드(Hidden)
		lgstrData = lgstrData & Chr(11) & iLoopCount2 + LngRow + 1
		lgstrData = lgstrData & Chr(11) & Chr(12)
		iLoopCount2 =  iLoopCount2 + 1

		lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(0),iItemNm))			'품목명 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(1),iBasicUnit))		'품목계정 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(2),DMO_CO+IMO_CO)	'재료비(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(3),DLO_CO+ILO_CO)	'노무비(내부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(4),DEO_CO+IEO_CO)	'경비(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(5),DMO_CO+IMO_CO+DLO_CO+ILO_CO+DEO_CO+IEO_CO)	'계(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(6),DMO_CO)			'직접재료비(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(7),DLO_CO)			'직접노무비(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(8),DEO_CO)			'직접경비(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(9),DMO_CO+DLO_CO+DEO_CO) '직접비계(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(10),IMO_CO)			'간접재료비(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(11),ILO_CO)			'간접노무비(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(12),IEO_CO)			'간접경비(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(13),IMO_CO+ILO_CO+IEO_CO) '간접비계(외부)
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(14),iItemCd) '품목코드(Hidden)
		lgstrData = lgstrData & Chr(11) & iLoopCount2 + LngRow + 1
		lgstrData = lgstrData & Chr(11) & Chr(12)
		iLoopCount2 =  iLoopCount2 + 1

		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(0),"")			
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(1),"SUM")		
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(2),DMI_CO+IMI_CO+DMO_CO+IMO_CO)	'재료비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(3),DLI_CO+ILI_CO+DLO_CO+ILO_CO)	'노무비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(4),DEI_CO+IEI_CO+DEO_CO+IEO_CO)	'경비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(5),DMI_CO+IMI_CO+DLI_CO+ILI_CO+DEI_CO+IEI_CO+DMO_CO+IMO_CO+DLO_CO+ILO_CO+DEO_CO+IEO_CO)	'계 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(6),DMI_CO+DMO_CO)			'직접재료비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(7),DLI_CO+DLO_CO)			'직접노무비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(8),DEI_CO+DEO_CO)			'직접경비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(9),DMI_CO+DLI_CO+DEI_CO+DMO_CO+DLO_CO+DEO_CO) '직접비계 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(10),IMI_CO+IMO_CO)			'간접재료비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(11),ILI_CO+ILO_CO)			'간접노무비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(12),IEI_CO+IEO_CO)			'간접경비 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(13),IMI_CO+ILI_CO+IEI_CO+IMO_CO+ILO_CO+IEO_CO) '간접비계 
		lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(14),iItemCd) '품목코드(Hidden)
		lgstrData = lgstrData & Chr(11) & iLoopCount2 + LngRow + 1
		lgstrData = lgstrData & Chr(11) & Chr(12)
		iLoopCount2 =  iLoopCount2 + 1
	END IF

    If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then                                            '☜: Check if next data exists
        lgPrevKey = ""
    End If
  	
  	
	rs0.Close
    Set rs0 = Nothing 
    
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Redim UNISqlId(0) 
                                                        '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	

	' 유효성 체크 
  	Call CommonQueryRs("PLANT_NM","B_PLANT","PLANT_CD = " & FilterVar(lgPlantCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)	
			'공장이 존재하지 않습니다.
		Call SetErrorStatus()
		Exit Sub
	Else
		lgPlantNm = Trim(Replace(lgF0,Chr(11),""))
	End if

	IF Trim(lgItemAcctCd) <> "" Then 
  		Call CommonQueryRs("MINOR_NM","B_MINOR","MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  AND MINOR_CD = " & FilterVar(lgItemAcctCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
			Call DisplayMsgBox("169952", vbInformation, "", "", I_MKSCRIPT)	
				'품목계정이 Minor 코드에 존재하지 않습니다.
			Call SetErrorStatus()
			Exit Sub
		Else
			
			lgItemAcctNm = Trim(Replace(lgF0,Chr(11),""))
		End if
	END IF
	
	IF Trim(lgCItemCd) <> "" Then 
  		Call CommonQueryRs("ITEM_NM","B_ITEM","ITEM_CD = " & FilterVar(lgCItemCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
			lgCItemNm = ""
			Exit Sub
		Else
			lgCItemNm = Trim(Replace(lgF0,Chr(11),""))
		End if
	END IF
	
	
    Redim UNIValue(0,2)

    UNISqlId(0) = "C2110MA101"

    UNIValue(0,0) = FilterVar(lgPlantCd, "''", "S")				'공장코드 
    UNIValue(0,1) = FilterVar(lgCItemCd, "''", "S")				'품목코드 
    UNIValue(0,2) = FilterVar(lgItemAcctCd ,"" & FilterVar("%", "''", "S") & "","S")			'품목계정 
	

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

	IF lgErrorStatus = "YES" Then
		Exit Sub
	END IF
	
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
		
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
     
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("232500", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else 
		Call  MakeSpreadSheetData()
    End If

End Sub





'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
    Parent.Frm1.txtPlantNm.Value = "<%=ConvSPChars(lgPlantNm)%>"                 
    Parent.Frm1.txtItemAccntNM.Value = "<%=ConvSPChars(lgItemAcctNm)%>"                 
    Parent.Frm1.txtCItemNM.Value = "<%=ConvSPChars(lgCItemNm)%>"                 
    
   
    If "<%=lgDataExist%>" = "Yes" Then
    
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.hPlantCd.Value = Parent.Frm1.txtPlantCd.Value                  'For Next Search
		  Parent.Frm1.hItemAccntCd.Value = Parent.Frm1.txtItemAccntCD.Value                  'For Next Search
		  Parent.Frm1.hCItemCd.Value = Parent.Frm1.txtCItemCd.Value                  'For Next Search	
       End If
    
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"            '☜ : Display data
       Parent.lgPageNo_A      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.lgPrevKey      =  "<%=lgPrevKey%>"               '☜ : Next next data tag
       Parent.DbQueryOk()
    End If   
</Script>	
