<!-- #Include file="../inc/CommResponse.inc" -->
<%
Const KEY_NAME = "DEFAULT_VALUE" ' 쿠키항목 
Const KEY_CNT = "KEY_CNT"        ' 쿠키 서브항목 
Const KEY_ITEM = "TXT"           ' 쿠키 이름 TXT1, TXT2 이렇게 나간다. 

' 3일후에 쿠키 지워지도록 
Dim MyDay

MyDay = Date()
MyDay = MyDay + 3
Response.cookies(KEY_NAME).path = "/"
Response.cookies(KEY_NAME).expires = MyDay

Dim nCnt, nRealCnt, nSeq, nNextCnt
Dim Key()

nNextCnt = Request.Cookies(KEY_NAME)(KEY_CNT)

If nNextCnt = "" Then
    nNextCnt = 0
End If

nCnt = Request("cnt")
Redim Key(nCnt)

nSeq = 0
nRealCnt = nNextCnt

For i = 0 to nCnt - 1
	bIsSame = False

	If Request("cb" & i) <> "" Then
		For Each ItemKey in Request.Cookies(KEY_NAME) 
			If Request.Cookies(KEY_NAME)(ItemKey) =  Request("cb" & i) Then
				bIsSame = True
			End If
		Next

		If Not bIsSame Then
			nSeq = nSeq + 1
			Response.cookies(KEY_NAME)(KEY_ITEM & (nNextCnt + nSeq)) = Request("cb" & i) & ""
			nRealCnt = nRealCnt + 1
		End If

	End If
Next

Response.cookies(KEY_NAME)(KEY_CNT) = nRealCnt
%>

<script language="vbscript">
	parent.close()
</script>
