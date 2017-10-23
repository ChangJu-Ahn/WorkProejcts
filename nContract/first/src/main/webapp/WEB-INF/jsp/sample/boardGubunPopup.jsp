<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf" %>
</head>
<body>
	<!-- <table>
	 <tr>
	 	<td><label style="font-weight: bold;">Filter : </label></td>
	 	<td><input type='text' id='txtFilter' onkeyup='{filter();return false}' onkeypress='javascript:if(event.keyCode==13){ filter(); return false;}'></td>
	 </tr>
	</table>
	<br/> -->
	
	<table class="board_list">
        <colgroup>
		<col width="15%"/>
		<col width="10%"/>
	</colgroup>
	<thead>
		<tr>
			<th scope="col">CODE</th>
			<th scope="col">CODE_NM</th>
		</tr>
	</thead>
	<tbody id="TBody">
 	<!-- CODE,
 		 CODE_NM,
 		 HIGH_CODE,
 		 LVL		 -->	 	
		<c:choose>
			<c:when test="${fn:length(list) > 0}">
				<c:forEach items="${list }" var="row">
					<div>
						<c:choose>
							<c:when test="${row.LVL == '1'}">
							<tr id = '${row.CODE_NM }'>
									<td><a href = "#this" name ="code" style="color: blue;">${row.CODE }</a></td>
									<td>${row.CODE_NM }</td>
							</tr>
							</c:when>
							<c:otherwise>
								<tr class = '${row.HIGH_CODE }' style="display: none;">
									<td><a href = "javascript:returnParent('${row.CODE }','${row.CODE_NM }')" style="color: red;">${row.CODE }</a></td>
									<td>${row.CODE_NM }</td>
								</tr>
							</c:otherwise>
						</c:choose>
					</div>
				</c:forEach>
			</c:when>
			<c:otherwise>
				<tr>
					<td colspan="2">조회된 결과가 없습니다.</td>
				</tr>
			</c:otherwise>
		</c:choose>
        </tbody>
    </table>    
</body>
<script type="text/javascript">
	$(document).ready(
		function() {
			$("a[name='code']").on("click", function(e) { //목록으로 버튼				
				e.preventDefault();
				fn_addGubun($(this));
			});
	});

	function fn_addGubun(obj){
		
		var trId = $("."+obj.html());
		
		if($(trId).css('display') == 'none'){			
			$(trId).show(function() {
	    		$(this).attr('hide', false);
	   	 	});
		}else{			
			$(trId).hide(function() {
	    		$(this).attr('hide', true);
	   	 	});
		}
		
		/* var str = "";
		var cnt = 1;
		
		var CODE = new Array();
		var CODE_NM = new Array();
		
		
		<c:forEach items="${list }" var="row">
			if("${row.HIGH_CODE}" == obj.html()){
					CODE[cnt] = "${row.CODE}";
					CODE_NM[cnt] = "${row.CODE_NM}";
					cnt ++;
			}
		</c:forEach>
		
		console.log(CODE); */
		
		/* for(var i=0; i < CODE.length ; i ++){
			if($("#"+CODE_NM[i]).attr('hide') == "true"){
				$("#"+CODE_NM[i]).show(function() {
		    		$(this).attr('hide', false);
		   	 	});
				//$("#"+CODE_NM[i]).hide();
				//chk[i] = "";
			}else{
				$("#"+CODE_NM[i]).hide(function() {
		    		$(this).attr('hide', true);
		   	 	});
				//$("#"+CODE_NM[i]).show();
				//chk[i] = CODE_NM[i];
			}
		}	 */
		
		/*
		if(chk == obj.parent().parent().attr("id")){
			
			for(var i = 1 ; i < CODE.length ; i++){
				$("#"+CODE_NM[i]).remove();
			}
			//$("#"+obj.parent().parent().attr("id")).remove(str);	
			
			chk = "";
		}else
		{
			for(var i = 1 ; i < CODE.length ; i++){
			    str += "<tr id = '"+CODE_NM[i]+"'>";	
				str += "<td><a href = '#this' style='color: red;'>"+CODE[i]+"</a></td>";
				str += "<td>"+CODE_NM[i]+"</td></tr>";
			}
			
			$("#"+obj.parent().parent().attr("id")).after(str);
			chk = obj.parent().parent().attr("id");
		}*/
				
	}

	function returnParent(CODE,CODE_NM){
		var returnValue = new Array();
		console.log(CODE);
		returnValue[0] = CODE;
		returnValue[1] = CODE_NM;
		
		window.opener.getReturnValue(returnValue);
		window.close();
	}
	
	function filter(){
		if($('#txtFilter').val()=="")
			$("#TBody tr").css('display','');
		else{
			$("#TBody tr").css('display','none');
			
			$("#TBody tr[id*='"+$('#txtFilter').val()+"']").css('display','');
		}
		return false;
	}
</script>
</html>