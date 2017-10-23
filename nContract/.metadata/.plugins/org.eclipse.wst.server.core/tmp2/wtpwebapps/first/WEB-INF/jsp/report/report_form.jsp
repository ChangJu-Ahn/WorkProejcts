<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf" %>
<style type="text/css">

.axis path, .axis line {
	fill: none;
	stroke: #ccc;
	shape-rendering: crispEdges;
}
/*
.axis text {
	font-size: 15px;
	font-family: dotum;
} */
.textlabel {
	font-size: 15px;
	font-weight: bold;
	font-family: dotum;
}

#Grid table {
	border-collapse: collapse;
	border: 2px black solid;
}

#Grid table th {
	border: 1px black solid;
	padding: 5px;
	text-align: center;
	background: aqua;
}

#Grid table tbody tr {
	padding: 5px;
	text-align: center;
}

#Grid table tbody td {
	border: 1px black solid;
	padding: 5px;
	text-align: center;
	font-size: 12px;
}

#Grid table tfoot td {
	padding: 5px;
	text-align: center;
	font-weight: bold;
	font-size: 12px;
}

.ui-datepicker-calendar {
	display: none;
}

.ui-datepicker-month {
	display: none;
}

.ui-datepicker-prev {
	display: none;
}

.ui-datepicker-next {
	display: none;
}

</style>
</head>
<body>	
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>	
	<sec:authorize access="hasRole('ROLE_ADMIN')" var="ROLE_ADMIN"/>
	
	<div style="left: 20px ;top: auto;">
		보고서 : 
		<select id = "gubun">
				<option value="">---------선 택---------</option>
				<option value="A">종합현황</option>
				<option value="B">기간별 계약 현황</option>
				<c:if test="${ROLE_ADMIN == true}">
					<option value="C">사업부별 체결 현황(종합)</option>
					<option value="D">기간별 현황(종합)</option>
				</c:if>
		</select>
		<br>
		<div id = "calendar" style="display: none;"> 
		날짜 :
    		<input name="year" id="year" class="date-picker" style="width:80px;"/>
    	</div>
    	<div id = "sort_group" style="display: none;"> 
		구분 :
    		<select id ="group_gubun">
    			<option value="simple">대분류</option>
    			<option value="detail">소분류</option>
    		</select>
    	</div>
    	<div id = "div_btn" style="display: none;">
			<button id="btn">조회</button>
		</div>
	</div>
	
	<div class ="report_Layer">
		<h2 id = "title"></h2>		
		<span id="Grid"></span> 
		<span id="Graph"></span>
	</div>

	<script type="text/javascript">
	$(function(){
		var userid = "<sec:authentication property='principal.username'/>";
		
		$('#gubun').change(function(){
	        
			init_control();
			
			var gubun = $('#gubun option:selected').val();
	        var title = $('#gubun option:selected').html();

	        var value = { val: gubun , userid : userid};
	        
	        if(gubun == "A"){
	        	$("#calendar").css("display","inline"); //16.09.05 종합현황도 날짜별로 볼 수 있도록 추가
	        	$("#sort_group").css("display","none");
	        	$("#div_btn").css("display","inline");  //16.09.05 종합현황도 날짜별로 볼 수 있도록 추가
	        	//fn_NtotalContract(title,value);
	        }else if(gubun == "B"){
	        	$("#calendar").css("display","inline");
	        	$("#sort_group").css("display","none");
	        	$("#div_btn").css("display","inline");
	        }else if(gubun == "C"){
	        	$("#calendar").css("display","none");
	        	$("#sort_group").css("display","inline"); //16.07.15, ahncj : 대분류, 소분류를 추가하기 위해 추가
	        	$("#div_btn").css("display","inline");
	        	//fn_AtotalContract(title,value);
	        }else if(gubun == "D"){
	        	$("#calendar").css("display","inline");
	        	$("#sort_group").css("display","none"); //16.07.15, ahncj : 대분류, 소분류를 추가하기 위해 추가
	        	$("#div_btn").css("display","inline");
	        }
	        else{
	        	$("#calendar").css("display","none");
	        	$("#sort_group").css("display","none"); //16.07.15, ahncj : 대분류, 소분류를 추가하기 위해 추가
	        	$("#div_btn").css("display","none");
	        }
	    });
		
		$('#btn').on('click',function(){
				 
			//컨트롤 초기화
			init_control();
			
	        var gubun = $('#gubun option:selected').val();
	        var title = $('#gubun option:selected').html();
	        var date = $("input[name='year']").val();
	        var group_gubun = $('#group_gubun option:selected').val();
	        
			//검색조건 확인
			if(gubun != "C"){
				if($('#year').val().length == 0) {
					//alert("날짜는 필수조건 입니다."); 
		    		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
		      	    		 {
		      	    		     msg: "날짜는 필수조건 입니다."
		      	    		 });
		    		return;
				}
			}
	        
	        console.log(date);
	        
	        var value = { val : gubun, year : date, userid : userid, G_gubun : group_gubun };
	        
	        switch(gubun){
	        	case "A":
	        		fn_NtotalContract(title,value);
	        		break;
	        	case "B":
	        		fn_NperiodContract(title,value);
	        		break;
	        	case "C":
	        		fn_AtotalContract(title,value);
	        		break;
	        	case "D":
	        		fn_AperiodContract(title,value);
	        		break;
	        }
	    });
		
	    $('.date-picker').datepicker( {
	        changeYear: true,
	        showButtonPanel: true,
	        dateFormat: 'yy',
	        onClose: function(dateText, inst) { 
	            var year = $("#ui-datepicker-div .ui-datepicker-year :selected").val();
	            $(this).datepicker('setDate', new Date(year, 1));
	        }
	    });
	});
	
	function init_control(){
		$("#title").html("");
        $("#Graph").html("");
        $("#Grid").html("");
	}
	     		
	</script>
</body>
</html>