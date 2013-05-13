$(document).ready(function() {
	$('html').bind('keydown', function(e) {
		var code = (e.keyCode ? e.keyCode : e.which);
		var hours = 9;
		var mins = 22;
		
		console.log(code);
		
		switch(code){
			case 37:
				//if(mins == 0) mins = 60;
				mins -= 1;
				break;
			case 38:
				//if(hours == 12) hours = 0;
				hours += 1;
				break;
			case 39:
				//if(mins == 60) mins = 0;
				mins += 1;				
				break;
			case 40:
				//if(hours == 0) hours = 12;
				hours -= 1;
				break;
			default:
				break;
		}

		var hdegree = hours * 30 + (mins / 2);
		var hrotate = "rotate(" + hdegree + "deg)";

		$("#hour").css({"-moz-transform" : hrotate, "-webkit-transform" : hrotate});

		var mdegree = mins * 6;
		var mrotate = "rotate(" + mdegree + "deg)";

		$("#min").css({"-moz-transform" : mrotate, "-webkit-transform" : mrotate});	 
	});
});