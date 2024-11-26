function removeBreaks(){
	var nopara = document.getElementById("noparagraphs").checked;
	var noBreaksText = document.getElementById("oldText").value;
	noBreaksText = noBreaksText.replace(/(\r\n|\n|\r)/gm,"\r");
	if(nopara == 1 || nopara ==  true){//Remove ALL line breaks
		noBreaksText = noBreaksText.replace(/\r/gm," ");
		noBreaksText = noBreaksText.replace(/\s+/gm,' ').trim();
	}else{//Remove some line breaks (preserve paragraphs)
		noBreaksText = noBreaksText.replace(/(\r\r)+/gm,"\n\n");
		noBreaksText = noBreaksText.replace(/\r/gm," ");
		noBreaksText = noBreaksText.replace(/ \n/gm,"\n");
		noBreaksText = noBreaksText.replace(/\n /gm,"\n");
		noBreaksText = noBreaksText.replace(/  +/gm," ").trim();	
	}
	document.getElementById("newText").value = noBreaksText;
}

//reset text areas
document.getElementById("clearText").addEventListener("click", function () {
	document.getElementById("oldText").value = "";
	document.getElementById("newText").value = "";
	document.getElementById("oldText").focus();
});

var textBox = document.getElementById("newText");
	textBox.onfocus = function() {
	textBox.select();
	textBox.onmouseup = function() {
		textBox.onmouseup = null;
		return false;
	};
};	