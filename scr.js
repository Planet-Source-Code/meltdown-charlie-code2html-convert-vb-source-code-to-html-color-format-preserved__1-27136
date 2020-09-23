function HiliteIt()
{
if (window.event.srcElement.className=="Normal") 						window.event.srcElement.className="Hover"
}

function UnHiliteIt()
{
if (window.event.srcElement.className=="Hover")
	window.event.srcElement.className="Normal"
}
document.onmouseover=HiliteIt;
document.onmouseout=UnHiliteIt;