  var a = b = c = d = e = 0;
  var time = addr = "";
  var surveyFlag = surveyFlag;
  var fav = true;
  var FavUrl = FavTitle = lc = Host = HostUrl = "";
  var Rate = 0;
  var DialogWidth = 300;
  var DialogHeight = 300;
  var first = "";

	function SetFirst(sName)
	{
		if (first == "") first = sName;
	}

  function fnAddToFavs()
  {
    c++;
    SetFirst('3');
    fav = false;
    window.external.AddFavorite(FavUrl, FavTitle);
  }

  function DoPopup()
  {
    try
    {
		  if (surveyFlag == null || surveyFlag == false)
		  {
			  if (Math.random() <= Rate)
			  {  
				  if (GetCookie("cars") == null)
				  {  
					  SetCookie("cars", "1");

					  var retVal = window.showModalDialog("/library/mnp/2/aspx/surveypopup.aspx?locale=" + lc,"","dialogHeight:" + DialogHeight + "px; dialogWidth:" + DialogWidth + "px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; scroll: No; status: No; unadorned: Yes; ");
					  if (retVal != null)
					  {
						  if (retVal != false)
						  {
							  time = retVal.substring(0, retVal.indexOf(";"));
							  addr =  retVal.substring(retVal.indexOf(";") + 1, retVal.length);
						  }
					  }
				  }
			  }
		  }
    }
    catch (e)
    {
      //noop
    }
  }

  function GetCookie(sName)
  {
    //cookies are separated by semicolons
    var aCookie = document.cookie.split("; ");
    CookieLength = aCookie.length;
    for (var i=0; i < CookieLength; i++)
    {
        // a name/value pair (a crumb) is separated by an equal sign
        var aCrumb = aCookie[i].split("=");
        if (sName == aCrumb[0]) return unescape(aCrumb[1]);
    }
    // a cookie with the requested name does not exist
    return null;
  }


  function SetCookie(sName, sValue)
  {
    var date = new Date();
    date.setDate(date.getDate() + 90);
    document.cookie = sName + "=" + escape(sValue) + "; expires=" + date.toGMTString() + "; path=/; domain=microsoft.com";
  }

  window.onunload = UnLoad;

  function UnLoad()
  {
    if (a > 0 || b > 0 || c > 0 || d > 0 || e > 0)
    {
      if (addr == "")
		ListenerFrame.navigate('/library/mnp/2/aspx/listener.aspx?s=' + encodeURIComponent(Host + '^' + HostUrl + '^' + a + '^' + b + '^' + c + '^' + d + "^" + e + "^" + time + "^" + addr + "^" + FavTitle + "^" + lc + "^" + FavUrl + "^" + first));
	  else
	    window.showModalDialog('/library/mnp/2/aspx/listener.aspx?s=' + encodeURIComponent(Host + '^' + HostUrl + '^' + a + '^' + b + '^' + c + '^' + d + "^" + e + "^" + time + "^" + addr + "^" + FavTitle + "^" + lc + "^" + FavUrl + "^" + first),"","dialogHeight:" + DialogHeight + "px; dialogWidth:" + DialogWidth + "px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: Yes; scroll: Yes; status: No; unadorned: Yes; ");
    }
  }
  
  window.onbeforeunload = BeforeUnload;

  function BeforeUnload()
  {
    if (a > 0 || b > 0 || c > 0 || d > 0 || e > 0)
    {
      if (fav)
	      DoPopup();
      else
	      fav = true;
    }
  }

  window.onbeforeprint = BeforePrint;

  function BeforePrint()
  {
    SetFirst('1');
    a++;
  }
function adjustIFrameSize (iframeWindow) {

    ratingHeight = document.getElementById("footerId");
    if(ratingHeight)ratingHeight.height = null;

if (iframeWindow.document.height) {var iframeElement = document.getElementById(iframeWindow.name);iframeElement.style.height = iframeWindow.document.height + 17 + 'px';}else if (document.all) {var iframeElement = document.all[iframeWindow.name];if (iframeWindow.document.compatMode && iframeWindow.document.compatMode != 'BackCompat') {iframeElement.style.height = iframeWindow.document.documentElement.scrollHeight + 5 + 'px';} else {iframeElement.style.height = iframeWindow.document.body.scrollHeight + 5 + 'px';}}

}
