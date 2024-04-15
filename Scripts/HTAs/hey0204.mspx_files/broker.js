// //
// Copyright 2009 comScore Networks. All rights reserved.

// Name: Microsoft-www
// Date: 2009-03-31


// Multiple script protection.
if (!window.SiteRecruit_Globals) {

// Create the configuration, globals, and constants namespaces.
var SiteRecruit_Config = new Object();
var SiteRecruit_Globals = new Object();
var SiteRecruit_Constants = new Object();

// Validation variables.
SiteRecruit_Globals.parseFlag = false;
SiteRecruit_Globals.empty = false;

// Browser information.
SiteRecruit_Constants.browser = new Object();
SiteRecruit_Constants.browser.internetExplorer = 'Microsoft Internet Explorer';
SiteRecruit_Constants.browser.mozilla = 'Netscape';
SiteRecruit_Constants.browser.opera = 'Opera';

// Check browser information.
SiteRecruit_Globals.browserName = navigator.appName; 
SiteRecruit_Globals.browserVersion = parseInt(navigator.appVersion);

// Initialize browser flags.
SiteRecruit_Globals.isInternetExplorer = false;
SiteRecruit_Globals.isMozilla = false;
SiteRecruit_Globals.isInternetExplorer7 = false;

// Check for Internet Explorer based browsers.
if (SiteRecruit_Globals.browserName == SiteRecruit_Constants.browser.internetExplorer)
{
    if (SiteRecruit_Globals.browserVersion > 3)
    {
        // Only 5.5 and above.
        var a = navigator.userAgent.toLowerCase();
        if (a.indexOf("msie 5.0") == -1 && a.indexOf("msie 5.0") == -1)
        {
            SiteRecruit_Globals.isInternetExplorer = true;
        }
        
        // Check for 7.
        if (a.indexOf("msie 7") != -1)
        {
            SiteRecruit_Globals.isInternetExplorer7 = true;
        }
    }
}

// Check for Mozilla based browsers.
if (SiteRecruit_Globals.browserName == SiteRecruit_Constants.browser.mozilla)
{
    if (SiteRecruit_Globals.browserVersion > 4)
    {
        SiteRecruit_Globals.isMozilla = true;
    }
}

// Since Opera 9.02, they no longer use 'Mozilla' in the browser name
if (SiteRecruit_Globals.browserName == SiteRecruit_Constants.browser.opera)
{
	SiteRecruit_Globals.isMozilla = true;	// treat the same as Mozilla
}



// Cookie lifetime.
SiteRecruit_Constants.cookieLifetimeType = new Object();
SiteRecruit_Constants.cookieLifetimeType.duration = 1;
SiteRecruit_Constants.cookieLifetimeType.expireDate = 2;
    
// Invitation type.
SiteRecruit_Constants.invitationType = new Object();
SiteRecruit_Constants.invitationType.standard = 0;
SiteRecruit_Constants.invitationType.email = 1;
SiteRecruit_Constants.invitationType.domainDeparture = 2;
SiteRecruit_Constants.invitationType.emailDomainDeparture = 3;
    
// Cookie type flags.
SiteRecruit_Constants.cookieType = new Object();
SiteRecruit_Constants.cookieType.alreadyAsked = 1;
SiteRecruit_Constants.cookieType.inProgress = 2;
SiteRecruit_Constants.cookieType.emailDomainDeparture = 3;

// Alignment types.
SiteRecruit_Constants.horizontalAlignment = new Object();
SiteRecruit_Constants.horizontalAlignment.left = 0;
SiteRecruit_Constants.horizontalAlignment.middle = 1;
SiteRecruit_Constants.horizontalAlignment.right = 2;
SiteRecruit_Constants.verticalAlignment = new Object();
SiteRecruit_Constants.verticalAlignment.top = 0;
SiteRecruit_Constants.verticalAlignment.middle = 1;
SiteRecruit_Constants.verticalAlignment.bottom = 2;

// Survey cookie configuration.
SiteRecruit_Config.cookieName = 'msresearch';
SiteRecruit_Config.cookieDomain = '.microsoft.com';
SiteRecruit_Config.cookiePath = '/';

SiteRecruit_Constants.cookieJoinChar = ':';

// Settings for cookie lifetime.
SiteRecruit_Config.cookieLifetimeType = 1;

    // Duration of the cookie in days.
    SiteRecruit_Config.cookieDuration = 90 * 86400000;

// Duration of the rapid cookie.
SiteRecruit_Config.rapidCookieDuration = 0 * 86400000;

// URL of the email domain dep server.
SiteRecruit_Config.listenerUrl = '';// //
// Copyright 2009 comScore Networks. All rights reserved.

// KeepAlive class definition.
function SiteRecruit_KeepAlive()
{
    // Time between page checks.
    this.keepAlivePollDelay = 1000;

    // Unique (well, sorta) ID for this page.
    this.id = Math.random();

    // Attach methods.
    this.attemptStart = KeepAlive_attemptStart;
    this.checkCookie = KeepAlive_checkCookie;
    this.inProgressCookieExists = KeepAlive_inProgressCookieExists;
    this.cookieExists = KeepAlive_cookieExists;

    // Check for either domain dep or email domain dep cookies and act accordingly.
    function KeepAlive_attemptStart()
    {
        if (this.inProgressCookieExists())
        {
            setInterval('SiteRecruit_Globals.keepAlive.checkCookie()', this.keepAlivePollDelay);
        }
        else
        {
            if (this.cookieExists(SiteRecruit_Constants.cookieType.emailDomainDeparture))
            {
                var c = document.cookie.toString();
                var index = c.indexOf(SiteRecruit_Config.cookieName + '=' + SiteRecruit_Constants.cookieType.emailDomainDeparture);
                var endc = c.length;
                c = c.substring(index, endc);
                var ind1 = c.indexOf(';');
                if (ind1 != -1) c = c.substring(0, ind1);        
                var ind2 = c.indexOf('=');
                c = c.substring(ind2 + 1);
                var values = c.split(':');
                if (values.length == 2)
                {
                    var url = SiteRecruit_Config.listenerUrl;
                    
                    // append ? to qs if needed
					url += (url.indexOf('?') == -1 ? '?' : '&');
					
                    url += 'action=log'
                        + '&user=' + values[1]
                        + '&location=' + escape(window.location);
                    
                    setTimeout("var i = new Image(); i.src = '" + url + "&' + (new Date()).getTime(); ", 5);
                }
            }
        }
    }
    
    // Check and update the cookie.
    function KeepAlive_checkCookie()
    {
        if (this.inProgressCookieExists())
        {
            var c = SiteRecruit_Config.cookieName + '=' + SiteRecruit_Constants.cookieType.inProgress
                + ':' + escape(document.location)
                + ':' + (new Date()).getTime()
                + ':' + this.id
                + '; path=' + SiteRecruit_Config.cookiePath;
            
            if (SiteRecruit_Config.cookieDomain != '')
            {
                c += '; domain=' + SiteRecruit_Config.cookieDomain;
            }
            
            document.cookie = c;
        }
    }

    // Return true if an in-progress cookie exists.
    function KeepAlive_inProgressCookieExists()
    {
        return this.cookieExists(SiteRecruit_Constants.cookieType.inProgress);
    }
    
    // Return true if a cookie of the given type exists.
    function KeepAlive_cookieExists(cookieType)
    {
        var c = SiteRecruit_Config.cookieName + '=' + cookieType;
        if (document.cookie.indexOf(c) != -1)
        {
            return true;
        }
        return false;
    }
}

// Create the KeepAlive if a suitable cookie exists.
SiteRecruit_Globals.keepAlive = new SiteRecruit_KeepAlive();
SiteRecruit_Globals.keepAlive.attemptStart();
// //
// Copyright 2009 comScore Networks. All rights reserved.

// Broker class definition.
function SiteRecruit_PageConfigurationBroker()
{
    // List of mappings, each contains [url, page, priority, prereqs]    
    this.m = [        
    ['//[\\w\\.-]+/athome/', '/library/svy/SiteRecruit_PageConfiguration_3331mt3.js', 0, null],
    ['//[\\w\\.-]+/atwork', '/library/svy/SiteRecruit_PageConfiguration_3331mt5.js', 0, null],
    ['//[\\w\\.-]+/australia/athome/', '/library/svy/SiteRecruit_PageConfiguration_p15466742-au-372.js', 0, null],
    ['//[\\w\\.-]+/australia/athome/default\\.mspx$', '/library/svy/SiteRecruit_PageConfiguration_p15466742-au-372-flashfix.js', 1, null],
    ['//[\\w\\.-]+/australia/business/', '/library/svy/SiteRecruit_PageConfiguration_p15466742-au-373.js', 0, null],
    ['//[\\w\\.-]+/australia/business/(default\\.aspx)?$', '/library/svy/SiteRecruit_PageConfiguration_p15466742-au-373-SB-FIXED.js', 1, null],
    ['//[\\w\\.-]+/australia/smallbusiness(/(?!sbrp)|$)', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Australia.js', 0, null],
    ['//[\\w\\.-]+/brasil/pequenasempresas/(?!msn/)', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Brazil.js', 0, null],
    ['//[\\w\\.-]+/business(/(?!executivecircle|appplat|bizinfra|coreinfra|process|performance|info|mobility|customerconnections|smb|peopleready|peopleready2|compliance|smb/en-my/|(.*tlp))|$)', '/library/svy/SiteRecruit_PageConfiguration_3331mt8.js', 0, null],
    ['//[\\w\\.-]+/business/peopleready/en-us/', '/library/svy/SiteRecruit_PageConfiguration_p32155930-299.js', 1, null],
    ['//[\\w\\.-]+/business/peopleready/en-us/.*ng=connect', '/library/svy/SiteRecruit_PageConfiguration_p32155930-771.js', 2, null],
    ['//[\\w\\.-]+/business/peopleready/en-us/.*ng=insights', '/library/svy/SiteRecruit_PageConfiguration_p32155930-768.js', 2, null],
    ['//[\\w\\.-]+/business/peopleready/en-us/.*ng=proof', '/library/svy/SiteRecruit_PageConfiguration_p32155930-770.js', 2, null],
    ['//[\\w\\.-]+/business/peopleready/en-us/.*ng=solutions', '/library/svy/SiteRecruit_PageConfiguration_p32155930-769.js', 2, null],
    ['//[\\w\\.-]+/canada/(smallbusiness|(smallbiz|smb)/(?!sbplus))', '/library/svy/SiteRecruit_PageConfiguration_2933mt_EnglishCanada.js', 0, null],
    ['//[\\w\\.-]+/china/(smallbusiness|smb)/.*\\.aspx', '/library/svy/SiteRecruit_PageConfiguration_2933mt_China-UTF8.js', 0, null],
    ['//[\\w\\.-]+/china/(smallbusiness|smb)/(?!.*\\.aspx)', '/library/svy/SiteRecruit_PageConfiguration_2933mt_China.js', 0, null],
    ['//[\\w\\.-]+/china/smb/default\\.aspx', '/library/svy/SiteRecruit_PageConfiguration_2933mt_China-SL.js', 0, null],
    ['//[\\w\\.-]+/de/de/(default\\.aspx)?$', '/library/svy/SiteRecruit_PageConfiguration_p15466742-Germany-HP.js', 0, null],
    ['//[\\w\\.-]+/downloads/(en/|.*?displaylang=en)', '/library/svy/SiteRecruit_PageConfiguration_3331mt13_NEW_751-753.js', 0, null],
    ['//[\\w\\.-]+/dynamics(/(?!dynamicsresearch.mspx|everyonegetsit)|$)', '/library/svy/SiteRecruit_PageConfiguration_3331mt14_NEW-750.js', 0, null],
    ['//[\\w\\.-]+/dynamics/asmartmove/default\\.mspx', '/library/svy/SiteRecruit_PageConfiguration_3331mt14-SL-fix_NEW-750.js', 3, null],
    ['//[\\w\\.-]+/dynamics/default\\.mspx$', '/library/svy/SiteRecruit_PageConfiguration_3331mt14_flashfix_NEW-750.js', 1, null],
    ['//[\\w\\.-]+/en/au/', '/library/svy/SiteRecruit_PageConfiguration_p15466742-AU-HP-369.js', 0, null],
    ['//[\\w\\.-]+/en/us/default\\.aspx', '/library/svy/SiteRecruit_PageConfiguration_p15394611-US-HP_NEW_248-748-749.js', 0, null],
    ['//[\\w\\.-]+/events', '/library/svy/SiteRecruit_PageConfiguration_3331mt16.js', 0, null],
    ['//[\\w\\.-]+/fr/fr/(default\\.aspx)?$', '/library/svy/SiteRecruit_PageConfiguration_p15466742-France-HP.js', 0, null],
    ['//[\\w\\.-]+/france/entrepreneur(/(?!plus)|$)', '/library/svy/SiteRecruit_PageConfiguration_2933mt_France.js', 0, null],
    ['//[\\w\\.-]+/france/windows/', '/library/svy/SiteRecruit_PageConfiguration_p15466742_21.js', 0, null],
    ['//[\\w\\.-]+/germany/kleinunternehmen(/(?!small-business-plus)|$)', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Germany.js', 0, null],
    ['//[\\w\\.-]+/germany/mittelstand/', '/library/svy/SiteRecruit_PageConfiguration_p11880082-German.js', 0, null],
    ['//[\\w\\.-]+/germany/server(/|$)', '/library/svy/SiteRecruit_PageConfiguration_DE-wss-p12038685.js', 0, null],
    ['//[\\w\\.-]+/germany/windows(/|$)', '/library/svy/SiteRecruit_PageConfiguration_DE-windows-p12038685.js', 0, null],
    ['//[\\w\\.-]+/italy/beit/', '/library/svy/SiteRecruit_PageConfiguration_p17637473_788.js', 1, null],
    ['//[\\w\\.-]+/italy/beit/($|default.aspx$|.*video=)', '/library/svy/SiteRecruit_PageConfiguration_p17637473_788-fix.js', 2, null],
    ['//[\\w\\.-]+/italy/pmi(/(?!plus)|$)', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Italy.js', 1, null],
    ['//[\\w\\.-]+/ja/jp/', '/library/svy/SiteRecruit_PageConfiguration_p15466742-Japan-HP.js', 0, null],
    ['//[\\w\\.-]+/japan/athome/', '/library/svy/SiteRecruit_PageConfiguration_JA-p15466742-athome.js', 0, null],
    ['//[\\w\\.-]+/japan/atwork/', '/library/svy/SiteRecruit_PageConfiguration_JA-p15466742-atwork.js', 0, null],
    ['//[\\w\\.-]+/japan/business/', '/library/svy/SiteRecruit_PageConfiguration_JA-p15466742-business.js', 0, null],
    ['//[\\w\\.-]+/japan/servers/', '/library/svy/SiteRecruit_PageConfiguration_JA-p15466742-servers.js', 0, null],
    ['//[\\w\\.-]+/japan/smallbiz((/)|(/((default)|(index))\\.((html?)|(aspx?)|(mspx))))?$', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Japan-Default.js', 1, null],
    ['//[\\w\\.-]+/japan/smallbiz', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Japan.js', 0, null],
    ['//[\\w\\.-]+/japan/technet/', '/library/svy/SiteRecruit_PageConfiguration_JA-p12038685-technet.js', 0, null],
    ['//[\\w\\.-]+/japan/users', '/library/svy/SiteRecruit_PageConfiguration_JA-p15466742-users.js', 0, null],
    ['//[\\w\\.-]+/japan/users/default\\.mspx$', '/library/svy/SiteRecruit_PageConfiguration_JA-p15466742-users-Flashfix.js', 1, null],
    ['//[\\w\\.-]+/japan/windows(/(?!(downloads/ie/au\\.mspx)|(downloads/ie/iedelete\\.mspx))|$)', '/library/svy/SiteRecruit_PageConfiguration_JA-p15466742-windows.js', 0, null],
    ['//[\\w\\.-]+/japan/windows/(digitallife|possibilities)/', '/library/svy/SiteRecruit_PageConfiguration_JA-p15466742-windows-digitallife.js', 1, null],
    ['//[\\w\\.-]+/learning((/)|(/default\\.asp)|(/default\\.mspx)|(/default\\.aspx))?$', '/library/svy/SiteRecruit_PageConfiguration_3331mt42.js', 0, null],
    ['//[\\w\\.-]+/licensing(/(?!(licensewise/product\\.aspx)|(licensewise/program\\.aspx)|(mla/select\\.aspx)))', '/library/svy/SiteRecruit_PageConfiguration_3331mt43.js', 0, null],
    ['//[\\w\\.-]+/mexico/pymes', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Mexico.js', 0, null],
    ['//[\\w\\.-]+/midsizebusiness', '/library/svy/SiteRecruit_PageConfiguration_3123mt.js', 0, null],
    ['//[\\w\\.-]+/netherlands/(ondernemers(/(?!plus)|$)|zeker_ondernemen|artikelen)', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Netherlands.js', 1, null],
    ['//[\\w\\.-]+/office/2007-rlt/en-us', '/library/svy/SiteRecruit_PageConfiguration_p32411396.js', 0, null],
    ['//[\\w\\.-]+/protect(/(?!computer/updates/bulletins)|$)', '/library/svy/SiteRecruit_PageConfiguration_3331mt4.js', 0, null],
    ['//[\\w\\.-]+/security', '/library/svy/SiteRecruit_PageConfiguration_3331mt49.js', 0, null],
    ['//[\\w\\.-]+/smallbusiness(/(?!small-business-plus)|$)', '/library/svy/SiteRecruit_PageConfiguration_2933mt_UnitedStates.js', 0, null],
    ['//[\\w\\.-]+/spain/empresas', '/library/svy/SiteRecruit_PageConfiguration_2933mt_Spain.js', 1, null],
    ['//[\\w\\.-]+/(sql|sqlserver)', '/library/svy/SiteRecruit_PageConfiguration_3331mt52.js', 0, null],
    ['//[\\w\\.-]+/sql/experience/(Default\\.aspx\\?loc=en)|(html/Default\\.aspx\\?loc=en)|(html/Events\\.aspx\\?loc=en)|(LearnSQL\\.aspx\\?h=t&loc=en)|(LearnSQL\\.aspx\\?loc=en)|(Events\\.aspx\\?loc=en)|(.*\\.wmv)', '/library/svy/SiteRecruit_PageConfiguration_blank.js', 2, null],
    ['//[\\w\\.-]+/technet/(?!mnp_utility\\.mspx/(framesmenu|quicksearch|masthead)\\?url)', '/library/svy/SiteRecruit_PageConfiguration_p15808382-p26386365.js', 0, null],
    ['//[\\w\\.-]+/technet/scriptcenter/', '/library/svy/SiteRecruit_PageConfiguration_p15808382-p26386365-TIER3.js', 1, null],
    ['//[\\w\\.-]+/technet/security/', '/library/svy/SiteRecruit_PageConfiguration_p15808382-p26386365-TIER2.js', 1, null],
    ['//[\\w\\.-]+/technet/(.*/subscriptions|support|community)/', '/library/svy/SiteRecruit_PageConfiguration_p15808382mt-technet.js', 1, null],
    ['//[\\w\\.-]+/uc/voipasyouare/default\\.((html)|(aspx)|(mspx))', '/library/svy/SiteRecruit_PageConfiguration_p12529714-Flash-VoIP.js', 2, null],
    ['//[\\w\\.-]+/uk/smallbusiness/', '/library/svy/SiteRecruit_PageConfiguration_3300mt-uk.js', 0, null],
    ['//[\\w\\.-]+/video/', '/library/svy/SiteRecruit_PageConfiguration_p23275586.js', 0, null],
    ['//[\\w\\.-]+/visualstudio/en-us/default\\.(mspx|htm)', '/library/svy/SiteRecruit_PageConfiguration_4075mt-781.js', 0, null],
    ['//[\\w\\.-]+/visualstudio/en-us/howtobuy/default\\.(mspx|htm)', '/library/svy/SiteRecruit_PageConfiguration_4075mt-784.js', 0, null],
    ['//[\\w\\.-]+/visualstudio/en-us/products/teamsystem/default\\.(mspx|htm)', '/library/svy/SiteRecruit_PageConfiguration_4075mt-782.js', 0, null],
    ['//[\\w\\.-]+/visualstudio/en-us/try/default\\.(mspx|htm)', '/library/svy/SiteRecruit_PageConfiguration_4075mt-783.js', 0, null],
    ['//(sr-www|wwwstaging|www\\.microsoft)[\\w\\.-]+/windows/', '/library/svy/SiteRecruit_PageConfiguration_p25328149.js', 0, null],
    ['//[\\w\\.-]+/windows/buy/windows-laptop-scout\\.aspx$', '/library/svy/SiteRecruit_PageConfiguration_p25328149-Buy-WLS.js', 0, null],
    ['//[\\w\\.-]+/windows/compatibility/', '/library/svy/SiteRecruit_PageConfiguration_p27207601.js', 1, null],
    ['//[\\w\\.-]+/windows/downloads/', '/library/svy/SiteRecruit_PageConfiguration_p25328149-downloads.js', 1, null],
    ['//[\\w\\.-]+/windows/downloads/ie/getitnow\\.mspx', '/library/svy/SiteRecruit_PageConfiguration_3331mt62-p25328149.js', 2, null],
    ['//[\\w\\.-]+/windows/internet-explorer/', '/library/svy/SiteRecruit_PageConfiguration_3331mt62-p25328149.js', 1, null],
    ['//[\\w\\.-]+/windows/internet-explorer/beta/videos\\.aspx$', '/library/svy/SiteRecruit_PageConfiguration_3331mt62-p25328149_SL-FIX.js', 2, null],
    ['//[\\w\\.-]+/windows/possibilities/', '/library/svy/SiteRecruit_PageConfiguration_p25328149_SL-FIX.js', 1, null],
    ['//[\\w\\.-]+/windows/products/winfamily/ie(/|$)', '/library/svy/SiteRecruit_PageConfiguration_3331mt62-p25328149.js', 1, null],
    ['//[\\w\\.-]+/windows/windows-laptop-scout/(default\\.aspx)?$', '/library/svy/SiteRecruit_PageConfiguration_p25328149_laptop-scout_SL-FIX.js', 1, null],
    ['//[\\w\\.-]+/windows/windows-vista(/|$)', '/library/svy/SiteRecruit_PageConfiguration_3331mt64-p25328149.js', 1, null],
    ['//[\\w\\.-]+/windows/windows-vista/discover/', '/library/svy/SiteRecruit_PageConfiguration_3331mt64-p25328149_SL-FX.js', 2, null],
    ['//[\\w\\.-]+/windowsembedded/en-us/', '/library/svy/SiteRecruit_PageConfiguration_3331mt174.js', 1, null],
    ['//[\\w\\.-]+/windowsmobile', '/library/svy/SiteRecruit_PageConfiguration_3331mt173.js', 0, null],
    ['//[\\w\\.-]+/windowsmobile/en-us/totalaccess/', '/library/svy/SiteRecruit_PageConfiguration_p30393194_3331mt173.js', 1, null],
    ];
    
    // Attach methods.
    this.start = PageConfigurationBroker_start;
    this.getConfigurationForPage = PageConfigurationBroker_getConfigurationForPage;
    this.loadConfiguration = PageConfigurationBroker_loadConfiguration;

    // Start the Broker.
    function PageConfigurationBroker_start(url)
    {        
        // Decide if there are surveys to run for this page.
        var configuration = this.getConfigurationForPage(url);
        
        // If there is, load id up.
        if (configuration != null)
        {
            this.loadConfiguration(configuration);
        }
    }
    
    // Return the appropriate configuration for the given URL, if any.
    function PageConfigurationBroker_getConfigurationForPage(url)
    {
        
        var currentScore = 0;
        var currentMatch = -1;
        
        // Iterate over each URL.
        for (var i = 0; i < this.m.length; i++)
        {
            if (this.m[i] != null)
            {
                // Do the reg exp match.
                var r = new RegExp(this.m[i][0], 'i');
                if (url.toString().search(r) != -1)
                {
                    var prereqs = this.m[i][3];
                    var matchPrereqs = true;
                    
                    // Now check the prereqs.
                    if (prereqs != null)
                    {
                        for (var j = 0; j < prereqs.length; j++)
                        {
                            var p = prereqs[j];
                            if (p != null)
                            {
                                if (p.elementType)
                                {
                                    var matchContent = false;
                                    var matchAttribute = false;
                                    
                                    var elements = document.getElementsByTagName(p.elementType);
                                    for (var k = 0; k < elements.length; k++)
                                    {
                                        if (p.contentValue != '')
                                        {
                                            if (elements[k].innerHTML.search(p.contentValue) != -1)
                                            {
                                                matchContent = true;
                                            }
                                        }
                                        else
                                        {
                                            matchContent = true;
                                        }
                                        
                                        if (p.attributeName != '')
                                        {
                                            var a = elements[k].attributes.getNamedItem(p.attributeName);
                                            if (a != null)
                                            {
                                                if (p.attributeValue != '')
                                                {
                                                    if (a.value.search(p.attributeValue) != -1)
                                                    {
                                                        matchAttribute = true;
                                                    }
                                                }
                                                else
                                                {
                                                    matchAttribute = true;
                                                }
                                             }
                                        }
                                        else
                                        {
                                            matchAttribute = true;
                                        }
                                    }
                                    
                                    if (!matchContent || !matchAttribute)
                                    {
                                        matchPrereqs = false;
                                        break;
                                    }
                                }
                                else if (p.cookieName)
                                {
                                    var cookieTemplate = p.cookieName + '=';
                                    if (p.cookieValue)
                                    {
                                        cookieTemplate += p.cookieValue;
                                    }
                                    if (document.cookie.indexOf(cookieTemplate) == -1)
                                    {
                                        matchPrereqs = false;
                                        break;
                                    }
                                }
                                else if (p.language)
                                {
                                    if (navigator.language && navigator.language.toString().toLowerCase().indexOf(p.language) == -1)
                                    {
                                        matchPrereqs = false;
                                        break;
                                    }
                                    else if (navigator.userLanguage && navigator.userLanguage.toString().toLowerCase().indexOf(p.language) == -1)
                                    {
                                        matchPrereqs = false;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        matchPrereqs = true;
                    }
                    
                    if (matchPrereqs)
                    {
                        // Take the current if the score is equal or better.
                        var newScore = this.m[i][2];
                        if (newScore >= currentScore)
                        {
                            currentMatch = i;
                            currentScore = newScore;
                        }
                    }
                }
            }
        }
        
        // If there was a match, return the appropriate page configuration.
        var page = null;
        if (currentMatch >= 0)
        {
            page = this.m[currentMatch][1];
        }
        
                
        return page;
    }
    
    // Dynamically loads the associated configuration.
    function PageConfigurationBroker_loadConfiguration(configuration)
    {
                
        document.write('<script language="JavaScript" src="' + configuration + '"></script>');
    }
}

try
{
    // Only run if the browser is supported.
    if (SiteRecruit_Globals.isInternetExplorer || SiteRecruit_Globals.isMozilla)
    {
        // Create and start the Broker for the current location.
        SiteRecruit_Globals.broker = new SiteRecruit_PageConfigurationBroker();
        SiteRecruit_Globals.broker.start(window.location);
    }
}
catch (e)
{
    // Suppress any errors.
}

SiteRecruit_Globals.parseFlag = true;

// Multiple script protection.
}