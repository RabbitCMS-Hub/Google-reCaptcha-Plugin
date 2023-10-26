<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
'**********************************************
' Set rsTempClass = New GooglereCaptcha
' 	rsTempClass.class_register()
' Set rsTempClass = Nothing

' Set reCaptcha = New GooglereCaptcha
Class GooglereCaptcha
	Private gc_SITEKEY, gc_SECRETKEY, gc_ACTIVE, gc_VERSION
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Variables
	'---------------------------------------------------------------
	'*/
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME, PLUGIN_AUTOLOAD
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Variables
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' REQUIRED: Register Class
	'---------------------------------------------------------------
	'*/
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		'/*
		'---------------------------------------------------------------
		' Check Register
		'---------------------------------------------------------------
		'*/
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		'/*
		'---------------------------------------------------------------
		' Plugin Settings
		'---------------------------------------------------------------
		'*/
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE&"_")
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "ClassName")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "0")
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		'/*
		'---------------------------------------------------------------
		' Register Settings
		'---------------------------------------------------------------
		'*/
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Register Class End
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Settings Panel
	'---------------------------------------------------------------
	'*/
	Public sub LoadPanel()
		'/*
		'--------------------------------------------------------
		' Sub Page
		'--------------------------------------------------------
		'*/
		If Query.Data("Page") = "SHOW:CachedFiles" Then
			Call PluginPage("Header")

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If
		'/*
		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		'*/
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			' .Write "    <div class=""col-lg-6 col-sm-12"">"
			' .Write 			PLUGIN_PANEL_INPUT(This, "select", "OPTION_1", "Buraya Title", "0#Seçenek 1|1#Seçenek 2|2#Seçenek 3", TO_DB)
			' ' .Write 			QuickSettings("select", ""& PLUGIN_CODE &"_OPTION_1", "Buraya Title", "0#Seçenek 1|1#Seçenek 2|2#Seçenek 3", TO_DB)
			' .Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", ""& PLUGIN_CODE &"_SITEKEY", "Site Anahtarı", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", ""& PLUGIN_CODE &"_SECRETKEY", "Gizli Anahtar", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("select", ""& PLUGIN_CODE &"_VERSION", "Kullanılacak Versin", "V1#V1|V2#V2|V3#V3", TO_DB)
			.Write "    </div>"
			.Write "</div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""https://www.google.com/recaptcha/about/"" class=""btn btn-sm btn-primary"">"
			.Write "        	reCaptcha Hakkında"
			.Write "        </a>"
			.Write "        <a open-iframe href=""https://developers.google.com/recaptcha/"" class=""btn btn-sm btn-danger"">"
			.Write "        	reCaptcha Geliştirici Klavuzu"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Settings Panel
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Initialize
	'---------------------------------------------------------------
	'*/
	Private Sub class_initialize()
		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: PluginTemplate Main Variables
		'-----------------------------------------------------------------------------------
		'*/
    	PLUGIN_CODE  			= "GOOGLE_RECAPTCHA"
    	PLUGIN_NAME 			= "Google reCaptcha Plugin"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/Google-reCaptcha-Plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-google"
    	PLUGIN_CREDITS 			= "@badursun Anthony Burak DURSUN"
    	PLUGIN_FOLDER_NAME 		= "Google-reCaptcha-Plugin"
    	PLUGIN_DB_NAME 			= "plugin_google_recaptcha"
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_AUTOLOAD 		= True
    	PLUGIN_ROOT 			= PLUGIN_DIST_FOLDER_PATH(This)
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
		'/*
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
		'*/
		gc_ACTIVE 		= Cint(GetSettings("GOOGLE_RECAPTCHA_ACTIVE","0"))
		gc_SITEKEY 		= GetSettings("GOOGLE_RECAPTCHA_SITEKEY","")
		gc_SECRETKEY	= GetSettings("GOOGLE_RECAPTCHA_SECRETKEY","")
		gc_VERSION 		= GetSettings("GOOGLE_RECAPTCHA_VERSION","V3")
		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: Register Plugin to CMS
		'-----------------------------------------------------------------------------------
		'*/
		class_register()
		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: Hook Plugin to CMS Auto Load Location WEB|API|PANEL
		'-----------------------------------------------------------------------------------
		'*/
		If PLUGIN_AUTOLOAD_AT("WEB") = True Then 
			' Cms.BodyData = Init()
			' Cms.FooterData = "<add-footer-html>Hello World!</add-footer-html>"
		End If
	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Initialize
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Terminate
	'---------------------------------------------------------------
	'*/
	Private sub class_terminate()

	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Terminate
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Manager Exports
	'---------------------------------------------------------------
	'*/
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property
	Public Property Get PluginAutoload() 	: PluginAutoload = PLUGIN_AUTOLOAD 			: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable, PluginAutoload)
	End Property
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Manager Exports
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' Verify Chapta Key
	'---------------------------------------------------------------
	'*/
	Public Property Get Verify(vCode)
        Verify = True

	    If gc_ACTIVE = 0 Then Exit Property

        If Len(gc_SECRETKEY) < 10 Then
            Verify = false
            Exit Property
        End If

        If Len(vCode) < 50 Then
            Verify = false
            Exit Property
        End If

        dataParam = "remoteip="& IPAdresi() &"&secret="& gc_SECRETKEY &"&response="& vCode
        Set httpGoogle = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
            httpGoogle.open "GET", CorsProxy("https://www.google.com/recaptcha/api/siteverify?"& dataParam &""), false
            httpGoogle.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
        	httpGoogle.setRequestHeader "origin", CORS_ORIGIN
            httpGoogle.setRequestHeader "Content-Length", Len(dataParam)
            httpGoogle.setRequestHeader "Content-type", "text/json; charset=utf-8"
            httpGoogle.send
        strStatus = httpGoogle.Status
        strRetval = httpGoogle.responseText
        Set httpGoogle = nothing

        Set recaptchaJson = JSON.parse(join(array( strRetval )))
        If Err <> 0 Then
            Call CMSErrorHandler("recaptchaVerify() json Data Error <!--["& strRetval &"] ["& vCode &"] ["& dataParam &"]-->") 
            Verify = false
        Else
	        If recaptchaJson.success = True Then
	            Verify = true
	        Else
	            Verify = false
	        End If
        End If
        Set recaptchaJson = Nothing
	End Property
	'/*
	'---------------------------------------------------------------
	' Verify Chapta Key
	'---------------------------------------------------------------
	'*/


	'/*
	'---------------------------------------------------------------
	' Chapta Widget
	'---------------------------------------------------------------
	'*/
	Public Property Get Widget()
		Response.Write IIf(gc_VERSION="V2", "<div class=""g-recaptcha"" data-callback=""CMScorrectCaptcha"" data-sitekey="""& gc_SITEKEY &""" /></div>", "")
		Response.Write "<input data-added=""reCaptcha.Class"" type=""hidden"" name=""g-recaptcha-response"" value="""" />"
	End Property
	'/*
	'---------------------------------------------------------------
	' Chapta Widget
	'---------------------------------------------------------------
	'*/
End Class 

%>