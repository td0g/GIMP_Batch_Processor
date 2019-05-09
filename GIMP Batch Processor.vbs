'GIMP Batch Editor
'Written by Tyler Gerritsen 2017-02-10

'Performs adjustments described in Settings.txt file to selected photos
'To build an example Settings.txt file, just run the script
'All lines commented out with '#' are ignored

'Any photos dropped onto the .vbs file will also be processed.
'If no photos are dropped onto the .vbs file with the .txt files, then any photos in the same folder as .vbs file will automatically get processed.

'	General Syntax
'		The following procedures are recognized: txt, resize, add, copy, backup, and strings
'		These procedures must be followed by their parameters which are preceded by a backslash
'		Examples can be seen in the example GIMP Batch Processor Settings.txt file, which is automatically generated when a .txt file is not found
'		Custom procedures can also be used by entering the procedure's name (eg. gimp-curves-spline) followed by its required parameters

'	Adding text: first word in .txt file must be '\txt'
'		Colour: text colour defaults to black.  Write 'txt \w' to make text colour white. \w \r \g \b
'		Position: \x10 = 10% from left side of image, \y-20 = 20% from bottom of page (default=5%)
'		Size: \s10 = 1.0% image width (default=30 --> 3.0%)

'	Overlay Image: first word in .txt file must be '\add "imagename.extension"'
'		Position (Top Left Corner): \x10 = 10% from left side of image, \y-20 = 20% from bottom of page (default=5%)
'								Negative numbers will flip which side the measurement is taken
'								eg. \x-15 = right side of overlayed image is 15% from right side of original image
'		Size: \X10 = 1.0% image width, \Y25 = 2.5% image height (default=30 --> 3.0%)
'		Opacity: \o67 = 67% Layer Opacity

'	Other adjustments:	
'		eg. '\gimp-curves-spline DRAWABLE HISTOGRAM-VALUE 6 #(0 0 20 70 255 255)'
'		Don't populate the IMAGE or DRAWABLE variables; leave as shown above

'	Debugging:
'		Create an emtpy file named 'debug.txt' in the folder
'		The script will enter debugging mode and save the shell command to the 'debug.txt' file

'###################################################################################

			'Script Configuration

'###################################################################################

			'GIMP Version (optional - will find most recent version automatically)
			'gimpVersion = "2.10"
			
			'Install folder for GIMP
			pgmLoc = "C:\Program Files\GIMP 2\bin\"

			'Maximum Simultaneous Commands Strings Executed
			maxCommands = 3

			'Minimum photos per Command String
			imgPerCommand = 5

			'Maximum length of command to run in shell
			maxStringLength = 30000

			'List of All Usable File Extensions
			dim fileExtList
			fileExtList = split("jpg,jpeg,jpe,jif,jfif,jfi,tif,tiff,png,gif")

			'Files Not Recognized By WINDOWS as Images
			dim jpgExtList
			jpgExtList = split("jpe, jif, jfif, jfi")

			'Example Settings File
			exampleSettingsFile =   "###### ADDING TEXT ######" & vbNewLine & _
									vbNewLine & _
									"#txt \r \y5 \x5 \s20 Some text to go into your photo" & vbNewLine & _
									"#  Adds text to the photo, 20% up from bottom, 50% right of left side, text size is 30% image dimension" & vbNewLine & _
									vbNewLine & _
									"#txt \1 \r \y-10 \x10 \s30 pic one with red text!" & vbNewLine & _
									"#txt \3 pic three :)" & vbNewLine & _
									"#  Adds text to the photo, 10% down from top, 10% right of left side, text size is 30% image dimension, applies specific text to specific photos" & vbNewLine & _
									"#  The following colours can be selected: w, r, g, b" & vbNewLine & _
									vbNewLine & _
									"###### IMAGE SCALING ######" & vbNewLine & _
									vbNewLine & _
									"#resize 50%" & vbNewLine & _
									"#  Scale to 50% in both directions, MAINTAINS ASPECT RATIO" & vbNewLine & _
									vbNewLine & _
									"#resize 50% 100%" & vbNewLine & _
									"#  Scales long dimension by 50%, short dimension by 100%, DOES NOT MAINTAIN ASPECT RATIO" & vbNewLine & _
									vbNewLine & _
									"#resize \2 1000p 2000p" & vbNewLine & _
									"#  Scale image NUMBER 2: Long dimension to 1000 pixels, short dimension to 2000 pixels" & vbNewLine & _
									vbNewLine & _
									"#resize d 1000p" & vbNewLine & _
									"#  Scale long dimension to 1000 pixels, decreases size ONLY, MAINTAINS ASPECT RATIO" & vbNewLine & _
									vbNewLine & _
									"###### OTHER PROCEDURES ######" & vbNewLine & _
									vbNewLine & _
									"#add \X50 \Y50 ""T:\Projects\2016-11-25 GIMP\Working 2017-02-13\TC0827 2017-02-03 002.JPG"" " & vbNewLine & _
									"#  Overlays image into current image (UNDER CONSTRUCTION - DOES NOT WORK)" & vbNewLine & _
									vbNewLine & _
									"#copy  copySuffixName" & vbNewLine & _
									"#  Creates copy of image instead of overwriting original.  File name will be suffixed by parameter." & vbNewLine & _
									vbNewLine & _
									"#gimp-curves-spline DRAWABLE HISTOGRAM-VALUE 6 #(0 0 41 91 255 255) " & vbNewLine & _
									"#plug-in-lens-distortion 1 IMAGE DRAWABLE 0 0 -100 0 0 0" & vbNewLine & _
									"#  Raw GIMP Batch Commands" & vbNewLine & _
									"#  Adjusts curve" & vbNewLine & _
									vbNewLine & _
									"#backup" & vbNewLine & _
									"#backup backup output folder" & vbNewLine & _
									"#  Copies files to backup folder before proceeding" & vbNewLine & _
									vbNewLine & _
									"#strings 5" & vbNewLine & _
									"#  Number of simultaneous instances to run"
			
			
'###################################################################################

			'Changelog

'###################################################################################


'v1.0
	'2017-02-10
	'Functional

'v1.01
	'2017-04-20
	'Added list of file extensions to check for (fileExtList) and file types windows does not recognize as photos (jpgExtList)
		'These files will be renamed to the .jpg file type using powershell, and will not be returned to their original file name
	'Changed syntax of User Settings File Input - Now use # to comment out lines

'v1.02
	'2017-05-03
	'Added 'copy' command
	'Fixed bug with changing text colour

'v1.03
	'2017-05-30
	'Added resize \d switch to resize command
	'If no operations performed on image, will not open & resave

'v1.04
	'2017-06-27
	'No practical limit on number of photos
	
'v1.05
	'2017-09-25
	'Removed use of .conf file - just use a .txt file sitting in the folder or dropped with the photos
	
'v1.06
	'2018-06-12
	'If no settings file is present, script will create the example file automatically
	'Added backup command with \backup optionalBackupSubfolderName
	'Added strings command
	'GIMP Version selectable

'v1.07
	'2018-07-25
	'GIMP Version selection optional
	'Added new parameter to SCALE and TEXT commands: specific images can be selected for editing
	
'v1.08
	'2018-09-21
	'Animated GIF Output
	'Minor improvements
	
'v1.09
	'2019-05-09
	'Removed GIF Output
	'	(Too much overhead, output was poor quality, better tools avaiable)

			
'###################################################################################

			'Load Arguments (Photos & Configuration File)

'###################################################################################

'Start Timer
sTime = Timer
	
'Variables and Objects		
	'Shell Object
	sFolder = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(sFolder)
	Set files = folder.Files
	
	'Settings
	dim rawArgs(100)
	rawArgCount = 0
	debugMode = 0
	allArg = ""
	txtArg = ""
	fonts = false
	addLayer = false
	dimensions = false
	oCompleteSuffix = ""
	dim txt(10000)
	dim rszA(10000)
	dim rszB(10000)
	dim rszAPix(10000)
	dim rszBPix(10000)
	oFolderName = ""
	backupDir = ""

	'Configuration File
	configFileExists = false
	configFileOpen = false

	'Photos
	fileListSize = 10000
	dim fileList (10000, 2)									'input, output filepaths
	totalImageCount=0																						'Get image dimensions
		dim imageDim()

		
		
'###################################################################################

			'Prepare Windows Shell

'###################################################################################
dim wShRen
Set wShRen = WScript.CreateObject("WScript.Shell")	'Shell to execute command

'Get GIMP executable path
if len(gimpVersion) = 0 then
	Set pgmFolder = fso.GetFolder(pgmLoc)
	Set pgmFiles = pgmFolder.Files
	for each file In pgmFiles
		if inStr(file.name, "gimp-console-") > 0 and right(file.name, 4) = ".exe" then
			if isNumeric(mid(file.name,inStr(file.name, "gimp-")+13,1)) then
			pgmFullPath = file.path
			end if
		end if
	next
else
	pgmFullPath = pgmLoc & "gimp-console-" & gimpVersion & ".exe"
end if



'###################################################################################

			'Check incoming arguments for photos and config file

'###################################################################################
If WScript.Arguments.Count > 0 Then 																				'Script was started by dropping files onto it
	For Each Arg in Wscript.Arguments 
		if oFolderName = "" then																					'Get output folder path for future use
			oFolderName = left(Arg,inStrRev(Arg,"\"))
		end if
		if ucase(right(Arg,4)) = ".TXT" and ucase(right(Arg,9)) <> "DEBUG.TXT" then
									'Argument contains settings - Copy them into config file
			
			'Now that config file is open, add settings to config file
			set argFile = fso.OpenTextFile(Arg)																		'Opens .txt settings argument
			configFileExists = true
			do until argFile.AtEndOfStream																			'Read one line at a time
				newArg = argFile.ReadLine
				if newArg <> "" then																				'Check that line is not blank
					if left(newArg, 1) <> "#" then 																	'Check that line is not commented out
						rawArgs(configActionCount) = newArg
						rawArgCount = rawArgCount + 1
					end if
				end if
			loop
			argFile.close
			
		
		elseif UBound(Filter(fileExtList, lcase(right(Arg,len(Arg) - instrrev(Arg,"."))))) > -1 then 				'Argument contains phtos - add to photo list
			fileList(totalImageCount, 0) = Replace(arg, "\", "/")
			if UBound(Filter(jpgExtList, lcase(right(Arg,len(Arg) - instrrev(Arg,"."))))) > -1 then					'Filetype not recognized by Windows
				fileList(totalImageCount, 1) = left(fileList(totalImageCount, 0),inStrRev(fileList(totalImageCount, 0), ".")) & "jpg"
					'Powershell rename file!
				wShRen.run "powershell rename-item " & Chr(34) & Chr(34) & Chr(34) & fileList(totalImageCount, 0) & Chr(34) & Chr(34) & Chr(34) & " " & Chr(34) & Chr(34) & Chr(34) & fileList(totalImageCount, 1) & Chr(34) & Chr(34) & Chr(34), 1, True
			else
				fileList(totalImageCount, 1) = fileList(totalImageCount, 0)											'Filetype recognized by Windows - just add to list
			end if
			totalImageCount = totalImageCount + 1
		end if
	Next 
else 																												'Script was started by double-clicking
	For each folderIdx In files																						'Loop through all files in folder and find photos
		if ucase(right(folderIdx.name,4)) = ".JPG" then
			if totalImageCount = fileListSize then
				fileListSize = fileListSize + 1000
				redim preserve fileList(fileListSize, 2)
			end if
			fileList(totalImageCount, 1) = Replace(folderIdx.path, "\", "/")
			totalImageCount = totalImageCount + 1
		else if ucase(folderIdx.name) = "DEBUG.TXT" then debugMode = 1
		end if
	next
end if

if not configFileExists then																										'Script was started by double-clicking
	For each folderIdx In files																						'Loop through all files in folder and find photos
		if ucase(right(folderIdx.name,4)) = ".TXT" and ucase(folderIdx.name) <> "DEBUG.TXT" and not configFileExists then
			set argFile = fso.OpenTextFile(sFolder & folderIdx.name)																		'Opens .txt settings argument
			configFileExists = true
			do until argFile.AtEndOfStream																			'Read one line at a time
				newArg = argFile.ReadLine
				if newArg <> "" then																				'Check that line is not blank
					if left(newArg, 1) <> "#" then 																	'Check that line is not commented out
						rawArgs(rawArgCount) = newArg
						rawArgCount = rawArgCount + 1
					end if
				end if
			loop
		end if
	next
end if
	
	
	
'###################################################################################

			'Check that images and config file are loaded before proceeding

'###################################################################################




eString = ""
If not configFileExists Then																						'If no configuration file found, stop program
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	outFile=sFolder & "GIMP Batch Processor Settings.txt"
	Set objFile = objFSO.CreateTextFile(outFile,True)
	objFile.write exampleSettingsFile
	objFile.Close
	msgbox "No config file loaded - Example file created"
	Wscript.Quit
end if

if rawArgCount = 0 then
	msgbox "No commands found in config file"
	Wscript.Quit
end if

if totalImageCount = 0 then 																						'Make sure that we have images before proceeding
	msgbox "No photos found"
	Wscript.Quit
end if

For each folderIdx In files																							'Search for debug.txt
	if ucase(folderIdx.name) = "DEBUG.TXT" then 
		debugMode = 1
		exit for
	end if
next



'###################################################################################

			'Close config file and report number of images + actions loaded

'###################################################################################



if configFileOpen then 
	oConfigFile.close
	oString = rawArgCount & " batch actions loaded"
	oString = oString & vbCrLf & totalImageCount & " photos loaded"
	msgbox oString
end if


'###################################################################################

			'Parse User Settings File

'###################################################################################

				'Default Text Positions
				txtXPos = 5
				txtYPos = 5
				txtSize = 30
				applyToPhoto = 0

				'Default Image Positions (Top left, 100% image size) & Opacity (%)
				imgXPos = 0
				imgYPos = 0
				imgXSize = 100
				imgYSize = 0
				imgOpacity = 100

				'Default Resize Parameters
				rszAPixtemp = 0
				rszBPixtemp = 0
				rszAtemp = 0
				rszBtemp = 0
				rszDecrease = 0
				dim rszDim()
				redim rszDim(totalImageCount,2)

for i = 0 to rawArgCount-1
	currArgParse = split(rawArgs(i))
	
'First check which photos to apply these settings to
	applyToPhoto = 0
	for j = 1 to ubound(currArgParse)
		if len(currArgParse(j)) > 1 then
			if left(currArgParse(j),1) = "\" and isnumeric(right(left(currArgParse(j),2),1)) and isnumeric(right(currArgParse(j),1)) then
				applyToPhoto = Cint(right(currArgParse(j),len(currArgParse(j))-1))
			end if
		end if
	next
	
'Insert Text
	if left(rawArgs(i), 3) = "txt" then																			
		txtCurr = ""
		getDimensions
		fonts = true		'Load fonts on GIMP startup
		for j = 1 to ubound(currArgParse)
			if left(currArgParse(j), 1) = "\" then
				select case left(currArgParse(j),2)
					case "\w": fgColour = "'(255 255 255)"	'Add argument checks here
					case "\r": fgColour = "'(255 0 0)"	'Add argument checks here
					case "\g": fgColour = "'(0 255 0)"	'Add argument checks here
					case "\b": fgColour = "'(0 0 255)"	'Add argument checks here
					case "\x": txtXPos = Cint(right(currArgParse(j),len(currArgParse(j))-2))
					case "\y": txtYPos = Cint(right(currArgParse(j),len(currArgParse(j))-2))
					case "\s": txtSize = Cint(right(currArgParse(j),len(currArgParse(j))-2))
				end select																'Text to display
			end if
		next
		for j = 1 to ubound(currArgParse)
			if left(currArgParse(j),1) <> "\" then
				txtCurr = txtCurr & currArgParse(j) & " "
			end if
		next
		if len(txtCurr) > 0 then txtCurr = left(txtCurr, len(txtCurr) - 1)	'Get rid of last space

		if applyToPhoto > 0 then
			txt(applyToPhoto - 1) = txtCurr
		else
			for j = 0 to totalImageCount - 1
				if len(txt(j)) = 0 then txt(j) = txtCurr
			next
		end if

'Overlaying Image
	elseif left(rawArgs(i), 3) = "add" then
		getDimensions
		addLayer = true
		newImgPath = mid(rawArgs(i),inStr(rawArgs(i),"""")+1,inStr(inStr(rawArgs(i),"""")+1,rawArgs(i),"""")-inStr(rawArgs(i),"""")-1)
		set fsoB = CreateObject("Scripting.FileSystemObject")
		if fsoB.FileExists(sFolder & newImgPath) then newImgPath = sFolder & newImgPath
			dim imageDimOL()
			redim imageDimOL(2)
	
				
		for j = 1 to ubound(currArgParse)
			commandLength = len(currArgParse(j))
			select case left(currArgParse(j),2)
				case "\x": imgXPos = Cint(right(currArgParse(j),len(currArgParse(j))-2))
				case "\y": imgYPos = Cint(right(currArgParse(j),len(currArgParse(j))-2))
				case "\X": imgXSize = Cint(right(currArgParse(j),len(currArgParse(j))-2))
				case "\Y": imgYSize = Cint(right(currArgParse(j),len(currArgParse(j))-2))
				case "\o": imgOpacity = Cint(right(currArgParse(j),len(currArgParse(j))-2))
			end select
		next

'Scaling
	elseif left(rawArgs(i), 6) = "resize" then
		getDimensions
		rszATemp = 0
		rszAPixTemp = 0
		rszBTemp = 0
		rszBPixTemp = 0
		for j = 1 to uBound(currArgParse)
			currArgParse(j) = replace(currArgParse(j),"\","")
			if lcase(currArgParse(j)) = "d" then																	'Decrease size only
			  rszDecrease = 1
			elseif rszAtemp = 0 and rszAPixtemp = 0 then
				if right(currArgParse(j),1) = "%" then 																'Size in percent
					rszAtemp = CInt(left(currArgParse(j),len(currArgParse(j))-1))
				elseif right(currArgParse(j),1) = "p" then 																							'Size in px
					rszAPixtemp = CInt(left(currArgParse(j),len(currArgParse(j))-1))
				end if
			else
				currArgParse(j) = replace(currArgParse(2),"\","")
				if right(currArgParse(j),1) = "%" then 
					rszBtemp = CInt(left(currArgParse(j),len(currArgParse(j))-1))
				else
					rszBPixtemp = CInt(currArgParse(j))
				end if
			end if
		next
		if rszBtemp = 0 then rszBtemp = rszAtemp
		
		if applyToPhoto > 0 then
			rszA(applyToPhoto - 1) = rszAtemp
			rszB(applyToPhoto - 1) = rszBtemp
			rszAPix(applyToPhoto - 1) = rszAPixtemp
			rszBPix(applyToPhoto - 1) = rszBPixtemp
		else
			for j = 0 to totalImageCount - 1
				if rszA(j) = 0 and rszB(j) = 0 and rszAPix(j) = 0 and rszBPix(j) = 0 then
					rszA(j) = rszAtemp
					rszB(j) = rszBtemp
					rszAPix(j) = rszAPixtemp
					rszBPix(j) = rszBPixtemp
				end if
			next
		end if


	elseif left(rawArgs(i), 4) = "copy" then 																		'Create Copies
		oCompleteSuffix = right(rawArgs(i),len(rawArgs(i)) - 5)
		
	elseif left(rawArgs(i), 7) = "strings" then
		maxCommands = CInt(currArgParse(1))
		
	elseif left(rawArgs(i),6) = "backup" then 		
		backupDir = ""
		if ubound(currArgParse) > 1 then
			for j = 1 to ubound(currArgParse)
					backupDir = backupDir & currArgParse(j) & " "																	'Text to display
			next	
			backupDir = left(backupDir,len(backupDir)-1)
		else
			backupDir = "backup"
		end if
		backupDir = oFolderName & backupDir
		if not fso.FolderExists(backupDir) then fso.createfolder backupDir 'Create Backups in \backup folder
		
	else 																											'Raw Command
		allArg = allArg & " -b " & Chr(34) & "(" & rawArgs(i) & ")" & Chr(34) & " "
	end if
next




'###################################################################################

			'Prepare & Execute Shell Command

'###################################################################################



dim oCmdString()
redim oCmdString(maxCommands)
oCmd = Chr(34) & pgmFullPath & Chr(34) & " "
B = " -b " & Chr(34)

if fonts = true then
	oCmd = oCmd & " -d --verbose --batch-interpreter plug-in-script-fu-eval " 		'Load fonts
else
	oCmd = oCmd & " -d -f --verbose --batch-interpreter plug-in-script-fu-eval "	'Don't load fonts
end if

if fgColour <> "" then oCmd = oCmd & B & "(gimp-context-set-foreground " & fgColour & ")" & Chr(34)	'Swap colors


'Universal Variables
commandStringNumber = 0
oCmdString(commandStringNumber) = oCmd
GIMPimage = 1
GIMPlayer = 2

'Calculate the number of simultaneous threads and images/thread
'NOTE: This could use some optimization
if totalImageCount > imgPerCommand then 
	imgPerCommand = int(totalImageCount / maxCommands)+1
else
	imgPerCommand = totalImageCount
end if

'Prepare strings for each thread
for i = 0 to totalImageCount - 1

	increment = false
	While not increment
		
		'GIMP expects forward slashes in file name
		iCompleteName = Replace(fileList(i, 1), "\", "/")
		iFileName = right(iCompleteName, len(iCompleteName) - inStrRev(iCompleteName,"/"))
		oCompleteName = Replace(fileList(i, 1), "\", "/")
		oCompleteName = left(oCompleteName, inStrRev(oCompleteName,".")-1) & oCompleteSuffix & right(oCompleteName, len(oCompleteName)-inStrRev(oCompleteName,".")+1)
		
		'Backup Images
		if backupDir <> "" then fso.CopyFile iCompleteName, backupDir & "\" & iFileName, True
		
		'Command to load image
		loadImage = B & "(gimp-file-load 1 \" & Chr(34) & iCompleteName & "\" & Chr(34) & " \" & Chr(34) & iCompleteName & "\" & Chr(34) & ")" & Chr(34) & " "
		
		'Populate IMAGE and DRAWABLE (layer) variables in raw GIMP commands
		oArg = replace(replace(allArg, "IMAGE", GIMPimage), "DRAWABLE",GIMPlayer)
		
		'Add image as new layer!
		if addLayer then		
			CURRimgXSize = round(imageDim(i,0) * imgXSize / 100)
			if imgYSize = 0 then	'Y size not defined, retain aspect ratio
				CURRimgYSize = round(CURRimgXSize *  imageDimOL(1) / imageDimOL(0))
			else
				CURRimgYSize = round(imageDim(i,1) * imgYSize / 100)
			end if
			if imgXPos > 0 then
				CURRimgXPos = imageDim(i,0) * imgXPos / 100
			else
				CURRimgXPos = imageDim(i,0) - imageDim(i,0) * -imgXPos / 100 - CURRimgXSize
			end if
			if imgYPos > 0 then
				CURRimgYPos = imageDim(i,1) * imgYPos / 100
			else
				CURRimgYPos = imageDim(i,1) - imageDim(i,1) * -imgYPos / 100 - CURRimgYSize
			end if
			
			overlayArg = overlayArg & B & "(gimp-file-load-layer 1 " & GIMPimage &" \" & Chr(34) & Replace(newImgPath, "\", "/") & "\" & Chr(34) & ")" & Chr(34)
			GIMPlayer = GIMPlayer + 3
			overlayArg = overlayArg & B & "(gimp-image-insert-layer " & GIMPimage & " " & GIMPlayer & " 0 -1)" & Chr(34)
			overlayArg = overlayArg & B & "(gimp-layer-scale " & GIMPlayer & " " & CURRimgXSize & " " & CURRimgYSize & " FALSE)" & Chr(34)
			if imgOpacity < 100 then overlayArg = overlayArg & B & "(gimp-layer-set-opacity " & GIMPlayer & " " & imgOpacity & ")" & Chr(34)
			overlayArg = overlayArg & B & "(gimp-layer-set-offsets " & GIMPlayer & " " & CURRimgXPos & " " & CURRimgYPos & ")" & Chr(34)
			overlayArg = overlayArg & B & "(gimp-image-flatten " & GIMPimage & ")" & Chr(34)
			GIMPlayer = GIMPlayer + 1
		else
			overlayArg = ""
		end if

		'Add text!
		if len(txt(i)) > 0 then
			if txtXPos > 0 then
				txtXPix = round(imageDim(i,0) * txtXPos / 100)
			else
				txtXPix = round(imageDim(i,0) * (100 + txtXPos) / 100)
			end if
			if txtYPos > 0 then
				txtYPix = round(imageDim(i,1) * txtYPos / 100)
			else
				txtYPix = round(imageDim(i,1) * (100 + txtYPos) / 100)
			end if
			
			txtSizePix = round(txtSize * imageDim(i,0) / 1000)
			
			txtArg = B & "(gimp-text-fontname " & GIMPimage & " -1 " & txtXPix & " " & txtYPix & " \" & Chr(34) & txt(i) & "\" & Chr(34) & " 0 1 " & txtSizePix & " 0 \" & Chr(34) & "Sans" & "\" & Chr(34) & ")" & Chr(34) 
			txtArg = txtArg & B & "(gimp-image-flatten " & GIMPimage & ")" & Chr(34)
			GIMPlayer = GIMPlayer + 2
		else
			txtArg = ""		'No text to add
		end if 
		
		'Resize Image!
		if rszDecrease = 1 and rszXPix => imageDim(i, 0) and rszYPix => imageDim(i, 1) then 
			resizeArg = ""
		elseif rszA(i) > 0 then
			if imageDim(i,0) > imageDim(i,1) then
				rszXPix = rszA(i) * imageDim(i,0)/100
				rszYPix = rszB(i) * imageDim(i,1)/100
			else
				rszYPix = rszA(i) * imageDim(i,1)/100
				rszXPix = rszB(i) * imageDim(i,0)/100
			end if
			resizeArg = B & "(gimp-image-scale " & GIMPimage & " " & round(rszXPix) & " " & round(rszYPix) & ")" & Chr(34) & " "
		elseif rszAPix(i) > 0 then
			if imageDim(i,0) > imageDim(i,1) then
				rszXPix = rszAPix(i)
				rszYPix = rszBPix(i)
				if rszYPix = 0 then rszYPix = rszXPix / imageDim(i,0) * imageDim(i,1)
			else
				rszYPix = rszAPix(i)
				rszXPix = rszYPix / imageDim(i,1) * imageDim(i,0)
			end if
			resizeArg = B & "(gimp-image-scale " & GIMPimage & " " & round(rszXPix) & " " & round(rszYPix) & ")" & Chr(34) & " "
		else
			resizeArg = ""
		end if
		
		'Save & Close commands
		saveImage = B & "(gimp-file-save 1 " & GIMPimage & " " & GIMPlayer & " \" & Chr(34) & oCompleteName & "\" & Chr(34) & " \" & Chr(34) & oCompleteName & "\" & Chr(34) & ")" & Chr(34) & " "
		closeImage = B & "(gimp-image-delete " & GIMPimage & ")" & Chr(34) & " "


		'Construct entire Shell Command for current image
		if  len(oArg & overlayArg & txtArg & resizeArg) = 0 then
			increment = true
		elseif GIMPimage > imgPerCommand or len(oCmdString(commandStringNumber)) > maxStringLength then		 'New command thread if images/thread exceeds set value or thread length exceeds the character limit
			commandStringNumber = commandStringNumber + 1
			redim preserve oCmdString(commandStringNumber)
			oCmdString(commandStringNumber) = oCmd
			GIMPimage = 0
			GIMPlayer = 0
		else
			increment = true
			oCmdStringNew = loadImage		
			oCmdStringNew = oCmdStringNew & oArg
			oCmdStringNew = oCmdStringNew & overlayArg
			oCmdStringNew = oCmdStringNew & txtArg
			oCmdStringNew = oCmdStringNew & resizeArg
			oCmdStringNew = oCmdStringNew & saveImage
			oCmdStringNew = oCmdStringNew & closeImage
			oCmdString(commandStringNumber) = oCmdString(commandStringNumber) & oCmdStringNew
			'Iteration-Specific Variables
			GIMPlayer = GIMPlayer + 2
			GIMPimage = GIMPimage + 1
		end if
	Wend
Next

'Output thread strings to file for debugging
if debugMode then
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	outFile=sFolder & "debug.txt"
	Set objFile = objFSO.CreateTextFile(outFile,True)
	for i = 0 to commandStringNumber
		objFile.write "   Shell Thread " & i & vbCrLf
		objFile.write Replace(oCmdString(i), "-b", vbCrLf & "-b")
		objFile.write vbCrLf & vbCrLf
	next
	objFile.Close
end if


'Execute shell command threads
dim wsh()
redim wsh(commandStringNumber)

for i = 0 to commandStringNumber
	if debugMode = false or commandStringNumber > 2 then oCmdString(i) = oCmdString(i) & B & "(gimp-quit 0)" & Chr(34) 			'Close shell immediately if not debugging or more threads queued
	oCmdString(i) = oCmdString(i) & " & pause"
	Set wsh(i) = WScript.CreateObject("WScript.Shell")									'Shell to execute command
	if i mod maxCommands = maxCommands - 1 then							
		wsh(i).run oCmdString(i), 1, True												'Wait to finish
	else
		wsh(i).run oCmdString(i), 1, False												'Don't wait to finish
	end if
	'WScript.Echo TypeName(wsh(i))
next



'###################################################################################

			'Public Procedures

'###################################################################################

Sub getDimensions()
	if not dimensions then
	dimensions = true	
		redim imageDim(totalImageCount,2)
		set oShell = CreateObject("Shell.Application")
		set oFolder = oShell.Namespace(replace(left(fileList(0, 1), inStrRev(fileList(i, 1),"/")),"/","\"))
		for ii = 0 to totalImageCount-1
			set oFolderItem = oFolder.parsename(right(fileList(ii, 1), len(fileList(ii, 1))-inStrRev(fileList(ii, 1),"/")))
			oString = oFolder.getdetailsof(oFolderItem,31)
			oStringParse = split(oString)
			imageDim(ii,0) = CInt(right(oStringParse(0), len(oStringParse(0))-1))
			imageDim(ii,1) = CInt(left(oStringParse(2), len(oStringParse(2))-1))
		next
	end if
End Sub