'%RunPerInstance
'@ DESCRIPTION
'@ Macro recorded by Synergy on 21-Jan-2016 at 16:26:53
SetLocale("en-us")
Dim SynergyGetter, Synergy
On Error Resume Next
Set SynergyGetter = GetObject(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SAInstance%"))
On Error GoTo 0
If (Not IsEmpty(SynergyGetter)) Then
	Set Synergy = SynergyGetter.GetSASynergy
Else
	Set Synergy = CreateObject("synergy.Synergy")
End If



'Dim minV, MaxV, tStep, tStepShw
Synergy.SetUnits "Metric"
'Set PlotMgr = Synergy.PlotManager()
'Set Viewer = Synergy.Viewer()
'Set Plot = PlotMgr.FindPlotByName2 ("Fill time", "Fill time")
'Viewer.ShowPlot Plot

'minV = Plot.GetMinValue
'maxV = Plot.GetMaxValue
'tStep = ((maxV - minV) / 4)
'tStepShw = minV + tStep

'Plot.SetMaxValue tStepShw
'Plot.Regenerate

'Declare PowerPoint Related variables	
Dim objPPT, ObjPresentation, ObjSlide, objSumTbl
'Declare file system object related variables
Dim fso, repDir

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True
' Create a New Presentation
Set objPresentation = objPPT.Presentations.Open _
	("E:\Projects\Training\Moldflow\API\" & _
		"Report Automation\ReportTemplate.pptx")
Set PlotManager = Synergy.PlotManager()
Set Viewer = Synergy.Viewer()
Set PropEd = Synergy.PropertyEditor()
call createRepDir(repDir)
call fillSummaryTable(meltTemp)
call resultsMeltFrontAdvancement(repDir)
call resultsMeltFrontContour(repDir)
call resultsFlowFrontTemp(repDir, meltTemp)
call resultsXferPressure(repDir, xPress)
call resultsXYPressure(repDir, xPress)
call resultsClampForce(repDir)
call resultsShearRate(repDir)
call resultsWeldLines(repDir, meltTemp)
call resultsVolShrink2D(repDir)
Call resultsSinkMarkDepth(repDir)
	
Sub createRepDir(repDir)
	'This sub creates a directory for the report within the current analysis folder
	Set PlotManager = Synergy.PlotManager()
	Set Plot = PlotManager.FindPlotByName("Plastic flow")
	Set Viewer = Synergy.Viewer()
	Set Project = Synergy.Project()
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set StudyDoc = Synergy.StudyDoc()
	SDYName = StudyDoc.StudyName
	SDYLen = Len(SDYName)
	SDYTrimLen = SDYLen - 4
	SDYFldr = Left(SDYName, SDYTrimLen)
	repDir = Project.Path & SDYFldr	
	'Check if directory already exists
	If (fso.FolderExists(repDir)) Then
		'Do Nothing
	Else
		fso.CreateFolder(repDir)
	End If
End Sub	

Sub fillSummaryTable(meltTemp)
	'Declare variables used 
	Dim injMethod, injTime, packMethod, packTime, packPressure, moldTemp, coolTime, shearRate
	Dim packArr, packSteps
	'Set the summary table slide
	Set objSlide = objPresentation.Slides(2)
	'Set the summary table shape
	Set objSumTbl = objSlide.Shapes("shpSumTbl")
	'Set object functions for synergy property editor
	set PropEd = Synergy.PropertyEditor()
	'set property to work with based on the tcodeset 30011 representing thermoplastic injection molding
	set Prop = PropEd.FindProperty(30011, 1)
	'Determine filling control method used in SDY file based on tcode 10109 Filling Control
	injMethod = Prop.FieldValues(10109).Val(0)
	if (injMethod = 1) Then
		msgBox "The filling control is set to automatic this" & _
			" is not allowed with a runner system. " & _
				"The report generator will now quit."
		WScript.Quit(0)
	elseif (injMethod = 2) Then
		'Enter the fill time used in the summary table
		injTime = Prop.FieldValues(10100).Val(0)
		objSumTbl.Table.Cell(2,2).Shape.TextFrame.TextRange.Text = injTime
	End If
	'Input the melt temperature in to the summary table using tcode 11002 in °F
	Synergy.SetUnits "English"
	meltTemp = Round(Prop.FieldValues(11002).Val(0))
	objSumTbl.Table.Cell(3,2).Shape.TextFrame.TextRange.Text = meltTemp
	'Input the mold temperature in to the summary table using tcode 11108 in °F
	objSumTbl.Table.Cell(4,2).Shape.TextFrame.TextRange.Text = Round(Prop.FieldValues(11108).Val(0))
	'Input the transfer pct in to the summary table using tcode 10308 for transfer control by %Volume
	objSumTbl.Table.Cell(5,2).Shape.TextFrame.TextRange.Text = Prop.FieldValues(10308).Val(0) & "%"
	'Determine packing control method used in SDY file based on tcode 10310
	'Value 4 is %Filling presure vs time and is tcode10702
	'Value 2 is packing pressure vs time and is tcode 10707
	packMethod = Prop.FieldValues(10704).Val(0)
	if (packMethod = 4) Then
		Set packArr = Prop.FieldValues(10702)
		packSteps = packArr.Size
		if (packSteps = 4) Then
			i = 2
			while i < packSteps
				'Pack time vs pressure is input as an array moldflow default example
				'Time Pressure
				'0     80
				'10    80
				'The first 0 indicates the time the solver will take to reduce from v/p pressure to 80% pressure
				'This time does not count as packing. Therefore start adding time from second time step
				'which is position 2 in the array.
				packTime = packTime + packArr.Val(i)
				i = i + 2
			Wend
			if (packArr.Val(1) = packArr.Val(3)) Then
				packPressure = packArr.Val(1)
				objSumTbl.Table.Cell(6,2).Shape.TextFrame. _ 
					TextRange.Text = packPressure & "% Fill Pressure/" & packTime & "s"
			Else
					objSumTbl.Table.Cell(6,2).Shape.TextFrame. _ 
						TextRange.Text = "Profiled/" & packTime & "s"
				
					objSumTbl.Table.Cell(6,3).Shape.TextFrame. _ 
						TextRange.Text = "see page details"
				
			end if
		end if
	end If
		

	'Input the cooling time using tcode 10102
	coolTime = Prop.FieldValues(10102).Val(0)
	objSumTbl.Table.Cell(7,2).Shape.TextFrame.TextRange.Text = 	coolTime
	'Input the total cycle time Fill + Pack + Cool
	objSumTbl.Table.Cell(8,2).Shape.TextFrame.TextRange.Text = (injTime + packTime + coolTime)
	
End Sub

Sub resultsMeltFrontAdvancement(repDir)
	'Declare variables if needed
	Dim img25, img50, img75, img100, fillTime
	'Set the slide for short shots
	Set objSlide = ObjPresentation.Slides(4)
	'Set Synergy Objects
	'Size the viewer for optimal image in PPT
	'Viewer uses pixels in width by height pixels are defined by the monitor being used. 
	'optimal pixels determined for HP LA2405wg
	Viewer.SetViewSize 884, 495
	'Define the Fill time plot
	Set Plot = PlotManager.FindPlotByName("Fill time")
	Viewer.ShowPlot Plot
	fillTime = Plot.GetMaxValue
	img25 = "\25pct.jpg"
	img50 = "\50pct.jpg"
	img75 = "\75pct.jpg"
	img100 = "\100pct.jpg"
	'Show and insert 25% filled
	Viewer.ShowPlotFrame Plot, 5
	'Wscript.Sleep 1000
	Viewer.SaveImage repDir & img25
	'Add picture reference https://msdn.microsoft.com/en-us/library/office/ff745953(v=office.14).aspx
	ObjSlide.Shapes.AddPicture repDir & img25, true, true, 36, 127, 288, 173
	'Show and insert 50% filled
	Viewer.ShowPlotFrame Plot, 11
	'Wscript.Sleep 1000
	Viewer.Saveimage repDir & img50
	ObjSlide.Shapes.AddPicture repDir & img50, true, true, 396, 127, 288, 173
	'Show and insert 75% Filled
	Viewer.ShowPlotFrame Plot, 17
	'Wscript.Sleep 1000
	Viewer.Saveimage repDir & img75
	ObjSlide.Shapes.AddPicture repDir & img75, true, true, 36, 325, 288, 173
	'Show and insert 100% Filled
	Viewer.ShowPlotFrame Plot, 23
	'Wscript.Sleep 1000
	Viewer.Saveimage repDir & img100
	ObjSlide.Shapes.AddPicture repDir & img100, true, true, 396, 325, 288, 173	
	'Input data to summaryTable
	Set objSlide = objPresentation.Slides(2)
	Set objSumTbl = objSlide.Shapes("shpSumTbl")
	objSumTbl.Table.Cell(2,3).Shape.TextFrame.TextRange.Text = FormatNumber(fillTime,2)
	
	
End Sub

Sub resultsMeltFrontContour(repDir)
	Dim img
	'Set working slide
	Set objSlide = ObjPresentation.Slides(5)
	'Show Contour Plot Assuming it exists
	Set Plot = PlotManager.FindPlotByName("Plastic flow")
	img = repDir & "\plastic_flow.jpg"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	ObjSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324

End Sub

Sub resultsFlowFrontTemp(repDir, meltTemp)
	dim img, minV, MaxV
	set objSlide = objPresentation.Slides(6)
	Synergy.SetUnits "English"
	Set Plot = PlotManager.FindPlotByName("Temperature at flow front")
	img = repDir & "\Flow_front_Temp.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	minV = Round(Plot.GetMinValue)
	MaxV = Round(Plot.GetMaxValue)
	if ((meltTemp - minV) < 40) Then
		objSlide.Shapes("shpFFTDesc").TextFrame.TextRange = "Flow front temperature in the main body of the part is " _
			& "uniform and meets the ideal target temperature drop of less than 20C/40F. Larger temperature drops " _
				& "will lead to flow marks or visible gloss variation in the part." & vbNewLine & _
					"Rises in temperature indicates shear effects inducing heating and large drops indicate the process injection " _
						& "rates may be too low"
	Else
		objSlide.Shapes("shpFFTDesc").TextFrame.TextRange = "Flow front temperature in the main body of the part is " _
			& "uniform and is outside of the ideal target temperature drop of less than 20C/40F. Larger temperature drops " _
				& "will lead to flow marks or visible gloss variation in the part." & vbNewLine & _
					"Rises in temperature indicates shear effects inducing heating and large drops indicate the process injection " _
						& "rates may be too low"		
	End if
	'Insert information on summary page	
	Set objSlide = objPresentation.Slides(2)
	Set objSumTbl = objSlide.Shapes("shpSumTbl")
	objSumTbl.Table.Cell(3,3).Shape.TextFrame.TextRange.Text = minV & " - " & maxV
End Sub

Sub resultsXferPressure(repDir, xPress)
	dim img, minV, MaxV
	set objSlide = objPresentation.Slides(7)
	Synergy.SetUnits "English"
	Set Plot = PlotManager.FindPlotByName("Pressure at V/P switchover")
	img = repDir & "\XferPress.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	maxV = FormatNumber(Round(plot.GetMaxValue), 0)
	
	If (maxV < 10000) Then
		objSlide.Shapes("shpPressDesc").TextFrame.TextRange = "Pressure at V/P switch is " & maxV & " psi." & vbNewLine _
			& "The pressure meets the general simulation target of less than 70 MPa/10,000 psi." & vbNewLine _
				& "will lead to flow marks or visible gloss variation in the part." & vbNewLine & _
					"This simulation does not account for pressure loss in the machine. Typical machines are capable of 18,000 psi or higher."
	Else
		objSlide.Shapes("shpPressDesc").TextFrame.TextRange = "Pressure at V/P switch is " & maxV & " psi." & vbNewLine _
			& "The pressure is outside of the general simulation target of less than 70 MPa/10,000 psi." & vbNewLine _
				& "will lead to flow marks or visible gloss variation in the part." & vbNewLine & _
					"This simulation does not account for pressure loss in the machine. Typical machines are capable of 18,000 psi or higher."
	End If
	'Insert information on summary page
	Set objSlide = objPresentation.Slides(2)
	Set objSumTbl = objSlide.Shapes("shpSumTbl")
	objSumTbl.Table.Cell(10,3).Shape.TextFrame.TextRange.Text = maxV
	xPress = maxV
End Sub

Sub resultsXYPressure(repDir, xPress)
	dim img, minV, MaxV
	set objSlide = objPresentation.Slides(8)
	Synergy.SetUnits "English"
	Set Plot = PlotManager.FindPlotByName("Pressure at injection location:XY Plot")
	img = repDir & "\XYPress.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	maxV = FormatNumber(Round(plot.GetMaxValue), 0)
	If(maxV < xPress) Then
		objSlide.Shapes("shpPeakPress").TextFrame.TextRange = "Peak pressure occurs at transfer and is " & maxV & " psi."
		objSlide.Shapes("shpPeakPRess").ZOrder msoSendToFront
	Else
		objSlide.Shapes("shpPeakPress").TextFrame.TextRange = "Peak pressure occurs before transfer due to gate sequencing and is " & maxV & " psi."
		objSlide.Shapes("shpPeakPRess").ZOrder msoSendToFront
	End If
	objSlide.Shapes("shpHoldPress").ZOrder msoSendToFront
	'Insert information on summary page
	Set objSlide = objPresentation.Slides(2)
	Set objSumTbl = objSlide.Shapes("shpSumTbl")
	'objSumTbl.Table.Cell(10,3).Shape.TextFrame.TextRange.Text = maxV
	'xPress = maxV
End Sub

Sub resultsClampForce(repDir)
	dim img, minV, MaxV, machClampSpec
	set objSlide = objPresentation.Slides(9)
	Synergy.SetUnits "English"
	Set Plot = PlotManager.FindPlotByName("Clamp force:XY Plot")
	img = repDir & "\XYClamp.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	maxV = FormatNumber(Round(plot.GetMaxValue), 0)
	machClampSpec = InputBox("Enter the machine clamp force in US tons. If clamp force is not provided enter '0'")
	machClampSpec = FormatNumber(machClampSpec,0)
	If (machClampSpec > 0) Then
		objSlide.Shapes("shpClamp").TextFrame.TextRange = "Peak clamp force is " & maxV &" US tons." _
			& vbNewLine & "The machine spec. is " & machClampSpec & " US tons."
		objSlide.Shapes("shpClamp").Zorder msoSendToFront
		Set objSlide = objPresentation.Slides(2)
		Set objSumTbl = objSlide.Shapes("shpSumTbl")
		objSumTbl.Table.Cell(12,3).Shape.TextFrame.TextRange.Text = maxV
		objSumTbl.Table.Cell(12,2).Shape.TextFrame.TextRange.Text = machClampSpec
		If (maxV > machClampSpec) Then
			objSumTbl.Table.Cell(12,4).Shape.TextFrame.TextRange.Text = ""
		End If
	Else
		objSlide.Shapes("shpClamp").TextFrame.TextRange = "Peak clamp force is " & maxV &" US tons." _
			& vbNewLine & "The machine spec. is unknown."
		objSlide.Shapes("shpClamp").Zorder msoSendToFront
		Set objSlide = objPresentation.Slides(2)
		Set objSumTbl = objSlide.Shapes("shpSumTbl")
		objSumTbl.Table.Cell(12,3).Shape.TextFrame.TextRange.Text = maxV
		objSumTbl.Table.Cell(12,2).Shape.TextFrame.TextRange.Text = "Unknown"
	End If
End Sub

Sub resultsShearRate(repDir)
	dim img, minV, maxV, matSpec
	set objSlide = objPresentation.Slides(10)
	Synergy.SetUnits "English"
	Set Plot = PlotManager.FindPlotByName("Shear rate, bulk")
	img = repDir & "\ShearRate.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	maxV = Round(Plot.GetMaxValue)
	call getMaterialData(matID, matSubID)
	Set Prop = PropEd.FindProperty(matID, matSubID)
	matSpec = Prop.FieldValues(1806).Val(0)
	If (maxV < matSpec) Then
		maxV = FormatNumber(maxV, 0)
		matSpec = FormatNumber(matSpec, 0)
		objSlide.Shapes("shpShearRateDesc").TextFrame.TextRange = "Maximum shear rate at the gate is " & maxV & " 1/sec." _
			& vbNewLine & "Shear rates are within the recommended material limit of " & matSpec & " 1/sec."
		objSlide.Shapes("shpShearRateDesc").Zorder msoSendToFront
		Set objSlide = objPresentation.Slides(2)
		Set objSumTbl = objSlide.Shapes("shpSumTbl")
		objSumTbl.Table.Cell(13,3).Shape.TextFrame.TextRange.Text = maxV
		objSumTbl.Table.Cell(13,2).Shape.TextFrame.TextRange.Text = matSpec	
	Else
		maxV = FormatNumber(maxV, 0)
		matSpec = FormatNumber(matSpec, 0)	
		objSlide.Shapes("shpShearRateDesc").TextFrame.TextRange = "Maximum shear rate at the gate is " & maxV & " 1/sec." _
			& vbNewLine & "Shear rates are outside the recommended material limit of " & matSpec & " 1/sec."
		objSlide.Shapes("shpShearRateDesc").Zorder msoSendToFront
		Set objSlide = objPresentation.Slides(2)
		Set objSumTbl = objSlide.Shapes("shpSumTbl")
		objSumTbl.Table.Cell(13,3).Shape.TextFrame.TextRange.Text = maxV
		objSumTbl.Table.Cell(13,2).Shape.TextFrame.TextRange.Text = matSpec
		objSumTbl.Table.Cell(13,4).Shape.TextFrame.TextRange.Text = ""
	End If

End Sub

Sub resultsWeldLines(repDir, meltTemp)
	dim img, minV, maxV, tolTemp
	set objSlide = objPresentation.Slides(11)
	Synergy.SetUnits "English"
	Set Plot = PlotManager.FindPlotByName("Weld lines")
	img = repDir & "\WeldLines.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	Plot.Regenerate
	minV = Plot.GetMinValue
	'msgBox minV
	tolTemp = meltTemp - 40
	objSlide.Shapes("shpWldLnDesc").TextFrame.TextRange = "Weld lines are shown with the temperatures at which they form." _
		& vbNewLine & "The process set melt temperature is " & meltTemp & "F." & vbNewLine & _
			"Cold welds occur at weld line locations with temperature drop of over 20C/40F and have the potential" _
				& " to be much more severe aesthetically and structurally."
	objSlide.Shapes("shpCldWld").Zorder msoSendToFront
	objSlide.Shapes("shpOval").Zorder msoSendToFront
	objSlide.Shapes("shpArrw").Zorder msoSendToFront


End Sub

Sub resultsAirTraps(repDir)
	'Place holder
	'needs message and pause for image manipulation
End Sub

Sub resultsVolShrink2D(repDir)
	dim img, minV, maxV, tolTemp
	set objSlide = objPresentation.Slides(13)
	Synergy.SetUnits "English"
	Set Plot = PlotManager.FindPlotByName("Volumetric shrinkage at ejection")
	img = repDir & "\VolShrink.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	objSlide.Shapes("shpCallout1").Zorder msoSendToFront
	objSlide.Shapes("shpCallout2").Zorder msoSendToFront
End Sub

Sub resultsSinkMarkDepth(repDir)
	dim img, minV, maxV, tolTemp
	set objSlide = objPresentation.Slides(14)
	Synergy.SetUnits "Metric"
	Set Plot = PlotManager.FindPlotByName("Sink marks, depth")
	img = repDir & "\sink.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	objSlide.Shapes("shpSinkCallOut").Zorder msoSendToFront
End Sub

Sub resultsTimeToReachEjectionTemp2D(repDir)
	dim img, minV, maxV, tolTemp
	set objSlide = objPresentation.Slides(16)
	Synergy.SetUnits "Metric"
	Set Plot = PlotManager.FindPlotByName("Time to reach ejection temperature")
	img = repDir & "\ejecttime.bmp"
	Viewer.SetViewSize 1075, 543
	Viewer.ShowPlot Plot
	Viewer.SaveImage img
	objSlide.Shapes.AddPicture img, true, true, 36, 168, 648, 324
	objSlide.Shapes("shpEjectCallout").Zorder msoSendToFront	
End Sub

Sub openRepDir(repDir)
	'After Completing report open the directory where everything has been saved.
	Set objShell = CreateObject("shell.application")
	objShell.Open(repDir)
End Sub


Sub getMaterialData(matID, matSubID)
' Get the relevent proccessing TSet data  from a study file
	MaterialID = 0
	MaterialSubID = 0
	' Find Appropriate Injection sets.
	Dim InjectionID, InjectionSubID, Injection2ID, Injection2SubID
	InjectionID = 40000
	InjectionSubID = 1
	Set Prop = PropEd.FindProperty(InjectionID, InjectionSubID)
	If Not Prop Is nothing Then
		Dim Field
		Field = Prop.GetFirstField()
		While Not Field = 0 And (MaterialID < 1)
			' Material reference Tcode
			If Field = 20020 Then
				Dim FieldValues
				Set FieldValues = Prop.FieldValues(Field)
				MaterialID = FieldValues.Val(0)
				MaterialSubID = FieldValues.Val(1)
			End If
			Field = Prop.GetNextField(Field)
		Wend
	End if
	matID = MaterialID
	matSubID = MaterialSubID


End Sub

