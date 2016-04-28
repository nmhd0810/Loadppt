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
	
	Dim objPPT, ObjPresentation, ObjSlide
	Set objPPT = CreateObject("PowerPoint.Application")
	objPPT.Visible = True
	' Create a New Presentation
	Set objPresentation = objPPT.Presentations.Open _
		("C:\Users\dhu\Desktop\TEST_C02539_FDJ_U554_LtchRd_JL7B78654A33A_R10.pptx")

	'Apply Ford Template
	'objPresentation
	'Set objPresentation = objPPT.Presentations.Active
	
	call fillSummaryTable(objPresentation, Synergy)
	'--------------------------------Start Di-------------------------------------------------------------
	ans=MsgBox("Are you satisfied with part position?",4,"Choose options")
	if vbYes Then
	Wscript.sleep.8000
	end if
	'----------------------------------End Di----------------------------------------------------------------------
	'Add Fill Time images
	call insFillPatternPlots(objPresentation, Synergy)
	


Sub fillSummaryTable(objPresentation, Synergy)
	'Declare variables used 
	Dim injMethod, injTime, packMethod, packTime, packPressure, meltTemp, moldTemp, coolTime, shearRate
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
	if injMethod = 1 Then
		msgBox "The filling control is set to automatic this" & _
			" is not allowed with a runner system. " & _
				"The report generator will now quit."
		WScript.Quit(0)
	elseif injMethod = 2 Then
		'Enter the fill time used in the summary table
		injTime = Prop.FieldValues(10100).Val(0)
		objSumTbl.Table.Cell(2,2).Shape.TextFrame.TextRange.Text = injTime
	End If
	'Input the melt temperature in to the summary table using tcode 11002 in °F
	Synergy.SetUnits "English"
	objSumTbl.Table.Cell(3,2).Shape.TextFrame.TextRange.Text = Round(Prop.FieldValues(11002).Val(0))
	'Input the melt temperature in to the summary table using tcode 11108 in °F
	objSumTbl.Table.Cell(4,2).Shape.TextFrame.TextRange.Text = Round(Prop.FieldValues(11108).Val(0))
	
	'Input the transfer pct in to the summary table using tcode 10308 for transfer control by %Volume
	objSumTbl.Table.Cell(5,2).Shape.TextFrame.TextRange.Text = Prop.FieldValues(10308).Val(0) & "%"
	'Determine packing control method used in SDY file based on tcode 10310'  '10310 is V/P switchover 0=Automatic 1=by volume%
	'Value 4 is %Filling presure vs time and is tcode10702
	'Value 2 is packing pressure vs time and is tcode 10707
	packMethod = Prop.FieldValues(10704).Val(0)
	'---------------------------------Start Di---------------------------------------------------
	if (packMethod = 4) Then
		Set packArr = Prop.FieldValues(10702)
		packSteps = packArr.Size
		if (packSteps = 4) Then
			i = 0
			packTime=0
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
		'profiled packing pressure using %volume with more than two lines
		if (packSteps>4) Then
			i = 0
			packTime=0
			while i < packSteps
			packTime = packTime + packArr.Val(i)
				i = i + 2
			Wend
			objSumTbl.Table.Cell(6,2).Shape.TextFrame. _ 
						TextRange.Text = "Profiled/" & packTime & "s"
				
					objSumTbl.Table.Cell(6,3).Shape.TextFrame. _ 
						TextRange.Text = "see page details"
		end if
	end If
'	----------------------------packing pressure vs time and is tcode 10707------------------
	if (packMethod = 2) Then
		Set packArr = Prop.FieldValues(10707)
		'Time Pressure(psi)
		'0     8000
		'10    8000
		packSteps = packArr.Size
		if (packSteps = 4) Then
			i = 0
			packTime=0
			while i < packSteps
			packTime = packTime + packArr.Val(i)
			i = i + 2
			Wend
			if (packArr.Val(1) = packArr.Val(3)) Then
				packPressure = packArr.Val(1)
				objSumTbl.Table.Cell(6,2).Shape.TextFrame. _ 
					TextRange.Text = packPressure & " psi/" & packTime & "s"
			Else
					objSumTbl.Table.Cell(6,2).Shape.TextFrame. _ 
						TextRange.Text = "Profiled/" & packTime & "s"
				
					objSumTbl.Table.Cell(6,3).Shape.TextFrame. _ 
						TextRange.Text = "see page details"
				
			end if
			'There are more than two lines
			if (packSteps>4) Then
			i = 0
			packTime=0
			while i < packSteps
			packTime = packTime + packArr.Val(i)
				i = i + 2
			Wend
			objSumTbl.Table.Cell(6,2).Shape.TextFrame. _ 
						TextRange.Text = "Profiled/" & packTime & "s"
				
					objSumTbl.Table.Cell(6,3).Shape.TextFrame. _ 
						TextRange.Text = "see page details"
			end if
		end if
	end if		
	
	

'--------------------------------------End Di-------------------------------------------------------------		

	'Input the cooling time using tcode 10102
	coolTime = Prop.FieldValues(10102).Val(0)
	objSumTbl.Table.Cell(7,2).Shape.TextFrame.TextRange.Text = 	coolTime
	'Input the total cycle time Fill + Pack + Cool
	objSumTbl.Table.Cell(8,2).Shape.TextFrame.TextRange.Text = (injTime + packTime + coolTime)
	
End Sub

Sub insFillPatternPlots(objPresentation, Synergy)
	'Declare variables if needed
	Dim imgDir, img25, img50, img75, img100
	'Set the slide for short shots
	Set objSlide = ObjPresentation.Slides(4)
	'Set Synergy Objects
	Set PlotManager = Synergy.PlotManager()
	Set Viewer = Synergy.Viewer()
	'Size the viewer for optimal image in PPT
	'Viewer uses pixels in width by height pixels are defined by the monitor being used. 
	'optimal pixels determined for HP LA2405wg
	Viewer.SetViewSize 884, 495
	'Define the Fill time plot
	Set Plot = PlotManager.FindPlotByName("Fill time")
	Viewer.ShowPlot Plot
	imgDir = "C:\Users\dhu\Desktop\pics"
	img25 = "\25pct.jpg"
	img50 = "\50pct.jpg"
	img75 = "\75pct.jpg"
	img100 = "\100pct.jpg"
	'Show and insert 25% filled
	Viewer.ShowPlotFrame Plot, 5
	'Wscript.Sleep 1000
	Viewer.SaveImage imgDir & img25
	'Add picture reference https://msdn.microsoft.com/en-us/library/office/ff745953(v=office.14).aspx
	ObjSlide.Shapes.AddPicture imgDir&img25, true, true, 36, 127, 288, 173
	'Show and insert 50% filled
	Viewer.ShowPlotFrame Plot, 11
	'Wscript.Sleep 1000
	Viewer.Saveimage imgDir & img50
	ObjSlide.Shapes.AddPicture imgDir & img50, true, true, 396, 127, 288, 173
	'Show and insert 75% Filled
	Viewer.ShowPlotFrame Plot, 17
	'Wscript.Sleep 1000
	Viewer.Saveimage imgDir & img75
	ObjSlide.Shapes.AddPicture imgDir & img75, true, true, 36, 325, 288, 173
	'Show and insert 100% Filled
	Viewer.ShowPlotFrame Plot, 23
	'Wscript.Sleep 1000
	Viewer.Saveimage imgDir & img100
	ObjSlide.Shapes.AddPicture imgDir & img100, true, true, 396, 325, 288, 173	
	
	
End Sub