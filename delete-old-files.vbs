Option Explicit

Dim CommandLineArguments

Dim TimeInterval
Dim FileMinimumAge
Dim FileExtension
Dim SearchDirectory

Dim Fso
Dim FolderObject
Dim FileObjectCollection
Dim FileObject

'Make sure the correct number of arguemnts were passed in.
Set CommandLineArguments = Wscript.Arguments

If CommandLineArguments.Count <> 4 Then
  Wscript.echo "Invalid arguments, expecting 3:" & vbcrlf  & vbcrlf & _
		"1) Time time interval to use when calculating  the age of the file." & vbcrlf & _
		"   a) yyyy = Year" & vbcrlf & _
		"   b) q = Quarter" & vbcrlf & _
		"   c) m = Month" & vbcrlf & _
		"   d) d = Day" & vbcrlf & _
		"   e) w = Weekday" & vbcrlf & _
		"   f) ww = Week of year" & vbcrlf & _
		"   g) h = Hour" & vbcrlf & _
		"   h) n = Minute" & vbcrlf & _
		"   i) s = Second" & vbcrlf & _
		"   j) More information available at http://msdn.microsoft.com/en-us/library/xhtyw595%28v=vs.84%29.aspx" & vbcrlf & vbcrlf & _
		"2) The minimum age (in the interval selected with the first argument) a file should be in order to delete it." & vbcrlf & _
		"   a) Uses the time interval from the first argument." & vbcrlf & _
		"   b) Must be a whole integer." & vbcrlf & vbcrlf & _
		"3) File extension to delete." & vbcrlf & _
		"   a) Can use * to indicate all files, regardless of extension, should be checked for deletion." & vbcrlf & vbcrlf & _
		"4) Directory path to search in."  & vbcrlf & _ 
		"   a) Wrap the directory in quotes if it contains spaces." & vbcrlf & _
		"   b) Include the trailing backslash."
	Wscript.Quit 99
End If

'set the arguments to vars
TimeInterval = CommandLineArguments(0)
FileMinimumAge = CommandLineArguments(1)
FileExtension = CommandLineArguments(2)
SearchDirectory = CommandLineArguments(3)

'define the directory objects
Set Fso = CreateObject("Scripting.FileSystemObject")
Set FolderObject = Fso.GetFolder(SearchDirectory)
Set FileObjectCollection = FolderObject.Files

'Loop through all the files and find the ones that have expired.
For each FileObject in FileObjectCollection
	If DateDiff(TimeInterval, FileObject.DateLastModified, Now) > cInt(FileMinimumAge) And _
		(UCase(Fso.GetExtensionName(FileObject.Name)) = UCase(FileExtension) Or _
			FileExtension = "*") Then
			
			FileObject.Delete(True)
		
	End If

Next
