VBScript delete old files
=========================

This script will scan a folder and find files that are old (based on your command line options) and delete them.

There are 4 command line arguments with the following options:

1. Time time interval to use when calculating  the age of the file.
  - yyyy = Year
  - q = Quarter
  - m = Month
  - d = Day
  - w = Weekday
  - ww = Week of year
  - h = Hour
  - n = Minute
  - s = Second
  - More information about valid time intervals is available at http://msdn.microsoft.com/en-us/library/xhtyw595%28v=vs.84%29.aspx
2. The minimum age (in the interval selected with the first argument) a file should be in order to delete it.
  - Uses the time interval from the first argument.
  - Must be a whole integer.
3. File extension to delete.
  - Can use * to indicate that all files, regardless of extension, should be checked for deletion.
4. Directory path to search in.
  - Wrap the directory in quotes if it contains spaces.
  - Include the trailing backslash.

The following example will look for all files with a "log" extension in "c:\my example folder\" that are older than 5 hours and delete them.

**delete-old-files.vbs h 5 log "c:\my example folder\"**

This script does not iterate through sub-folders. If you'd like to contribute, that would be a great place to start. Please add the command line option to allow for iteration through sub-folders and add the code that will delete files in sub-folders.


