Attribute VB_Name = "vbXCopy"
'Written by Ormond Stock Mar 2001.  This version updated July 2002.
'E&OE. Not fully tested. Comments, corrections, improvements please to ormond@osa.co.nz.
Option Explicit
Global SlashA As Boolean
Global SlashD As Boolean
Global SlashE As Boolean
Global SlashEX As Boolean
Global SlashH As Boolean
Global SlashK As Boolean
Global SlashM As Boolean
Global SlashR As Boolean
Global SlashS As Boolean
Global SlashT As Boolean
Global SlashU As Boolean
Global SlashY As Boolean
Global CopyDate As Date
Global ExcludePattern As String
Global IncludePattern As String
Global Const ATTR_DIRECTORY = 16
Public CopyError As String
Function CopyFiles(srcPath1 As String, dstPath1 As String) As Integer
Dim DirReturn As String
Dim CurrFile$
Dim CurrDir$
Dim dstPathBackup As String
Dim ResetArchive As Boolean
Dim FileExists As Boolean
Dim ThisOneReadOnly As Boolean
Dim ThisOneHidden As Boolean
Dim ThisOneSystem As Boolean
Dim ThisOneArchive As Boolean
Dim AttributesThisOne As Integer
Dim ThisFileDate As Date
Dim FirstPos As Integer
Dim f%

'On Error GoTo DirErr
CurrDir$ = CurDir$
CopyFiles = -1
'Find all files
DirReturn = Dir(srcPath1 & IncludePattern, 7)
Do While DirReturn <> ""
    ' Establish attributes of this file
    AttributesThisOne = GetAttr(srcPath1 & DirReturn)

    If ((AttributesThisOne And vbReadOnly) <> 0) Then
        ThisOneReadOnly = True
        Else
        ThisOneReadOnly = False
    End If
    If ((AttributesThisOne And vbHidden) <> 0) Then
        ThisOneHidden = True
        Else
        ThisOneHidden = False
    End If
    If ((AttributesThisOne And vbSystem) <> 0) Then
        ThisOneSystem = True
        Else
        ThisOneSystem = False
    End If
    If ((AttributesThisOne And vbArchive) <> 0) Then
        ThisOneArchive = True
        Else
        ThisOneArchive = False
    End If

    'Make sure we don't get a dir
    If Not ((AttributesThisOne And ATTR_DIRECTORY) = ATTR_DIRECTORY) Then
        'It's a file.  Establish whether it exists already.
        On Error Resume Next
        f% = FreeFile
        Open dstPath1 & DirReturn For Input As #f%
        Close #f%
        If Err = 0 Then
            FileExists = True
        Else
            FileExists = False
        End If
        
        'Check if /A or /M is set and archive attribute isn't
        If (SlashA Or SlashM) Then
            If Not ThisOneArchive Then GoTo BogOut
        End If
        
        'Check if /M means archive bit is to be reset.
        If SlashM Then
            ResetArchive = True
        Else
            ResetArchive = False
        End If
        
        'Check if hidden or system without /H
        If Not SlashH And (ThisOneSystem Or ThisOneHidden) Then
            GoTo BogOut
        End If
        
        'Check if /D is set
        If SlashD Then
            If Len(CopyDate) = 0 Then 'date wasn't entered, so only copy if newer than destination
                CopyDate = FileDateTime(dstPath1 & DirReturn)
            End If
            ThisFileDate = FileDateTime(srcPath1 & DirReturn)
            ThisFileDate = Left(Str(ThisFileDate), 11)
            If InStr(ThisFileDate, "/") = 2 Then
                ThisFileDate = "0" & ThisFileDate
            End If
            If ThisFileDate < CopyDate Then GoTo BogOut
        End If
        
        'Is this file one of the ones excluded by file pattern?
        If SlashEX And InStr(dstPath1 & DirReturn, ExcludePattern) > 0 Then
            GoTo BogOut
        End If
        
        'Are we required to only copy files which exist?
        If SlashU And Not FileExists Then
            GoTo BogOut
        End If
               
        'Does the file exist and we require prompting?
        If FileExists And Not SlashY Then
            f% = MsgBox("The file " & dstPath1 & DirReturn & " already exists. Do you wish to overwrite it?", 4 + 32 + 256)
            If f% <> 6 Then
                GoTo BogOut
            End If
        End If
        
        'If FileExists check if Read only and are we allowed to overcopy
        If FileExists Then
            If ((GetAttr(dstPath1 & DirReturn) And vbReadOnly) <> 0) And Not SlashR Then
                GoTo BogOut
            End If
        End If
        
        'If we've got this far we can copy!
        FileCopy srcPath1 & DirReturn, dstPath1 & DirReturn
        
        If ResetArchive Then  'reset the archive bit
                SetAttr (srcPath1 & DirReturn), AttributesThisOne - vbArchive
        End If
        
        'If /K is set and original file is ReadOnly then need to set new file as ReadOnly
        If SlashK And ThisOneReadOnly Then
                SetAttr (dstPath1 & DirReturn), GetAttr(dstPath1 & DirReturn) + vbReadOnly
        End If

End If  'End of loop
        
BogOut:
      DirReturn = Dir
Loop

ExitFunc:
    ChDir CurrDir$
    CopyFiles = 1
Exit Function

DirErr:
'   CopyFiles = False
Resume ExitFunc

End Function
Function ProcessDirectory(srcPath As String, dstPath As String) As String
Dim dcount As Integer
Dim d() As String
Dim Result1 As Integer
Dim Result2 As Integer
Dim DirReturn As String
Dim chkPath As String
Dim newPath As String
Dim NewDirName As String
Dim DirEmpty As Integer
Dim i As Integer

On Error GoTo ProcessError

'First, copy all the files in this directory if /T is not set.
If Not SlashT Then
    Result1 = CopyFiles(srcPath, dstPath)
End If

'If /E or /S or /T then we have to do subdirectories.
If SlashE Or SlashS Or SlashT Then
    'Count number of subdirectories (dCount)
    DirReturn = Dir(srcPath & "*.*", ATTR_DIRECTORY)
    Do While DirReturn <> ""
    'Make sure we bypass "." and ".."
    If DirReturn <> "." And DirReturn <> ".." Then
         If (GetAttr(srcPath & DirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY Then
             ' It's a dir. Add it to dirlist
            dcount = dcount + 1
        End If
    End If
    DirReturn = Dir
    Loop

    'Now set size of array to cope with all directories
    'Initialize var to hold filenames
    ReDim d(dcount + 1) As String
    dcount = 0

    DirReturn = Dir(srcPath & "*.*", ATTR_DIRECTORY)
   
    'Find all subdirs and add names to d()
    Do While DirReturn <> ""
        ' Make sure we don't do anything with "." and "..", they aren't really files
        If DirReturn <> "." And DirReturn <> ".." Then
            If (GetAttr(srcPath & DirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY Then
                ' It's a dir. Add it to dirlist
                dcount = dcount + 1
                d(dcount) = srcPath & DirReturn
            End If
        End If
        DirReturn = Dir
    Loop
   
' Now do all subs
For i = 1 To dcount
    ' Check if Directory is empty
    NewDirName = Right$(d(i), Len(d(i)) - Len(srcPath)) & "\"
    chkPath = srcPath & NewDirName
    newPath = dstPath & NewDirName
    DirEmpty = CheckDirEmpty(chkPath)
    If DirEmpty = 0 Then 'Directory is empty
        If SlashE Or SlashT Then
            'Make directory anyway
            MkDir newPath
        End If
    Else
    'Directory is not empty, make directory and process it using this routine again.
    MkDir newPath
    Result2 = ProcessDirectory(chkPath, newPath)
    End If
Next
End If
ProcessDirectory = 1
GoTo ExitFunction
ProcessError:
ProcessDirectory = -1

ExitFunction:
End Function
Function CheckDirEmpty(PathToBeChecked)
Dim fcount As Integer
Dim DirReturn As String
fcount = 0
    'Count number of subdirectories (dCount)
    DirReturn = Dir(PathToBeChecked & "*.*", 23)
    Do While DirReturn <> ""
    'Make sure we bypass "." and ".."
    If DirReturn <> "." And DirReturn <> ".." Then
         fcount = fcount + 1
    End If
    DirReturn = Dir
    Loop
CheckDirEmpty = fcount
End Function
