'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Environment
Imports System.IO
Imports System.Linq

'This module contains this program's core procedures.
Public Module CoreModule
   'This structure defines a file's data.
   Private Structure FileStr
      Public FileName As String   'Defines a file's name.
      Public Offset As Integer    'Defines a file's offset.
   End Structure

   Private ReadOnly BYTES_TO_TEXT As Func(Of List(Of Byte), String) = Function(Bytes As List(Of Byte)) New String((From ByteO In Bytes Select ToChar(ByteO)).ToArray())   'This procedure converts the specified bytes to text.
   Private ReadOnly INVALID_CHARACTERS() As Char = {"*"c, "/"c, "<"c, ">"c, "?"c, "["c, "\"c, "]"c, "|"c, " "c}                                                           'Defines characters that are invalid in file names in MS-DOS.
   Private ReadOnly PADDING As Char = ToChar(&H0%)                                                                                                                        'Defines the character used to pad file names.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim SourceFile As String = GetCommandLineArgs.Last
         Dim Data As New List(Of Byte)(File.ReadAllBytes(SourceFile))
         Dim EndOfFileList As New Integer?
         Dim FileName As String = Nothing
         Dim FileOffset As New Integer
         Dim Files As New List(Of FileStr)
         Dim NextOffset As New Integer
         Dim Offset As Integer = 0
         Dim TargetFile As String = Nothing
         Dim TargetDirectory As String = Path.Combine(Path.GetDirectoryName(SourceFile), $"{Path.GetFileNameWithoutExtension(SourceFile)}_DATA")

         If GetCommandLineArgs.Count < 2 Then Throw New Exception("Specify an input file.")

         Do
            FileName = GetString(Data, Offset, 12, AdvanceOffset:=True)
            If FileName.Contains(PADDING) Then FileName = FileName.Substring(0, FileName.IndexOf(PADDING))
            If FileName.Intersect(INVALID_CHARACTERS).Count > 0 Then
               Throw New Exception("Invalid characters found in filename.")
            End If

            FileOffset = BitConverter.ToInt32(GetBytes(Data, Offset, &H4%, AdvanceOffset:=True).ToArray(), 0)

            If EndOfFileList Is Nothing Then
               EndOfFileList = FileOffset
            ElseIf Offset >= EndOfFileList Then
               Exit Do
            End If

            Files.Add(New FileStr With {.FileName = FileName, .Offset = FileOffset})
         Loop

         If Not Directory.Exists(TargetDirectory) Then
            Directory.CreateDirectory(TargetDirectory)
         End If

         For Index As Integer = 0 To Files.Count - 1
            With Files(Index)
               TargetFile = Path.Combine(TargetDirectory, .FileName)
               Console.WriteLine(TargetFile)
               NextOffset = If(Index >= Files.Count - 1, Data.Count, Files(Index + 1).Offset)
               If NextOffset > Data.Count Then
                  Throw New Exception("The offset cannot be beyond the end of the data file.")
               ElseIf NextOffset < .Offset Then
                  Throw New Exception("The offset of the next file cannot be less than the current file's offset.")
               End If
               File.WriteAllBytes(TargetFile, GetBytes(Data, .Offset, NextOffset - .Offset).ToArray())
            End With
         Next Index
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure returns the specified number of bytes at the specified position.
   Private Function GetBytes(Data As List(Of Byte), ByRef Offset As Integer, Count As Integer, Optional AdvanceOffset As Boolean = False) As List(Of Byte)
      Try
         Dim Bytes As New List(Of Byte)(Data.GetRange(Offset, Count))

         If AdvanceOffset Then Offset += Count

         Return Bytes
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns a string of the length specified in bytes at the specified position.
   Private Function GetString(Data As List(Of Byte), ByRef Offset As Integer, Count As Integer, Optional AdvanceOffset As Boolean = False) As String
      Try
         Return BYTES_TO_TEXT(GetBytes(Data, Offset, Count))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure handles any errors that occur.
   Private Sub HandleError(ExceptionO As Exception)
      Try
         Console.WriteLine($"ERROR: {ExceptionO.Message}")
         [Exit](0)
      Catch
         [Exit](0)
      End Try
   End Sub
End Module
