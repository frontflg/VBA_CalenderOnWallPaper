Attribute VB_Name = "Module1"
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Sub InsImg()
   Dim vntFileName As Variant
   Dim shp         As Shape
   Dim cropPoints  As Long
   Dim lngHeight   As Long
   Dim orgHeight   As Long
   Dim hiritu      As Single

   ChDrive ThisWorkbook.Path
   ChDir ThisWorkbook.Path
   ChDir "..\Wallpaper"       '�ǎ��t�H���_

 ' �t�@�C���I���_�C�A���O�ŉ摜�t�@�C����I�����A�ǂݍ���
   vntFileName = Application.GetOpenFilename(Title:="�ǎ��摜��I��", MultiSelect:=False)

   If vntFileName <> False Then
    ' �w�i�摜������΍폜
      If ActiveSheet.Shapes.Count > 4 Then
        On Error Resume Next
        ActiveSheet.Shapes("Background").Delete
      End If
      Range("J1").Select
      With ActiveSheet.Pictures.Insert(vntFileName)
      ' ���̒����i�g��E�k���j
        .ShapeRange.LockAspectRatio = msoTrue
        .ShapeRange.Width = GetSystemMetrics(0) * 75 / 100

      ' �����̒����i�g���~���O�j
         hiritu = GetSystemMetrics(1) / GetSystemMetrics(0)
         lngHeight = .ShapeRange.Width * hiritu
         If .ShapeRange.Height > lngHeight Then
            orgHeight = .ShapeRange.Height
            cropPoints = (orgHeight - lngHeight) / 2
           .ShapeRange.PictureFormat.CropTop = cropPoints
           .ShapeRange.PictureFormat.CropBottom = cropPoints
           .ShapeRange.IncrementTop (orgHeight - .ShapeRange.Width)
         End If

        .ShapeRange.ZOrder msoSendToBack
        .Name = "Background"
      End With
      
   End If
End Sub

Sub WrtImg()
   ChDrive ThisWorkbook.Path
   ChDir ThisWorkbook.Path

 ' �󎚗̈�̉摜�������o��(1366�~768)
   With ActiveWorkbook.PublishObjects.Add(xlSourcePrintArea, _
        ActiveWorkbook.Path & "\CalenderImg.htm", "Sheet1", "", xlHtmlStatic _
       , "", "")
       ActiveCell.Activate
       .Publish (True)
       .AutoRepublish = False
   End With

 ' �s�v�t�@�C�����폜
   Kill "CalenderImg.htm"

   Dim FSO As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")
   FSO.MoveFile ActiveWorkbook.Path & "\CalenderImg.files\*image003.png", ActiveWorkbook.Path
   FSO.DeleteFile ActiveWorkbook.Path & "\CalenderImg.files\*.*"
   Set FSO = Nothing
   
   RmDir (ActiveWorkbook.Path & "\CalenderImg.files\")

   MsgBox "�摜�t�@�C�����쐬���܂����B"
End Sub

Sub SetCalender()
   Dim myDate As Date
   If Day(Date) > 16 Then
      myDate = DateAdd("m", 1, Date)        '����
   Else
      myDate = Date                         '����
   End If
   Dim i As Integer
   Dim j As Integer: j = 4
   Dim m As Integer: m = Month(myDate)
   Dim y As Integer: y = Year(myDate)
   Dim k As Integer: k = Weekday(DateSerial(y, m, 1))
   Dim n As Integer: n = Day(DateSerial(y, m + 1, 0))

   Cells(2, 1).Value = m
   Cells(2, 3).Value = Format(myDate, "mmmm")
   Cells(2, 7).Value = y
   Range("B4:H9").ClearContents
   For i = 1 To n
      Cells(j, k + 1).Value = i
      If k > 1 And k < 7 Then
         If IsHoliday(DateSerial(y, m, i)) Then
            Cells(j, k + 1).Font.Color = vbMagenta
         Else
            Cells(j, k + 1).Font.Color = vbWhite
         End If
      End If
      k = k + 1
      If k > 7 Then
         j = j + 1
         k = 1
      End If
   Next i
   Cells(1, 1).Select
End Sub

'https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html �̕\�����[�N�V�[�g�ɓ\��t����
' ���t�����́AB5�`B26(2019�N�̏ꍇ)
Private Function IsHoliday(myDate As Date) As Boolean
   Dim ws As Worksheet: Set ws = Worksheets("�����̏j��")
   IsHoliday = False
 ' �j��
   If Not IsError(Application.Match(CLng(myDate), ws.Range("B5:B26"), 0)) Then
      IsHoliday = True
 ' �U�֋x���H
   ElseIf Weekday(myDate) = vbMonday And _
      Not IsError(Application.Match(CLng(DateAdd("d", -1, myDate)), ws.Range("B5:B26"), 0)) Then
      IsHoliday = True
   End If
End Function
