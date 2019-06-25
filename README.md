# VBA Macro - Image resize & arrange

## Usage
#### 메인코드
```vba
Attribute VB_Name = "resizeNarrange"
' resizeNarrange(430,3,5,10)
'      changeHeight : Image Height
'      cols : column count
'      gapH : holizontal gap
'      gapV : vertical gap
Option Explicit
Public Sub resizeNarrange(changeHeight As Integer, cols As Integer, gapH As Integer, gapV As Integer)

    Dim idx As Integer
    Dim rIdx As Integer
    Dim cIdx As Integer
    'Dim cols As Integer
    'Dim changeHeight As Integer
    Dim changeWidth As Integer

    Dim startTop As Integer
    Dim startLeft As Integer

    Dim initTop As Integer
    Dim initLeft As Integer

    'Dim gapH, gapV As Integer

    'changeHeight = 430

    'cols = 3
    'gapH = 5
    'gapV = 10
    cIdx = 0
    Dim Shp As Shape

    For Each Shp In ActiveSheet.Shapes
        If Not Shp Is Nothing And Shp.Type = 13 Then
            Shp.Height = changeHeight
            rIdx = Int(idx / cols)
            If idx = 0 Then
                changeWidth = Shp.Width
                initTop = Shp.Top
                initLeft = Shp.Left
                startTop = initTop
                startLeft = initLeft
            Else
                If cIdx = 0 Then
                    startLeft = initLeft
                Else
                    'startLeft = ((changeWidth) * (cIdx)) + 50
                    startLeft = startLeft + Shp.Width + gapH
                End If
            End If

            If rIdx <> 0 Then
                startTop = initTop + ((rIdx) * changeHeight) + gapV
            End If

            Shp.Top = startTop

            'startLeft = startLeft + Shp.Width + 20
            Shp.Left = startLeft

            idx = idx + 1
            cIdx = cIdx + 1

            If idx Mod cols = 0 Then
                startLeft = initLeft
                cIdx = 0
            End If

        End If
    Next
End Sub

```
