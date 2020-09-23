Attribute VB_Name = "mdlFilterWithPaintOption5"
Public Function fncNoLuminence() As Long
Dim a As Integer
    If rAdj < gAdj Then
    a = rAdj
    Else
    a = gAdj
    End If
    
    If a < bAdj Then
    a = a
    Else
    a = bAdj
    End If

fncNoLuminence = RGB(rAdj - a, gAdj - a, bAdj - a)
End Function

Public Sub rtnFilterWithPaintOption5()
Dim X As Integer, Y As Integer, xx As Integer, UnChkdForAdjPxlCount As Long, BeingChkdForAdkPxlCount As Long
Dim rTrace As Byte, gTrace As Byte, bTrace As Byte
Dim rPxl As Integer, gPxl As Integer, bPxl As Integer, SumPxl As Integer
Dim w As Long, lngPixelPointer As Long
Dim colortrace As Long
Dim intPxlX As Integer 'Main Working Variable
Dim intPxlY As Integer 'Main Working Variable
Dim intBW As Integer 'Holds Black And White RGB value
Dim rDiff As Integer, gDiff As Integer, bDiff As Integer
'***************************************************************
'***********Summary Option 3 Black And White********************
'**Paint picture, pixel selection according to bandpass filter,*
'***********two passes, RGB-Luminence, RGB******************************
'***************************************************************
'We enter with our picture box blank and arrays holding
'the colors and XY coordinates of each pixel.
'The XY array is in random order.
'We also already know the user entered value of our filter.
'A filter value of one will pass only one color.
'Any larger value will pass a respective bandpass of colors.
'This routine will go through each pixel in random order.
'Let's say we are at our very first pixel.
'*****Start*****
'We get its color and test it to see if this pixel
'has been painted. If yes then we jump to the next random pixel.
'If it has not been painted then our main processing begins.
'We paint and mark this pixel as painted.
'Before we paint it though we take away the luminece portion.
'No filter checking on this first pixel.
'This will be our "trace" color used in our filter.
'Now we check all eight of its immediately adjacent pixels.
'We check if not already painted and if not we check too see
'if they pass the filter test.If yes they are marked and
'painted. As before we take away the luminece portion.
'Of those painted and marked as painted, their eight adjacent
'pixels too are checked for painted status & with the filter test.
'When no more adjacent pixels can be found that both, are not
'painted and pass the filter test then the next random pixel is
'selected and we start the process over. (from *****Start*****)
'This will continue until we have randomly gone through each
'picture pixel.
'***************************************************************
'***************************************************************
'***************************************************************

With frmMain
.picMain.AutoRedraw = False
intColorFilter = Val(.Text1.Text)
'Scan all pixels one by one
For lngPixelPointer = 1 To lngTotalPixelCount
'Get them from previously loaded 2 dimension array
'Remember that the array holds the correct X,Y coordinates
'but in random order so as to scan nonsequentially
intPxlX = intRandomXYs(cX, lngPixelPointer)
intPxlY = intRandomXYs(cY, lngPixelPointer)
'Check to see if current pixel has already been painted
'If it is then jump to next one
If lngPixelColors(intPxlX, intPxlY) = cAlreadyPainted Then GoTo lblAlreadyPainted:
'If it has not been painted then here we begin our Main processing
'First we calculate its RGB colors and
'add the RGB values together to use for our filter testing
color = lngPixelColors(intPxlX, intPxlY)
colortrace = color
rtnGetRGBColors
SumPxl = rAdj + bAdj + gAdj
intColorToFilter = SumPxl
'Paint without the luminece portion.
.picMain.PSet (intPxlX, intPxlY), fncNoLuminence()
'And we mark this pixel as painted
lngPixelColors(intPxlX, intPxlY) = cAlreadyPainted
'Now we store it.
'This arrAdjPxlsToChk array will hold all the pixel XYs that have
'been marked as painted but their eight surrounding pixels still need to be tested
arrAdjPxlsToChk(1, cX) = intPxlX: arrAdjPxlsToChk(1, cY) = intPxlY
'This variable will tell us how many pixels are in the arrAdjPxlsToChk array
'all of which still need to have their adjacent pixels checked
UnChkdForAdjPxlCount = 1
    'Do while there are still unchecked adjacent pixels
    Do While UnChkdForAdjPxlCount > 0
    'Thansfer count to out work variable
    BeingChkdForAdkPxlCount = UnChkdForAdjPxlCount
    'Reset so we can see how many adjacent pixels will pass filter test
    'and thus become unchecked pixels needing their adjacents checked
    UnChkdForAdjPxlCount = 0
        'Transfer pixels to our work array
        For xx = 1 To BeingChkdForAdkPxlCount
        arrAdjPxlsBeingChkd(xx, cX) = arrAdjPxlsToChk(xx, cX)
        arrAdjPxlsBeingChkd(xx, cY) = arrAdjPxlsToChk(xx, cY)
        Next xx
  
        'Scan through each pixel in work array
        'so we can check its adjacent pixels
        For xx = 1 To BeingChkdForAdkPxlCount
       'Get its XY coordinates
        X = arrAdjPxlsBeingChkd(xx, cX): Y = arrAdjPxlsBeingChkd(xx, cY)
        
        'Get color of First adjacent pixel
        color = lngPixelColors(X + 1, Y)
            'If not already painted then
            If color <> cAlreadyPainted Then
            'get RGB colors and add them together to use to check against filter
            rtnGetRGBColors
            intColorToFilter = rAdj + bAdj + gAdj
                'Check if it passes user entered filter. If yes then
                If intColorToFilter > SumPxl - intColorFilter And intColorToFilter < SumPxl + intColorFilter Then
                'Mark it
                lngPixelColors(X + 1, Y) = cAlreadyPainted
                'Paint without the luminece portion.
                .picMain.PSet (X + 1, Y), fncNoLuminence()
                'Update/increment unchecked for adjacents pixel count
                UnChkdForAdjPxlCount = UnChkdForAdjPxlCount + 1
                'Update unchecked for adjacents pixel array
                arrAdjPxlsToChk(UnChkdForAdjPxlCount, cX) = X + 1:    arrAdjPxlsToChk(UnChkdForAdjPxlCount, cY) = Y
                End If
            End If
    
        'Get color of Second adjacent pixel
        color = lngPixelColors(X, Y + 1)
            'If not already painted then
            If color <> cAlreadyPainted Then
            'get RGB colors and add them together to use to check against filter
            rtnGetRGBColors
            intColorToFilter = rAdj + bAdj + gAdj
                'Check if it passes user entered filter. If yes then
                If intColorToFilter > SumPxl - intColorFilter And intColorToFilter < SumPxl + intColorFilter Then
                'Mark it
                lngPixelColors(X, Y + 1) = cAlreadyPainted
                'Paint without the luminece portion.
                .picMain.PSet (X, Y + 1), fncNoLuminence()
                'Update/increment unchecked for adjacents pixel count
                UnChkdForAdjPxlCount = UnChkdForAdjPxlCount + 1
                'Update unchecked for adjacents pixel array
                arrAdjPxlsToChk(UnChkdForAdjPxlCount, cX) = X:     arrAdjPxlsToChk(UnChkdForAdjPxlCount, cY) = Y + 1
                End If
            End If
    
        'Get color of Third adjacernt pixel
        color = lngPixelColors(X + 1, Y + 1)
            'If not already painted then
            If color <> cAlreadyPainted Then
            'get RGB colors and add them together to use to check against filter
            rtnGetRGBColors
            intColorToFilter = rAdj + bAdj + gAdj
                'Check if it passes user entered filter. If yes then
                If intColorToFilter > SumPxl - intColorFilter And intColorToFilter < SumPxl + intColorFilter Then
                'Mark it
                lngPixelColors(X + 1, Y + 1) = cAlreadyPainted
                'Paint without the luminece portion.
                .picMain.PSet (X + 1, Y + 1), fncNoLuminence()
                'Update/increment unchecked for adjacents pixel count
                UnChkdForAdjPxlCount = UnChkdForAdjPxlCount + 1
                'Update unchecked for adjacents pixel array
                arrAdjPxlsToChk(UnChkdForAdjPxlCount, cX) = X + 1:    arrAdjPxlsToChk(UnChkdForAdjPxlCount, cY) = Y + 1
                End If
            End If
    
        'Get color of Fourth adjacent pixel
        color = lngPixelColors(X - 1, Y - 1)
            'If not already painted then
            If color <> cAlreadyPainted Then
            'get RGB colors and add them together to use to check against filter
            rtnGetRGBColors
            intColorToFilter = rAdj + bAdj + gAdj
                'Check if it passes user entered filter. If yes then
                If intColorToFilter > SumPxl - intColorFilter And intColorToFilter < SumPxl + intColorFilter Then
                'Mark it
                lngPixelColors(X - 1, Y - 1) = cAlreadyPainted
                'Paint without the luminece portion.
                .picMain.PSet (X - 1, Y - 1), fncNoLuminence()
                'Update/increment unchecked for adjacents pixel count
                UnChkdForAdjPxlCount = UnChkdForAdjPxlCount + 1
                'Update unchecked for adjacents pixel array
                arrAdjPxlsToChk(UnChkdForAdjPxlCount, cX) = X - 1:   arrAdjPxlsToChk(UnChkdForAdjPxlCount, cY) = Y - 1
                End If
            End If
    
        'Get color of Fifth adjacent pixel
        color = lngPixelColors(X - 1, Y)
            'If not already painted then
            If color <> cAlreadyPainted Then
            'get RGB colors and add them together to use to check against filter
            rtnGetRGBColors
            intColorToFilter = rAdj + bAdj + gAdj
                'Check if it passes user entered filter. If yes then
                If intColorToFilter > SumPxl - intColorFilter And intColorToFilter < SumPxl + intColorFilter Then
                'Mark it
                lngPixelColors(X - 1, Y) = cAlreadyPainted
                'Paint without the luminece portion.
                .picMain.PSet (X - 1, Y), fncNoLuminence()
                'Update/increment unchecked for adjacents pixel count
                UnChkdForAdjPxlCount = UnChkdForAdjPxlCount + 1
                'Update unchecked for adjacents pixel array
                arrAdjPxlsToChk(UnChkdForAdjPxlCount, cX) = X - 1:    arrAdjPxlsToChk(UnChkdForAdjPxlCount, cY) = Y
                End If
            End If

        'Get color of Sixth adjacent pixel
        color = lngPixelColors(X, Y - 1)
            'If not already painted then
            If color <> cAlreadyPainted Then
            'get RGB colors and add them together to use to check against filter
            rtnGetRGBColors
            intColorToFilter = rAdj + bAdj + gAdj
                'Check if it passes user entered filter. If yes then
                If intColorToFilter > SumPxl - intColorFilter And intColorToFilter < SumPxl + intColorFilter Then
                'Mark it
                lngPixelColors(X, Y - 1) = cAlreadyPainted
                'Paint without the luminece portion.
                .picMain.PSet (X, Y - 1), fncNoLuminence()
                'Update/increment unchecked for adjacents pixel count
                UnChkdForAdjPxlCount = UnChkdForAdjPxlCount + 1
                'Update unchecked for adjacents pixel array
                arrAdjPxlsToChk(UnChkdForAdjPxlCount, cX) = X:     arrAdjPxlsToChk(UnChkdForAdjPxlCount, cY) = Y - 1
                End If
            End If
    
        'Get color of Seventh adjacent pixel
        color = lngPixelColors(X - 1, Y + 1)
            'If not already painted then
            If color <> cAlreadyPainted Then
            'get RGB colors and add them together to use to check against filter
            rtnGetRGBColors
            intColorToFilter = rAdj + bAdj + gAdj
                'Check if it passes user entered filter. If yes then
                If intColorToFilter > SumPxl - intColorFilter And intColorToFilter < SumPxl + intColorFilter Then
                'Mark it
                lngPixelColors(X - 1, Y + 1) = cAlreadyPainted
                'Paint without the luminece portion.
                .picMain.PSet (X - 1, Y + 1), fncNoLuminence()
                'Update/increment unchecked for adjacents pixel count
                UnChkdForAdjPxlCount = UnChkdForAdjPxlCount + 1
                'Update unchecked for adjacents pixel array
                arrAdjPxlsToChk(UnChkdForAdjPxlCount, cX) = X - 1:    arrAdjPxlsToChk(UnChkdForAdjPxlCount, cY) = Y + 1
                End If
            End If

        'Get color of Eighth adjacent pixel
        color = lngPixelColors(X + 1, Y - 1)
            'If not already painted then
            If color <> cAlreadyPainted Then
            'get RGB colors and add them together to use to check against filter
            rtnGetRGBColors
            intColorToFilter = rAdj + bAdj + gAdj
                'Check if it passes user entered filter. If yes then
                If intColorToFilter > SumPxl - intColorFilter And intColorToFilter < SumPxl + intColorFilter Then
                'Mark it
                lngPixelColors(X + 1, Y - 1) = cAlreadyPainted
                'Paint without the luminece portion.
                .picMain.PSet (X + 1, Y - 1), fncNoLuminence()
                'Update/increment unchecked for adjacents pixel count
                UnChkdForAdjPxlCount = UnChkdForAdjPxlCount + 1
                'Update unchecked for adjacents pixel array
                arrAdjPxlsToChk(UnChkdForAdjPxlCount, cX) = X + 1:     arrAdjPxlsToChk(UnChkdForAdjPxlCount, cY) = Y - 1
                End If
            End If
        Next xx
   
    DoEvents
    'Speed
    For w = 1 To lngDelay
    Next w
    Loop

lblAlreadyPainted:
Next lngPixelPointer

Exit Sub
End With

End Sub



