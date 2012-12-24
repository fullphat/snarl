
Function style_Init()

  ' // provide information about our style

  With style.info
    .format = "display"
    .name = "Dialogue Scripted"
    .description = "Looks a bit like a login window (powered by VBScript)"
    .copyright = "© 2012 full phat products"
    .version = 0
    .revision = 2
    .icon = style.path & "icon.png"
    .schemes = "Standard"
    .supporturl = "http://getsnarl.info"

  End With

End Function

Sub style_Draw()
Dim pbIcon
Dim xEdge
Dim yTitle
Dim sz
Dim prText
Dim prMeter
Dim pr
Dim cyMeter
Dim cyText
Dim i

Const RX = 6
Const TITLE_TEXT_GAP = 2
Const POPUP_WIDTH = 300

  ' /* Must create these as it seems VBScript can't instantiate an object when it's referenced as an argument */

  Set prText = new_BRect(0,0,0,0)
  Set prMeter = new_BRect(0,0,0,0)

  With view 
    .SizeTo POPUP_WIDTH, 80

    Set pbIcon = load_image_obj(style.notification.ValueOf("icon"))
    xEdge = 6
    If is_valid_image((pbIcon)) Then _
      xEdge = xEdge + 32 + 4			' // Add on icon width and gap

    If style.notification.ValueOf("title") <> "" Then _
      yTitle = 20

    If style.notification.ValueOf("text") <> "" Then
      .SetFont "Arial", 9
      .MeasureString style.notification.ValueOf("text"), new_BRect(xEdge, 0, .Width - 6, 16384), (prText)
      cyText = prText.Height + 6

    End If

    ' /* meter? */

    If style.notification.exists("value-percent") Then
      Set prMeter = new_BRect(xEdge + 3, 0, .Width - 6 - 6, 12 - 1)
      cyMeter = prMeter.Height + 6

    End If

    ' /* final size */

    .SizeTo POPUP_WIDTH, MAX(60, cyText + cyMeter + 6 + yTitle + 6)	' // 6 = 2xMargin, 10=Space

    ' /* background */

    .EnableSmoothing True
    .TextMode = 4			'// MFX_TEXT_ANTIALIAS

    Select Case style.notification.ValueOf("scheme")
    Case "light"

    Case Else

    End Select


    i = string_as_long(style.notification.ValueOf("priority"))
    If i > 0 Then
      ' /* priority */
      .SetHighColour rgba(236, 59, 0)

    ElseIf i < 0 Then
      ' /* low */
      .SetHighColour rgba(154, 154, 154)

    Else
      ' /* normal */
      .SetHighColour rgba(89, 109, 149)

    End If

    ' /* background */

    Set pr = .Bounds()
    .FillRoundRect (pr), RX, RX

    ' /* shading */

    .SetHighColour rgba(0, 0, 0, 0)
    .SetLowColour rgba(0, 0, 0, 64)
    .FillRoundRect (pr), RX, RX, 3

    ' /* edge */

    .SetHighColour rgba(0, 0, 0, 76)
    .StrokeRoundRect (pr), RX, RX

    ' /* content area */

    pr.InsetBy 3, 3
    pr.Top = pr.Top + yTitle
    .SetHighColour rgba(255, 255, 255, 230)
    .FillRoundRect (pr), RX, RX

    .SetHighColour rgba(0, 0, 0, 46)
    .StrokeRoundRect (pr), RX, RX

    ' /* title */

    If yTitle > 0 Then
      .SetFont "Arial", 9, True            
      .SetHighColour rgba(255, 255, 255)
      .SetLowColour rgba(0, 0, 0, 76)

      Set pr = .Bounds.InsetByCopy(8, 2)
      pr.Bottom = pr.Top + yTitle
      .DrawString .GetFormattedText(style.notification.ValueOf("title"), (pr), True), (pr), &H222	' // H-Centre/V-Centre/Outline

    End If

    ' /* icon */

    If (style.notification.ValueOf("text") = "") And (cyMeter = 0) Then
      ' /* no meter or text, so center horizontally */
      i = Fix((.Width - 32) / 2)

    Else
      i = 3 + 4

    End If

    .DrawScaledImage (pbIcon), new_BPoint(i, 3 + yTitle + 4), new_BPoint(32, 32)

    ' /* text */

    .SetFont "Arial", 9
    Select Case style.notification.ValueOf("scheme")
    Case "light"

    Case Else
      .SetHighColour rgba(255, 255, 255, 198)

    End Select

    .SetHighColour rgba(0, 0, 0, 247)

    prText.OffsetBy 0, yTitle + 3 + 4
    .DrawString style.notification.ValueOf("text"), (prText), 0

    ' /* meter */

    If cyMeter > 0 Then

      prMeter.OffsetBy 0, yTitle + 3 + 6 + cyText

      .SetHighColour rgba(0, 0, 0, 24)
      .FillRect (prMeter)
      .SetHighColour rgba(0, 0, 0, 32)
      .StrokeRect (prMeter)

      ' /* meter fill */

      i = string_as_long(style.notification.ValueOf("value-percent"))
      If i > 0 Then
        If string_as_long(style.notification.ValueOf("priority")) > 0 Then
          ' /* priority */
          .SetHighColour rgba(236, 59, 0)

        Else
          .SetHighColour rgba(89, 109, 149)

        End If

        prMeter.Right = prMeter.Left + ((i / 100) * prMeter.Width)
        .FillRect (prMeter)

        .SetHighColour rgba(0, 0, 0, 32)
        .StrokeRect (prMeter)

      End If

    End If

  End With

End Sub