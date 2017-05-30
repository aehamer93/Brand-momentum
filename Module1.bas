Attribute VB_Name = "Module1"
Dim ie As InternetExplorerMedium
Dim h As MSHTML.HTMLDocument
Sub GooglePlus()
Dim Last_Row As Long, i As Long, j As Long
Dim span_tags As Object, span_length As Long, span_class As String

'Application.ScreenUpdating = False
Worksheets("Google+").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Activate
ActiveCell.Value = Now()
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
        ActiveCell.Offset(1, 0).Select
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            Application.Wait (Now + TimeValue("0:00:20"))
'            Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
            Set h = ie.document
            Set s = h.getElementsByClassName("C98T8d GseqId b12n5")
            ActiveCell.Value = Pick_number(s(0).innerText)
            ActiveCell.Value = ActiveCell.Value / 1000
'            Set span_tags = h.getElementsByTagName("span")
'            span_length = span_tags.Length
'                For i = 0 To span_length - 1
'                    span_class = span_tags(i).className
'                        If span_class = "C98T8d GseqId b12n5" Then
'                            ActiveCell.Value = Pick_number(span_tags(i).innerText)
'                            ActiveCell.Value = ActiveCell.Value / 1000
'                            Exit For
'                        End If
'                Next
        Else
            ActiveCell.FillRight
        End If
    Next
ie.Quit
Set span_tags = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub

Sub Facebook()
Application.ScreenUpdating = False
Worksheets("Facebook").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.Value = Now()
Dim Last_Row As Long, i As Long, j As Long
Dim div_tags As Object, div_length As Long, div_class1 As String, div_class2_innertext As String
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
        ActiveCell.Offset(1, 0).Select
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
            Set h = ie.document
            Set div_tags = h.getElementsByTagName("div")
            div_length = div_tags.Length
                For i = 0 To div_length - 1
                    div_class1 = div_tags(i).className
                    div_class2_innertext = div_tags(i + 1).innerText
                        If div_class1 = "_50f6 _50f7 _5tfx" And div_class2_innertext = "Total Page Likes" Then
                            ActiveCell.Value = Pick_number(div_tags(i).innerText)
                            ActiveCell.Value = ActiveCell.Value / 1000
                            Exit For
                        End If
                Next
           Else
              ActiveCell.FillRight
        End If
    Next
ie.Quit
Set div_tags = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub

Sub Twitter()
Application.ScreenUpdating = False
Worksheets("Twitter").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.Value = Now()
Dim Last_Row As Long, i As Long, j As Long
Dim td_tags As Object, td_length As Long, td_class As String, td_class2_innertext As String
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
        ActiveCell.Offset(1, 0).Select
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            Application.Wait (Now + TimeValue("0:00:10"))
            Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
            Set h = ie.document
            Set td_tags = h.getElementsByTagName("td")
            td_length = td_tags.Length
                For i = 0 To td_length - 1
                    td_class = td_tags(i).className
                        If td_class = "stat stat-last" Then
                            ActiveCell.Value = Pick_number(td_tags(i).getElementsByTagName("div")(0).innerText)
                            ActiveCell.Value = ActiveCell.Value / 1000
                            Exit For
                        End If
                Next
        Else
            ActiveCell.FillRight
        End If
       
    Next
ie.Quit
Set h = Nothing
Set td_tags = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub


Sub YouTube()
Application.ScreenUpdating = False
Worksheets("YouTube").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.Value = Now()
Dim Last_Row As Long, i As Long, j As Long
Dim span_tags As Object, span_length As Long, span_class As String
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
        ActiveCell.Offset(1, 0).Select
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            Application.Wait (Now + TimeValue("0:00:10"))
            'Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
            Set h = ie.document
            Set span_tags = h.getElementsByTagName("span")
            span_length = span_tags.Length
                For i = 0 To span_length - 1
                    span_class = span_tags(i).className
                        If span_class = "yt-subscription-button-subscriber-count-branded-horizontal subscribed yt-uix-tooltip" Then
                            ActiveCell.Value = Pick_number(span_tags(i).innerText)
                            ActiveCell.Value = ActiveCell.Value / 1000
                            Exit For
                        End If
                Next
        Else
            ActiveCell.FillRight
        End If
    Next
ie.Quit
Set span_tags = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub

Sub Instagram()
Application.ScreenUpdating = False
Worksheets("Instagram").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.Value = Now()
Dim Last_Row As Long, i As Long, j As Long
Dim span_tags As Object, span_length As Long, span_class As String
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
        ActiveCell.Offset(1, 0).Select
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
             Application.Wait (Now + TimeValue("0:00:15"))
            Set h = ie.document
            Set span_tags = h.getElementsByTagName("span")
            span_length = span_tags.Length
                For i = 0 To span_length - 1
                    span_class = span_tags(i).className
                        If span_class = "-cx-PRIVATE-FollowedByStatistic__count" Then
                            ActiveCell.Value = Pick_number(span_tags(i).Title)
                            ActiveCell.Value = ActiveCell.Value / 1000
                            Exit For
                        End If
                Next
        Else
            ActiveCell.FillRight
        End If
    Next
ie.Quit
Set span_tags = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub

Sub Pintrest()
Application.ScreenUpdating = False
Worksheets("Pinterest").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.Value = Now()
Dim Last_Row As Long, i As Long, j As Long
Dim div_tags As Object, div_length As Long, div_class As String
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
        ActiveCell.Offset(1, 0).Select
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
            Set h = ie.document
            Set div_tags = h.getElementsByTagName("div")
            div_length = div_tags.Length
                For i = 0 To div_length - 1
                    div_class = div_tags(i).className
                        If div_class = "FollowerCount Module" Then
                            ActiveCell.Value = Pick_number(div_tags(i).getElementsByTagName("span")(0).innerText)
                            ActiveCell.Value = ActiveCell.Value / 1000
                            Exit For
                        End If
                Next
        Else
            ActiveCell.FillRight
        End If
    Next
ie.Quit
Set div_tags = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub
Sub Weibo()
Application.ScreenUpdating = False
Worksheets("Weibo").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.Value = Now()
Dim Last_Row As Long, i As Long, j As Long
Dim td_tags As Object, td_length As Long, td_span_innertext As String
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
        ActiveCell.Offset(1, 0).Select
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            'Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
            Application.Wait (Now + TimeValue("0:00:20"))
            Set h = ie.document
            Set td_tags = h.getElementsByTagName("td")
            td_length = td_tags.Length
                For i = 0 To td_length - 1
                    td_span_innertext = td_tags(i).getElementsByTagName("span")(0).innerText
                        If td_span_innertext = Range("H1").Value Then
                            ActiveCell.Value = Pick_number(td_tags(i).getElementsByTagName("strong")(0).innerText)
                            ActiveCell.Value = ActiveCell.Value / 1000
                            Exit For
                        End If
                Next
        Else
            ActiveCell.FillRight
        End If
    Next
ie.Quit
Set td_tags = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub

Sub Tmall()
Application.ScreenUpdating = False
Worksheets("# products on TMall").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.Value = Now()
Dim Last_Row As Long, i As Long, j As Long
Dim Search_Box As Object, Search_Button As Object, Button_Text As String, P_Tag As Object, P_ClassNm As String, P_TagCount As Long
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
ie.navigate ("http://www.tmall.com/")
Do While ie.Busy = True Or ie.readyState <> 4 Or InStr(1, ie.statusText, "Waiting for"): DoEvents: Loop
Application.Wait (Now + TimeValue("0:00:10"))
Set h = ie.document
ActiveCell.Offset(1, 0).Select
    For j = 4 To Last_Row
      If Range("A" & j).Value <> "Company" And Range("A" & j).Value <> "" Then
        Set Search_Box = h.getElementById("mq")
        Set Search_Button = h.getElementsByTagName("button")(0)
        Search_Box.innerText = Range("B" & j).Value
        Search_Button.Click
        Do While ie.Busy = True Or ie.readyState <> 4 Or InStr(1, ie.statusText, "Waiting for"): DoEvents: Loop
        Application.Wait (Now + TimeValue("0:00:10"))
        Set h = ie.document
        Set P_Tag = h.getElementsByTagName("p")
        P_TagCount = P_Tag.Length
            For i = 0 To P_TagCount - 1
                P_ClassNm = P_Tag(i).className
                If P_ClassNm = "crumbTitle j_ResultsNumber" Then
                    ActiveCell.Value = Pick_number(P_Tag(i).innerText)
                    Exit For
                End If
            Next
      Else
            ActiveCell.FillRight
      End If
        ActiveCell.Offset(1, 0).Select
    Next
ie.Quit
Set P_Tag = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub
Sub Similarweb_Rank()
Application.ScreenUpdating = False
Worksheets("SimilarWeb Global Rank").Activate
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.FillRight
Dim Last_Row As Long, i As Long, j As Long
Dim span_tags As Object, span_length As Long, span_class As String, A As Integer
Dim CurrentCell As Range
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
        ActiveCell.Offset(1, 0).Select
        Set CurrentCell = ActiveCell
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            Do While ie.LocationName = "Security Screen"
            A = MsgBox(prompt:="Please enter the captcha code and click ok once the webpage is visible", Buttons:=vbOKOnly, Title:="Captcha screen is reached")
             If A = vbOK Then
                CurrentCell.Select
             End If
            Loop
            Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
            Application.Wait (Now + TimeValue("0:00:10"))
            Set h = ie.document
            Set span_tags = h.getElementsByTagName("span")
            span_length = span_tags.Length
                For i = 0 To span_length - 1
                    span_class = span_tags(i).className
                    
                        If span_class = "rankingItem-value js-countable" Then
                            ActiveCell.Value = Pick_number(span_tags(i).innerText)
                            Exit For
                        End If
                Next
        Else
            ActiveCell.FillRight
        End If
    Next
ie.Quit
Set span_tags = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub
Sub Similarweb_VISITS()
Application.ScreenUpdating = False
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).Offset(0, 2).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(0, 1).EntireColumn.Insert
Range("A3").End(xlToRight).Offset(0, 1).Select
ActiveCell.FillRight
Dim Last_Row As Long, i As Long, j As Long, answer As Integer
Dim span_tags As Object, span_length As Long, span_class As String, A As Integer
Dim CurrentCell As Range
Last_Row = Range("A1048576").End(xlUp).Row
Set ie = New InternetExplorer
ie.Visible = True
    For j = 4 To Last_Row
     ActiveCell.Offset(1, 0).Select
     Set CurrentCell = ActiveCell
        If Range("C" & j) <> "NA" And Range("C" & j) <> "" Then
            ie.navigate (Range("C" & j).Value)
            Do While ie.Busy = True Or ie.readyState <> 4: DoEvents: Loop
            Application.Wait (Now + TimeValue("0:00:05"))
            Set h = ie.document
            Do While ie.LocationName = "Security Screen"
            A = MsgBox(prompt:="Please enter the captcha code and click ok once the webpage is visible", Buttons:=vbOKOnly, Title:="Captcha screen is reached")
             If A = vbOK Then
                CurrentCell.Select
             End If
            Loop
           
            Set span_tags = h.getElementsByTagName("span")
            span_length = span_tags.Length
                For i = 0 To span_length - 1
                    span_class = span_tags(i).className
                    
                        If span_class = "engagementInfo-value engagementInfo-value--large u-text-ellipsis" Then
                            If Right(span_tags(i).innerText, 1) = "K" Then
                                ActiveCell.Value = Replace(span_tags(i).innerText, "K", "")
                            ElseIf Right(span_tags(i).innerText, 1) = "M" Then
                                ActiveCell.Value = Replace(span_tags(i).innerText, "M", "")
                                ActiveCell.Value = ActiveCell.Value * 1000
                            Else
                                ActiveCell.Value = Pick_number(span_tags(i).innerText)
                            End If
                            Exit For
                        End If
                Next
        Else
            ActiveCell.FillRight
        End If
    Next
ie.Quit
Set span_tags = Nothing
Set h = Nothing
Set ie = Nothing
Application.ScreenUpdating = True
End Sub

Sub Update_Chart()
Application.ScreenUpdating = False
Range("A3").End(xlToRight).EntireColumn.Insert
Range("A3").End(xlToRight).End(xlToRight).EntireColumn.Cut
Range("A3").End(xlToRight).Offset(-2, 1).Insert
Range("XFD3").End(xlToLeft).Offset(0, 1).Select
ActiveCell.Value = "=TEXT(EOMONTH(" & Replace(Range("XFD3").End(xlToLeft).Address, "$", "") & "," & Month(Range("XFD3").End(xlToLeft).Value) - Month(Range("XFD3").End(xlToLeft).Offset(0, -1).Value) & "),""mmm-yy"")"
Application.ScreenUpdating = True
End Sub
