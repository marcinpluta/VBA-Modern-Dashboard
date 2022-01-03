VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function Change_lnfo_Box_Visibility(lnfo_Button As Object, lnfo_Box As Object)

If lnfo_Button.Fill.ForeColor.RGB = RGB(255, 255, 255) Then
    lnfo_Button.Fill.ForeColor.RGB = RGB(255, 255, 0)
    lnfo_Box.Visible = True
Else
    lnfo_Button.Fill.ForeColor.RGB = RGB(255, 255, 255)
    lnfo_Box.Visible = False
End If
End Function

Sub Change_lnfo_Box_All()


With ActiveSheet

Call Change_lnfo_Box_Visibility(.Shapes("Icon_Info"), .Shapes("Box_Info"))
End With

End Sub

Sub Display_Tab_Sales_Revenue()

    With ActiveSheet
        'Visibility of Tab Button
        With .Shapes("Tab_Button_Sales_Revenue")
            .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
            .Fill.Transparency = 0#
        End With
        
        With .Shapes("Tab_Button_Sales_Units")
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.Transparency = 1#
        End With

        .Shapes("Map_Chart_Sales_Revenue").Visible = True
        .Shapes("Line_Chart_Sales_Revenue").Visible = True
        .Shapes("Map_Chart_Sales_Units").Visible = False
        .Shapes("Line_Chart_Sales_Units").Visible = False
 
    End With

End Sub

Sub Display_Tab_Sales_Units()

    With ActiveSheet
    
       
        With .Shapes("Tab_Button_Sales_Revenue")
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.Transparency = 1#
        End With
        
        With .Shapes("Tab_Button_Sales_Units")
            .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
            .Fill.Transparency = 0#
        End With

        .Shapes("Map_Chart_Sales_Revenue").Visible = False
        .Shapes("Line_Chart_Sales_Revenue").Visible = False
        .Shapes("Map_Chart_Sales_Units").Visible = True
        .Shapes("Line_Chart_Sales_Units").Visible = True

    End With

End Sub

'Settings Menu Button

Sub Change_Settings_Visibility()
    
    With ActiveSheet
        If .Shapes("Settings_Button").Fill.Transparency = 1# Then
        
            'Visibility Settings Button
            .Shapes("Settings_Button").Fill.Transparency = 0#
            
            'Visibility Settings Menu
            .Shapes("Settings_Menu_Frame").Visible = True
            .Shapes("Theme_Button_1").Visible = True
            .Shapes("Theme_Button_2").Visible = True
            .Shapes("Theme_Button_3").Visible = True
            .Shapes("Theme_Button_4").Visible = True
            .Shapes("Toggle_Background_Tabs").Visible = True
            .Shapes("Toggle_Circle_Tabs").Visible = True
            .Shapes("Toggle_Textbook_Tabs").Visible = True
            .Shapes("Toggle_Background_Info").Visible = True
            .Shapes("Toggle_Circle_Info").Visible = True
            .Shapes("Toggle_Textbook_Info").Visible = True
            
         Else
        
            'Visibility Settings Button
            .Shapes("Settings_Button").Fill.Transparency = 1#
            
            'Visibility Settings Menu
            .Shapes("Settings_Menu_Frame").Visible = False
            .Shapes("Theme_Button_1").Visible = False
            .Shapes("Theme_Button_2").Visible = False
            .Shapes("Theme_Button_3").Visible = False
            .Shapes("Theme_Button_4").Visible = False
            .Shapes("Toggle_Background_Tabs").Visible = False
            .Shapes("Toggle_Circle_Tabs").Visible = False
            .Shapes("Toggle_Textbook_Tabs").Visible = False
            .Shapes("Toggle_Background_Info").Visible = False
            .Shapes("Toggle_Circle_Info").Visible = False
            .Shapes("Toggle_Textbook_Info").Visible = False
            
        End If
        
    End With
    
End Sub




'Color Theme Buttons

Function Change_Radio_Button_Status(Radio_Button As Object, Active As Boolean)
    With Radio_Button
        If Active = True Then
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.Transparency = 0#
        Else
            .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
            .Fill.Transparency = 1#
        End If
    End With
End Function

Sub Display_Theme_1()
    With ActiveSheet
        'Dashboard Themes Visibility
        .Shapes("Dashboard_Theme_1").Visible = True
        .Shapes("Dashboard_Theme_2").Visible = False
        .Shapes("Dashboard_Theme_3").Visible = False
        .Shapes("Dashboard_Theme_4").Visible = False

        'Color Theme Buttons Status
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_1"), True)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_2"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_3"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_4"), False)
    End With
End Sub

Sub Display_Theme_2()
    With ActiveSheet
        'Dashboard Themes Visibility
        .Shapes("Dashboard_Theme_1").Visible = False
        .Shapes("Dashboard_Theme_2").Visible = True
        .Shapes("Dashboard_Theme_3").Visible = False
        .Shapes("Dashboard_Theme_4").Visible = False

        'Color Theme Buttons Status
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_1"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_2"), True)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_3"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_4"), False)
    End With
End Sub

Sub Display_Theme_3()
    With ActiveSheet
        'Dashboard Themes Visibility
        .Shapes("Dashboard_Theme_1").Visible = False
        .Shapes("Dashboard_Theme_2").Visible = False
        .Shapes("Dashboard_Theme_3").Visible = True
        .Shapes("Dashboard_Theme_4").Visible = False

        'Color Theme Buttons Status
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_1"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_2"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_3"), True)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_4"), False)
    End With
End Sub

Sub Display_Theme_4()
    With ActiveSheet
        'Dashboard Themes Visibility
        .Shapes("Dashboard_Theme_1").Visible = False
        .Shapes("Dashboard_Theme_2").Visible = False
        .Shapes("Dashboard_Theme_3").Visible = False
        .Shapes("Dashboard_Theme_4").Visible = True

        'Color Theme Buttons Status
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_1"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_2"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_3"), False)
        Call Change_Radio_Button_Status(.Shapes("Theme_Button_4"), True)
    End With
End Sub

'Toggle Buttons

Function Change_Toggle_Button_Status(Toggle_Circle As Object, Toggle_Background As Object, Toggle_Textbook As Object, Active As Boolean)
    If Active = True Then
        With Toggle_Circle
            .Fill.ForeColor.RGB = RGB(0, 176, 80) 'dark green
            .Left = Toggle_Circle.Left + 10 'move Toggle_Circle 17 pt to the right
        End With
        Toggle_Background.Fill.ForeColor.RGB = RGB(0, 204, 0) 'light green
        Toggle_Textbook.TextFrame.Characters.Text = "On"
    Else
        With Toggle_Circle
            .Fill.ForeColor.RGB = RGB(255, 255, 255) 'white
            .Left = Toggle_Circle.Left - 10 'move Toggle_Circle 17 pt to the left
        End With
        Toggle_Background.Fill.ForeColor.RGB = RGB(169, 169, 160) 'light grey

        Toggle_Textbook.TextFrame.Characters.Text = "Off"
    End If
End Function

Sub Change_Info_Visibility()
    With ActiveSheet
    
        If .Shapes("Toggle_Circle_Info").Fill.ForeColor.RGB = RGB(255, 255, 255) Then
            'Visibility Info Elements
            With .Shapes("Icon_Info")
                .Visible = True
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
            End With
            .Shapes("Box_Info").Visible = False
            'Toggle Button Status
            Call Change_Toggle_Button_Status(.Shapes("Toggle_Circle_Info"), .Shapes("Toggle_Background_Info"), .Shapes("Toggle_Textbook_Info"), True)
                                            
        Else
            'Visibility Info Elements
            .Shapes("Icon_Info").Visible = False
            .Shapes("Box_Info").Visible = False
            'Toggle Button Status
            Call Change_Toggle_Button_Status(.Shapes("Toggle_Circle_Info"), .Shapes("Toggle_Background_Info"), .Shapes("Toggle_Textbook_Info"), False)
                                                                                 
        End If
    End With
End Sub


Sub Change_Tabs_Visibility()

    With ActiveSheet
    
        If .Shapes("Toggle_Circle_Tabs").Fill.ForeColor.RGB = RGB(255, 255, 255) Then
        
           .Shapes("Tab_Button_Sales_Revenue").Visible = True
           .Shapes("Tab_Button_Sales_Units").Visible = True
            
            With .Shapes("Tab_Button_Sales_Revenue")
                .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
                .Fill.Transparency = 0#
            End With
        
            With .Shapes("Tab_Button_Sales_Units")
                .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
                .Fill.Transparency = 1#
            End With
        
            .Shapes("Map_Chart_Sales_Revenue").Visible = True
            .Shapes("Line_Chart_Sales_Revenue").Visible = True
            .Shapes("Map_Chart_Sales_Units").Visible = False
            .Shapes("Line_Chart_Sales_Units").Visible = False
            
            Call Change_Toggle_Button_Status(.Shapes("Toggle_Circle_Tabs"), .Shapes("Toggle_Background_Tabs"), .Shapes("Toggle_Textbook_Tabs"), True)
        
        Else
        
            .Shapes("Tab_Button_Sales_Revenue").Visible = False
            .Shapes("Tab_Button_Sales_Units").Visible = False
            
                 
            .Shapes("Map_Chart_Sales_Revenue").Visible = True
            .Shapes("Line_Chart_Sales_Revenue").Visible = True
            .Shapes("Map_Chart_Sales_Units").Visible = False
            .Shapes("Line_Chart_Sales_Units").Visible = False
            
            Call Change_Toggle_Button_Status(.Shapes("Toggle_Circle_Tabs"), .Shapes("Toggle_Background_Tabs"), .Shapes("Toggle_Textbook_Tabs"), False)
        End If
        
        
    End With
        
End Sub








 


