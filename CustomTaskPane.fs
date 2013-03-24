namespace XamlCustomTaskPane
open System.Windows.Forms
open System.Windows.Forms.Integration
open FSharpx
open ExcelDna.Integration.CustomUI

// based on :
// Docking WPF (XAML) UserControls in Task Pane
//    http://www.clear-lines.com/blog/post/Docking-WPF-controls-in-the-VSTO-Task-Pane.aspx
// Building a CustomTaskPane with ExcelDna
//    http://exceldna.codeplex.com/SourceControl/changeset/view/78811#1198220
// FSharpx XAML TypeProvider
//    http://www.navision-blog.de/2012/03/22/wpf-designer-for-f/


/// FSharpx XAML TypeProvider class for our custom UserControl
type TaskPaneContent = XAML<"TaskPaneContent.xaml">

/// our custom WPF UserControl wrapped in a Windows Form UserControl
type public MyUserControl() as this =
    inherit UserControl()
    let _content = TaskPaneContent()
    do
        let wpfElementHost = new ElementHost(Dock = DockStyle.Fill)
        this.Controls.Add(wpfElementHost)
        wpfElementHost.HostContainer.Children.Add(_content.Root) |> ignore
        _content.MyButton.Click.Add(fun _ -> this.MyButton_Click())

    member this.Content with get() = _content
    member this.MyButton_Click() = MessageBox.Show("You clicked the button.") |> ignore  



/// helper module to assist in showing/hiding the custom task pane
module CTPManager =
    let mutable ctp:CustomTaskPane option = None

    let ctp_DockPositionStateChange (ctp:CustomTaskPane) = 
        (ctp.ContentControl :?> MyUserControl).Content.MyLabel.Content <- 
            "Moved to " + ctp.DockPosition.ToString()

    let ctp_VisibleStateChange (ctp:CustomTaskPane) =
        MessageBox.Show("Visibility changed to " + ctp.Visible.ToString()) |> ignore
    
    let ToggleCTP(visible:bool) =
        if ctp.IsNone && visible then
            ctp <- 
              CustomTaskPaneFactory.CreateCustomTaskPane(typeof<MyUserControl>, "My Custom Task Pane")
              |> fun c -> c.Visible <- true
                          c.DockPosition <- MsoCTPDockPosition.msoCTPDockPositionLeft
                          c.add_DockPositionStateChange(fun arg -> ctp_DockPositionStateChange arg)
                          c.add_VisibleStateChange(fun arg -> ctp_VisibleStateChange arg)
                          Some(c)
        elif ctp.IsSome then
            ctp.Value.Visible <- visible



/// CustomUI Ribbon class that uses ribbon XML included in the .dna file
type public MyRibbon() =
    inherit ExcelRibbon()
    member x.OnToggleCTP(control:IRibbonControl, isPressed:bool) = CTPManager.ToggleCTP(isPressed)