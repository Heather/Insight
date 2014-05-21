#light (*
	exec fsharpi --exec $0 $@
*)

#r "Fakelib.dll"
#r "System.Data.dll"

#I @"packages\ModernUI\lib\net40"
#I @"packages\FSharp.Charting\lib\net40"

#r "MetroFramework.dll"
#r "FSharp.Charting.dll"

open System
open System.IO
open System.Text
open System.Drawing
open System.Windows.Forms

open System.Data
open System.Data.SqlClient

open Fake.StringHelper
open FSharp.Charting
open FSharp.Charting.ChartTypes

open MetroFramework.Forms
open MetroFramework.Controls

let form = new MetroForm()
form.Width  <- 800
form.Height <- 750
form.Text   <- "Insight"
form.Font   <- new Font( "Lucida Console"
                       , 12.0f
                       , FontStyle.Regular,GraphicsUnit.Point )

let b1 = new MetroButton();
b1.Location <- Point(40, 680); b1.Size <- Size(150, 50); b1.Text <- "Exit"
b1.Anchor <- (AnchorStyles.Bottom ||| AnchorStyles.Left)
let b2 = new MetroButton();
b2.Location <- Point(580, 680); b2.Size <- Size(150, 50); b2.Text <- "Go"
b2.Anchor <- (AnchorStyles.Bottom ||| AnchorStyles.Right)

let menuStrip = new ContextMenuStrip();
let fileMenu = new ToolStripMenuItem();
fileMenu.Text <- "File";
let aboutMenu = new ToolStripMenuItem();
aboutMenu.Text <- "About";

let exitM = new ToolStripMenuItem();
exitM.Text <- "Exit";

fileMenu.DropDownItems.Add exitM
menuStrip.Items.Add fileMenu
menuStrip.Items.Add aboutMenu
form.ContextMenuStrip <- menuStrip

let data1 = [for x in 0.0 .. 0.1 .. 6.0 -> sin x + cos (2.0 * x)]
let data2 = [for x in 0.0 .. 0.1 .. 6.0 -> cos x + sin (2.0 * x)]

let myChart = 
    Chart.Combine(
        [   Chart.Line(data1,Name="data1-1")
            Chart.Line(data2,Name="data2-2") ])

let myChartControl = new ChartControl(myChart, Dock=DockStyle.Top);
myChartControl.Height <- 500

aboutMenu.Click.Add (fun _ -> 
    MessageBox.Show("Insight v.0.0.1") |> ignore
)

exitM.Click.Add (fun _ -> ignore <| form.Close())
b1.Click.Add    (fun _ -> ignore <| form.Close())
b2.Click.Add    (fun _ -> ())

form.Controls.AddRange [|myChartControl; b1; b2|]
Application.Run(form)
