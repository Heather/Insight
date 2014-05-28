#light (*
	exec fsharpi --exec $0 $@
*)

#r "Fakelib.dll"
#r "System.Data.dll"

#I @"packages\ModernUI\lib\net40"
#I @"packages\FSharp.Charting\lib\net40"
#I @"packages\NPOI\lib\net40"

#r "MetroFramework.dll"
#r "FSharp.Charting.dll"
#r "NPOI.dll"

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

open NPOI.HSSF.UserModel

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

let make(fname, shName, varCol,startRow,endRow) =
        let cXL name =  
            if name <> "" then
               (name.ToLower().ToCharArray()
                |> Seq.map (fun char -> Convert.ToInt32 char - 96)
                |> Seq.sumBy(fun x -> x + 25)) - 26
            else 0
        if [|fname;varCol;startRow;endRow|]
            |> Seq.forall(fun x -> (x <> "" && x <> null)) then 
            using(new FileStream(fname, FileMode.Open, FileAccess.Read))<| fun fs ->
                let templateWorkbook = new HSSFWorkbook(fs, true)
                let sheet = templateWorkbook.GetSheet(shName)
                let cvar  = cXL varCol
                let doMathAndSave sr er =
                    [ for i in sr..er -> try Double.Parse(sheet.GetRow(i-1).GetCell(cvar).ToString())
                                         with _ -> 0.0 ]
                let sr = try Int32.Parse startRow 
                         with _ -> 0
                let er = match endRow with
                            | "0" -> let rec counter cn =
                                        try ignore <| sheet.GetRow(cn).GetCell(cvar)
                                            counter (cn+1)
                                        with _ -> (cn-1) 
                                     counter 0
                            | _ -> try Int32.Parse endRow
                                   with _ -> 0 
                doMathAndSave sr er
        else [] (*
            using(new MemoryStream()) <| fun ms ->  
                templateWorkbook.Write(ms)         
                let msA = ms.ToArray()
                using(new FileStream((@"X.xls"), FileMode.OpenOrCreate , FileAccess.Write))
                <| fun newF ->
                    try newF.Write(msA,0,msA.Length)
                        MessageBox.Show( "X.xls created, check the result" ) |> ignore
                    with _ -> MessageBox.Show( "Can't write to file" ) |> ignore
                            *)


                        
let data1 = make("olya.xls", "olya", "F", "2", "101")
let data2 = make("marina.xls", "marina", "F", "2", "101")

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
