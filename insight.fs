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

(* power pipe *)
let inline (/>) f a = f a

let form = new MetroForm()
form.Width  <- 650
form.Height <- 800
form.Text   <- "Insight"
form.Font   <- new Font( "Lucida Console"
                       , 12.0f
                       , FontStyle.Regular,GraphicsUnit.Point )

let b1 = new MetroButton();
b1.Location <- Point(40, 730); b1.Size <- Size(150, 50); b1.Text <- "Exit"
b1.Anchor <- (AnchorStyles.Bottom ||| AnchorStyles.Left)
let b2 = new MetroButton();
b2.Location <- Point(450, 730); b2.Size <- Size(150, 50); b2.Text <- "Export"
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

let make fname varCol startRow endRow =
        let cXL name =  
            if name <> "" then
               (name.ToLower().ToCharArray()
                |> Seq.map   /> fun char -> Convert.ToInt32 char - 96
                |> Seq.sumBy (fun x -> x + 25)) - 26
            else 0
        if [|fname;varCol;startRow;endRow|]
            |> Seq.forall(fun x -> (x <> "" && x <> null)) then 
            using(new FileStream(fname, FileMode.Open, FileAccess.Read)) <| fun fs ->
                let templateWorkbook = new HSSFWorkbook(fs, true)
                let sheet = templateWorkbook.GetSheetAt(0) //.GetSheet(shName)
                let cvar  = cXL varCol
                let getData sr er =
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
                getData sr er
        else []

let groupConsecutive sq = 
    (Array.ofSeq sq, [])
    ||> Array.foldBack /> fun x -> function
        | [] -> [[x]]
        | xs :: xss -> if x = List.head xs then (x :: xs) :: xss
                                           else [x] :: xs :: xss

let sameData d = groupConsecutive d
                 |> Seq.groupBy  /> fun s -> Seq.length s
                 |> Seq.filter   /> fun (n, _) -> n > 1
                 |> Seq.map      /> fun (n, g) -> (n, (Seq.length g))

let dh (data : seq<list<float>>) a6 = 
    let result = data |> Seq.map        /> fun l -> sameData l
                      |> Seq.concat
                      |> Seq.groupBy    /> fun (n, _)  ->  n
                      |> Seq.map        /> fun (nn, s) -> (nn, (s |> Seq.map(fun (_, c) -> c)
                                                                  |> Seq.sum ))
                      |> Seq.map        /> fun (n, c)  -> (n.ToString(), c)
    Chart.Doughnut(result, a6)
    
let sameDataX d = groupConsecutive d
                 |> Seq.groupBy  /> fun s -> Seq.length s
                 |> Seq.filter   /> fun (n, _) -> n > 1
                 |> Seq.map      /> fun (n, g) -> (n, (g |> Seq.nth 0).[0], (Seq.length g))
let dhX (data : seq<list<float>>) a6 = 
    let result = data |> Seq.map        /> fun l -> sameDataX l
                      |> Seq.concat
                      |> Seq.groupBy    /> fun (n, v, _)  -> n
                      |> Seq.map        /> fun (nn, s) -> (nn, (s |> Seq.map(fun (_, v, _) -> v))
                                                             , (s |> Seq.map(fun (_, _, c) -> c)
                                                                  |> Seq.sum ))
                      |> Seq.map        /> fun (n, vv, c)  -> vv  |> Seq.map(fun v -> 
                                                                                (n.ToString() + ": " + v.ToString()), c)
                      |> Seq.concat
    Chart.Doughnut(result, a6)

let sameDataY d = (groupConsecutive (d |> Seq.concat)) |> Seq.map /> fun s -> ( Seq.length s, s |> Seq.nth 0 )
let herLogicsThere d = let dataY    = sameDataY (d : seq<float list>)
                       let dataYlen = Seq.length dataY
                       dataY
                      |> Seq.mapi   /> fun i (n, v) ->
                                        if n > 1 && i + 3 < dataYlen then
                                            let (nn, vn)   = dataY |> Seq.nth (i + 1)
                                            let (nnn, vnn) = dataY |> Seq.nth (i + 2)
                                            let (nnnn, vnnn) = dataY |> Seq.nth (i + 3)
                                            (n, v, vn, vnn, vnnn)
                                        else (n, v, 0.0, 0.0, 0.0)
                      |> Seq.filter /> fun (n, _, _, _, _) -> n > 1
                      |> Seq.mapi   /> fun i (n, v1, v2, v3, v4) -> Chart.SplineArea ([v1; v2; v3; v4], "dataH" + i.ToString())
                      |> Seq.toList
    
let xls = (new DirectoryInfo(".")).GetFiles()
          |> Seq.filter /> fun f -> f.Name.EndsWith(".xls")
          |> Seq.map    /> fun f -> f.Name
let xlsData col start fin = xls |> Seq.map /> fun n -> make n col start fin

let xlsRF = xlsData "F" "2" "101"
let xlsRD = xlsData "D" "2" "101"

let hData = herLogicsThere xlsRD 
let Fdata = xlsRF |> Seq.mapi /> fun i d -> Chart.Line (d, "dataF" + i.ToString())
                  |> Seq.toList
let Ddata = xlsRD |> Seq.mapi /> fun i d -> Chart.SplineArea( d, "dataD" + i.ToString())
                  |> Seq.toList

let ds xs = new ChartControl(Chart.Combine xs, Dock=DockStyle.Top)

let dataset2 = ds [ dhX <| xlsData "F" "2" "101"
                        <| "DH"
    ]
let dataset1 = ds <| Fdata @ Ddata
let dataset3 = ds <| hData

let Graphs : Control array = [|
    dataset1
    dataset2
    dataset3
|]

let he() = (form.Height - 200) / Graphs.Length
Graphs |> Array.iteri(fun i g ->
    let cc = g :?> ChartControl
    cc.Height <- he()
    cc.Resize.Add /> fun _ -> 
        cc.Height <- he()
)

aboutMenu.Click.Add /> fun _ -> 
    MessageBox.Show("Insight v.0.0.1") |> ignore

exitM.Click.Add /> fun _ -> ignore <| form.Close()
b1.Click.Add    /> fun _ -> ignore <| form.Close()
b2.Click.Add    /> fun _ -> using(new MemoryStream()) <| fun ms ->  
                            let templateWorkbook = new HSSFWorkbook()
                            templateWorkbook.Write(ms)         
                            let msA = ms.ToArray()
                            using(new FileStream((@"X.xls"), FileMode.OpenOrCreate , FileAccess.Write))
                            <| fun newF ->
                                try newF.Write(msA,0,msA.Length)
                                    MessageBox.Show( "X.xls created, check the result" ) |> ignore
                                with _ -> MessageBox.Show( "Can't write to file" )       |> ignore

form.Controls.AddRange <| Array.concat [Graphs; [|b1; b2|]]
Application.Run(form)
