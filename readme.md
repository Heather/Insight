blah blah blah everybody sing along
-----------------------------------

```fsharp
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
```