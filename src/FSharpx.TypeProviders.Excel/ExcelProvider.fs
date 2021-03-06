﻿module FSharpx.TypeProviders.ExcelProvider

open System.IO
open System
open Samples.FSharp.ProvidedTypes
open Microsoft.FSharp.Core.CompilerServices
open Microsoft.Office.Interop
open System.Runtime.InteropServices
open FSharpx.TypeProviders.Helper
open System.Collections.Generic
open ClosedXML.Excel

//let ApplyMoveToRange (rg:Excel.Range) (move:Excel.XlDirection) = rg.Worksheet.Range(rg, rg.End(move))

// Simple type wrapping Excel data
type  ExcelFileInternal(filename:string, rowsToSkip:int) =
    //TK: document
    //do printfn "filename in constructor = %s" filename
    let dict = new Dictionary<string,obj[][]>()
    let sheetsData = match Path.GetExtension(filename) with
                        |".xlsx"-> let wb = new XLWorkbook(filename) //doc is >= 2007
                                   let mysheets = wb.Worksheets
                                   //mysheets |> Seq.iter (fun s -> printfn "Sht %s" s.Name)
                                   let defnames = wb.NamedRanges
                                   //defnames |> Seq.iter (fun rng -> printfn "Rng %s" rng.Name)
                                   let getData (rng:IXLRange) = 
                                            let sheetData = seq { for r in rng.RowsUsed() do
                                                                        yield seq {for c in r.Cells() do
                                                                                    yield c.GetValue()} |> Array.ofSeq
                                                                } |> Array.ofSeq
                                            sheetData
                                   //add Sheets
                                   for sht in mysheets do
                                       let rng  = sht.RangeUsed()
                                       if rng <> null then
                                           //printfn "getData %s" sht.Name
                                           let data = getData rng
                                           if data.Length > rowsToSkip then
                                                dict.Add(sht.Name,data |> Seq.skip rowsToSkip |> Array.ofSeq)
                                   //add named ranges//TK: refactor for multiple ranges within a namedrange
                                   for namedrng in defnames do
                                      let rng  = namedrng.Ranges |> Seq.exactlyOne
                                      if rng <> null then
                                          //printfn "getData %s" namedrng.Name
                                          let data = getData rng
                                          if data.Length > rowsToSkip then
                                            dict.Add(namedrng.Name,data |> Seq.skip rowsToSkip |> Array.ofSeq)
                                   //dict.Keys |> Seq.iter (fun k -> printfn "%s" k)
                                   dict
                        |".xls"-> let xlApp = new Excel.ApplicationClass()//doc is < 2007, have to use offfice interop. Ho hum ...
                                  xlApp.Visible <- false
                                  xlApp.ScreenUpdating <- false
                                  xlApp.DisplayAlerts <- false;
                                  let xlWorkBookInput = xlApp.Workbooks.Open(filename)
                                  let _mysheets = xlWorkBookInput.Worksheets
                                  let mysheets = seq { for  sheet in _mysheets do yield sheet :?> Excel.Worksheet }
                                  let _names = xlWorkBookInput.Names
                                  let names = seq { for name in _names  do yield name :?> Excel.Name}
                                  let getData (xlRangeInput:Excel.Range) = 
                                    let rows_data = let _rows = xlRangeInput.Rows
                                                    seq { for row  in _rows do
                                                            yield row :?> Excel.Range }
                                    let res = seq { for line_data in rows_data do 
                                                        yield ( seq { for cell in line_data.Columns do
                                                                        if (cell  :?> Excel.Range).Value2 <> null && (cell  :?> Excel.Range).Value2.ToString() <> String.Empty then
                                                                            if (((cell  :?> Excel.Range).NumberFormat).ToString().Contains("d")) then //this is a date with a dd in the format.
                                                                                match Double.TryParse(((cell  :?> Excel.Range).Value2).ToString()) with
                                                                                |true,dtDbl -> yield box (DateTime.FromOADate(dtDbl))
                                                                                |false,_    -> yield (cell  :?> Excel.Range).Value2 
                                                                            else
                                                                                yield (cell  :?> Excel.Range).Value2
                                                                    }
                                                                     |> Seq.filter (fun c -> c.ToString() <> String.Empty) |>Seq.toArray
                                                               )
                                                  }
                                               |> Seq.toArray |> Array.filter (fun r-> r.Length > 0)
                                    res
                                  for sht in mysheets do
                                        let xlRangeInput = sht.UsedRange
                                        if xlRangeInput <> null then
                                            let data = getData xlRangeInput
                                            if data.Length > rowsToSkip then
                                                dict.Add(sht.Name,data |> Seq.skip rowsToSkip |> Array.ofSeq)
                                  for rng in names do
                                        let xlRangeInput = rng.RefersToRange
                                        if xlRangeInput <> null then
                                            let data = getData xlRangeInput
                                            if data.Length > rowsToSkip then
                                                dict.Add(rng.Name,data |> Seq.skip rowsToSkip |> Array.ofSeq)
                                  xlWorkBookInput.Close()
                                  // Cleanup
                                  GC.Collect();
                                  GC.WaitForPendingFinalizers();
                                  Marshal.FinalReleaseComObject(_names) |> ignore
                                  Marshal.FinalReleaseComObject(_mysheets) |> ignore
                                  xlWorkBookInput.Close(Type.Missing, Type.Missing, Type.Missing)
                                  Marshal.FinalReleaseComObject(xlWorkBookInput) |> ignore
                                  xlApp.Quit()
                                  Marshal.FinalReleaseComObject(xlApp) |> ignore
                                  //dict.Keys |> Seq.iter (printfn "%s")
                                  dict
                        |_     -> failwithf "%s is not a valid path for a spreadsheet " filename
    member __.SheetAndRangeNames = dict.Keys |> Array.ofSeq
    member __.SheetData(name:string)  = sheetsData.[name]

(*type internal ReflectiveBuilder = 
       static member Cast<'a> (args:obj) =
          args :?> 'a
       static member BuildTypedCast lType (args: obj) = 
             typeof<ReflectiveBuilder>
                .GetMethod("Cast")
                .MakeGenericMethod([|lType|])
                .Invoke(null, [|args|])
*)
type internal GlobalSingleton private () =
    static let mutable instance = Dictionary<_, _>()
    static member Instance = instance

let internal memoize f =
      //let cache = Dictionary<_, _>()
      fun x ->
         if (GlobalSingleton.Instance).ContainsKey(x) then (GlobalSingleton.Instance).[x]
         else let res = f x
              (GlobalSingleton.Instance).[x] <- res
              res

let internal typExcel(cfg:TypeProviderConfig) =
      // Create the main provided type
      let excTy = ProvidedTypeDefinition(System.Reflection.Assembly.GetExecutingAssembly(), rootNamespace, "ExcelFile", Some(typeof<ExcelFileInternal>))
      do excTy.AddXmlDoc("The main provided type - static parameters of filename:string, forcestring:bool, headerrow:int. \n If forcestring, all data will be coerced to string")
      let defaultHeaderRow = 1
      // Parameterize the type by the file to use as a template
      let filename = ProvidedStaticParameter("filename", typeof<string>)
      let forcestring = ProvidedStaticParameter("forecstring", typeof<bool>,false)
      let headerRow = ProvidedStaticParameter("headerrow", typedefof<int>, defaultHeaderRow)
      let staticParams = [filename
                          forcestring
                          headerRow]
      do excTy.DefineStaticParameters(staticParams, fun tyName paramValues ->
        let filename,forcestring,headerRow = 
                                   match paramValues with
                                   | [| :? string  as filename;   :? bool as forcestring; :? int as headerRow|] -> (filename, forcestring,headerRow)
                                   | [| :? string  as filename;   :? bool as forcestring|] -> (filename, forcestring, defaultHeaderRow)
                                   | [| :? string  as filename|] -> (filename, false, defaultHeaderRow)
                                   | _ -> ("no file specified to type provider",  true, defaultHeaderRow)
        let resolvedFilename = Path.Combine(cfg.ResolutionFolder, filename)
        //printfn "resFileName = %s; filename = %s" resolvedFilename filename
        let ex = ExcelFileInternal(resolvedFilename, (headerRow - 1) )
        // define the provided type, erasing to excelFile
        let ty = ProvidedTypeDefinition(System.Reflection.Assembly.GetExecutingAssembly(), rootNamespace, tyName, Some(typeof<ExcelFileInternal>))
        
        // add a parameterless constructor
        ty.AddMember(ProvidedConstructor([], InvokeCode = fun [] -> <@@  ExcelFileInternal(resolvedFilename,(headerRow - 1) ) @@>))
        //ty.AddMember(ProvidedConstructor([], InvokeCode = fun [] -> <@@  ex @@>))
        ty.AddMember(ProvidedConstructor([ProvidedParameter("filename", typeof<string>)], InvokeCode = fun [filename] -> <@@  ExcelFileInternal(%%filename,(headerRow - 1) ) @@>))
        //for each worksheet (with data), add a property of provided type shtTyp
        for sht in ex.SheetAndRangeNames do
            let shtTyp = if  forcestring then
                            ProvidedTypeDefinition(sht,Some typeof<string[][]>,HideObjectMethods = true)
                         else
                            ProvidedTypeDefinition(sht,Some typeof<obj[][]>,HideObjectMethods = true)
            do shtTyp.AddXmlDoc(sprintf "Type for data in %s" sht)
            let data = ex.SheetData(sht)
            let rowTyp = ProvidedTypeDefinition("Row", 
                                                (if forcestring then 
                                                    Some typeof<string[]>
                                                else 
                                                    Some typeof<obj[]>), 
                                                HideObjectMethods = true)
            shtTyp.AddMember(rowTyp)
            let rowsProp = ProvidedProperty(propertyName = "Rows",
                                            propertyType = typedefof<seq<_>>.MakeGenericType(rowTyp),
                                            GetterCode = if forcestring then 
                                                            (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:string[][])
                                                                                                  |> Seq.skip 1
                                                                                                  |> Array.ofSeq 
                                                                                                  //|> Array.map ( fun row -> row |> Array.map (fun cel -> cel.ToString())) 
                                                                                                 @@>)
                                                         else
                                                            (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:obj[][])
                                                                                                  |> Seq.skip 1
                                                                                                  |> Array.ofSeq 
                                                                                                 @@>)
                                                         )
            let  colHdrs = data.[0]
            colHdrs |> Array.iteri (fun j col -> let propName = match col.ToString() with
                                                                |"" -> "Col" + j.ToString()
                                                                |_  ->  col.ToString()
                                                 let valueType, gettercode  = if forcestring then typeof<string>,(fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:string[]).[j] @@>)
                                                                              else
                                                                              match data.[1].[j] with
                                                                              | :? bool      -> typeof<bool>,    (fun (args:Quotations.Expr list) -> <@@ bool.Parse(((%%args.[0]:obj[]).[j]).ToString()) @@>)
                                                                              | :? string    -> typeof<string>,  (fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:obj[]).[j]).ToString() @@>)
                                                                              | :? DateTime  -> typeof<DateTime>,(fun (args:Quotations.Expr list) -> <@@ DateTime.Parse(((%%args.[0]:obj[]).[j]).ToString()) @@>)
                                                                              | :? float     -> typeof<float>,   (fun (args:Quotations.Expr list) -> <@@ Double.Parse(((%%args.[0]:obj[]).[j]).ToString()) @@>)
                                                                              |_             -> typeof<obj>,     (fun (args:Quotations.Expr list) -> <@@  (%%args.[0]:obj[]).[j] @@>)
                                                 let colp = ProvidedProperty(propertyName = propName,
                                                                             propertyType = valueType,
                                                                             GetterCode= gettercode)
                                                 rowTyp.AddMember(colp))
            data |> Array.iteri (fun i r -> if i > 0 then //skip header col
                                                let rowTyp =  if  forcestring then
                                                                ProvidedTypeDefinition("Row" + i.ToString(),Some typeof<string[]>,HideObjectMethods = true)
                                                              else
                                                                ProvidedTypeDefinition("Row" + i.ToString(),Some typeof<obj[]>,HideObjectMethods = true)
                                                let getCode = if forcestring then
                                                                (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:string[][]).[i] @@>)
                                                              else
                                                                (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:obj[][]).[i] @@>)
                                                let rowp = ProvidedProperty(propertyName = "Row" + i.ToString(),
                                                                            propertyType = rowTyp,
                                                                            GetterCode = getCode
                                                                            )
                                                colHdrs |> Array.iteri (fun j col -> let propName = match col.ToString() with
                                                                                                    |"" -> "Col" + j.ToString()
                                                                                                    |_  ->  col.ToString()
                                                                                     let valueType, gettercode  = if forcestring then typeof<string>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:string[])).[j] @@>)
                                                                                                                  else
                                                                                                                  match r.[j] with
                                                                                                                  | :? bool      -> typeof<bool>,    (fun (args:Quotations.Expr list) -> <@@ bool.Parse(((%%args.[0]:obj[]).[j]).ToString()) @@>)
                                                                                                                  | :? string    -> typeof<string>,  (fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:obj[]).[j]).ToString() @@>)
                                                                                                                  | :? DateTime  -> typeof<DateTime>,(fun (args:Quotations.Expr list) -> <@@ DateTime.Parse(((%%args.[0]:obj[]).[j]).ToString()) @@>)
                                                                                                                  | :? float     -> typeof<float>,   (fun (args:Quotations.Expr list) -> <@@ Double.Parse(((%%args.[0]:obj[]).[j]).ToString()) @@>)
                                                                                                                  |_             -> typeof<obj>,     (fun (args:Quotations.Expr list) -> <@@  (%%args.[0]:obj[]).[j] @@>)
                                                                                     let colp = ProvidedProperty(propertyName = propName,
                                                                                                                 propertyType = valueType,
                                                                                                                 GetterCode= gettercode)
                                                                                     colp.AddXmlDoc(sprintf "Value for Cell in Col%d in Row%d in range %s" j i sht)
                                                                                     rowTyp.AddMember(colp)
                                                                       )
                                                shtTyp.AddMember(rowTyp)
                                                rowp.AddXmlDoc(sprintf "Data for Row%d in range %s" i sht)
                                                shtTyp.AddMember(rowp)
                                )
            //data |> Array
            let shtGetCode = if forcestring then
                                (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:ExcelFileInternal).SheetData(sht) |> Array.map ( fun row -> row |> Array.map (fun cel -> cel.ToString())) @@>)
                             else
                                (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:ExcelFileInternal).SheetData(sht) @@>)
            let shtp = ProvidedProperty(propertyName = sht, 
                                        propertyType = shtTyp,
                                        GetterCode= shtGetCode
                                       )
            do shtp.AddXmlDoc(sprintf "Data in %s" sht)
            shtTyp.AddMember(rowsProp)
            ty.AddMember(shtTyp)
            ty.AddMember(shtp)
        ty
        //(memoize ProvidedTypeDefinitionExcelCall)(filename, sheetorrangename ,  forcestring)
        )
      excTy

[<TypeProvider>]
type public ExcelProvider(cfg:TypeProviderConfig) as this =
    inherit TypeProviderForNamespaces()
    do this.AddNamespace(rootNamespace,[typExcel cfg])

[<TypeProviderAssembly>]
do ()