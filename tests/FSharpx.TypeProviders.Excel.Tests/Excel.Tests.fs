module FSharpx.TypeProviders.Tests.ExcelTests

open NUnit.Framework
open FSharpx
open FsUnit

open System
open System.IO

type BookTestXLSForceString = ExcelFile<"BookTest.xls", true>
type BookTestXLSXForceString = ExcelFile<"BookTest.xlsx", true>

type BookTestXLS = ExcelFile<"BookTest.xls", false>
type BookTestXLSX= ExcelFile<"BookTest.xlsx", false>

type HeaderTest = ExcelFile<"BookTestWithHeader.xlsx", true,2>

let xlsFileForceString = BookTestXLSForceString()
let xlsxFileForceString = BookTestXLSXForceString()

let xlsFile = BookTestXLS()
let xlsxFile = BookTestXLS()
//let row1 = file.Data |> Seq.head 

[<Test>]
let ``Can access first row as string excel data by interop``() = 
    let row1 = xlsFileForceString.Sheet1.Row1
    row1.SEC |> should equal "ASI"
    row1.BROKER |> should equal "TFS Derivatives HK"

[<Test>]
let ``Can access typed excel row data by closedXML``() = 
    let row1 = xlsxFile.Sheet2.Row1
    row1.Date |> should equal (new DateTime(2013,1,1))
    row1.Bool |> should equal true

[<Test>]
let ``Can access typed excel row data by interop``() = 
    let row1 = xlsxFile.Sheet2.Row1
    row1.Date |> should equal (new DateTime(2013,1,1))
    row1.Bool |> should equal true

[<Test>]
let ``Can access typed row as string  data by closedXML``() = 
    let row1 = xlsxFileForceString.Sheet1.Row1
    row1.SEC |> should equal "ASI"
    row1.BROKER |> should equal "TFS Derivatives HK"


[<Test>]
let ``Can pick an arbitrary header row``() =
    let file = HeaderTest()
    let row = file.Sheet1.Row1
    row.SEC |> should equal "ASI"
    row.BROKER |> should equal "TFS Derivatives HK"

[<Test>]
let ``Can load data from spreadsheet using filename .ctor``() =
    let file = Path.Combine(Environment.CurrentDirectory, "BookTestDifferentData.xls")

    let otherBook = BookTestXLSForceString(file)
    let row = otherBook.Sheet1.Row1

    row.SEC |> should equal "TASI"
    row.STYLE |> should equal "B"
    row.``STRIKE 1`` |> should equal "3"
    row.``STRIKE 2`` |> should equal "4"
    row.``STRIKE 3`` |> should equal "5"
    row.VOL |> should equal "322"
