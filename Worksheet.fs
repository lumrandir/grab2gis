namespace Grab2Gis

open System
open Microsoft.Office.Interop.Excel
open FSharp.Data

module Worksheet =
  let extractString (value : JsonValue) : string =
    match value with
    | JsonValue.String s -> s
    | _                  -> ""

  let extractProperty (prop : string) (jsn : JsonValue) : string =
    match jsn.TryGetProperty(prop) with
    | Some s -> s.AsString()
    | None   -> ""

  let extractRecord (rc : JsonValue) : (string * string) [] =
    match rc with
    | JsonValue.Record r -> Array.map (fun (tp, jval) -> tp, extractString jval) r
    | _                  -> [||]

  let contains (field : string) (ary : (string * string) []) : bool =
    Array.exists (fun (tag, value) -> tag = "type" && value = field) ary

  let filterByAttr (field : string) (ary : (string * string) [] []) : (string * string) [] [] =
    Array.filter (contains field) ary

  let extractAttr (field : string) (ary : (string * string) []) : string =
    match Array.find (fun (tag, _) -> tag = "value") ary with
    | _, value -> value

  let extractContacts (lst : JsonValue) (field : string) : string =
    match lst with
    | JsonValue.Array a ->
      Array.map extractRecord a |> filterByAttr field
        |> Array.map (extractAttr field) |> String.concat ", "
    | _                 -> ""

  let saveExcel (data : JsonValue []) (filename : string) : unit =
    let xlApp = new ApplicationClass()
    let xlOutBook = xlApp.Workbooks.Add()
    let xlOutSheet = xlOutBook.Worksheets.[1] :?> Worksheet
    xlOutSheet.Name <- "Результаты"

    xlOutSheet.Cells.Range("A1", "E1").Value2 <- Array.map box [| "Адрес"; "Email"; "Сайт"; "Название"; "Телефон" |]
    for i in 0 .. data.Length - 1 do
      let num = string (i + 2)
      let cgroups = match data.[i].["contact_groups"] with
                    | JsonValue.Array [||] -> [||]
                    | JsonValue.Array ar   -> ar
                    | _                    -> [||]
      if cgroups.Length > 0 then
        let contacts = cgroups.[0].["contacts"]
        let rowData = [|
          extractProperty "address_name" data.[i];
          extractContacts contacts "email";
          extractContacts contacts "website";
          extractProperty "name" data.[i];
          extractContacts contacts "phone"
        |]
        xlOutSheet.Cells.Range("A" + num, "E" + num).Value2 <- Array.map box rowData
      else
        ()

    xlOutBook.SaveAs(box(filename), XlFileFormat.xlWorkbookDefault, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing,
                     XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing)
    xlApp.Visible <- true
    ()
