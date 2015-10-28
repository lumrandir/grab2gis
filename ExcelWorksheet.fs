namespace Grab2Gis

open System
open FSharp.Data
open ExcelGenerator.SpreadSheet
open Data

module ExcelWorksheet =
  let saveExcel (data : JsonValue []) (filename : string) : unit =
    let wb = new Workbook()
    let ws = new Worksheet("Результаты")
    let titleRow = new Row()
    List.map (fun t -> new Cell(t)) ["Адрес"; "Email"; "Сайт"; "Название"; "Телефон"]
      |> titleRow.Cells.AddRange
    ws.Rows.Add(titleRow)

    for elem in data do
      let cgroups = match elem.["contact_groups"] with
                    | JsonValue.Array [||] -> [||]
                    | JsonValue.Array ar   -> ar
                    | _                    -> [||]
      if cgroups.Length > 0 then
        let contacts = cgroups.[0].["contacts"]
        let rowData  = [
          extractProperty "address_name" elem;
          extractContacts contacts "email" "value";
          extractContacts contacts "website" "text";
          extractProperty "name" elem;
          extractContacts contacts "phone" "value";
        ]
        let row = new Row()
        List.map (fun x -> new Cell(x)) rowData
          |> row.Cells.AddRange
        ws.Rows.Add(row)

    wb.Worksheets.Add(ws)
    wb.save(filename)

