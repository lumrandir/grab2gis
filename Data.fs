namespace Grab2Gis

open System.Collections
open System.Net
open FSharp.Data
open FSharp.Data.HttpRequestHeaders

module Data =
  type City(name0 : string, regionId0 : int) =
    let name     = name0
    let regionId = regionId0
    member this.Name   with get() = name
    member this.Region with get() = regionId

  let cities : ArrayList =
    new ArrayList([| new City("Новосибирск", 1);
                     new City("Екатеринбург", 9);
                     new City("Москва", 32);
                     new City("Санкт-Петербург", 38) |])

  let key = "asdfsdf"
  let fields = "items.adm_div,items.contact_groups,items.flags,items.address,items.name_ex,items.external_content,items.org"
  let url = "http://catalog.api.2gis.ru/2.0/catalog/branch/search"

  let rec fetch0 (q : string) (id : int) (page : int) (items : JsonValue []) : JsonValue [] option =
    let queryParams = [ "page", (string page); "q", q; "page_size", (string 50);
                        "output", "json"; "key", key; "region_id", (string id);
                        "fields", fields; "sort", "rating" ]
    let result = JsonValue.Parse(Http.RequestString(url, httpMethod = "GET", query = queryParams,
                                                    headers = [ Accept HttpContentTypes.Json ]))
    match result.["meta"].["code"] with
    | JsonValue.Number x when x = decimal 200 ->
      match result.["result"].["items"] with
      | JsonValue.Array jvs -> fetch0 q id (page + 1) (Array.append items jvs)
      | _                   -> Some(items)
    | JsonValue.Number x when x = decimal 403 ->
      None
    | _                                       ->
      Some(items)

  let fetch (query : string) (regionId : int) : JsonValue [] option =
    fetch0 query regionId 1 [||]
