namespace Grab2Gis

open System
open System.Windows.Forms
open System.Drawing
open System.IO
open Data
open Worksheet

type MainForm() as form =
  inherit Form()

  let queryLabel = new Label(Text = "Запрос", Location = new Point(10, 10))
  let queryText  = new TextBox(Location = new Point(10, 35), Size = new Size(260, 100))
  let cityLabel  = new Label(Text = "Город", Location = new Point(10, 60))
  let cityCombo  = new ComboBox(Location = new Point(10, 85),
                                Size = new Size(260, 100), ValueMember = "Region",
                                DisplayMember = "Name", DataSource = cities)
  let processBtn = new Button(Location = new Point(190, 130),
                              Size = new Size(80, 30), Text = "Граб")
  let debugLabel = new Label(Text = "", Location = new Point(10, 160))

  do form.InitializeForm

  member this.onProcessBtnClick =
    let query    = queryText.Text
    let regionId = unbox<int> cityCombo.SelectedValue
    let saveFile = new SaveFileDialog()
    saveFile.FileName <- query
    saveFile.Filter   <- "Excel 2007 files (*.xlsx)|*.xlsx"
    saveFile.InitialDirectory <- Directory.GetCurrentDirectory()
    if saveFile.ShowDialog(form) = DialogResult.OK
      then saveExcel (fetch query regionId) saveFile.FileName
      //then saveExcel [||] saveFile.FileName
      else ()

  member this.InitializeForm =
    this.Width           <- 300
    this.Height          <- 300
    this.Text            <- "Скорограббер"
    this.MaximizeBox     <- false
    this.FormBorderStyle <- FormBorderStyle.FixedDialog
    processBtn.Click.Add(fun _ -> do form.onProcessBtnClick)
    this.Controls.AddRange(
      [| (queryLabel :> Control);
         (queryText  :> Control);
         (cityLabel  :> Control);
         (cityCombo  :> Control);
         (processBtn :> Control);
         (debugLabel :> Control)
      |] )
