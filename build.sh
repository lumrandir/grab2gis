#!/bin/sh

exec fsharpc Data.fs ExcelWorksheet.fs MainForm.fs Main.fs -r:FSharp.Data --staticlink:FSharp.Data -r:ExcelGenerator --staticlink:ExcelGenerator --standalone --target:winexe
