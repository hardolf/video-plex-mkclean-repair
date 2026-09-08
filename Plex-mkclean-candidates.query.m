// Power Query for the Plex-mkclean-candidates workbooks.
//
// Kept here as text because the query otherwise exists only inside a binary
// .xlsx, where git cannot see it and a lost workbook takes it with it.
//
// To put it into a workbook: Data > Get Data > Launch Power Query Editor >
// select the query > Home > Advanced Editor > replace everything with this.
//
// The workbook needs a sheet named Config carrying two named cells:
//   WorkbookFolder  =LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)
//   WorkbookName    =MID(CELL("filename"),FIND("[",CELL("filename"))+1,FIND("]",CELL("filename"))-FIND("[",CELL("filename"))-1)
//
// Columns the query does not name are passed straight through by Power Query,
// so a report gaining a column still loads. The five remux columns are named
// anyway, so they arrive typed and with the same two-line headers as the rest.

let
    // The source always follows this workbook: a workbook named <name>.xlsx
    // reads <name>.csv beside it, and saving under a new name takes the source
    // along.
    Folder = Excel.CurrentWorkbook(){[Name="WorkbookFolder"]}[Content]{0}[Column1],
    WorkbookFile = Excel.CurrentWorkbook(){[Name="WorkbookName"]}[Content]{0}[Column1],
    CsvFile = Text.BeforeDelimiter(WorkbookFile, ".", {0, RelativePosition.FromEnd}) & ".csv",
    Source = Csv.Document(
        File.Contents(Folder & CsvFile),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Promoted Headers",
        {
            {"ScanStatus", type text},
            {"Candidate", type logical},
            {"StrictCandidate", type logical},
            {"IsMkclean", type logical},
            {"HasHevc", type logical},
            {"HevcTrackCompressed", type logical},
            {"CodecPrivateScope2", type logical},
            {"UnknownHevcProfile", type logical},
            {"ContentCompression", type logical},
            {"HevcProfile", type text},
            {"HevcTrackId", type text},
            {"CompressionAlgorithms", type text},
            {"WritingApplication", type text},
            {"MkvMergeExitCode", Int64.Type},
            {"MkvInfoExitCode", Int64.Type},
            {"SizeGiB", type number},
            {"SizeBytes", Int64.Type},
            {"LastWriteTime", type datetime},
            {"RemuxStatus", type text},
            {"RemuxTime", type datetime},
            {"RemuxHevcProfile", type text},
            {"RemuxSizeBytes", Int64.Type},
            {"BackupPath", type text},
            {"Root", type text},
            {"FileName", type text}
        },
        "en-US"
    ),
    #"Renamed Columns" = Table.RenameColumns(
        #"Changed Type",
        {
            {"ScanStatus", "Scan#(lf)Status"},
            {"Candidate", "Candidate"},
            {"StrictCandidate", "Strict#(lf)Candidate"},
            {"IsMkclean", "Is#(lf)Mkclean"},
            {"HasHevc", "Has#(lf)HEVC"},
            {"HevcTrackCompressed", "HEVC Track#(lf)Compressed"},
            {"CodecPrivateScope2", "Codec Private#(lf)Scope 2"},
            {"UnknownHevcProfile", "Unknown HEVC#(lf)Profile"},
            {"ContentCompression", "Content#(lf)Compression"},
            {"HevcProfile", "HEVC#(lf)Profile"},
            {"HevcTrackId", "HEVC#(lf)Track ID"},
            {"CompressionAlgorithms", "Compression#(lf)Algorithms"},
            {"WritingApplication", "Writing#(lf)Application"},
            {"MkvMergeExitCode", "MkvMerge#(lf)Exit Code"},
            {"MkvInfoExitCode", "MkvInfo#(lf)Exit Code"},
            {"SizeGiB", "Size#(lf)GiB"},
            {"SizeBytes", "Size#(lf)Bytes"},
            {"LastWriteTime", "Last Write#(lf)Time"},
            {"RemuxStatus", "Remux#(lf)Status"},
            {"RemuxTime", "Remux#(lf)Time"},
            {"RemuxHevcProfile", "Remux HEVC#(lf)Profile"},
            {"RemuxSizeBytes", "Remux Size#(lf)Bytes"},
            {"BackupPath", "Backup Path"},
            {"Root", "Root"},
            {"FileName", "File Path Under Root"}
        }
    )
in
    #"Renamed Columns"
