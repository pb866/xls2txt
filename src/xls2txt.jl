module xls2txt

# Import Julia packages
import LoggingExtras; const logg = LoggingExtras
import ExcelFiles; const xls = ExcelFiles
import CSV
import DataFrames.DataFrame
import ProgressMeter; const pm = ProgressMeter
import Juno.input

export logg, correctXLS


readXLS(filename, sheet::String="Tabelle1") = DataFrame(xls.load(filename, sheet))


"""
    findFiles(inventory::Vector{String}, folder::String, filetypes::String...) -> inventory

Load inventory files saved in `folder` to the `inventory` holding the file names
and locations ending with `filetypes` as a vector of strings.
"""
function findFiles(inventory::Vector{String}, folders::Vector{String},
  folder::String, basefolder::String, filetypes::String...)
  # Construct Regex of file endings from filetypes
  fileendings = Regex(join(filetypes,'|'))
  # Scan directory for files and folders and save directory
  dir = readdir(folder); path = normpath(folder)
  for cwd in dir
    # Save current directory/file
    file = joinpath(path, cwd)
    if endswith(file, fileendings)
      # Save files of correct type
      push!(inventory, basename(file))
      push!(folders, relpath(dirname(file), basefolder))
    elseif isdir(file)
      # Step into subdirectories and scan them, too
      inventory, folders = findFiles(inventory, folders, file, basefolder, filetypes...)
    end
  end

  return inventory, folders
end # function findcsv

function prepare_outfolder(outfolder::String)
  if isdir(outfolder)
    println("\33[95mOutput folder already exists. Overwrite content?")
    confirm = input("(\33[4mY\33[0m\33[95mes/\33[4mN\33[0m\33[95mo)\33[0m ")
    if startswith(lowercase(confirm), "y")
      rm(outfolder, recursive=true); mkpath(outfolder)
    else
      @error "Cancel overwrite"
    end
  else
    println("$outfolder created."); mkpath(outfolder)
  end
end #function prepare_outfolder


function processXLS(files::Vector{String}, folders::Vector{String},
  infolder::String, outfolder::String, sheetname::String)
  mkpath.(joinpath.(outfolder, unique(folders)))
  pm.@showprogress 0.1 "convert xls to text..." for (file, folder) in zip(files, folders)
    try xlsdata = readXLS(joinpath(infolder, folder, file), sheetname)
      valid = @. (!ismissing(xlsdata.Latitude)) & !(xlsdata.Latitude isa String) &
        (!ismissing(xlsdata.Longitude)) & !(xlsdata.Longitude isa String) &
        (!ismissing(xlsdata.feet)) & !(xlsdata.feet isa String)
      xlsdata = xlsdata[valid,:]
      xlsdata.Latitude ./= 1e4
      xlsdata.Longitude ./= 1e4
      xlsdata.feet = [x < 50 ? x * 1000 : x for x in xlsdata.feet]
      CSV.write(splitext(joinpath(outfolder, folder, file))[1]*".dat", xlsdata, delim='\t')
    catch
      @warn "Error in reading data from $folder/$file. Possibly empty Excel sheets. Data skipped."
    end
  end
end #function processXLS


"""
    correctXLS(infolder, outfolder, sheetname::String="Tabelle1")

Loop over Excel files in `infolder` and save `.dat` files in `outfolder` with the
same name as the Excel files. The `sheetname` specifies the name of the Excel sheet
that is read in.
"""
function correctXLS(infolder::String, outfolder::String,
  sheetname::String="Tabelle1")
  prepare_outfolder(outfolder)
  files = String[]; folders = String[]
  files, folders = findFiles(files, folders, "data/", infolder, ".xlsx")
  processXLS(files, folders, infolder, outfolder, sheetname)
end #function correctXLS

end # module
