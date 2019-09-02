module xls2txt

# Import Julia packages
# import ExcelReaders; const xls = ExcelReaders
import ExcelFiles; const xls = ExcelFiles
import CSV
import DataFrames.DataFrame
import ProgressMeter; const pm = ProgressMeter


readXLS(filename, sheet::String="Tabelle1") = DataFrame(xls.load(filename, sheet))

"""
    correctXLS(infolder, outfolder, sheetname::String="Tabelle1")

Loop over Excel files in `infolder` and save `.dat` files in `outfolder` with the
same name as the Excel files. The `sheetname` specifies the name of the Excel sheet
that is read in.
"""
function correctXLS(infolder, outfolder, sheetname::String="Tabelle1")
  folders = readdir(infolder)[isdir.(infolder.*"/".*readdir(infolder))]
  pm.@showprogress 1 "convert xls to text..." for folder in folders
    files = readdir(joinpath(infolder, folder))
    mkdir(joinpath(outfolder, folder))
    for file in files
      try xlsdata = readXLS(joinpath(infolder, folder, file))
        CSV.write(splitext(joinpath(outfolder, folder, file))[1]*".dat", xlsdata, delim='\t')
      catch
        @warn "Error in reading data from $folder/$file. Possibly empty Excel sheets. Data skipped."
      end
    end
  end
end #function correctXLS

correctXLS("data/in", "data/out/")



end # module
