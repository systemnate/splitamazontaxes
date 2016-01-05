This is a simple script that will take Amazon's FBA Tax Report and produce a .xls file where each tab represents a state.

The usage is simple: the first argument is the CSV file from Amazon and the second argument is the output file.  Then simply call the #run:

SplitTaxes.new("TaxFile.csv", "output.xls").run

You can see an example at the bottom of the source.
