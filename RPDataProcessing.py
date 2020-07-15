#RPDataProcessing.py Python app to process Reck Peterson Lab Data in XLSX format
#Author: James Holcomb
#Last Revised: July 4th, 2020

#Library imports
import os                                   #imports functions for interacting with O/S - allows Command Line 
import re                                   #imports regular expression operations
import openpyxl                             #imports Python Library to read/write 2010 Excel files
import requests                             #imports gives ability to send HTTP/1.1 request
import tkinter as tk                        #imports interface to the Tk GUI toolkit
from tkinter import filedialog
import xml.etree.ElementTree as ET

def process_bp_output(FullBPOutput):
    print("Called process_bp_output")
    
    tree = ET.parse(FullBPOutput)
    root = tree.getroot()
    for child in root:
        print (child.tag, child.attrib)
    
    print(root.BlastOutput_query)

    #handle if no matches and re-set number of matches            
            #if numMatches == 0:
            #    wsOut.cell(row=outRow, column = NCBI_geneCol).value = NCBIgene 
            #    wsOut.cell(row=outRow, column = log2FoldCol).value = log2Fold
            #    wsOut.cell(row=outRow, column = log10PValueCol).value = log10PValue 
             #   wsOut.cell(row=outRow, column = uniprotEntryHumanCol).value = currentGene
            #    wsOut.cell(row=outRow, column = possibleMatchCol).value = "No Potential Matches Found"
            #    outRow = outRow + 1
           # else:
            #    numMatches = 0
    exit


#Files used later in App
currentFastaFile = "currentFasta.txt"
blastpFile = "blastpResults.out"
logFileName = "logFile.out"

#Set up paths
root = tk.Tk()
root.withdraw()
wbName = filedialog.askopenfilename()
print("input name " + wbName)
wbName = os.path.normpath(wbName)
print("input name after OS normalization " + wbName)
outDir = os.path.dirname(wbName) + "\\"
print("outdir is " + outDir)
outName = os.path.basename(wbName)
print("outName is " + outName)
(outName,extension) = os.path.splitext(outName)
print("outName is now " + outName)
outName = outName + ".csv"
print("outName is " + outName)
try:
    outputFile = open(outName, "w")
except:
    print("The output file " + outName + " already exists!")
    exit()

print("workbook name " + wbName)
wsName = "FAM160A1-N"
wsOutName = wsName +" Processed Data"
print("worksheet name " + wsName)

#URL and Command Query 
uniprotURLString1 = "https://www.uniprot.org/uniprot/?query="
uniprotURLString2 = "&fil=organism:\"Homo+sapiens+(Human)+[9606]\"&sort=score&columns=id&format=tab"  
fastaURLString1 = "https://www.uniprot.org/uniprot/"
fastaURLString2 = ".fasta"

blastpQuery = 'cmd /c "dir & blastp -db nr -query  "' + outDir + currentFastaFile + '" -entrez_query \"Aspergillus Nidulans[ORGN]\" -out "' + outDir + blastpFile + '" -remote -qcov_hsp_perc 20 -outfmt "7 sseqid qcovs" & dir"'
#-------------------------------------------------------------------------------------------------------------------------
print ("new bp query is " + blastpQuery)
#Hard-coded Parameters used later in App
foldCheck = 1.6     #foldCheck and pvalueCheck are the conditional formatting for the highlighting of cells
pvalueCheck = 1.3
scoreMin = 20
eValueMax = 1

#Initialization of Parameters used later in App
outRow = 1    #set initial data row
log2Fold = 0
log10PValue = 0

#set delimiters and delimiter lengths for blastp text
ridTxt = "RID: "
ridTxtLength = len(ridTxt)

#Open and set logFile (for debugging purposes)
logFile = open(outDir+logFileName, "w")

#output column names to outputfile
outputFile.write("'NCBI_GENE','log2(fold)','log10(pvalue)','Uniprot Entry (Human)','BlastP Order','Possible Match','Score','Expect','Positives','Positives Out Of','Gaps', 'Gaps Out Of'\n")

#Open and set Spreadsheet up
wb = openpyxl.load_workbook(wbName)
ws = wb[wsName]
wsOut = wb[wsOutName]

#start processing rows of spreadsheet data
topValue = int(input("enter the first row you want processed: ")) + 1
botValue = int(input("enter the last row you want processed: ")) + 1
for i in range(topValue, botValue): #ws.max_row):               #skip header row start with 2
    
    #Need to check if gene is empty or whitespace;
    if (not (str(ws.cell(row=i, column=1).value).isspace() or ws.cell(row=i, column=1).value is None)): 

        NCBIgene = ws.cell(row=i, column=1).value   #extract NCBI Gene information

        #extract log2(fold) and log10(pvalue) data
        try:
            log2Fold = float(ws.cell(row=i, column=2).value)
        except ValueError:
            logFile.write("ERROR - Expected log2Fold Value\n")
            logFile.write("x: "+ x + "\n")
            break
        
        try:
            log10PValue = float(ws.cell(row=i, column=3).value)
        except ValueError:
            logFile.write("ERROR - Expected log10PValue Value\n")
            logFile.write("x: "+ x + "\n")
            break
        
        #Check if log2(fold) and log10(pvalue) data falls within bounds to check for potential match           
        if (log2Fold > foldCheck and log10PValue > pvalueCheck):     #Highlighting is conditional so checking value for Columns B and C
            
            #generate uniprot URL, and extract gene
            uniprotURL = uniprotURLString1 + ws.cell(row=i, column=1).value+uniprotURLString2
            uniprotResponse = requests.get(uniprotURL)
            currentGene = uniprotResponse.text.split('\n')[1]
            print (currentGene)
            
            #generate URL to get FASTA, and write to output file
            fastaURL = fastaURLString1 + currentGene + fastaURLString2
            fastaResponse = requests.get(fastaURL)  
            fasta = fastaResponse.text
            with open(outDir+currentFastaFile, "w") as fastafile:
                fastafile.write(fasta)
    
            #run Command Line to query blastp for
            logFile.write("Processing Gene: "+ currentGene +"\n")
            logFile.close()    #closing because os.system command messes up open file

            #Command Line blastP Query
            os.system(blastpQuery)
            logFile = open(outDir+logFileName, "a")

            #read blastp results from output file
            with open(outDir+blastpFile, "r")as blastp:
                blpl = blastp.readlines()
        
                for x in blpl:
                
                    #check if blank line
                    if not x.isspace():
                                            
                        #check for RID
                        if "RID:" in x:
                            RID = x[x.find(ridTxt) + ridTxtLength:]
                            RID = RID.strip()
                            
                            # How to get the results file from the RID
                            # Extract the RID from the original blastp output then replace the RID in the following command.
                            # The output file name will be the RID-Alignment.txt
                            # https://blast.ncbi.nlm.nih.gov/Blast.cgi?RESULTS_FILE=on&RID=GR1RE74P014&FORMAT_TYPE=Text&FORMAT_OBJECT=Alignment&DESCRIPTIONS=100&ALIGNMENTS=100&CMD=Get&DOWNLOAD_TEMPL=Results_All&ADV_VIEW=on
                            # Alternatively you can get the XML file
                            xmlOutput1 = "https://blast.ncbi.nlm.nih.gov/Blast.cgi?RESULTS_FILE=on&RID="
                            xmlOutput2 = "&FORMAT_TYPE=XML&FORMAT_OBJECT=Alignment&CMD=Get"
                            xmlURL = xmlOutput1 + RID + xmlOutput2
                            xmlOutputResponse = requests.get(xmlURL, allow_redirects=True)
                            RIDoutput = outDir + RID + "_output.xml"
                            open(RIDoutput, 'wb').write(xmlOutputResponse.content)
                            print(xmlURL)
                            print(RID)
                            # Then parse the XML file with https://www.geeksforgeeks.org/xml-parsing-python/

                            #call function to process XML output
                            process_bp_output(RIDoutput)
                            break #RID found and processed. Leaves for loop.

                if RID == "" :
                    print("RID " + ridTxt + " not found in blast output file")
                
            #finished processing sequence                     
            logFile.write("Done Processing Sequence\n")
        else:
            logFile.write(NCBIgene + " did not meet the minimum requirements and was skipped.")

        logFile.write("Done Processing Gene: "+ currentGene +"\n")
        print("Done Processing Gene: "+ currentGene +"\n")
        
    wb.save(filename = wbOutName)   #save output spreadsheet data after processing each row
logFile.close() 


