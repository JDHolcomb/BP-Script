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

#Files used later in App
currentFastaFile = "currentFasta.txt"
blastpFile = "blastpResults.out"
logFileName = "logFile.out"

#Spreadsheet Workbook and Worksheet Names used later in App
root = tk.Tk()
root.withdraw()
wbName = filedialog.askopenfilename()
print("input name " + wbName)
wbName = os.path.normpath(wbName)
print("input name after OS normalization " + wbName)
outDir = os.path.dirname(wbName) + "\\"
print("outdir is " + outDir)
wbOutName = "C:\\RPScript\\Test Data Out.xlsx"  #eventually can save to same Workbook but don't want to overwrite till I know it works
print("workbook name " + wbName)
wsName = "Hook1-C"
wsOutName = wsName +" Processed Data"

#URL and Command Query 
uniprotURLString1 = "https://www.uniprot.org/uniprot/?query="
uniprotURLString2 = "&fil=organism:\"Homo+sapiens+(Human)+[9606]\"&sort=score&columns=id&format=tab"  
fastaURLString1 = "https://www.uniprot.org/uniprot/"
fastaURLString2 = ".fasta"
blastpQuery = 'cmd /c "dir & blastp -db nr -outfmt 10qcovs -query ' + outDir + currentFastaFile + ' -entrez_query \"Aspergillus Nidulans[ORGN]\" -out ' + outDir + blastpFile + ' -remote & dir"'
print ("new bp query is " + blastpQuery)
#Hard-coded Parameters used later in App
foldCheck = 1.6     #foldCheck and pvalueCheck are the conditional formatting for the highlighting of cells
pvalueCheck = 1.3
scoreMin = 40
eValueMax = 1

#Initialization of Parameters used later in App
outRow = 1    #set initial data row
log2Fold = 0
log10PValue = 0

#set delimiters and delimiter lengths for blastp text
matchTxt = ">"                  #Delimiter to begin detailed potential match info
matchTxtLength = len(matchTxt)
endMatchNameTxt = "Length="     #Delimiter to end detailed potential match name                
scoreTxt = "Score = "           #Delimiter to begin score info
scoreTxtLength = len(scoreTxt)
eValueTxt = "Expect = "         #Delimiter to begin eValue info 
eValueTxtLength = len(eValueTxt)
positivesTxt = "Positives = "   #Delimiter to begin positives info 
positivesTxtLength = len(positivesTxt)
gapsTxt = "Gaps = "             #Delimiter to begin gaps info 
gapsTxtLength = len(gapsTxt)

#Open and set logFile (for debugging purposes)
logFile = open(outDir+logFileName, "w")

#Open and set Spreadsheet up
wb = openpyxl.load_workbook(wbName)
ws = wb[wsName]
wsOut = wb[wsOutName]

#set column header titles and indices
colNames = ['NCBI_GENE','log2(fold)','log10(pvalue)','Uniprot Entry (Human)','BlastP Order','Possible Match','Score','Expect','Positives','Positives Out Of','Gaps', 'Gaps Out Of']
for col,name in enumerate(colNames,1):
    wsOut.cell(row=outRow, column=col).value = name
outRow = outRow + 1
NCBI_geneCol = colNames.index('NCBI_GENE') + 1       #Set Column Indices NOTE:OpenPyXL columns start with 1 not 0
log2FoldCol = colNames.index('log2(fold)') + 1
log10PValueCol = colNames.index('log10(pvalue)') + 1
uniprotEntryHumanCol = colNames.index('Uniprot Entry (Human)') + 1
blastPOrderCol = colNames.index('BlastP Order') + 1
possibleMatchCol = colNames.index('Possible Match') + 1
scoreCol = colNames.index('Score') + 1
expectCol = colNames.index('Expect') + 1
positivesCol = colNames.index('Positives') + 1
positivesOutOfCol = colNames.index('Positives Out Of') + 1
gapsCol = colNames.index('Gaps') + 1
gapsOutOfCol = colNames.index('Gaps Out Of') + 1

#start processing rows of spreadsheet data
for i in range(4, 6): #ws.max_row):               #skip header row start with 2
    
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

                #initialize loop variables and flags for blastp
                sigFlag = False                 #checks for header of significant protein match table
                sigTableEndedFlag = False       #checks if significant protein match table had ended
                matchStartFlag = False          #checks if have reached first detailed protein match
                foundFlag = False               #checks if have found data in detailed protein match
                numMatches = 0
                eValue = 0
                score = 0
                matchStringList = []    #matchStringList - collects potential match name string to join together
                positives = []
                gaps = []
        
                for x in blpl:
                
                    #check if blank line
                    if not x.isspace():
                    
#might want to skip significant table check because not using anything in table                        
                        #check if have reached header for significant protein match table
                        if "significant" in x and (not sigFlag) :          
                            sigFlag = True
                            
                        #check if after significant protein match table header and before match information (within significant protein match table)
                        elif (sigFlag and not(sigTableEndedFlag)):    
                        
                            #check if significant protein match table ended and reached first detailed sequence match
                            if matchTxt in x[0]:    #check if first character matches matchTxt             
                                sigTableEndedFlag = True
                                matchStartFlag = True
                                matchStringList.append(x[x.find(matchTxt) + matchTxtLength:])
            

                        #within information on detailed match information section
                        elif sigTableEndedFlag:
#might want to skip significant table check because not using anything in table (and get rid of sigTableEndedFlag

                            #if within detailed match information    
                            if matchStartFlag:
                                
                                #check if match sequence name has ended
                                if endMatchNameTxt in x:
                                    matchStartFlag = False
                                #otherwise append match sequence name
                                else:
                                    matchStringList.append(x)
                                    
                            #check if starting next new detailed match
                            else:
                                if matchTxt in x[0]:    #check if first character matches matchTxt                
                                    matchStartFlag = True
                                    matchStringList.append(x[x.find(matchTxt) + matchTxtLength:])

                                #if finished match sequence name check for score
                                if scoreTxt in x:
                                    for element in x[x.find(scoreTxt) + scoreTxtLength:].split():

                                        #Get score for sequence from detailed match information section
                                        try:
                                            score = float(element)
                                            break
                                        except ValueError:
                                            logFile.write("ERROR - Expected Score Value\n")
                                            logFile.write("x: "+ x[x.find(scoreTxt) + scoreTxtLength:] + "\n")
                            
                                #if finished match sequence name check for eValue
                                if eValueTxt in x:
                                    for element in x[x.find(eValueTxt) + eValueTxtLength:].replace(',', '').split():  #remove "," after number

                                        #Get eValue for sequence from detailed match information section
                                        try:
                                            eValue = float(element)
                                            break
                                        except ValueError:
                                            logFile.write("ERROR - Expected eValue Value\n")
                                            logFile.write("x: "+ x[x.find(eValueTxt) + eValueTxtLength:] + "\n")

                                #if finished match sequence name check for positives
                                if positivesTxt in x:
                                    
                                    #Get positives for sequence from detailed match information section                                    
                                    positives = re.findall('\d+',x[x.find(positivesTxt) + positivesTxtLength:].split()[0])  #looking at substring after match split into first element                                        
                                    if not positives:
                                        logFile.write("ERROR - Expected Positives Value\n")
                                        logFile.write("x: "+ x[x.find(positivesTxt) + positivesTxtLength:] + "\n")

                                #if finished match sequence name check for gaps
                                if gapsTxt in x:

                                    #Get gaps for sequence from detailed match information section
                                    gaps = re.findall('\d+',x[x.find(gapsTxt) + gapsTxtLength:].split()[0]) #looking at substring after match split into first element

                                    if not gaps:
                                        logFile.write("ERROR - Expected Gaps Value\n")
                                        logFile.write("x: "+ x[x.find(gapsTxt) + gapsTxtLength:] + "\n")
    
                                    #check if protein match is significant for R-P
                                    if (score >= scoreMin and eValue <=  eValueMax):
        
                                        #increment number of potential matches for gene
                                        numMatches = numMatches + 1
                                    
                                        #output potential match data
                                        wsOut.cell(row=outRow, column = NCBI_geneCol).value = NCBIgene 
                                        wsOut.cell(row=outRow, column = log2FoldCol).value = log2Fold
                                        wsOut.cell(row=outRow, column = log10PValueCol).value = log10PValue 
                                        wsOut.cell(row=outRow, column = uniprotEntryHumanCol).value = currentGene
                                        wsOut.cell(row=outRow, column = blastPOrderCol).value = numMatches
                                        wsOut.cell(row=outRow, column = possibleMatchCol).value = " ".join(matchStringList)
                                        wsOut.cell(row=outRow, column = scoreCol).value = score
                                        wsOut.cell(row=outRow, column = expectCol).value = eValue
                                        wsOut.cell(row=outRow, column = positivesCol).value = positives[0]
                                        wsOut.cell(row=outRow, column = positivesOutOfCol).value = positives[1]
                                        wsOut.cell(row=outRow, column = gapsCol).value = gaps[0]
                                        wsOut.cell(row=outRow, column = gapsOutOfCol).value = gaps[1]
                                        logFile.write("From Matches: Possible Match: " + " ".join(matchStringList) + " score: " + str(score)
                                            + " e-value: " + str(eValue) + " positives: " + str(positives[0])+"/" + str(positives[1])
                                            + " gaps: " + str(gaps[0])+"/" + str(gaps[1]) +"\n")
                                        outRow = outRow + 1

                                    else:
                                        logFile.write(NCBIgene + " not a match.\n")
                                        
                                    #done processing this sequence so reset match name list
                                    matchStringList = []
                                    logFile.write("Done Processing Sequence\n")

            #handle if no matches and re-set number of matches            
            if numMatches == 0:
                wsOut.cell(row=outRow, column = NCBI_geneCol).value = NCBIgene 
                wsOut.cell(row=outRow, column = log2FoldCol).value = log2Fold
                wsOut.cell(row=outRow, column = log10PValueCol).value = log10PValue 
                wsOut.cell(row=outRow, column = uniprotEntryHumanCol).value = currentGene
                wsOut.cell(row=outRow, column = possibleMatchCol).value = "No Potential Matches Found"
                outRow = outRow + 1
            else:
                numMatches = 0
                
            #finished processing sequence                     
            logFile.write("Done Processing Sequence\n")
        logFile.write("Done Processing Gene: "+ currentGene +"\n")
        print("Done Processing Gene: "+ currentGene +"\n")
    wb.save(filename = wbOutName)   #save output spreadsheet data after processing each row
logFile.close()                              
