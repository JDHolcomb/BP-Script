# BP-Script
A script to automatically take a list of proteins and process them through Blast P and return results.

# General overview of the program:
This program takes large batches of Human protein data from a spreadsheet and processes it through uniprot then through blastp servers (or through a local NCBI database if you have it downloaded) to result in a list of proteins that are homologs found in Aspergillus Nidulans to your original proteins found in humans. 

# Environment set up required:
1. Install Python at https://www.python.org/downloads/
  a. Run the downloaded executable as administrator
  b. Select Custom install
    i. Select install for all users
    ii. Select update path / environment variables
2. Install Openpyxl
  a. Download Openpyxl at https://pypi.org/project/openpyxl/ openpyxl-3.0.4-py2.py3-none-any.whl (note: version numbers may change)
  b. Open a command window as administrator
  c. Pip3 install open[tab]
3. Install Tkinter (optional if you’d rather pass in the excel spreadsheet through the command line)
  a. Tkinter installation tutorial at https://tkdocs.com/tutorial/install.html.

4. Download blastp+ executable on whatever system you have 
  a. At this link ftp://ftp.ncbi.nlm.nih.gov/blast/executables/LATEST/ (only tested on windows)


5. Download the actual program at github link: https://github.com/JDHolcomb/BP-Script 
If you want to request permission to access the private repo just email me your github username at “jdholcom@ucsd.edu”.


 6. Go to File with script and run it 
  a. (Optional) If you would like to edit the script then I recommend visual studio code and install a python extension. A link to download Visual studio code: https://code.visualstudio.com/ which is available for Windows, Mac, or Linux.



# Running the program:
1. Open a command prompt and go to the directory where you downloaded the repo code and enter “RPDataProcessing.py”.
2. Program will ask if you would like to process your files using a local database or the remote blastp database, choose whichever you want if you have the local database available.
3. Program will ask for the input excel file, navigate and select the file you want processed.
4. Next the program will display a list of the tabs inside of that file and ask you to select the one you want processed, do that.
5. Choose the lines within your input file that you want processed, make note that this program will append data to the same results file if you process the same file again. This means if you include lines that you already processed the results will show up twice.
6. The results file will be in the same directory as your input file, named  “yourinputfile_results.csv”.


# Installing a local database:
1. In order to run the program locally (which can be much faster than running it remotely) you must first install the Blastp local database (nr database). It is rather large (around 578 gb) so you may experience difficulties downloading it. 
https://www.ncbi.nlm.nih.gov/books/NBK537770/
This links you to the blast help page that talks about downloading the local databases and provides its own link to them directly.
2. I had trouble downloading the databases fully which was resolved by using an open-source third party ftp client: https://filezilla-project.org/
  a. Download the individual volumes of the nr database (currently 38) then uncompress and untar them and then run the alias tool to aggregate the volumes of nr databases into one index file for the Blastp to use. The alias command can be found here: https://www.ncbi.nlm.nih.gov/books/NBK279693/.
  b. Run this command from within the directory that you have the databases downloaded in: blastdb_aliastool -title nrVolumes -num_volumes 38 -out nr -dbtype prot




# If you have any questions that were not answered by this help page feel free to email me at jdholcom@ucsd.edu



