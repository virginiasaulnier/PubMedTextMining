
import xlrd
import xlwt


#this script takes an excel file and extracts search terms from cells in column 4
# the search terms are added to a premade search with key words and mesh terms
# top 10 pubmed results are recorded

from Bio import Entrez
Entrez.email = "virginia.saulnier@gmail.com"


# open excel file to get search terms
workbook= xlrd.open_workbook("/Users/virginiasaulnier/Documents/emaadversereactions.xlsx")

# search terms are on the first sheet
sheet = workbook.sheet_by_index(0)
i=1
workbook2 = xlwt.Workbook()
#workbook2.save('newadversereactions.xls')
sheet2 = workbook2.add_sheet('Sheet_1')

#this iterates through the sheet and does a pubmed search for all applicable TNF and RA terms plus the search term
while sheet.cell(i,3).value != xlrd.empty_cell.value:
    # effect is the search term
    effect = sheet.cell(i,3).value

    handle = Entrez.esearch(db="pubmed",retmax=10, term="(tumour necrosis factor[All Fields]"
                                            "OR tumor necrosis factor-alpha[MeSH Terms] OR (tumor[All Fields])"
                                            "AND necrosis[All Fields] AND factor-alpha[All Fields]) OR "
                                            "tumor necrosis factor-alpha[All Fields] OR (tumor[All Fields] AND "
                                            "necrosis[All Fields] AND factor[All Fields]) OR "
                                            "tumor necrosis factor[All Fields]) AND "
                                            "inhibitor[All Fields] AND "+str(effect)+"[All Fields]))")
    record = Entrez.read(handle)


    record["Count"]
    record["IdList"]
    #this prints the pubmed IDs
    print("The first 10 are\n{}".format(record['IdList']))
    PubMedIDs = dict({"+str(effect)+": ".format(record['IdList']"})
    i=i+1
    print(i)
    break



print(PubMedIDs)
