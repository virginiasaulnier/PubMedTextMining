
from Bio import Entrez
import xlrd
import pandas as pd

# this script takes an excel file and extracts search terms from cells in column 4
# the search terms are added to a predefined advanced pubmed search with key words and mesh terms
# top 10 pubmed ID results for each term are then searched again to get article info
# article info is recorded and saved to a new file with the corresponding search term


Entrez.email = "virginia.saulnier@gmail.com"


workbook= xlrd.open_workbook("/Users/virginiasaulnier/Documents/emaadversereactions.xlsx")
sheet = workbook.sheet_by_index(0)
#file = open('/Users/virginiasaulnier/Documents/articleRecomendations.csv', 'w')
i=1
mylist=[]
#This is bad juju*****, need a better iteration method
while i<561:
    effect = sheet.cell(i, 3).value
    handle = Entrez.esearch(db="pubmed", retmax=10, term="(tumour necrosis factor[All Fields]"
                                                         "OR tumor necrosis factor-alpha[MeSH Terms] OR (tumor[All Fields])"
                                                         "AND necrosis[All Fields] AND factor-alpha[All Fields]) OR "
                                                         "tumor necrosis factor-alpha[All Fields] OR (tumor[All Fields] AND "
                                                         "necrosis[All Fields] AND factor[All Fields]) OR "
                                                         "tumor necrosis factor[All Fields]) AND "
                                                         "inhibitor[All Fields] AND " + str(effect) + "[All Fields]))")
    record = Entrez.read(handle)
    #creat a list of the pubmed IDs for each search term
    lists=record['IdList']
    print(lists)
    # for each pubmed ID get the title, Date, Journal
    for list in lists:
        print(list)
        handle = Entrez.esummary(db="pubmed", id=list, rettype="abstract", retmode="xml")
        record= Entrez.read(handle)
        print(record)
        title=record[0]['Title']
        pubmedID=record[0]['Id']
        date=record[0]['PubDate']
        journal=record[0]['Source']
        newrow=[title, journal, date, pubmedID, effect]

        print(newrow)
        #save article info to mylist
        mylist.append(newrow)

    #print(myList)
    i= i+1
#save each mylist row to a data frame
my_df = pd.DataFrame(mylist)
print("arg")
#put data frame in a csv file
my_df.to_csv('/Users/virginiasaulnier/Documents/articleRecomendations.csv', index=False, header=False)

