import csv
from Bio import Entrez
import xlrd
import xlwt
import bs4
import numpy
import pandas as pd


Entrez.email = "virginia.saulnier@gmail.com"


workbook= xlrd.open_workbook("/Users/virginiasaulnier/Documents/emaadversereactions.xlsx")
sheet = workbook.sheet_by_index(0)
#file = open('/Users/virginiasaulnier/Documents/articleRecomendations.csv', 'w')
i=1
mylist=[]
#while sheet.cell(i, 3).value != xlrd.empty_cell.value:
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
    #print(format(record['IdList']))
    #lists = (', '.join(map(str, record['IdList'])))
    lists=record['IdList']
    print(lists)
    for list in lists:
        print(list)
        handle = Entrez.esummary(db="pubmed", id=list, rettype="abstract", retmode="xml")
        record= Entrez.read(handle)
        print(record)
        #print(handle.readline().strip())
        #title=record['PubMedArticle']['MedlineCitation']['ArticleTitle']
        #title=record[0]['PubMedArticle'][0]['MedlineCitation'][0]['ArticleTitle']
        title=record[0]['Title']
        pubmedID=record[0]['Id']
        date=record[0]['PubDate']
        journal=record[0]['Source']
        #print(record[0])
        #abstract=record[0]['PubMedArticle']['MedlineCitation']['ArticleTitle']
        #mesh =record[0]['MedlineCitation']['MeshHeadingList']
        #descriptors = ','.join(term['DescriptorName'] for term in mesh)
        #file = open('/Users/virginiasaulnier/Documents/articleRecomendations.csv', 'w')
        #newrow = [1, 2, 3]
        #A = numpy.concatenate((A, newrow))
        newrow=[title, journal, date, pubmedID, effect]

        #X = np.empty(shape=[0, n]
        #mylist=numpy.empty(shape=[400, 4])
        print(newrow)
        #mylist=[]
        #if len(newrow)==4:
        mylist.append(newrow)

        #mylist=numpy.append(mylist, [title, journal, date, pubmedID, effect])
        #mylist.append([title, journal, date, pubmedID, effect])
        #writer = csv.writer(file, delimiter='\t', lineterminator='\n', )
        #writer.writerow([title, journal, date, pubmedID, effect])
    print(i)
    #print(myList)
    i= i+1
my_df = pd.DataFrame(mylist)
print("arg")
my_df.to_csv('/Users/virginiasaulnier/Documents/articleRecomendations.csv', index=False, header=False)

#thefile = open('/Users/virginiasaulnier/Documents/articleRecomendations.csv', 'w')
#for item in mylist:
 # thefile.write("%s\n" % item)
#numpy.savetxt('/Users/virginiasaulnier/Documents/articleRecomendations.csv', mylist, delimiter=",")
# with open('/Users/virginiasaulnier/Documents/articleRecomendations.csv', 'w') as csvfile:
#     print("in open loop")
#     writer = csv.writer(csvfile, delimiter='\t', lineterminator='\n', )
#     #dv = dd[:, 0]
#     for r in mylist:
#      writer.writerow([r])
#with open('/Users/virginiasaulnier/Documents/articleRecomendations.csv','w') as csv_file:
 #   csv_file.write('\n'.join(map(','.join, mylist)+'\n'))
