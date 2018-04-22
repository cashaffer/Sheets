import pandas as pd
import numpy as np
import hashlib
import xlwt

def hashValue(string, size):
    hash = 0
    string = string.lower();
    string.replace(u'\xa0', ' ')
    for x in string:
        hash = hash + ord(x)
    return(hash % size)


def searchHashTable(string, hashtable):
    x = 0;
    temp = string
    collisions = 0
    found = False
    if (string == string):
        string = string.lower();
        string.replace(u'\xa0', ' ')
        slot = hashValue(string, len (hashtable))
        while not found:
            x = x + 1;

            possible = hashtable[slot]
            if possible is None:
                return 'Not Found'
            else:
                if string == possible[0]:
                    found = True
                    return possible[1]
                else:
                    collisions = collisions + 1
                    slot = (slot + (collisions**2))%len(hashtable)

    else:
        return 'Not Found'

def addHashTable(pair, hashtable):
        collisions = 0
        stop = False
        string = pair[0]
        if (string == string):
            slot = hashValue(string, len (hashtable))
            while not stop:
                if hashtable[slot] == None:
                    hashtable[slot] = pair
                    stop = True
                else:
                    slot = (slot + (collisions**2))%len(hashtable)
                    collisions = collisions + 1



def searchHashTableProducts(pair, hashtable):
    collisions = 0
    found = False
    if (pair[0] == pair[0]):
        string = pair[0]
        slot = hashValue(string, len (hashtableProducts))
        while not found:
            possible = hashtable[slot]
            if possible is None:
                return 'Not Found'
            if possible[0] == pair[0] and possible[1] == pair[1]:
                found = True
                return possible[1]
            else:
                slot = (slot + (collisions**2))%len(hashtableProducts)
                collisions = collisions + 1
    else:
        return 'Not Found'


#creating the hashtable
df = pd.read_excel('ProductSynonymList.xlsx', 'Synonyms')
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")
sheet2 = book.add_sheet("Sheet 2")
sheet3 = book.add_sheet("Sheet 3")
sheet4 = book.add_sheet("Sheet 4")
#get the values for a given column
#values = df['IUPAC Name'].values
#get a data frame with selected columns
#df_selected = df[FORMAT]
hashtable = dict.fromkeys((range(10000)))
hashtableProducts = dict.fromkeys((range(20000)))
numCanonicalNames = 0;
h = 1;
for i, row in df.iterrows():
    space = False

    value = row['IUPAC Name']
    if (value == value):
        sheet2.write(i, 0, value)
    if (value == value):
        sheet3.write(numCanonicalNames, 0, value)
        j = i
        numAliasNames = 0;
        while not space:
            try:
                alias = df.iloc[j, 2]
            except Exception:
                pass
            if (alias == alias and isinstance(alias, basestring)):
                alias = alias.lower();
                alias.replace(u'\xa0', ' ')
                pair = (alias, value)
                duplicate = searchHashTable(alias, hashtable)
                if duplicate == 'Not Found':
                    numAliasNames = numAliasNames + 1
                    addHashTable(pair, hashtable)
                    sheet2.write(j, 1, alias)
                else:
                    if duplicate != numCanonicalNames:
                        sheet4.write(h, 1, alias)
                        h = h + 1
                j = j + 1
            else:
                space = True
        pairSame = (value, value)
        addHashTable(pairSame, hashtable)
        sheet3.write(numCanonicalNames, 1, numAliasNames)
        numCanonicalNames = numCanonicalNames + 1;





#creating the second excel document
totalProducts = 0;

af = pd.read_excel('ProductSynonymList.xlsx', 'Product List');
#a = af.sort_values(by='Product');
for j, row in af.iterrows():
    productName = row[0]
    if productName is np.nan:
        productName = ' '
    chemicalName = row[1]
    canonicalName = searchHashTable(chemicalName, hashtable)
    productChemicalPair = (productName, canonicalName)
    found = searchHashTableProducts(productChemicalPair, hashtableProducts)
    if found == 'Not Found':
        sheet1.write(j, 0, productName)
        addHashTable(productChemicalPair, hashtableProducts)
        if not productName.isspace():
            sheet1.write(j, 1, searchHashTable(chemicalName, hashtable))
            totalProducts = totalProducts + 1;
sheet1.write(j + 1, 0, totalProducts)





book.save("trial.xls")
