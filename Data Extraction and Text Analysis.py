import urllib.request
import xlrd
import requests
import re
from bs4 import BeautifulSoup
from nltk.corpus import stopwords
from collections import Counter

#.................. import files...............
report={}
string={}
num_words=0
uncertainty_dictionary={}
constraining_dictionary={}
loc_uncertainty = (r'C:\Users\SANJEEV\Desktop\uncertainty_dictionary.xlsx')
loc_constraining = (r'C:\Users\SANJEEV\Desktop\constraining_dictionary.xlsx')
loc = (r'C:\Users\SANJEEV\Desktop\cik_list.xlsx')
report[1] = "MANAGEMENT'S DISCUSSION AND ANALYSIS"
report[2] = "QUANTITATIVE AND QUALITATIVE DISCLOSURES ABOUT MARKET RISK"
report[3] = "RISK FACTORS"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
wb1 = xlrd.open_workbook(loc_uncertainty)
sheet1 = wb1.sheet_by_index(0)
wb2 = xlrd.open_workbook(loc_constraining)
sheet2 = wb2.sheet_by_index(0)

#..............Program to extract a particular row value.....
def main():
    # We can change the value of i and can have any number of loop we want or the number SEFCNAME we want from Clk_list 
    for i in range (1,2):
     string1= 'https://www.sec.gov/Archives/'
     string2= sheet.cell_value(i,5)
     string[i] = string1+string2
     with urllib.request.urlopen(string[i]) as response:
      string[i] = response.read()
      soup = str(BeautifulSoup(string[i], 'html.parser'))
      
      # ................filtering_data..........................................
      for j in range (1,4):         
       Temp_filter = re.search("ITEM\s\w\.\s+"+report[j]+".+\s+.+[0-9]",soup)
       string[i]=Temp_filter
       if j==1:
           Note='mda '
       if j==2:
           Note='qqdmr '
       if j==3:
           Note='rf '
       if not Temp_filter:
           word_count=0
           positive_score=0
           negative_score=0
           Polarity_score=0
           Subjectivite_score=0
           Complex_Word_Count=0
           Percentage_Complex_words=0
           uncertainty_score=0
           constraining_score=0
           print(Note+'word_count:=',word_count)
           print(Note+'positive_score:=',positive_score)
           print(Note+'negative_score:=',negative_score)
           print(Note+'polarity_score:=',Polarity_score)
           print(Note+'Subjectivite_score:=',Subjectivite_score)
           print(Note+'Complex_Word_Count:=',Complex_Word_Count)
           print(Note+"Percentage_Complex_words",Percentage_Complex_words)
           print(Note+"uncertainty_score:=",uncertainty_score)
           print(Note+"constraining_score:=",constraining_score)
           continue
       Temp_filter1 = Temp_filter.group(0)
       Temp_filter2 = re.search("\.+\s+\d+",Temp_filter1)
       Temp_filter3 = re.search("\d+",Temp_filter2.group(0))
       Temp_filter4 = int(Temp_filter3.group(0))
       f1 = re.search(str(Temp_filter4-1)+"[\s]+<page>",soup) 
       f2= re.search(str(Temp_filter4)+"[\s]+<page>",soup)  
       string[i] = soup[f1.span(0)[1]:f2.span(0)[1]].lower()
       string[i] = re.sub(r'[^\w\s]','',string[i])
       string[i]=string[i].split()
       stop_words = set(stopwords.words('english'))
       string[i]=[word for word in string[i] if word not in stop_words]
       total_word=string[i]
       
       # ..................word count..............................
       word_count=len(string[i])
       print(Note+'word_count:=',word_count)

       # ....................positive words........................
       pos_word=open(r'C:\Users\SANJEEV\Desktop\pos.txt')
       pos_word=((pos_word.read()).lower()).split()
       string_pos=[word for word in string[i] if word not in pos_word]
       positive_score=word_count-len(string_pos)
       print(Note+'positive_score:=',positive_score)

       # ...................negative words........................
       neg_word=open(r'C:\Users\SANJEEV\Desktop\neg.txt')
       neg_word=((neg_word.read()).lower()).split()
       string_neg=[word for word in string[i] if word not in neg_word]
       negative_score=(-(word_count-len(string_neg)))
       print(Note+'negative_score:=',(negative_score))

       # ...................polarity score........................
       Polarity_score = (positive_score-negative_score)/((positive_score+negative_score)+0.000001)
       print(Note+'polarity_score:=',Polarity_score)

       # ...................subjective score.....................
       Subjectivite_score = (positive_score + negative_score)/ ((word_count) + 0.000001)
       print(Note+'Subjectivite_score:=',Subjectivite_score)

       # ..................Complex_Word_Count....................
       vowels = "aeiouyAEIOUY"
       count=0
       Complex_Word_Count=0
       for a in range(len(string[i])):
           temp_word=total_word[a]
           b=0
           for b in range(len(temp_word)):
             if temp_word[b] in vowels:
                    count=count+1
                    if temp_word.endswith('es') or temp_word.endswith('ed'):
                        count=count-1
           if (count>=3):
                 Complex_Word_Count=Complex_Word_Count+1
           count=0
       print(Note+'Complex_Word_Count:=',Complex_Word_Count)
       
       # ..................Percentage Complex words.................
       Percentage_Complex_words = Complex_Word_Count/word_count
       print(Note+"Percentage_Complex_words:=",Percentage_Complex_words)

       #.................. uncertainty_dictionary...............
       uncertainty_score=0
       for k in range (1,298):
           uncertainty_dictionary[k]= sheet1.cell_value(k,0)
           uncertainty_dictionary[k]=uncertainty_dictionary[k].lower()
       for g in range(1,len(uncertainty_dictionary)):
           for h in range(1,298):
               if (Counter(total_word[g]) == Counter(uncertainty_dictionary[h])):
                   uncertainty_score=uncertainty_score+1
       print(Note+"uncertainty_score:=",uncertainty_score)
       
       #.................. constraining_dictionary...............
       constraining_score=0
       for k in range (1,185):
           constraining_dictionary[k]= sheet2.cell_value(k,0)
           constraining_dictionary[k]=constraining_dictionary[k].lower()
       for g in range(1,len(constraining_dictionary)):
           for h in range(1,185):
               if (Counter(total_word[g]) == Counter(constraining_dictionary[h])):
                   constraining_score=constraining_score+1
       print(Note+"constraining_score:=",constraining_score)

# main program
if __name__ == "__main__":
          main()
