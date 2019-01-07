import stringdist as sd
import numpy as np
import matplotlib
import xlsxwriter
from PIL import Image
import math as m
from operator import itemgetter
from wordcloud import WordCloud


file = open(r'name-list.txt')
data = file.read()
data = data.split("\n")


for i in range(len(data)):
    data[i] = data[i].lower()


visited = []
clusters = []
cluster = []
frequency = {}
unique_names = []
for i in range(len(data)):
    if data[i]!="":
        if data[i] in frequency:
            frequency[data[i]] += 1
        else:
            frequency[data[i]] = 1
Cluster_book = xlsxwriter.Workbook('names-result.xlsx')
worksheet1 = Cluster_book.add_worksheet('frequency')
worksheet1.write('A1', 'Serial Number')
worksheet1.write('B1', 'Name')
worksheet1.write('C1', 'Frequency')
sorted_frequency ={}
no = 1
for key, value in sorted(frequency.items(), key=itemgetter(1), reverse=True):
    unique_names.append(key)
    worksheet1.write('A' + str(2 + no), no)
    worksheet1.write('B' + str(2+no), key)
    worksheet1.write('C' + str(2 + no), value)
    no += 1
frequency_clusters = []
for i in range(len(data)):
    if data[i] not in visited and data[i] != "":
        cluster.append(data[i])
        for j in range(i+1, len(data)):
            dist = sd.levenshtein(data[i], data[j])

            if dist <= m.ceil(len(data[i]) * .4):
                cluster.append(data[j])
                if data[j] not in visited:
                    visited.append(data[j])

        if len(cluster) != 0:
            frequency_clusters.append(cluster)
        cluster = []



count = {}



cluster_sheet = Cluster_book.add_worksheet('clusters')
cluster_sheet.write('A1', 'Serial Number')
cluster_sheet.write('B1','frequency-A')
cluster_sheet.write('C1', 'Cluster Name')
cluster_sheet.write('D1', 'frequency-B')
cluster_sheet.write('E1', 'Other Member Name')
cluster_sheet.write('F1', 'Length Of  Cluster Name')
cluster_sheet.write('G1', 'LD between B and C')
cluster_sheet.write('H1', 'Ceil of 40% of LD')
cluster_sheet.write('I1', 'Decision')

t = 1

for i in range(len(unique_names)):

    cluster.append(unique_names[i])
    if i+1 < 10:
        s = "00"
    elif i+1 > 9 and i+1<100 :
        s = "0"
    else:
        s = ""
    workbook = xlsxwriter.Workbook(r"E:\Text Mining\Unique Excell\\" + s+str(i+1)+"__"+unique_names[i] + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Serial Number')
    worksheet.write('B1', 'Cluster Name')
    worksheet.write('C1', 'Other Member Name')
    worksheet.write('D1', 'Length Of ' + unique_names[i])
    worksheet.write('E1', 'LD between B and C')
    worksheet.write('F1', 'Ceil of 40% of LD')
    worksheet.write('G1', 'Decision')
    k = 1
    for j in range(len(unique_names)):
        if unique_names[j] != unique_names[i]:
            dist = sd.levenshtein(unique_names[i], unique_names[j])
            worksheet.write('A' + str(3 + k), k)
            worksheet.write('B' + str(3 + k), unique_names[i])
            worksheet.write('C' + str(3 + k), unique_names[j])
            worksheet.write('D' + str(3 + k), len(unique_names[i]))
            worksheet.write('E' + str(3 + k), dist)
            worksheet.write('F' + str(3 + k), m.ceil(len(unique_names[i]) * .4))

            if dist <= m.ceil(len(unique_names[i]) * .4):
                cluster.append(unique_names[j])
                worksheet.write('G' + str(3 + k), 'True')
                cluster_sheet.write('A' + str(2 + t), t)
                cluster_sheet.write('B' + str(2+t),frequency[unique_names[i]])
                cluster_sheet.write('C' + str(2 + t), unique_names[i])
                cluster_sheet.write('D' + str(2 + t), frequency[unique_names[j]])
                cluster_sheet.write('E' + str(2 + t), unique_names[j])
                cluster_sheet.write('F' + str(2 + t), len(unique_names[i]))
                cluster_sheet.write('G' + str(2 + t), dist)
                cluster_sheet.write('H' + str(2 + t), m.ceil(len(unique_names[i]) * .4))
                cluster_sheet.write('I' + str(2 + t), 'True')
                t += 1
                if unique_names[j] not in visited:
                    visited.append(unique_names[j])
            else:
                worksheet.write('G' + str(3 + k), 'False')
            k += 1

    workbook.close()

    if len(cluster) != 0:
        clusters.append(cluster)
    cluster = []

text = ""

Cluster_book.close()


#
# for i in range(len(clusters)):
#     if len(clusters[i]) != 0:
#         f = open(clusters[i][0]+".txt", "w")
#         for j in range(len(clusters[i])):
#             f.write(clusters[i][j]+"\n")
#
#
# f = open("names of cluster.txt","w")
# for i in range(len(clusters)):
#     if len(clusters[i]) != 0:
#         f.write(clusters[i][0]+"\n")
#
#
#
#
# count = {}
#
# right_spell = []
# wrong_spell = []
#
# for i in range(len(clusters)):
#     for j in range(len(clusters[i])):
#         if clusters[i][j] in count:
#             count[clusters[i][j]] += 1
#         else:
#             count[clusters[i][j]] = 1
#
#     mirror = clusters[i].copy()
#     mirror = list(set(mirror))
#     total_length = len(clusters[i])
#     mirror_len = len(mirror)
#
#     for j in range(mirror_len):
#         if count[mirror[j]] / total_length <= .35 and mirror[j] not in right_spell:
#             wrong_spell.append(mirror[j])
#         elif mirror[j] not in right_spell:
#             right_spell.append(mirror[j])
#
#
# print("right Spell words :", right_spell)
# print("wrong Spell words :", wrong_spell)
# r = open("right spell.txt","w")
# w = open("wrong spell.txt","w")
#
# for i in range(len(right_spell)):
#     r.write(right_spell[i]+'\n')
#
# for i in range(len(wrong_spell)):
#     w.write(wrong_spell[i]+'\n')






