import pandas as pd
import numpy as np
import math
from PIL import Image
from wordcloud import WordCloud, ImageColorGenerator
import xlsxwriter

cluster_name = 'malaysia'
data = pd.read_excel('names-result.xlsx', sheet_name='clusters')
data = data.dropna()
print(list(data))
df = data.where(data['Cluster Name'] == cluster_name)
frequency = []
qatar_f = set(df['frequency-A'].dropna().tolist())
frequency.append(list(qatar_f)[0])
other_f = list(df['frequency-B'].dropna())
for i in range(len(other_f)):
    frequency.append(other_f[i])

name = list(set(df['Cluster Name'].dropna()))
other_name = df['Other Member Name'].dropna().tolist()



for i in range(len(other_name)):
    name.append(other_name[i])

word_frequency = {}

for i in range(len(name)):
    word_frequency[name[i]] = math.log(int(frequency[i]))+1


# text = text[:-1]
# print(text)
mask = np.array(Image.open('color.jpg'))
wc = WordCloud(font_path="BebasNeue Bold.otf",background_color='white',width=500,height=150, prefer_horizontal=1, min_font_size=7,
               repeat=False, mask=mask)
wc.generate_from_frequencies(word_frequency)
image_colors = ImageColorGenerator(mask)
wc.recolor(color_func=image_colors)


wc.to_file(cluster_name+".jpg")