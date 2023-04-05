import numpy as np
import pandas as pd
import xlwings as xw
import csv
# 0.1
wb=xw.Book('себестоимостьА_в1.xlsx')
sh0=wb.sheets[0]
sh0.range('Q5:Q6').value='Себестоимость'
sh0.range('Q5:Q6').formula='=TEXT('
color=sh0.range('G5:G6').color
sh0.range('Q5:Q6').color=color
sh0.range('Q5:Q6').autofit()
sh0.range('Q5:Q6').api.Font.Bold=True
sums=sum(list(filter(None, sh0.range('G7:O7').value)))
sh0.range('Q7').options(index = False).value = sums
sums1=sum(list(filter(None, sh0.range('G8:O8').value)))
sh0.range('Q8').options(index = False).value = sums1
sums2=sum(list(filter(None, sh0.range('G9:O9').value)))
sh0.range('Q9').options(index = False).value = sums2
sums3=sum(list(filter(None,sh0.range('G10:O10').value)))
sh0.range('Q10').options(index = False).value = sums3
# 1
reviews=pd.read_csv('reviews_sample.csv')
recipes=pd.read_csv('recipes_sample.csv')
recipes=recipes.drop(columns='contributor_id',axis=1)
recipes=recipes.drop(columns='n_steps',axis=1)
print(recipes.shape[0])
print(reviews.shape[0])
# 2
x,x1=recipes.sample(int((recipes.shape[0]/100)*5)),reviews.sample(int((reviews.shape[0]/100)*5))
wb=xw.Book('Рецепты.xlsx')
sh,sh1=wb.sheets[0],wb.sheets[1]
sh.name,sh1.name='Рецепты','Отзывы'
sh.range('A1').options(index = False).value = x
sh1.range('A1').options(index = False).value = x1
print(np.array(sh.range('A1').value))
# 3
p=pd.DataFrame(np.array(sh.range('C2:C1501').value)*60)
sh.range('G1').expand('down').options(index = False).value = p
sh.range('G1').value='seconds_assign'
sh.range('G1').api.Font.Bold=True
sh.range('G1').api.HorizontalAlignment=-4108
# 4
f=sh.range('H2').formula=f'=C2*60'
sh.range('H2:H1501').formula=f
sh.range('H1').value='seconds_formula'
sh.range('H1').api.Font.Bold=True
sh.range('H1').api.HorizontalAlignment=-4108
# 6
n=np.array(sh.range('C2').expand('down').value)
for i,el in enumerate(n):
    if el<=5:
        sh.range(f"C{i+2}").color=(0,255,0)
    elif 5<el<10:
        sh.range(f"C{i+2}").color =(255, 255, 0)
    else:
        sh.range(f"C{i+2}").color = (255, 0, 0)
# 7
wb.sheets[1].activate()
x=pd.DataFrame(np.array(sh1.range('C2:C1501'))).count()
print(x)
# 8
def validate():
    wb.sheets[1].activate()
    rating=np.array(sh1.range('E1').expand('down'))
    wb.sheets[0].activate()
    rating2=np.array(sh.range('B1').expande('down'))
    for i,el in enumerate(rating):
        if el in [1,2,3,4,5] and el in rating2:
            pass
        else:
            sh1.range(f"E{i+2}").color=(255, 0, 0)
validate()
# 9
with open('recipes_model.csv','r',encoding='utf-8') as f:
    r=csv.reader(f,delimiter='\t')
    r=pd.DataFrame(r)
sh3=wb.sheets.add(name='Модель',after=wb.sheets[1].name)
sh3.range('A2').options(index = False).value = r
f=sh.range('J2').formula=f'=B2,C2'
sh.range('H2:H1501').formula=f
sh.range('J1').value='SQL'
# 11
sh3.range("A2:I2").color=(0, 204, 255)
sh3.range("A2:I2").autofit()
sh3.range('A2:I2').api.Font.Bold=True
# 12
import matplotlib.pyplot as plt
from collections import Counter
r=Counter(r)
x =set(r.values())
y =set(r.keys())
fig, ax = plt.subplots()
ax.bar(x, y)
ax.set_facecolor('seashell')
fig.set_facecolor('floralwhite')
fig.set_figwidth(12)
fig.set_figheight(6)
sh4=wb.sheets.add(name='Модель',after=wb.sheets[2].name)
sh4.pictures.add(fig, name='Статистика', update=True)
