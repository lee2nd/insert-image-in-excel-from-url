import pandas as pd
import xlwings as xw
import numpy as np
from PIL import Image
import requests
from io import BytesIO
import matplotlib.pyplot as plt

def fig_trans(url): 
    # get the url and transform into image
    
    response = requests.get(url)
    img = Image.open(BytesIO(response.content))
    img_array = np.array(img)
    fig, ax = plt.subplots(figsize=(1, 1))
    ax.imshow(img_array,cmap='gray')
    ax.axis('off')
    
    return fig
  
df = pd.read_csv("raw.csv")

wb = xw.Book()
ws = wb.sheets[0]

ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = df
img_lst = list(df["Uploadpath"])

i = 1
for img in img_lst:
    # transform all the urls and paste the output images in the worksheet directly
    ws.pictures.add(fig_trans(img),left=ws.range('S'+str(i+1)).left, top=ws.range('S'+str(i+1)).top)
    # auto adjust the row height
    ws.range("A1:A"+str(len(df)+1)).row_height = 60
    i+=1

wb.save('demo.xlsx')
wb.close()
