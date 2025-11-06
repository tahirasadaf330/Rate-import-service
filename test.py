import pandas as pd

df =  pd.read_excel(r'C:\Users\Tahira Sadaf\Desktop\projects\rate-import-service\attachments\rates_at_evox.fr_20251106_091941\CPL_HAYOTEL_DEU-2025116-43120_.xls',engine="calamine")

print(df.shape)
# print(df.head())