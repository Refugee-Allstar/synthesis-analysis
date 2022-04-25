import pandas as pd
import numpy as np

df_products = pd.read_csv('products.tsv',sep='\t')
df_reviews = df_products = pd.read_csv('reviews.tsv',sep='\t')

df_products.to_excel('products.xlsx')
df_reviews.to_excel('reviews.xlsx')
def check_data_type(df):

    return df.dtypes

 
def product_data(products):
    #Read in Product Data
    df_products = pd.read_csv(products,sep='\t')
    #Rename product ID column to match other datafram
    df_products.rename(columns={'product_id':'productId'}, inplace=True)
    return df_products

def review_data(reviews):
    #Read in Review Data
    df_reviews = pd.read_csv(reviews, sep='\t')
    return df_reviews

def merged_data():
    #merge dataframes based on product ID and split the first column apart for more granular analysis
    df_merged = pd.merge(product_data('products.tsv'), review_data('reviews.tsv'), on="productId")
    
    #The first two delimiters were always brand and material
    df_merged[['Brand', 'Material']] = df_merged['product_name'].str.split(',', n=1, expand=True)
    #Separating out the other descriptors from the material column
    df_merged['Material'] = df_merged['Material'].str.split(',', n=0, expand=True)[0]
    #Standarizing the Material Type for Easier Analysis
    df_merged.loc[df_merged['Material'].str.contains('KN95'), 'Material'] = "KN95"
    df_merged.loc[df_merged['Material'].str.contains('Reus'), 'Material'] = "Resuable"
    df_merged.loc[df_merged['Material'].str.contains('Disposable'), 'Material'] = "Disposable"
    #The final delimiter was always the quantity, strip all characters and change the data type to float
    df_merged['Quantity'] = df_merged['product_name'].str.split(',').str[-1]
    df_merged['Quantity'] = df_merged['Quantity'].str.replace(r'\D', '', regex=True)
    df_merged['Quantity'] = df_merged['Quantity'].astype(float)
    #get price per unit for comparison
    df_merged['Price Per Unit'] = df_merged['product_price']/df_merged['Quantity']
    #Some had an additional description
    df_merged['Descriptor'] = df_merged['product_name'].str.split(',').str[1:4]
    df_merged['Descriptor'] = df_merged['Descriptor'].astype(str)
    df_merged = df_merged.drop(['product_name'],axis=1)
    return df_merged

#Group based on various characteristics to look for trends and output to a single excel file
def avg_by_descriptors():
    #avg ratings by brand and style
    df = merged_data()
    df_brand = df.groupby(['Brand','Descriptor']).count()
    df_by_country = df.groupby(['Brand','languageCode'].mean())
    df_brand.to_excel('analysis.xlsx')
    df_by_country.to_excel('analysis_country.xlsx')

#Column 1 has a lot of interesting data -- worth splitting this into multiple columns to be able run analysis by pack size, type etc

 

def word_frequency():
    df = merged_data()
    df.loc[df['languageCode.1'] == 'en-US', 'Translated'] = df['reviewText']
    df['translation.reviewText'] = df['translation.reviewText'].astype(str)
    df['Translated'] = df['Translated'].astype(str)
    df['All Reviews'] = df[['translation.reviewText', 'Translated']].agg(''.join, axis=1)
    df['All Reviews'] = df['All Reviews'].str.replace('nan','')

    df.to_excel('consolidated.xlsx')
    
def crosstabreport(dataframe, categories, droppedcolumns='none'):

    if droppedcolumns != 'none':
        dataframe = dataframe.drop(droppedcolumns,axis=1)
    else:
        pass
    dataframe['ratingValue'] = dataframe['ratingValue'].apply(lambda x: x/10)
    report = dataframe.groupby(categories).agg({'Descriptor': 'count', 'ratingValue': 'mean'}).sort_values(by=['ratingValue'], axis=0,ascending=False)
    report = report.round(decimals=2)
    print(report)
    return report

   
    

def create_excel_report():
    writer = pd.ExcelWriter('Facemask Market Analysis.xlsx')
    workbook = writer.book
    brandReport = crosstabreport(merged_data(), ['Brand'], ['productId','product_price',
                                                       'abuseCount','helpfulNo','helpfulYes',
                                                       'imagesCount','profileInfo.ugcSummary.answerCount',
                                                       'reviewed','score'])
    brandReport.to_excel(writer, sheet_name="Brand Analysis",startrow=1, startcol=0)
    brandDescriptorReport = crosstabreport(merged_data(), ['Brand','Descriptor'], ['productId','product_price',
                                                       'abuseCount','helpfulNo','helpfulYes',
                                                       'imagesCount','profileInfo.ugcSummary.answerCount',
                                                       'reviewed','score'])
    brandDescriptorReport.to_excel(writer, sheet_name="Brand Analysis",startrow=1, startcol=5)                                                
    writer.save()
word_frequency()