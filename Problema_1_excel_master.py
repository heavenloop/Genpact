# -*- coding: utf-8 -*-
"Build a solution, which monitor a folder looking for new files."
#def buscar_excels_en_carpeta(path):
def Get_listOfFiles(path):
    import os
    listOfFiles_list=os.listdir(path)
    #Si existe el archivo removerlo
    if 'MasterBook.xlsx' in listOfFiles_list:
        os.remove('C:\\Users\\danie\\Desktop\\Genpac folder\\Problema 1\\MasterBook.xlsx') 
    return listOfFiles_list

"Each time a file is found, it should verify if is an excel file (.xls* files)"
". If is true, it should take each sheet on it and consolidate it on a master workbook file (make a copy from each sheet to the master file)."
def buildMasterExcel(listOfFiles_list,path):
    #Creacion del master book
    import pandas as pd
    import os
    import shutil
    if os.path.isdir(path+"\\Processed") is False:
        os.makedirs(path+"\\Processed")
        os.makedirs(path+"\\Not_applicable") 
    Total_sheets={}
    j=0
    for i in listOfFiles_list:
        #Si termina en .xls
        if i[-5:-1] ==".xls" and i != 'Masterbook.xlsx':
            #leer el archivo y estraer las 3 hojas como data frames
            sheets_df=pd.read_excel('C:\\Users\\danie\\Desktop\\Genpac folder\\Problema 1\\'+i, sheet_name=None)
            #asignar un nombre a cada hoja para evitar nombres repetidos
               
            #para cada titulo de las 3 hojas
            for titulos in sheets_df.keys():
                j+=1
                Nuevos_titulos='Sheet'+str(j)
                print(Nuevos_titulos)
                #Cambio a nuevo titulo de hoja sheet j
                Total_sheets[Nuevos_titulos]=sheets_df[titulos]
                #del sheets_df[titulos]           
    
    
            #Escribir los sheets extraidos como diccionarios en el destino masterbook
            writer=pd.ExcelWriter('Masterbook.xlsx',engine='xlsxwriter')
            #para cada hoja, en el dataframe sheets_df..
            for sheet_name in Total_sheets.keys():
               #tomar a partir de los nombres de las hojas y escribir en excel master     
                Total_sheets[sheet_name].to_excel(writer,sheet_name=sheet_name,index=False)
            writer.save()
            shutil.move('C:\\Users\\danie\\Desktop\\Genpac folder\\Problema 1\\'+i , path+"\\Processed")

        elif i[-5:-1] !=".xls" and i != 'Masterbook.xlsx' and i != 'Processed':
            shutil.move('C:\\Users\\danie\\Desktop\\Genpac folder\\Problema 1\\'+i , path+"\\Not_applicable")
    "Every file found should be moved to 2 different folders depending if was or not a excel file"
    #o Processed
    #o Not applicable."        
  
    return Total_sheets

path=input('Please introduce the folders path:')
listOfFiles_list=Get_listOfFiles(path)
buildMasterExcel(listOfFiles_list,path)



        
        


        

        
    
    

