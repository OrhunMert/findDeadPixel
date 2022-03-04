import pandas as pd
import numpy as np
import xlsxwriter
import math

"""
3 tane sorum var. 

1.Oklid distance matrisinde 0 olan yerlerde Vest nasıl hesaplanacak? Ben 0 ise 0 olarak kabul edip, o değeri hesabaa katmadım.
2.si ise avg deki formul m = 8 mi yoksa m = 9 mu? G5 orta noktası o toplama dahil m?
3.Vth hiçbir zaman Vdif'ten büyük olmayacak? Bunu nasıl gideririz?Gerekli olan koşul bu mu? Vdif < Vth
"""

FileName = "10_output_C.xlsx"

sheetName = "means"

isPrint = True

def ReadExcel(FileName , sheetName):
    
    df = pd.read_excel(FileName , sheet_name = sheetName,engine = "openpyxl")
    
    return np.array(df)

def WriteExcel(G5_matrix , Vest_matrix , Vdif_matrix , Vavg_matrix , Vth_matrix , countDeadPixel):
    
    """
    All parameter's size must be same. if it isn't same, The function won't execute correctly.
    """
    
    workbook = xlsxwriter.Workbook("Block_Matrix_output.xlsx") 
    data_format = workbook.add_format({'num_format': '0.00'})
    
    worksheet_G5   = workbook.add_worksheet("G5")
    worksheet_Vest = workbook.add_worksheet("Vest")
    worksheet_Vdif = workbook.add_worksheet("Vdif")
    worksheet_Vavg = workbook.add_worksheet("Vavg")
    worksheet_Vth  = workbook.add_worksheet("Vth")
    
    worksheet_G5.write(0,0,"count * =  ")
    worksheet_G5.write(0,1,countDeadPixel)
    
    for i in range(0 , len(G5_matrix)):
        for j in range(0 ,len(G5_matrix[0])):
            
            worksheet_G5.write(i+1,j,G5_matrix[i][j] , data_format)  
            worksheet_Vest.write(i,j,Vest_matrix[i][j] , data_format)  
            worksheet_Vdif.write(i,j,Vdif_matrix[i][j] , data_format)  
            worksheet_Vavg.write(i,j,Vavg_matrix[i][j] , data_format)  
            worksheet_Vth.write(i,j,Vth_matrix[i][j] , data_format)  
            
    workbook.close()

def WriteFile(File , indexX , indexY , Value):
    
    File.write(str(indexX)+".row "+str(indexY)+".column  G5 Value:"+str(Value)+"\n")    

def FindDimensions(matrix):
    
    rowNum = len(matrix)
    colNum = len(matrix[0])
    
    return rowNum , colNum

def calculateG5(block_matrix):
    
    """
    G5 is mid value of block_matrix.
    block_matrix's size must be (3x3)
    
    """
    return block_matrix[1][1]

def calculateVest(block_matrix , Distance_Matrix):
    
    """
    block_matrix must be 3x3
    Distance_Matrix must be 3x3
    
    """
    
    Vest    = 0
    value_1 = 0
    value_2 = 0
    
    
    for i in range(0 , len(block_matrix)):
        
        for j in range(0 , len(block_matrix[0])):
            
            #  if Distance_Matrix[i][j] == 0 ? --> what is the result it ?
            
            if Distance_Matrix[i][j] == 0:
                value_1+=0
                value_2+=0
            
            else:
                value_1+=block_matrix[i][j]/Distance_Matrix[i][j]
                value_2+=1/Distance_Matrix[i][j]
                    
                
    Vest = value_1/value_2
    
    return Vest

def calculateVdif(Vest , G5):
    
    """
    Vest is returned value from calculateVest function.
    G5 is mid point of block matrix
    
    """
    
    return abs(Vest - G5)

def calculateVavg(block_matrix , G5):
    
    """
    block_matrix must be 3x3
    G5 is mid point of block matrix
    
    """
    countSum = 0
    
    for i in range(0 , len(block_matrix)):
        for j in range(0 , len(block_matrix[0])):
            
            countSum += block_matrix[i][j]
    
    Transaction = countSum/(len(block_matrix)*len(block_matrix[0]))
    Vavg = (Transaction + G5)/2
    
    return Vavg
    
def calculateVth(Vavg , V0 = 0):

    """
    Vavg is returned value from calculateVavg function.
    V0 is between with 0 and 255 but our value is fixed now.
    
    """
    
    return abs(Vavg - V0)

def isDeadPixcel(Vdif , Vth):
    
    control_bool = True
    
    if Vdif < Vth:
        control_bool = False
    
    return control_bool

def main(FileName , sheetName , isPrint = True):
    
    file = open("notInterval_output.txt" , "w+")
    
    OklitDistance_matrix = [
        [math.sqrt(2) , 1  ,math.sqrt(2)],
        [   1 ,   0 ,    1],
        [math.sqrt(2) , 1 , math.sqrt(2)]
        ]
    
    
    Matrix = ReadExcel(FileName , sheetName)
    rowNum , colNum = FindDimensions(Matrix)
    
    """
    G5_matrix   = [[0 for j in range(colNum-2)] for i in range(rowNum-2)]
    Vest_matrix = [[0 for j in range(colNum-2)] for i in range(rowNum-2)]
    Vdif_matrix = [[0 for j in range(colNum-2)] for i in range(rowNum-2)]
    Vavg_matrix = [[0 for j in range(colNum-2)] for i in range(rowNum-2)]
    Vth_matrix  = [[0 for j in range(colNum-2)] for i in range(rowNum-2)]
    
    """
    
    G5_matrix   = [[0 for j in range(200)] for i in range(100)]
    Vest_matrix = [[0 for j in range(200)] for i in range(100)]
    Vdif_matrix = [[0 for j in range(200)] for i in range(100)]
    Vavg_matrix = [[0 for j in range(200)] for i in range(100)]
    Vth_matrix  = [[0 for j in range(200)] for i in range(100)]
    
    countDeadPixel = 0
    
    for i in range(0 , 100):
        for j in range(0 , 200):
            
             block_matrix = Matrix[i:3+i , j:3+j] # a matrix is created and this matrix's size : 3x3 
             G5   = calculateG5(block_matrix)
             Vest = calculateVest(block_matrix, OklitDistance_matrix)
             Vdif = calculateVdif(Vest, G5)
             Vavg = calculateVavg(block_matrix, G5)
             Vth  = calculateVth(Vavg)
             
             
             if isPrint:
                
                 print("\nblock matrix:\n"+str(block_matrix)+"\n")
                 print("G5: "+str(G5))
                 print("Vest: "+str(Vest))
                 print("Vdif: "+str(Vdif))
                 print("Vavg: "+str(Vavg))
                 print("Vth: "+str(Vth))
                
             
             control_bool = isDeadPixcel(Vdif, Vth) 
             
             if control_bool:
                 
                 print("\nG5 is not between avg and DiF!!! ")
                 print("Block Matrix's indexes is:\nrow number(starting from 0):"+str(i)+" column number(starting from 0):"+str(j))   
                 
                 G5_matrix[i][j]   = str(G5) + "*"
                 Vest_matrix[i][j] = Vest
                 Vdif_matrix[i][j] = Vdif
                 Vavg_matrix[i][j] = Vavg
                 Vth_matrix[i][j]  = Vth
                  
                 countDeadPixel+=1
                    
                 WriteFile(file, i, j, G5)
            
             else:
                
                G5_matrix[i][j]   = str(G5) + ""
                Vest_matrix[i][j] = Vest
                Vdif_matrix[i][j] = Vdif
                Vavg_matrix[i][j] = Vavg
                Vth_matrix[i][j]  = Vth
                
    WriteExcel(G5_matrix , Vest_matrix , Vdif_matrix , Vavg_matrix , Vth_matrix , countDeadPixel)            

#-------- running ------- 
  
main(FileName , sheetName , isPrint)

#------------------------
