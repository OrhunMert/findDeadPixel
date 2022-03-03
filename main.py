import pandas as pd
import numpy as np
import xlsxwriter

FileName = "1_output_C.xlsx" # input file name as excel type.

sheetName = "means"

isPrint = True # True or False. if you want, can see GH,GL,DiF,G5,avg and block matrix.

def ReadExcel(File_Name ,sheetName):
    
    df = pd.read_excel(File_Name ,sheet_name = sheetName ,engine='openpyxl')
    matrix = np.array(df)
    
    return matrix

def WriteExcel(avg_matrix , DiF_matrix , G5_matrix , countInterval):
    
    """
    All parameter's size must be same. if it isn't same, The function won't execute correctly.
    """
    
    workbook = xlsxwriter.Workbook("block_matrix_output.xlsx") 
    data_format = workbook.add_format({'num_format': '0.00'})
    
    worksheet_avg = workbook.add_worksheet("avg")
    worksheet_DiF = workbook.add_worksheet("DiF")
    worksheet_G5 = workbook.add_worksheet("G5")
    
    worksheet_G5.write(0,0,"count '*' =  ")
    worksheet_G5.write(0,1,countInterval)
    
    for i in range(0 , len(avg_matrix)):
        for j in range(0 ,len(avg_matrix[0])):
            
            worksheet_avg.write(i,j,avg_matrix[i][j] , data_format)
            worksheet_DiF.write(i,j,DiF_matrix[i][j] , data_format)
            worksheet_G5.write(i+1,j,G5_matrix[i][j] , data_format)
            
    workbook.close()

def WriteFile(File , indexX , indexY , Value):
    
    File.write(str(indexX)+".row "+str(indexY)+".column  G5 Value:"+str(Value)+"\n")    
    
    
    
def FindDimensions(matrix):
    
    rowNumber = len(matrix)
    colNumber = len(matrix[0])
    
    return rowNumber , colNumber


def findSecondMax(matrix):
    
    # matrix must be a numpy array.
    
    max_value = np.max(matrix)
    
    flat=matrix.flatten()
    flat.sort()
   
    # we are controling ; is the max value equals secondMax value? if = true--> you pass previous value for secondMax value if = false-->you found your secondMax value.
    # For Examle: Your data are : [ 1 2 3 4 5 5] in the this point, Your second max value is 5 but it's not true. We are finding value 4 with this while loop.
    i = -2
    while(flat[i] == max_value):
        i-=1
        
        
   # it is second max value.
   
    return flat[i]

def findSecomdMin(matrix):
    
    # matrix must be a numpy array.
    
    min_value = np.min(matrix)
    
    flat=matrix.flatten()
    flat.sort()
    
    secondMin = 0
    
    for i in range(1 , len(flat)):
        
        if flat[i] != min_value:
            secondMin = flat[i]
            break
    
    return secondMin

def findMidPointofMatrix(matrix):
    
    """
    matrix must be a numpy array. We expect matrix's size is 3x3. if it's size isn't 3x3 , The found value will be incorrect.
    
    """
    
    return matrix[1][1]
 

def controlIntervalBounds(avg , DiF , G5):
    
    """
    G5 is mid point of 3x3 matrix
    
    """
    interval_upperBound = avg + DiF
    interval_lowerBound = avg - DiF
    
    control_bool = True
    
    # aralıklar eşitlik olacak mı ?
    if G5 >= interval_lowerBound and G5 <= interval_upperBound:
        control_bool = False
        
    return control_bool

def firstTransaction(GH , GL):
    return GH - GL

def secondTransaction(matrix , GH , GL , G5):
    
    """
    matrix must be a numpy array and size' 3x3
    GH is second max value of matrix
    GL is second min value of matrix
    G5 is mid value of matrix. G5 always mid point for 3x3 matrix.

    """
    sum_value = matrix.sum()
    avg = (sum_value - (GH + GL + G5)) / 6
    
    
    """
    row_matrix = matrix.flatten()
    
    for i in range(0 , len(row_matrix)):
        
        calculate_sum = row_matrix[i] - (G5 + GH + GL)
        
        sum_value+=calculate_sum
    
    avg = sum_value / 6
    """
    
    return avg

def main(FileName ,sheetName , isPrint = True):
    
    file = open("notInterval_output.txt" , "w+")
    
    Matrix = ReadExcel(FileName,sheetName)
    rowNum , colNum = FindDimensions(Matrix)
    
    avg_matrix = [[0 for j in range(colNum-2)] for i in range(rowNum-2)] # colNum - 2 rowNum - 2
    DiF_matrix = [[0 for j in range(colNum-2)] for i in range(rowNum-2)]
    G5_matrix  = [[0 for j in range(colNum-2)] for i in range(rowNum-2)]

    print("Matrix:\n"+str(Matrix))
    print("\nrow number: "+str(rowNum)+" col number: "+str(colNum))
    
    countInterval = 0 # pixel data's count not in interval
    
    for i in range(0 , rowNum-2): #rowNum - 2
        for j in range(0 , colNum-2): #colNum - 2
           
            print("\n"+str(i)+".row "+str(j)+".column")
           
            block_matrix = Matrix[i:3+i , j:3+j] # a matrix is created and this matrix's size : 3x3
            GH  = findSecondMax(block_matrix)
            GL  = findSecomdMin(block_matrix)
            DiF = firstTransaction(GH, GL)
            G5  = findMidPointofMatrix(block_matrix)
            avg = secondTransaction(block_matrix, GH, GL, G5)
            
            if isPrint:
                
                print("\nblock matrix:\n"+str(block_matrix))
                print("\nGH: "+str(GH))
                print("\nGL: "+str(GL))
                print("\nDiF: "+str(DiF))
                print("\nG5: "+str(G5))
                print("\navg: "+str(avg))
            
            control_bool = controlIntervalBounds(avg, DiF, G5)
            
            if control_bool:
                print("\nG5 is not between avg and DiF!!! ")
                print("Block Matrix's indexes is:\nrow number(starting from 0):"+str(i)+" column number(starting from 0):"+str(j))
                
                avg_matrix[i][j] = avg
                DiF_matrix[i][j] = DiF
                G5_matrix[i][j]  = str(G5)+"*"
                countInterval+=1
                
                WriteFile(file, i, j, G5)
                
            
            else:
                avg_matrix[i][j] = avg
                DiF_matrix[i][j] = DiF
                G5_matrix[i][j]  = str(G5)+""
            
    WriteExcel(avg_matrix, DiF_matrix, G5_matrix, countInterval)
   
    
    file.close()
    
    
#------- running -------

main(FileName ,sheetName , isPrint)

#-----------------------
    
