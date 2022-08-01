import win32com.client
import uuid
import os.path
import time
from tqdm.auto import tqdm
import fnmatch
from multiprocessing import Pool

def listdir(dirname, pattern="*"):
    return fnmatch.filter(os.listdir(dirname), pattern)
    
 
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True
direc ="./input/"
inputDir = listdir(direc,"*.pptx")
print(inputDir)
com = 0


mbar = tqdm(inputDir , leave=False, position=1)
for pptx in mbar:
        
        file = direc + pptx
        #tqdm.write(file)
        
        output = "./output/" + pptx[:-5]
        Presentation = Application.Presentations.Open(os.path.abspath(file))
        #Presentation.Slides[1].Export(os.path.abspath(output) +"1.jpg", "JPG", 800, 600);
        num = Presentation.Slides.count
        #print(num)
        com = num + com
        for i in tqdm(range(num), position=0):
            #tqdm.write(os.path.abspath(output)  + str(i) +".jpg")
            
            Presentation.Slides[i].Export(os.path.abspath(output)+ "-"  + str(i+1) +".jpg", "JPG", 1440, 1882)
            mbar.refresh()
        
        Presentation.close()
        time.sleep(1)
        

os.system("TASKKILL /F /IM powerpnt.exe")
mbar.update(1)
print("\n",com)          
