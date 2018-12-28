import comtypes.client
import os 
def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] == 'ppt':
        outputFileName = outputFileName[0:-4] + ".pdf"
        print(outputFileName)
    if outputFileName[-4:] == 'pptx':
        outputFileName = outputFileName[0:-5] + ".pdf"
        print(outputFileName)
    deck = powerpoint.Presentations.Open(inputFileName,WithWindow=False)
    deck.SaveAs(outputFileName,formatType)                          # formatType = 32 for ppt to pdf
    deck.Close()

def convert_files_in_folder(powerpoint, folder):
    files = os.listdir(folder)                                      #回指定文件夹包含的文件或文件夹名字的列表
    pptfiles = [f for f in files if f.endswith((".ppt",".pptx"))]   #使用循环批量转换
    for pptfile in pptfiles:
        fullpath = os.path.join(cwd,pptfile)                        #将多个路径组合后返回
        fullpath2 =os.path.join(cwd,"pdf",pptfile)
        ppt_to_pdf(powerpoint, fullpath, fullpath2)

if __name__ == "__main__":
    powerpoint = init_powerpoint()
    cwd = os.getcwd()  # 返回当前进程的目录
    convert_files_in_folder(powerpoint, cwd)
    powerpoint.Quit()
