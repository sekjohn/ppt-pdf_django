from django.http import JsonResponse
from rest_framework.decorators import api_view
from api_main.settings import buckets 
from pythoncom import CoInitializeEx
from pythoncom import CoUninitialize
import os 
import shutil
import json

import comtypes.client

def PPTtoPDF(inputFileName, outputFileName):
    res = CoInitializeEx(0)
    formatType=32
    WithWindow=False 
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf" 

    if inputFileName and outputFileName:

        deck = powerpoint.Presentations.Open(inputFileName)
        deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdfs
        deck.Close()
        powerpoint.Quit()
    CoUninitialize()

    
@api_view(["POST"])
def conversion_main(request):
    file_src = json.loads(request.body)
    uploadPath = os.getcwd() +"\src\\"

    if not os.path.exists(uploadPath):
        os.makedirs(uploadPath)
    
    try:
        if file_src:
            
            FileName = str(file_src['src_path']).replace('src/','').replace('.pptx','')
            outputFileName = f"{uploadPath}{FileName}.pdf"
            inputFileName = f"{uploadPath}{FileName}.pptx"
            print(outputFileName )
            s3_name = f"{file_src['src_path']}"
            up_s3_name = f"{file_src['dst_path']}"
        
            buckets.download_file(s3_name , inputFileName)
            PPTtoPDF(inputFileName=inputFileName,outputFileName=outputFileName)
                        
            with open(outputFileName, 'rb') as data:
                buckets.upload_file(data.name, up_s3_name)

            shutil.rmtree(uploadPath)
            
            return JsonResponse({"message":True}, status=200)

    except Exception as e:
        print(f"[ERROR] {e}")
        return JsonResponse({"변환 오류"})
