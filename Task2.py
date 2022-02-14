import os
import subprocess
import re
import json
from openpyxl import Workbook
import openpyxl


print("What format are you need? Please, write the number. \n" + "1. xlsx \n" + "2. json")
format = input()
while format:
    if format is "1" or format is "2":
        break
    else:
        print("Error format! Please, try again.")
        format = input()



if format is "1":
    wb = Workbook()
    ws = wb.active
    ws.title = "New Title"
    ws['A1'] = "#N"
    ws['B1'] = "Channel"
    ws['C1'] = "Address"
    ws['D1'] = "Resolution"
    ws['E1'] = "V-codec"
    ws['F1'] = "A-codec"
else:
    json_f = open("m9.json", "w")
    json_f.close()




i = 151
#check = 0
while i != 152:
    i += 1


    mCast = "http://207.110.52.50:4022/udp/233.166.172." + str(i) + ":1234"



    
    output = (subprocess.run(["docker", "run", "-it", "--rm", "--network", "host", "nfs01.techstudio.tv/ffprobe:latest", "-v", "quiet", "-print_format", "json", "-show_format", "-show_streams", "-i", mCast], stdout=subprocess.PIPE))
    output = str(output).replace("\\r\\n", "")  
    output = output.replace(" ", "")
    #print(output)
    if i < 10:
        output = output[253:-2]
    if i >=10 and i <100:
        output = output[254:-2]
    if i >=100:
        output = output[255:-2]
    print(output)
    data = json.loads(output)
    if not data:
        continue


    if format is "1":
        ws['A' + str(i + 1)] = i
        ws['C' + str(i + 1)] = mCast
        aud_check = 0
        for check in data["streams"]:

            if check["codec_type"] not in data["streams"]:
                continue
            if aud_check != 0:
                if re.match("audio", check["codec_type"]):
                    ws['G' + str(i + 1)] = check["codec_name"]
                    aud_check = 0
            elif check["codec_type"] == "video":
                ws['E' + str(i + 1)] = check["codec_name"]
                resolution = str(check["width"]) + "x" + str(check["height"])
            elif check["codec_type"] == "audio":
                ws['F' + str(i + 1)] = check["codec_name"]
                aud_check += 1
            elif check["codec_type"] == "subtitle":
                ws['H' + str(i + 1)] = check["codec_name"]
            

        #if data["streams"][check]["codec_type"] == "audio":
        #    ws['F' + str(i + 1)] = data["streams"][check]["codec_name"]
        #    check += 1

        #    if data["streams"][check]["codec_type"] == "audio":
        #        ws['F' + str(i + 1)] = data["streams"][check]["codec_name"]
        #        check += 1
        #        if data["streams"][check]["codec_type"] != "video":
        #            ws['F' + str(i + 1)] = data["streams"][check]["codec_name"]
        #            resolution = "--------"
        #if data["streams"][check]["codec_type"] == "video": 
        #    ws['E' + str(i + 1)] = data["streams"][check]["codec_name"]
        #    if check == 0:
        #        ws['F' + str(i + 1)] = data["streams"][check + 1]["codec_name"]
        #    resolution = str(data["streams"][check]["width"]) + "x" + str(data["streams"][check]["height"])
        #else:
        #    resolution = "--------"

        if not resolution:
            resolution = "--------"
        ws['D' + str(i + 1)] = resolution
    else:
        json_f = open("m9.json", "a")
        json_f.write(json.dumps(data))
        json_f.close



    
j = 11
while j < 10:
    i += 1
    j += 1

    mCast = "http://207.110.52.50:4022/udp/233.166.173." + str(j) + ":1234"



    output = str(subprocess.run(["docker", "run", "-it", "--rm", "--network", "host", "nfs01.techstudio.tv/ffprobe:latest", "-v", "quiet", "-print_format", "json", "-show_format", "-show_streams", "-i", mCast], stdout=subprocess.PIPE))
    output = str(output).replace("\\r\\n", "")  
    output = output.replace(" ", "")
    print(output)
    if j < 10:
        output = output[253:-2]
    if j >=10 and j <100:
        output = output[254:-2]
    if j >=100:
        output = output[255:-2]
    data = json.loads(output)



    if format is "1":
        ws['A' + str(i + 1)] = i
        ws['C' + str(i + 1)] = mCast
        if not data:
            continue
        check = 0
        if data["streams"][check]["codec_type"] == "audio":
            ws['F' + str(i + 1)] = data["streams"][check]["codec_name"]
            check += 1

            if data["streams"][check]["codec_type"] == "audio":
                ws['F' + str(i + 1)] = data["streams"][check]["codec_name"]
                check += 1
                if data["streams"][check]["codec_type"] != "video":
                    ws['F' + str(i + 1)] = data["streams"][check]["codec_name"]
                    resolution = "--------"
        if data["streams"][check]["codec_type"] == "video": 
            ws['E' + str(i + 1)] = data["streams"][check]["codec_name"]
            if check == 0:
                ws['F' + str(i + 1)] = data["streams"][check + 1]["codec_name"]
            resolution = str(data["streams"][check]["width"]) + "x" + str(data["streams"][check]["height"])
        else:
            resolution = "--------"

        ws['D' + str(i + 1)] = resolution
    else:
        json_f = open("m9.json", "a")
        json_f.write(json.dumps(data))
        json_f.close


if format is "1":
    wb.save('Channels.xlsx')
