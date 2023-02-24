
# Opening JSON file
import json

myjsonfile=open("C:\\Users\\PRAJWAL\\PycharmProjects\\WebScapping\\excel\\Passion_Process.json","r")
jsondata=myjsonfile.read()

obj=json.loads(jsondata)

li=obj['Process_areas']
process_areas=[]
for i in range(len(li)):
    process_areas.append(li[i].get("Process"))
print(process_areas)