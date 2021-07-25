import urllib.request
import json
from docxtpl import DocxTemplate

def get_data():
    with urllib.request.urlopen("https://raw.githubusercontent.com/wiki/djay/covidthailand/cases_briefings") as url:
        latest_data = json.loads(url.read().decode())[-1]

    return latest_data

def data_processing():
    raw_data = get_data()
    keywords = ["Date","Cases","Deaths","Deaths Age Max","Deaths Age Min","Deaths Male","Deaths Female","Recovered"]
    th_kw = {
        "Date":"date","Cases":"cases","Deaths":"deaths","Deaths Age Max":"dmx",
        "Deaths Age Min":"dmn","Deaths Male":"dml","Deaths Female":"dfm","Recovered":"recovered"
    }
    data = {}
        
    for btf in keywords:
        thshit = th_kw[btf]
        if isinstance(raw_data[btf], str):
            data[thshit] = str(raw_data[btf])
        if isinstance(raw_data[btf], float):
            data[thshit] = str(int(raw_data[btf]))

    return data

def gen_file():
    print("[*] Fetching Data...")
    raw_data = get_data()

    print("[+] Data Fetched!!!")
    print("[*] Processing Data...")
    context = data_processing()
    filename = "Result_" + raw_data['Date'] + ".docx"
    print("[+] Data Processed!!!")
    print("[*] Outputting result to a (docx) file...")
    doc = DocxTemplate("template.docx")
    doc.render(context)
    doc.save(filename)

    print("[+] Result file is generated as",filename)

if __name__ == "__main__":
    print("Should be done in seconds!!")
    gen_file()
