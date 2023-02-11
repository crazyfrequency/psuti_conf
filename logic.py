from PSUTI import *
from lib import *
import os, json, dataclasses
import asyncio

class CustomJSONEncored(json.JSONEncoder):
    def default(self, o):
        if dataclasses.is_dataclass(o):
            return dataclasses.asdict(o)
        return super().default(o)

def check_path(path: str="C:/")->bool:
    try:
        if(os.path.exists(os.path.join(path,"psuti_conf_config.json"))):return True
        if(len(os.listdir(path))==0):return True
        return False
    except:
        return False

class Download:
    def __init__(self, path:str="", psuti:PSUTI=None) -> None:
        self.path = path
        self.psuti = psuti
        self.stopped = False

    def start(self, func, end):
        self.func = func
        main_file = open(os.path.join(self.path,"psuti_conf_config.json"),"w")
        sections = self.psuti.get_all_sections()
        count=0
        now_count=0
        for section in sections:
            count+=section.count
            try:
                os.mkdir(os.path.join(self.path,section.title))
            except:pass
        main_file.write(json.dumps(sections,cls=CustomJSONEncored))
        main_file.close()
        for section in sections:
            participants = self.psuti.get_all_participants(section.url)
            file = open(os.path.join(self.path,section.title,"psuti_section_config.json"),"w")
            file.write(json.dumps(participants,cls=CustomJSONEncored))
            file.close()
            full_data = []
            for participant in participants:
                if(self.stopped):return end()
                participant_data = self.psuti.get_participant(participant.url)
                full_data.append(participant_data)
                self.psuti.download(participant_data.document.url,os.path.join(self.path,section.title,participant_data.document.title))
                now_count+=1
                func(now_count,count)
            file = open(os.path.join(self.path,section.title,"psuti_section_config_full.json"),"w")
            file.write(json.dumps(full_data,cls=CustomJSONEncored))
            file.close()
        end()

    def stop(self):
        self.stopped = True
        
class MakeConf(MSWord):
    def __init__(self) -> None:
        pass
    def start(self, func, mode1:bool, mode2:bool, data_path:str, res_path:str, first_file_path:str, end, cancel, error):
        super().__init__()
        self.func = func
        self.mode = mode2
        self.data_path = data_path.replace("/","\\")
        self.res_path = res_path.replace("/","\\")
        self.first_file_path = first_file_path.replace("/","\\")
        self.stopped = False
        self.on_end = end
        self.on_cancel = cancel
        self.error = error
        try:
            name = os.path.join(self.data_path,"temp.docx").replace("/","\\")
            self.new(name)
            self.set_def_format(name)
            self.open(self.first_file_path)
            self.copy(self.first_file_path,name)
            self.close(self.first_file_path)
            if(mode1):self._start_mode1()
            else:self._start_mode2()
        except Exception as err:
            self.word.Quit(True)
            self.error(f'{type(err)}\n{err}\n{err.args}')

    def _start_mode1(self):
        name = os.path.join(self.data_path,"temp.docx").replace("/","\\")
        doc=self.word.Documents(name)
        doc.Range(doc.Content.End-1,doc.Content.End-1).InsertBreak()
        doc.Sections(1).Footers(1).PageNumbers.Add(1,True)
        doc.Sections.Add(doc.Range(doc.Content.End-1,doc.Content.End-1),0)
        doc.Sections(doc.Sections.Count).Headers(1).LinkToPrevious=False
        file = open(os.path.join(self.data_path,"psuti_conf_config.json"))
        sections:list[Sections] = json.loads(file.read())
        file.close()
        count = sum([i['count'] for i in sections])
        now_count = 0
        pages = {}
        for section in sections:
            file = open(os.path.join(self.data_path,section['title'],"psuti_section_config_full.json"))
            all_thesis:list[Thesis] = json.loads(file.read())
            file.close()
            doc.Sections(doc.Sections.Count).Headers(1).Range.Text=section['title']
            doc.Sections(doc.Sections.Count).Headers(1).Range.ParagraphFormat.Alignment=2
            doc.Range(doc.Content.End-1,doc.Content.End-1).Text=section['title']+"\r"
            doc.Range(doc.Content.End-len(section['title'])-2,doc.Content.End-1).Font.Size = 16
            doc.Range(doc.Content.End-len(section['title'])-2,doc.Content.End-1).Font.Bold = 1
            doc.Range(doc.Content.End-len(section['title'])-2).ParagraphFormat.Alignment=1
            for thesis in all_thesis:
                if(self.stopped):
                    self.on_cancel()
                    return self.end()
                now_count+=1
                self.func(now_count,count)
                path = os.path.join(self.data_path,section['title'],thesis['document']['title']).replace("/","\\")
                self.open(path)
                start_page = doc.ActiveWindow.Panes(1).Pages.Count
                self.copy(path,name)
                end_page = doc.ActiveWindow.Panes(1).Pages.Count
                self.close(path)
                if(start_page==end_page):page = f', {start_page}'
                else:page = f', {start_page}-{end_page}'
                for author in thesis['participants']:
                    fio = f"{author['surname']} {author['name'][0]}. {author['patronymic'][0]}."
                    if(fio in pages):pages[fio]+= page
                    else:pages[fio] = "с. " + page[2:]
            doc.Sections.Add(doc.Range(doc.Content.End-1,doc.Content.End-1),0)
            doc.Sections(doc.Sections.Count).Headers(1).LinkToPrevious=False
        if(self.mode):
            pages_keys = list(pages.keys())
            pages_keys.sort()
            doc.Range(doc.Content.End-1,doc.Content.End-1).InsertBreak()
            doc.Range(doc.Content.End-1,doc.Content.End-1).Text = "Участники:\r"
            doc.Sections(doc.Sections.Count).Headers(1).Range.Text="Участники\r"
            doc.Range(doc.Content.End-12,doc.Content.End-1).Font.Size = 16
            doc.Range(doc.Content.End-12,doc.Content.End-1).Font.Bold = 1
            doc.Range(doc.Content.End-12).ParagraphFormat.Alignment=1
            for i in pages_keys:
                doc.Range(doc.Content.End-1,doc.Content.End-1).Text = f'{i}:\t{pages[i]}\r'
        doc.Save()
        save_type = 17 if self.res_path.endswith(".pdf") else 16
        doc.SaveAs2(self.res_path,save_type)
        doc.Close()
        self.end()
        self.on_end()

        
    def _start_mode2(self):
        name = os.path.join(self.data_path,"temp.docx").replace("/","\\")
        doc=self.word.Documents(name)
        doc.Range(doc.Content.End-1,doc.Content.End-1).InsertBreak()
        doc.Sections(1).Footers(1).PageNumbers.Add(1,True)
        doc.Sections.Add(doc.Range(doc.Content.End-1,doc.Content.End-1),0)
        doc.Sections(doc.Sections.Count).Headers(1).LinkToPrevious=False
        path, files = getfiles(self.data_path)
        count = sum([len(files[i]) for i in files])
        now_count = 0
        for folder in path:
            doc.Sections(doc.Sections.Count).Headers(1).Range.Text=folder
            doc.Sections(doc.Sections.Count).Headers(1).Range.ParagraphFormat.Alignment=2
            doc.Range(doc.Content.End-1,doc.Content.End-1).Text=folder+"\r"
            doc.Range(doc.Content.End-len(folder)-2,doc.Content.End-1).Font.Size = 16
            doc.Range(doc.Content.End-len(folder)-2,doc.Content.End-1).Font.Bold = 1
            doc.Range(doc.Content.End-len(folder)-2).ParagraphFormat.Alignment=1
            for file in files[folder]:
                if(self.stopped):
                    self.on_cancel()
                    return self.end()
                now_count+=1
                self.func(now_count,count)
                self.open(file)
                self.copy(file,name)
                self.close(file)
            doc.Sections.Add(doc.Range(doc.Content.End-1,doc.Content.End-1),0)
            doc.Sections(doc.Sections.Count).Headers(1).LinkToPrevious=False
        doc.Save()
        save_type = 17 if self.res_path.endswith(".pdf") else 16
        doc.SaveAs2(self.res_path,save_type)
        doc.Close()
        self.end()
        self.on_end()

    def stop(self):
        self.stopped = True