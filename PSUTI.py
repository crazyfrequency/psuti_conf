from dataclasses import dataclass
import requests, re, os
from lxml import html
from bs4 import BeautifulSoup


def check_url(url:str="")->bool:
    if(len(url)<5):return False
    if(not re.search(r'(https:\/\/|http:\/\/)conf.(psuti|psati).ru\/',url)):
        return False
    return True

@dataclass
class Sections():
    title:str = None
    url:str   = None
    count:int = None

@dataclass
class Participants():
    title:str = None
    url:str   = None

@dataclass
class Participant():
    surname:str          = None
    name:str             = None
    patronymic:str       = None
    degree:str           = None
    post:str             = None
    country:str          = None
    city:str             = None
    number:str           = None
    email:str            = None
    data:list[list[str]] = None

@dataclass
class Document():
    title:str = None
    url:str   = None
    size:float  = 0

@dataclass
class Thesis():
    participants:list[Participant] = None
    title:str                     = None
    section_code:int              = 0
    document:Document             = None



class PSUTI:
    def __init__(self, url:str) -> None:
        self.url = url[0:-1] if url[-1]=="/" else url
        self.email = ""
        self._password = ""
        self.session = requests.Session()

    def login(self, email:str=None, password: str=None) -> bool:
        """
        Функция для входа в акаунт администратора.
        Возвращает True, если удалось.
        """
        try:
            if(email!=None):self.email = email
            if(password!=None):self._password = password
            self.session.get("https://conf.psuti.ru/login")
            r = self.session.post(
                "https://conf.psuti.ru/login",
                data=f"YII_CSRF_TOKEN={self.session.cookies['YII_CSRF_TOKEN']}&"+
                f"LoginForm[email]={self.email}&LoginForm[password]={self._password}",
                headers={"Content-Type":"application/x-www-form-urlencoded"})
            if(r.status_code==200 and len(r.history)!=0):
                if(r.history[0].status_code==302):
                    return True
            return False
        except:
            return False
    
    def get_all_sections(self) -> list[Sections]|None:
        try:
            sections=[]
            r = self.session.get(self.url+"/participants")
            tree = BeautifulSoup(r.text, features="lxml")
            tree = tree.find("table",{"class":"table"})
            sections_tree = tree.find_all("tr",{"class":"ordered"})[0:-2]
            for i in sections_tree:
                t = i.find("a")
                title=t.decode_contents()
                count = int(i.find_all("td",{"class":"center"})[1].decode_contents())
                if(count):
                    sections.append(Sections(
                        title,
                        t.attrs["href"],
                        count
                    ))
            return sections
        except:
            return None
    
    def get_all_participants(self, conf_url:str) -> list[Participants]|None:
        try:
            participants = []
            r = self.session.get("https://conf.psuti.ru"+conf_url)
            tree = BeautifulSoup(r.text, features="lxml")
            tree = tree.find("table",{"class":"table"})
            participants_tree = tree.find_all("tr",{"class":"ordered"})
            for i in participants_tree:
                t = i.find_all("td")
                if(t[3].find("img").attrs["alt"] == "принят"):
                    a = t[1].find_all("a")[1]
                    participants.append(Participants(
                        a.decode_contents(),
                        a.attrs["href"]
                    ))
            return participants
        except:
            return None

    def get_participant(self, url:str) -> Thesis|None:
       
            r = self.session.get("https://conf.psuti.ru"+url)
            tree = BeautifulSoup(r.text, features="lxml")
            tree = tree.find("div",{"class":"conf-column"})
            title = tree.find("h2").decode_contents()
            tree = tree.find("div",{"class":"form participant-view"})
            code = int(tree.find("select",{"id":"Participant_topic_id"}).find("option",{"selected":"selected"}).attrs["value"])
            list_of_participant = tree.find_all("a",{"href":"#"},text="подробнее")
            index_list = [tree.index(e) for e in list_of_participant]
            children=[e for e in tree.children]
            names = [children[i-1].split() for i in index_list]
            data = tree.find_all("div",{"class":"author-info"})
            authors = []
            for i in range(len(data)):
                sub_data = [x.replace("    ","").replace("\t","").replace("\n","").split(": ") for x in data[i].decode_contents().split("<br/>")][0:-1]
                new_data = {}
                for key, value in sub_data:
                    new_data[key] = value
                authors.append(Participant(
                    *names[i][-3:],
                    degree= new_data["Ученая степень"] if "Ученая степень" in new_data else None,
                    post= new_data["Должность"] if "Должность" in new_data else None,
                    country= new_data["Страна"],
                    city= new_data["Город"],
                    number= new_data["Телефон"],
                    email=None,
                    data=new_data
                ))
            
            doc = tree.find_all("table")[-1]
            doc = doc.find("a")
            return Thesis(
                title=title,
                section_code=code,
                participants=authors,
                document=Document(doc.decode_contents(), doc.attrs["href"])
            )
    
    def download(self, url:str, path:str) -> None:
        res = self.session.get("https://conf.psuti.ru"+url)
        file = open(path, "bw")
        file.write(res.content)
        file.flush()
        file.close()
