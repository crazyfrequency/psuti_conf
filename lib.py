from os import close, name, path, walk, getcwd
import win32com.client as cl

class MSWord:
    """Класс microsoft word"""
    def __init__(self, visible:bool=False) -> None:
        self.word = cl.DispatchEx("Word.Application")
        self.word.Visible = visible
        self.word.ScreenUpdating = visible
        self.const = cl.constants
    
    def open(self, file:str=None)->None|Exception:
        """Открытие файла word"""
        if file==None:return 'Файл не указан'
        try:
            self.word.Documents.Open(file)
        except Exception as err:
            return err

    def close(self, file:str=None)->None|Exception:
        """Закрытие файла word"""
        if file==None:return 'Файл не указан'
        try:
            self.word.Documents(file).Close()
        except Exception as err:
            return err

    def getnames(self,file:str=None):
        if file==None:return [1,'Файл не указан']
        res=[[],[]]
        format=0
        if self.word.Documents(file).Paragraphs.Count>2:
            for i in self.word.Documents(file).Paragraphs:
                if i.Range.Alignment==2 and i.Range.Italic==-1:
                    words=[h for h in i.replace('\t',' ').replace('\r','').replace('.','. ').replace(',',', ').split(' ') if h]
                    for i in range(len(words)-2):
                        if len(words[i+2])<3 and len(words[i+1])<3 and len(words[i])>2:
                            if words[i][0].isupper() and words[i+1][0].isupper() and words[i+2][0].isupper():
                                res[0].append(words)
                else:break

    def set_def_format(self,file:str=None):
        if file==None:return [1,'Файл не указан']
        doc = self.word.Documents(file)
        doc.Range().Font.Name = 'Times New Roman'
        for i in range(doc.OMaths.Count):
            try:
                doc.OMaths.Item(i+1).ConvertToMathText()
            except:None
        doc.Range().Font.Size = 10
        doc.Paragraphs.FirstLineIndent = 28.350000381469727
        doc.PageSetup.LeftMargin = 56.70000076293945
        doc.PageSetup.RightMargin = 56.70000076293945
        doc.PageSetup.TopMargin = 56.70000076293945
        doc.PageSetup.BottomMargin = 56.70000076293945
        doc.Paragraphs.LineSpacingRule = 0

    def new(self,file:str=None):
        if file==None:return [1,'Файл не указан']
        doc = self.word.Documents.Add()
        doc.SaveAs2(file)
        return file

    def copy(self,file1:str=None,file2:str=None):
        if (not file1)or(not file2):return [1,'Файл не указан']
        doc1=self.word.Documents(file1)
        doc2=self.word.Documents(file2)
        doc1.Range().Copy()
        doc2.Range(doc2.Content.End-1,doc2.Content.End-1).Paste()
        last_math_or_table = max([
                doc2.OMaths(doc2.OMaths.Count).Range.End if doc2.OMaths.Count!=0 else 0,
                doc2.Tables(doc2.Tables.Count).Range.End if doc2.Tables.Count!=0 else 0
            ])

        doc2.Range(doc2.Content.End-1,doc2.Content.End-1).Text
        
        try:
            last_end = doc2.Content.End+1
            while doc2.Content.End>last_math_or_table:
                if doc2.Range(doc2.Content.End-2,doc2.Content.End-1).Text not in ["\r"," ","\t","\x0b"] or doc2.Content.End==last_end:
                    break
                doc2.Range(doc2.Content.End-2,doc2.Content.End-1).Text=""
                last_end = doc2.Content.End
        except:print("bug")

        start_page = doc2.ActiveWindow.Panes(1).Pages.Count
        for i in range(9):
            doc2.Range(doc2.Content.End-1,doc2.Content.End-1).Text = "\r"
            if doc2.ActiveWindow.Panes(1).Pages.Count!=start_page:
                break
        if(i==8):
            doc2.Range(doc2.Content.End-6,doc2.Content.End-1).Text = ""

    def end(self):
        self.word.Visible = True
        self.word.ScreenUpdating = True
        self.word.Quit(True)
        
def getallfiles(pathtofiles:str='.')->list:
    """
    Возвращает список содержащий пути ко
    всем файлам word в папке\n
    files=getfiles('.')
    """
    lib=path.abspath(pathtofiles)
    foundedfiles=[]
    for root, dirs, files in walk(lib):
        for file in files:
            if (file[-5:]=='.docx' or file[-4:]=='.doc')and file[0:2]!='~$':
                foundedfiles.append(root+'\\'+file)
    return foundedfiles

def getfiles(pathtofiles:str='.')->list:
    """
    Возвращает список содержащий папки(ключи)
    и словарь содержащий пути к файлам word\n
    path,files=getfiles('.')
    """
    lib=path.abspath(pathtofiles)
    resfolders=[]
    resfiles={}
    for root,dirs,files in walk(lib):
        for file in files:
            if (file[-5:]=='.docx' or file[-4:]=='.doc')and file[0:2]!='~$':
                name_of_file=root+'\\'+file[0:file.rfind("_")]+file[file.rfind("."):]
                
                if root[root.rfind('\\')+1:] in resfolders:
                    if(not name_of_file in resfiles[root[root.rfind('\\')+1:]]):
                        resfiles[root[root.rfind('\\')+1:]].append(root+'\\'+file)
                    else:
                        resfiles[root[root.rfind('\\')+1:]][resfiles[root[root.rfind('\\')+1:]].index(name_of_file)]=root+'\\'+file
                else:
                    resfolders.append(root[root.rfind('\\')+1:])
                    resfiles[root[root.rfind('\\')+1:]]=[root+'\\'+file]
    now_path = pathtofiles
    if pathtofiles.endswith("\\"):
        now_path = pathtofiles[0:-1]
    if now_path in resfolders:
        resfolders.remove(now_path)
    if now_path in resfiles:
        del resfiles[now_path]
    return resfolders,resfiles