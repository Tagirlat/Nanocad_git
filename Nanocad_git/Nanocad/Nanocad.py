import win32com.client
import json


class NanocadApp:
    def __init__(self):
        self.__nanocad_app = None
        self.__nanocad_doc = None

    def init_nanocad(self):
        self.__nanocad_app = win32com.client.Dispatch("nanoCAD.Application")

    def get_app_params(self):
        app_params = {
            "LocaleId": self.__nanocad_app.LocaleId,
            "Version": self.__nanocad_app.Version,
            "Caption": self.__nanocad_app.Caption,
            "Documents": [i.Name for i in self.__nanocad_app.Documents]
        }
        return app_params

    def doc(self, name: str = None, new: bool = False, default: bool = True):
        self.__nanocad_doc = self.__nanocad_app.ActiveDocument
        if default:
            with open('files\\setting.json', 'r', encoding='utf-8') as f:
                loaded_dict = json.load(f)
            self.__nanocad_doc = self.__nanocad_doc.Open(loaded_dict["Файл-шаблон"])
        else:
            if new:
                self.__nanocad_doc = self.__nanocad_doc.New(1)
            else:
                self.__nanocad_doc = self.__nanocad_doc.Open(name)
        doc = NanocadDoc(self.__nanocad_doc)
        return doc


class NanocadDoc:
    def __init__(self, doc):
        self.__doc = doc
        self.__layouts = {}
        self.__objects = []
        pass

    def get_layouts(self):
        for i, layout in enumerate(self.__doc.Layouts):
            self.__layouts[f"{layout.Name}"] = layout
        # print(self.__layouts)
        return self.__layouts.keys()

    def add_m_text(self, layout: str, text: str, coordinates: list):
        mtext_obj = self.__layouts[f"{layout}"].Block.AddMText(coordinates, 10, text)
        mtext_obj.Height = 2
        # mtext_obj.Document.HorizontalAlignment = "acHorizontalAlignmentLeft"
        self.__objects.append(mtext_obj)

    def add_text(self, layout: str, text: str, coordinates: list):
        text_obj = self.__layouts[f"{layout}"].Block.AddText(text, coordinates, 2)
        text_obj.Height = 3
        text_obj.TextGenerationFlag = 0
        # text_obj.TextAlignmentPoint = "acAlignmentTopLeft"
        self.__objects.append(text_obj)

    def replace_text(self, old_text: str, new_text: str, layout: str = None, all_doc: bool = False):
        if all_doc:
            for slide in self.__layouts.values():
                for entity in slide.Block:
                    if entity.ObjectName == "AcDbText" or entity.ObjectName == "AcDbMText":
                        text_obj = win32com.client.CastTo(entity, "IAcadText")
                        if old_text in text_obj.TextString:
                            text_obj.TextString = text_obj.TextString.replace(old_text, new_text)
        else:
            for entity in self.__layouts[f"{layout}"].Block:
                if entity.ObjectName == "AcDbText" or entity.ObjectName == "AcDbMText":
                    text_obj = win32com.client.CastTo(entity, "IAcadText")
                    if old_text in text_obj.TextString:
                        text_obj.TextString = text_obj.TextString.replace(old_text, new_text)