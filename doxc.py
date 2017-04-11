# coding: utf8
#есть смысл переделать под пи3
from docx import Document
document = Document('1.docx')
fileProblems = open('bad.txt','w')
filedoctype = open('doctype.txt','w')
fileautors = open('autors.txt','w')
fileautorsUkazatel = open('autorsUkazatel.txt','w')
filelastautorsUkazatel = open('lastautorsUkazatel.txt','w')
intNumOfZapis = 0
intColOfDonFindZapis = 0
testingString = ""
dataOfAllFindUniqKeysOfAll = {}
listOfKeyWords = []
listOfAuthors = []
nonSortedSlovarOfAUtors = {}
flagIsAllZapisRegonuzeBuType = False
flagIsAllZapisRegonuzeBuAutorsField = False
flagIsPresenBetterKeyForRedactors = False


class Zapis:    
    """обьект записи """
    def CheckRazdeliteli():
        """описывает правила для разделителей"""
        pass

    def CheckTypeOfDoc():
        """описывает типы  док-та"""
        pass


    docAuthorRazdelitel = {
        "razdelitel": ", ",
        "prinstavka": "",
        "prinstavkaAvtoraIRedactora": "авт. "
    }

    """словарь разделителей для поля авторов и типоов авторов (автор, ред, науч сотр. и т.д.)"""
    docRedactorRazdelitel = {

        "sostavitelBiblio":{
            "razdelitel": " ; сост. библиогр. указ. ",
            "prinstavka": "сост."
        },
        "sostavitel": {
            "razdelitel": "сост.: " ,
            "prinstavka": "сост."
        },
        "sostavitelBegin": {
            "razdelitel": "сост. ",
            "prinstavka": "сост."
        },
        "sostavitelUkr": {
            "razdelitel": "уклад.: ",
            "prinstavka": "уклад."
        },
        "sostavitelUkr2": {
            "razdelitel": "уклад.: ",
            "prinstavka": "уклад."
        },
        "uporiadUkr": {
            "razdelitel": "упоряд.: ",
            "prinstavka": "упоряд."
        },

        "redactor": {
            "razdelitel": " ; под ред. ",
            "prinstavka": "ред."
        },
        "redactor2": {
            "razdelitel": " ; под ред. ",
            "prinstavka": "ред."
        },
        "redactor3": {
            "razdelitel": " ; ред.: ",
            "prinstavka": "ред."
        },
        "redactor4": {
            "razdelitel": "под ред. ",
            "prinstavka": "ред."
        },
        "redactorUkr": {
            "razdelitel": " ; під ред. ",
            "prinstavka": "ред."
        },
        "redactorUkrInBegin": {
            "razdelitel": "під ред. ",
            "prinstavka": "ред."
        },
        "redactorUkrInBegin2": {
            "razdelitel": "за ред. ",
            "prinstavka": "ред."
        },

        "totalredactor": {
            "razdelitel": " ; под общ. ред. ",
            "prinstavka": "ред."
        },
        "pidzagallredactor": {
            "razdelitel": "під заг. ред. ",
            "prinstavka": "ред."
        },
        "pidzagallredactorBegin": {
            "razdelitel": "за заг. ред. ",
            "prinstavka": "ред."
        },
        "zagallredactor": {
            "razdelitel": " ; за заг. ред.",
            "prinstavka": "ред."
        },

        "nauchKonsult": {
            "razdelitel": " ; науч. консультант ",
            "prinstavka": "науч. консультант"
        },
        "naukovKonsult": {
            "razdelitel": " ; наук. консультант ",
            "prinstavka": "наук. консультант"
        },
        "nauchRuk": {
            "razdelitel": " ; науч. рук. ",
            "prinstavka": "науч. рук."
        },
        "nauchRukUkr": {
            "razdelitel": " ; наук. кер. ",
            "prinstavka": "наук.кер."
        },

        "editedInBegin": {
            "razdelitel": ", ed. ",
            "prinstavka": "ed."
        },

        "editedsInBegin": {
            "razdelitel": "eds.: ",
            "prinstavka": "ed."
        },
        "editedsInBegin2": {
            "razdelitel": "ed. ",
            "prinstavka": "ed."
        },
        "edited": {
            "razdelitel": " ; ed. by",
            "prinstavka": "ed."
        },
        "edited3": {
            "razdelitel": " ; ed. by ",
            "prinstavka": "ed."
        },
        "edited3": {
            "razdelitel": "ed. by ",
            "prinstavka": "ed."
        },
        "edited2": {
            "razdelitel": " ; ed. ",
            "prinstavka": "ed."
        },
        "vstupUkr": {
            "razdelitel": " ; вступ. ст. ",
            "prinstavka": "вступ. ст."
        },


    }
    docFieldRazdelitel={
        # -' / '  -- ' ; '
                "patent":{
                    "docIndex": {
                        "begin": ("Пат."),
                        "end": ("Україна,")
                    },
                    "codes": {
                        "begin": ("Україна,"),
                        "end": (". ")
                    },
                    "name":{
                        "begin": (". "),
                        "end": (" / ")
                    },
                    "author":{
                        "begin":(" / "),
                        "end":(" ; ")
                    },
                    "sobstvennik":{
                        "begin": (" ; "),
                        "end": (" – ")
                    },
                    "regNumber": {
                        "begin": (" – "),
                        "end": (" ; ")
                    },
                    "dateOfPublish":{
                        "begin": (" ; заявл. "),
                        "end": (" ; ")
                    },
                    "dateOfZayavki": {
                        "begin": (" ; опубл. "),
                        "end": (", ")
                    },
                    "numOfBul": {
                        "begin": (", "),
                        "end": (".")
                    }
                },
        # -' / '  -- '. – '
                "dissertation":{
                    "name": {
                        "end": (" : дис. ...")
                    },
                    "type": {
                        "begin": (", "),
                        "end": ("техн. наук : ")
                    },
                    "shifrSpecialnosti": {
                        "begin": (" : "),
                        "end": (" – ")
                    },
                    "nazvanieSpecialnosti": {
                        "begin": (" – "),
                        "end": (" / ")
                    },
                    "author": {
                        "begin": (" / "),
                        "end": (". – ")
                    },
                    "nauchniy": {
                        "begin": ("; науч."),
                        "end": (". – ")
                    },
                    "placeYear": {
                        "begin": ("; науч."),
                        "end": (". – ")
                    },
                    "pages": {
                        "begin": (". –"),
                        "end": ("с.", "p.")
                    },
                },
        # -' / '  -- '. – '
                "posobie":{
                    "author": {
                         "begin": (" / "),
                        "end": (". – ")
                    },
                    "remark": {},
                },
        # -' / '  -- '. – '
                "conference": {
                    "author": {
                        "begin": (" / "),
                        "end": (" // ")
                    },
                },
        # -' / '  -- '. – '
                "monografy":{
                    "author": {
                        "begin": (" / "),
                        "end": (". – "),
                        "contentRedactor": ("– Content of:", "– Из содерж.:", "– Зі змісту:")
                    },
                    "remark": {},
                },
        # -' / '  -- '. – '
                "textbook":{
                    "author": {
                        "begin": (" / "),
                        "end": (". – ")
                    },
                },
        # -' / '  -- ' // '
                "journal":{
                    "author": {
                        "begin": (" / "),
                        "end": (" // ")
                    },
                },
        # -' / ' -- '. – '       или ' // '
                "spravochniki":{
                    "author": {
                        "begin": (" / "),
                        "end": (". – "),
                        "iskluchenie": (" // ")
                    },
                }
    }
    """словарь разделителей областей встречающихся в разных типах док-тов(для определения полей док-та)"""

    """словарь разделителей+словосочетаний встречающихся в разных типах док-тов (для определения типа док-та)"""
    docTypeRazdelitelOLD={

             "patent": ("Пат. ", ", Бюл. №"),

             "dissertation": (": дис. ... д-ра техн. наук", ": дис. ... канд. техн. наук"),

             "posobie": (": учеб. пособие", ": практикум", ": навч. посіб.", ": метод. рек.",
                         ": лекц. материал", ": сб. задач",": сб. лаб. работ", ": reader", ": training guide",
                         ": [учеб. пособие]", ": [лекц. материал]", ": tutorial", ": coursebook.",
                         ": зб. практ. робіт", ": зб. задач", " : посіб. для соц. працівників", ": lectures",
                         ": метод. реком", ": навч. наоч. посіб.", ": teacher’s aid", ": teaching-aid book",
                         ": lecture course man.", ": конспект лекций", ": конспект лекцій", ": курс лекций",
                         ": рабочая тетр.", ": практ. посіб.", ": практ. рук-во к лекц. курсу ",
                         ": зб. вправ", ": cб. заданий для", ": практ. руководство для",
                         ": опор. конспект лекцій", ": консп. лекций",
                         ": рук. к решению задач", ": метод. вказівки до", ": man. to lab. works",
                         ": нав. посіб.", ": навч.-метод.", ": контрол.-тест. завдання",
                         ": guidance man. for", ": [навч. посіб.]", ": лаб. практикум",
                         ": зб. текстів і задач", ": [посібник]", ": manual",
                         ": synopsis", ": study guide for", ": educational supply", ": зб. техн. текстів",
                         ": сб. практ. заданий", ": practical work", ": lab. classes tutorial",
                         ": метод. указания", ": метод. указ.", ": учеб.-метод. пособие",
                         ": cб. практ. заданий", ": lecture synopsis", ": зб. завдань до самост.",
                         ": summary lectures", ": the laboratory works man.", "Сборник задач", ": course-book ",
                         ": workbook for lab.", ": [workbook for lab. course]", "Laboratory experiments" ,
                         ": teхtbook" #здесь костыль - в слове кирилическая 'х'
                          ),

             "monografy": (": монографія", ": монография", ": очерк", ": колект. моногр.", ": monogr. of",
                           ": [монография]", ": [монографія]", ": [monograph]",  ": воспоминания, библиогр. указ.",),

             "textbook": (": учебник", ": підручник", ": учеб. для вузов", ": textbook.", ": підруч.",
                          ": [підручник]", ": [учебник]"),

             "journal": (" журнал.", " журнал ", "// Журнал", "Вестник", "Вісті", ", №", "– №", "Спец. вип.",
                         " – Спец. вып" , ": альманах", "– Vol.", "Темат. вип. №", ", Iss.",
                         ": digest", "– Спецвипуск", [ " [Electronic resource]", "May. –" ] , [ " [Electronic resource]", "Sept. –" ],
                          [ " [Electronic resource]", "March. –" ]),

             "spravochniki": (": справ. пособие", ": биобиблиогр. указ.", ": библиогр. указ.",
                              ": рук. пользователя", ": англ.-рус.", ": энциклопед. изд.",
                              "// Енциклопедія сучасної України", ": еnglish-russian-ukrainian lexicon",
                              ": биобиблиогр. сб.", ": рос.-укр. навч. енциклопедія", ": [довідник]",
                              ": биобиблиогр. очерк", "Довідкові матеріали")
            }
    """словарь разделителей+словосочетаний встречающихся в разных типах док-тов (для определения типа док-та)"""
    docTypeRazdelitel={
             "conference": ("– Вып.", "– Вип.", ", Вип.", ", Вып.", "– Спец. Вип.", "‒ Вип.", " ‒ Вып.", "Кн.", "Т.", "[Нац. аэрокосм.", "[Nat. aerospace",
                            "[Харківська медична" , "[Нац. аерокосм.", "[Нац. техн.", "[сб. науч. тр.]", "Назва з екрану", "Нац. косм.", "[Запорож. нац.",
                            "[Ін-т наук.", "[S. l.]", "сonf.", "IEEE", "[Харків. нац.", "[Інститут", "Ін-т", "нац.", "[Нац.", "конф", "proc",
                            "рroc", "Proc", "Conf.", "НЮАУ", "сб. работ", "– Вып", "сост.:", "Электронный ресурс", "зб. тез", "[Seoul, Korea]",
                            "зав. каф. філософії", "[S.l.]", "– Спец. вип.", ", вип.", "зб. наук.")

            }

print ''
print '-----------------------------------------------------------------------------------------------------------'

# вывод в файл разделителей типов док-та в читаемом виде + паралельное формир. списка разделителей+словосочетаний
fileProblems.write('\n'+'ниже перечисленны словосочитания , по которым программа пытается определить тип док-та в записи :' + '\n' + '\n')
for docType in Zapis.docTypeRazdelitel.keys():
    for docNames in Zapis.docTypeRazdelitel[docType]:
        fileProblems.write(docType+" - \""+str(docNames)+"\""+'\n')
        listOfKeyWords.append(docNames)

fileProblems.write('------------------------------------------------------------------------'+'\n' + '\n' + '\n')

#проверка уникальности ключа Не вхождкнии строки ключа в строку другого ключа ключи-списки не проверяем на вхождение
for i in listOfKeyWords :
    if listOfKeyWords.count(i) > 1 :
        fileProblems.write("------данный ключ повторяется"+'\n')
        fileProblems.write(i+'\n')
    if isinstance(i, str):
        for j in listOfKeyWords:
            if isinstance(j, str):
                if i !=j and j.find(i) == 0:
                    fileProblems.write("------данный ключ являтся тоставной частью другого ключа "+'\n')
                    fileProblems.write(i+'<---->'+j+'\n')
                
indexOfZapBegin = 3284
intNumOfZapis = indexOfZapBegin
for p in document.paragraphs:
    #в записи ВСЕГДА должен присутствовать слэш
    # помни ! Unicode! https://habrahabr.ru/post/135913/
    if p.text.encode('utf-8').count('/') > 0:
        
        intNumOfZapis = intNumOfZapis +1
        #инициальзация главного словаря первого уровня где ключ-номер найденной записи 
        dataOfAllFindUniqKeysOfAll[intNumOfZapis]={}
        testingString = p.text.encode('utf-8')
        # убрать эту хуйню при правильном форматировании
        testingString = testingString.replace('\n', " ")
        
        if dataOfAllFindUniqKeysOfAll[intNumOfZapis].has_key("Text") == False:
            dataOfAllFindUniqKeysOfAll[intNumOfZapis]["Text"] = ""
            dataOfAllFindUniqKeysOfAll[intNumOfZapis]["Text"] = testingString
        
        testingString = testingString.replace('  ', " ")
        testingString = testingString.replace('  ', " ")
        testingString = testingString.replace(' ', ' ')



        #перебор ро типам
        for docType in Zapis.docTypeRazdelitel.keys():
            #перебор по уникальным разделителям
            for typeRazdelitel in Zapis.docTypeRazdelitel[docType]:
                #   добавить определнние типа Zapis.docTypeRazdelitel[docType] для перебора либо ключей-строк
                #   либо ключей-списков(для случаев определения через обязательное вхождение всех эл.
                #     списка-ключа)

                if isinstance(typeRazdelitel, list) :
                    FlagIsAllIclude = True ;
                    for i in range(0,len(typeRazdelitel)):
                        if testingString.find(typeRazdelitel[i].replace(' ',' ')) < 0:
                            FlagIsAllIclude = False
                    if FlagIsAllIclude == True :
                        #print str(intNumOfZapis) +'<-- в этой записи сработал список-ключ'
                        if dataOfAllFindUniqKeysOfAll[intNumOfZapis].has_key("DocType") == False:
                            dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"] = {}
                            dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"][docType] = []
                        if dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"].has_key(docType) == False:
                            dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"][docType] = []
                        dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"][docType].append(typeRazdelitel)
                else:
                    if testingString.find(typeRazdelitel.replace(' ',' ')) >=0:
                        if dataOfAllFindUniqKeysOfAll[intNumOfZapis].has_key("DocType") == False:
                            dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"] = {}
                            #инициальзация списка сработавших разделителей по типу док-та(может сработать несколько разделителей
                            #описанных для определенного типа док-та )
                            dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"][docType] = []
                        #если возм. опред. док-т, но найден разделитель для другого типа док-та - создаем новый словарь с найденным
                        #типом док. и списком разделителей
                        if dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"].has_key(docType) == False:
                            dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"][docType] = []
                        # и только теперь дописыпаем найденный разд. в словарь типов. док. в словаре "DocType" в словаре
                        #со значением intNumOfZapis(массивом записей )
                        dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"][docType].append(typeRazdelitel)
        
        #print str(intNumOfZapis)+'. '+p.text.encode('utf-8')
        
        if dataOfAllFindUniqKeysOfAll[intNumOfZapis].has_key("DocType") == False:
            print str(intNumOfZapis)+". ----------не определена по типу--------"
            fileProblems.write(str(intNumOfZapis) + '. ' + ". ----------не определена по типу--------" + '\n')
            fileProblems.write(str(intNumOfZapis)+'. '+p.text.encode('utf-8')+'\n')
            intColOfDonFindZapis = intColOfDonFindZapis + 1
        else:
            #блок корректировки определения типов док-та при наличии разделителей разных видов док-та
            if dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"].has_key("patent") == True and dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"].has_key("journal") == True:
                del dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"]["journal"]
            
            for typeindex in dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"].keys():
                
                if len(dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"].keys()) > 1:
                    #print "found two or more razdelitel for different types"
                    fileProblems.write(str(intNumOfZapis)+'. ----------найдено несколько разделителей описывающих разные типы док-та----------'+'\n')
                    fileProblems.write(str(intNumOfZapis)+'. '+p.text.encode('utf-8')+'\n')
                else:    
                    for razdelindex in dataOfAllFindUniqKeysOfAll[intNumOfZapis]["DocType"][typeindex]:
                        filedoctype.write(str(intNumOfZapis)+". - "+str(typeindex)+" key is - "+str(razdelindex)+'\n')
                        #print str(intNumOfZapis)+". - "+str(typeindex)+" key is - "+str(razdelindex)
                        #pass
    else:
        if len(p.text) >0:
            #вывод НЕ записей (для проверки)
            
            #print intNumOfZapis
            #print p.text
            
            pass



print '>>>>> найденно '+str(intNumOfZapis-indexOfZapBegin)+'  записей <<<<<<'
fileProblems.write('>>>>>> найденно '+str(intNumOfZapis-indexOfZapBegin)+'  записей <<<<<<' + '\n' + '\n')
if intColOfDonFindZapis > 0:
    print '>>>>> НЕ определенно по типу '+str(intColOfDonFindZapis)+'  записей <<<<<'
    fileProblems.write('>>>>>> НЕ определенно по типу '+str(intColOfDonFindZapis)+'  записей <<<<<'+'\n' + '\n')
else:
    fileProblems.write('>>>>>> все записи были определены по типу док-та <<<<<' + '\n' + '\n')
    flagIsAllZapisRegonuzeBuType = True

if  flagIsAllZapisRegonuzeBuType == True :
    flagIsAllZapisRegonuzeBuAutorsField = True
    #обработка массива записей с поиском авторов по типу док-та
    for numOfZapp in range(1,len(dataOfAllFindUniqKeysOfAll) + 1) :
        numOfZap=numOfZapp + indexOfZapBegin
        docType = str(dataOfAllFindUniqKeysOfAll[numOfZap]["DocType"].keys()[0])
        if Zapis.docFieldRazdelitel[docType]['author'].has_key('begin') == True :
            #print  Zapis.docFieldRazdelitel[docType]['author']['begin']
            beginPlace=dataOfAllFindUniqKeysOfAll[numOfZap]["Text"].find(Zapis.docFieldRazdelitel[docType]['author']['begin'])
        if Zapis.docFieldRazdelitel[docType]['author'].has_key('end') == True :
            #print  Zapis.docFieldRazdelitel[docType]['author']['end']
            endPlace = dataOfAllFindUniqKeysOfAll[numOfZap]["Text"].find(Zapis.docFieldRazdelitel[docType]['author']['end'])
        if Zapis.docFieldRazdelitel[docType]['author'].has_key('iskluchenie') == True:
            #print  Zapis.docFieldRazdelitel[docType]['author']['iskluchenie']
            isklucheniePlace = dataOfAllFindUniqKeysOfAll[numOfZap]["Text"].find(Zapis.docFieldRazdelitel[docType]['author']['iskluchenie'])
            if isklucheniePlace > 0 and isklucheniePlace < endPlace:
                endPlace = isklucheniePlace


        if beginPlace < 0 or endPlace < 0 :
            #print str(numOfZap) + '. ' + dataOfAllFindUniqKeysOfAll[numOfZap]["Text"]
            if beginPlace < 0 :
                #print "не могу найти начальный разделитель области авторов"
                fileProblems.write( str(numOfZap) + '. ' + dataOfAllFindUniqKeysOfAll[numOfZap]["Text"]+'\n')
                fileProblems.write("----------не могу найти начальный разделитель области авторов----------" + '\n' + '\n')
            if endPlace < 0 :
                #print "не могу найти конечный разделитель области авторов"
                fileProblems.write(str(numOfZap) + '. ' + dataOfAllFindUniqKeysOfAll[numOfZap]["Text"] + '\n')
                fileProblems.write("----------не могу найти конечный разделитель области авторов----------" + '\n' + '\n')
        if dataOfAllFindUniqKeysOfAll[numOfZap]["Text"][beginPlace+4:endPlace].find('–') > 0 :
            fileProblems.write(str(numOfZap) + '. ' + dataOfAllFindUniqKeysOfAll[numOfZap]["Text"] + '\n')
            fileProblems.write("----------не корректный конечный разделитель области авторов----------" + '\n' + '\n')


        if beginPlace > 0 and endPlace > beginPlace :

            authors = ""
            redactors = ""
            #цикл обработки записи с ссылкой на другие источники(в источниках не может быть редакторов)
            if Zapis.docFieldRazdelitel[docType]['author'].has_key('contentRedactor') == True:
                for i in Zapis.docFieldRazdelitel[docType]['author']['contentRedactor'] :
                    if dataOfAllFindUniqKeysOfAll[numOfZap]["Text"].find(i) > 0:
                        strOfGlavs = dataOfAllFindUniqKeysOfAll[numOfZap]["Text"][dataOfAllFindUniqKeysOfAll[numOfZap]["Text"].find(i):len(dataOfAllFindUniqKeysOfAll[numOfZap]["Text"])]
                        while strOfGlavs.count(Zapis.docFieldRazdelitel[docType]['author']['begin']) > 0:
                            authors = authors + strOfGlavs[strOfGlavs.find(Zapis.docFieldRazdelitel[docType]['author']['begin'])+4:strOfGlavs.find(Zapis.docFieldRazdelitel[docType]['author']['end'])]+', '
                            strOfGlavs = strOfGlavs[strOfGlavs.find(Zapis.docFieldRazdelitel[docType]['author']['end'])+4:len(strOfGlavs)]



            redactors =  dataOfAllFindUniqKeysOfAll[numOfZap]["Text"][beginPlace+4:endPlace]
            redactors = redactors.replace('[et al.]','')
            redactors = redactors.replace('[', '')
            redactors = redactors.replace(']', '')

            #print str(numOfZap)+ ' ----->' + authors
            fileautors.write( str(numOfZap))
            if len(authors) > 0 :
                fileautors.write( ' authors string----->' + authors )
            if len(redactors) > 0 :
                fileautors.write(' redact. string--->'+redactors+'\n')
            dataOfAllFindUniqKeysOfAll[numOfZap]["textAutors"] = ""
            dataOfAllFindUniqKeysOfAll[numOfZap]["textAutors"] = authors
            dataOfAllFindUniqKeysOfAll[numOfZap]["authors"] = {}
            dataOfAllFindUniqKeysOfAll[numOfZap]["textRedactors"] = ""
            dataOfAllFindUniqKeysOfAll[numOfZap]["textRedactors"] = redactors
            dataOfAllFindUniqKeysOfAll[numOfZap]["redactors"] = {}
            if numOfZap == 1169 :
                pass
            #    цикл разбора строки по редакторам СРОЧНО ПЕРЕДЕЛАТЬ

            # прогоняем строку редакторов на наличие включения строки-поля редакторов - если нет включений то все авторы
            nameOfRazdelitel = ''
            nameOfNextRazdelitel = ''
            nameOfRedactorField = ''
            for findingRazdelitelRedactors in Zapis.docRedactorRazdelitel.keys() :
                if Zapis.docRedactorRazdelitel[findingRazdelitelRedactors].has_key('razdelitel') == True :
                    # находим ближайшее включение
                    if redactors.find(Zapis.docRedactorRazdelitel[findingRazdelitelRedactors]['razdelitel']) >= 0:
                        # находим ближайшее включение
                        # выбираем более длинный ключ из списка ключей
                        if nameOfRazdelitel == '' or redactors.find(nameOfRazdelitel) > redactors.find(Zapis.docRedactorRazdelitel[findingRazdelitelRedactors]['razdelitel']) :
                            nameOfRazdelitel = Zapis.docRedactorRazdelitel[findingRazdelitelRedactors]['razdelitel']
                            nameOfRedactorField = findingRazdelitelRedactors
                        if  redactors.find(nameOfRazdelitel) == redactors.find(Zapis.docRedactorRazdelitel[findingRazdelitelRedactors]['razdelitel'])  and  len(nameOfRazdelitel) < len(Zapis.docRedactorRazdelitel[findingRazdelitelRedactors]['razdelitel']) :
                            nameOfRazdelitel = Zapis.docRedactorRazdelitel[findingRazdelitelRedactors]['razdelitel']
                            nameOfRedactorField = findingRazdelitelRedactors
            # если индекс ключа <> 0 то :  обрезаем строку от начала до индекса ключа - это авторы
            if len(nameOfRazdelitel) > 0 and redactors.find(nameOfRazdelitel) > 0 :
                authors = authors + Zapis.docAuthorRazdelitel['razdelitel'] + redactors[0:redactors.find(nameOfRazdelitel)]
                redactors = redactors[redactors.find(nameOfRazdelitel):len(redactors)]

            #строка начинается с поля описыв. редактров
            if len(nameOfRazdelitel) > 0 and redactors.find(nameOfRazdelitel) == 0:
                while len(redactors) > 2 :
                    #перебираем строку на возможные многоуровневые включения
                    for findingRazdelitelRedactors2 in Zapis.docRedactorRazdelitel.keys():
                        if Zapis.docRedactorRazdelitel[findingRazdelitelRedactors2].has_key('razdelitel') == True:
                            # находим ближайшее включение больше нуля и за пределами поля описыв. редакторов
                            if redactors.find(Zapis.docRedactorRazdelitel[findingRazdelitelRedactors2]['razdelitel']) > len(nameOfRazdelitel):
                                if nameOfNextRazdelitel == '' or redactors.find(Zapis.docRedactorRazdelitel[findingRazdelitelRedactors2]['razdelitel']) < redactors.find(nameOfNextRazdelitel):
                                    nameOfNextRazdelitel = Zapis.docRedactorRazdelitel[findingRazdelitelRedactors2]['razdelitel']
                                    nameOfRedactorField = findingRazdelitelRedactors2
                                if redactors.find(nameOfNextRazdelitel) == redactors.find(Zapis.docRedactorRazdelitel[findingRazdelitelRedactors2]['razdelitel']) and len(nameOfNextRazdelitel) < Zapis.docRedactorRazdelitel[findingRazdelitelRedactors2]['razdelitel'] :
                                    nameOfNextRazdelitel = Zapis.docRedactorRazdelitel[findingRazdelitelRedactors2]['razdelitel']
                                    nameOfRedactorField = findingRazdelitelRedactors2

                    dataOfAllFindUniqKeysOfAll[numOfZap]["authors"][nameOfRedactorField] = []
                    if nameOfNextRazdelitel == '' :
                        dataOfAllFindUniqKeysOfAll[numOfZap]["authors"][nameOfRedactorField] = redactors[len(nameOfRazdelitel):len(redactors)].split(Zapis.docAuthorRazdelitel['razdelitel'])
                        redactors = ''
                    else :
                        dataOfAllFindUniqKeysOfAll[numOfZap]["authors"][nameOfRedactorField] = redactors[len(nameOfRazdelitel):redactors.find(nameOfNextRazdelitel)].split(Zapis.docAuthorRazdelitel['razdelitel'])
                        redactors = redactors[redactors.find(nameOfNextRazdelitel):len(redactors)]


                    nameOfRazdelitel = nameOfNextRazdelitel
                    nameOfNextRazdelitel = ''

            #если в строке редакторов что то осталось - то это авторы проверка на размер для обрезки пробеллов и запятых
            # которые могли остаться в строке после поика редакторов
            if len(redactors) > 2 :
                authors = authors + Zapis.docAuthorRazdelitel['razdelitel']+redactors
            if len(authors) > 0 :
                dataOfAllFindUniqKeysOfAll[numOfZap]["authors"]["authors"] = []
                dataOfAllFindUniqKeysOfAll[numOfZap]["authors"]["authors"] = authors.split(Zapis.docAuthorRazdelitel['razdelitel'])

        else:
            flagIsAllZapisRegonuzeBuAutorsField = False
            fileautors.write(str(numOfZap) + '----------не могу распознать область авторов----------' + '\n')
            fileProblems.write(str(numOfZap) + '. ' + dataOfAllFindUniqKeysOfAll[numOfZap]["Text"] + '\n')
            fileProblems.write("----------не могу распознать область авторов----------" + '\n' + '\n')

for numOfZapp in range(1,len(dataOfAllFindUniqKeysOfAll)+1) :
    numOfZap = numOfZapp + indexOfZapBegin

    if dataOfAllFindUniqKeysOfAll[numOfZap].has_key('authors') == True :
        if len(dataOfAllFindUniqKeysOfAll[numOfZap]['authors']) >0 :
            for spiski in dataOfAllFindUniqKeysOfAll[numOfZap]['authors'].keys() :

                if isinstance(dataOfAllFindUniqKeysOfAll[numOfZap]['authors'][spiski], list):

                    for authorsNames in dataOfAllFindUniqKeysOfAll[numOfZap]['authors'][spiski] :

                        testingString = str(authorsNames)

                        testingString = testingString.replace('  ', ' ')
                        if testingString[0:1] == ' ':
                            testingString = testingString[1:len(testingString)]
                        testingString.replace(' ', ' ')
                        if testingString.rfind('. ') > 0:
                            testingString = testingString[
                                            testingString.rfind('. ') + 3:len(testingString)] + ' ' + testingString[0:testingString.rfind('. ') + 3]
                        if testingString.rfind('. ') > 0:
                            testingString = testingString[testingString.rfind('. ') + 2:len(testingString)] + ' ' + testingString[0:testingString.rfind('. ') + 1]
                       # filelastautorsUkazatel.write(testingString + '\n')

                        if nonSortedSlovarOfAUtors.has_key(authorsNames) ==  False :
                            nonSortedSlovarOfAUtors[authorsNames] = {}
                            nonSortedSlovarOfAUtors[authorsNames]['perchislenieZapisey'] = ""
                            nonSortedSlovarOfAUtors[authorsNames]['pravilnoeRaspolojenie'] = {}
                            nonSortedSlovarOfAUtors[authorsNames]['pravilnoeRaspolojenie'] = testingString
                        if len(authorsNames) > 1 :
                            if spiski == 'authors' :
                                nonSortedSlovarOfAUtors[authorsNames]['perchislenieZapisey'] = nonSortedSlovarOfAUtors[authorsNames]['perchislenieZapisey'] + ', '+str(numOfZap)
                                fileautorsUkazatel.write(authorsNames+' '+str(numOfZap) + ' ' + Zapis.docAuthorRazdelitel['prinstavka']+'\n')

                            else :
                                nonSortedSlovarOfAUtors[authorsNames]['perchislenieZapisey'] = nonSortedSlovarOfAUtors[authorsNames]['perchislenieZapisey'] + ', ' + str(numOfZap)+' ('+Zapis.docRedactorRazdelitel[spiski]['prinstavka']+')'

                                fileautorsUkazatel.write(authorsNames+' '+str(numOfZap)+' ('+Zapis.docRedactorRazdelitel[spiski]['prinstavka']+')'+'\n')




        else :
            print str(numOfZap)+" не определены авторы"
    else :
        print str(numOfZap)+" не определено поле авторов"

for i in nonSortedSlovarOfAUtors.keys():
    if listOfAuthors.count(nonSortedSlovarOfAUtors[i]['pravilnoeRaspolojenie']+nonSortedSlovarOfAUtors[i]['perchislenieZapisey']) == 0 :
        listOfAuthors.append(nonSortedSlovarOfAUtors[i]['pravilnoeRaspolojenie']+nonSortedSlovarOfAUtors[i]['perchislenieZapisey'])

listOfAuthors.sort()
for i in listOfAuthors:
    print i
    filelastautorsUkazatel.write(i+'\n')



if flagIsAllZapisRegonuzeBuAutorsField == True :
    for numOfZap in range(1, len(dataOfAllFindUniqKeysOfAll) + 1):
        pass
        # print numOfZap
        # print str(dataOfAllFindUniqKeysOfAll[numOfZap]["textAutors"])
else:
     print "NOT all authors is find"







#проверка на правильность заполнения для связанных ключей(ACess mode - url, ed by - content off    и т.д)

#проверка срабатывания ключа(отчистка словаря от не исп. ключей) на основе обработки массива
# for numOfZap in range(1,len(dataOfAllFindUniqKeysOfAll)) :
#     if dataOfAllFindUniqKeysOfAll[numOfZap].__len__() > 0 :
#         if dataOfAllFindUniqKeysOfAll[numOfZap].has_key("DocType") == True :
#             for typeOfZap in dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'].keys() :
#                 colOfKeyWords = 1
#                 if len(dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'][typeOfZap][0]) > 1 :
#                     colOfKeyWords=len(dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'][typeOfZap])
#                 #print str(numOfZap) + '. '
#                 for i in range(0,colOfKeyWords) :
#
#                     if isinstance(dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'][typeOfZap][i], list):
#                         FlagIsAllIclude = True;
#                         print "vcfvdfgdgd"
#                         for i in range(0, len(dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'][typeOfZap][i])):
#                             if testingString.find(dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'][typeOfZap][i][i].replace(' ', ' ')) < 0:
#                                 FlagIsAllIclude = False
#                         if FlagIsAllIclude == True:
#                             print str(intNumOfZapis) + '<-- в этой записи сработал список-ключ'
#                     else:
#                         if listOfKeyWords.count(dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'][typeOfZap][i]) > 0 :
#                             #print dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'][typeOfZap][i]
#                             listOfKeyWords.remove(dataOfAllFindUniqKeysOfAll[numOfZap]['DocType'][typeOfZap][i])
#
# if len(listOfKeyWords) > 0:
#     for i in listOfKeyWords:
#         fileProblems.write(' ----------найден не используемый ключ для определения типа док-та----------' + '\n')
#         fileProblems.write('---> '+str(i)+' <---'+'\n')
#





print '-----------------------------------------------------------------------------------------------------------'
print ''

fileProblems.close()
filedoctype.close()
fileautorsUkazatel.close()
filelastautorsUkazatel.close()
fileautors.close()
#document.save('test.docx')
#p.text = p.text.encode('utf-8').replace('Пат.', 'Патент').decode('utf8')


