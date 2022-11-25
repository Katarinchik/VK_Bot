import vk_api
import conf
from openpyxl import load_workbook
from vk_api.longpoll import VkLongPoll, VkEventType
from vk_api import keyboard

vk_session = vk_api.VkApi(token = conf.settings['TOKEN'])
session_api = vk_session.get_api()
longpool = VkLongPoll(vk_session)
d = {}
number_of_mes =0;
fn1 = r"./stat.xlsx"
wb = load_workbook(fn1)
ws1 = wb["Лист1"]
def send_some_msg(id, some_text):
    vk_session.method("messages.send", {"user_id":id, "message":some_text,"random_id":0})
letters = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N',];
for event in longpool.listen():
    if event.type == VkEventType.MESSAGE_NEW:
        if event.to_me:
            msg = event.text.lower()
            id = event.user_id
            if str(id) in d:
                row = d[str(id)];
            else:
                d[str(id)]  =len(ws1['A'])+1
            r= d[str(id)];
            set = dict(one_time=False, inline=True)
            
            if(number_of_mes ==9):
                
                send_some_msg(id, "Спасибо за уделенное время!")
                
                popped_value =d.pop(str(id))
                cell_coord = letters[number_of_mes-1]+ str(r)
                print(cell_coord)
                ws1[cell_coord].value = msg
                number_of_mes = 0;
            else:
                print(number_of_mes)
                if(number_of_mes<6 and number_of_mes>1):
                    session_api.messages.send(  user_id=event.user_id, message=conf.questions[number_of_mes],keyboard = open("./keyboard2.json", encoding = 'ANSI', mode = 'r').read(),random_id=0 )
                if(number_of_mes==7):
                    session_api.messages.send(  user_id=event.user_id, message=conf.questions[number_of_mes],keyboard = open("./keyboard3.json", encoding = 'ANSI', mode = 'r').read(),random_id=0 )
                if(number_of_mes==0):
                    session_api.messages.send(  user_id=event.user_id, message=conf.questions[number_of_mes],keyboard = open("./keyboard1.json", "r", encoding="ANSI").read(),random_id=0 )
                if(number_of_mes==6 or number_of_mes== 1or number_of_mes== 8):
                    session_api.messages.send(  user_id=event.user_id, message=conf.questions[number_of_mes],random_id=0 )
                if(number_of_mes>0):  
                        cell_coord = letters[number_of_mes-1]+ str(r)
                        print(cell_coord)
                        ws1[cell_coord].value = msg
                number_of_mes =number_of_mes+1;
    wb.save('./stat.xlsx')

    


    

  