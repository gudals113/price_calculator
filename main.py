import sys
import os
from PyQt5.QtWidgets import *
from PyQt5 import uic
import openpyxl
import json

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

form = resource_path('untitled.ui')
form_class = uic.loadUiType(form)[0] 


#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.

#form_class = uic.loadUiType("untitled.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :

    def readExcel(self) :
        def getColChar(i):
            if i < 26:
                return chr(i+65)
            elif i < 26*2:
                return 'A'+chr(i-26+65)
            else:
                return 'B'+chr(i-26*2+65)

        self.frameData = {}
        self.doorData = {} # 분체도어단가
        self.designData = {}
        self.lammaData= {}
        
        try:
            filename="price.xlsx"
            book=openpyxl.load_workbook(filename)
            
            #후램 가격 불러오기
            doorsheet=book.worksheets[0]
            stop = False
            for j, row in enumerate(doorsheet.rows):
                if row[0].value is not None and all([len(item) == 4 and item[-1] == '바' for item in row[0].value.split('\n')]):
                    for bar in row[0].value.split('\n'):
                        if bar in self.frameData:
                            stop = True
                            break
                        self.frameData[bar.strip()] = []
                        for i, column in enumerate(row):
                            if i%2 == 1 and column.value is not None:
                                block = {'type':-1, 'price':[], 'exact': not '↓' in column.value}
                                cnt = 0
                                for item in column.value.replace('↓', '\n').split('\n'):
                                    if cnt == 0 and item.isdigit():
                                        block['w'] = int(item)
                                        cnt += 1
                                    elif cnt == 1 and item.isdigit():
                                        block['h'] = int(item)
                                        cnt += 1
                                    elif cnt == 2 and item.strip() == '양개':
                                        block['type'] = 1
                                    elif cnt == 2 and item.strip() == '편개':
                                        block['type'] = 0
                                block['price'].append(row[i+1].value)
                                block['price'].append(doorsheet[getColChar(i+1)+str(j+2)].value)
                                block['price'].append(doorsheet[getColChar(i+1)+str(j+3)].value)
                                if cnt != 0:
                                    if block['type'] == -1:
                                        block_a = json.loads(json.dumps(block))
                                        block_a['type'] = 0
                                        block_b = json.loads(json.dumps(block))
                                        block_b['type'] = 1
                                        self.frameData[bar].append(block_a)
                                        self.frameData[bar].append(block_b)
                                    else:
                                        self.frameData[bar].append(block)
                    if stop:
                        break
                    
            #람마용후램 가격 불러오기

            doorsheet=book.worksheets[1]
            for j, row in enumerate(doorsheet.rows):
                if row[0].value is not None and all([len(item) == 4 and item[-1] == '바' for item in row[0].value.split('\n')]):
                    for bar in row[0].value.split('\n'):
                        if bar in self.lammaData:
                            stop = True
                            break
                        self.lammaData[bar.strip()] = []
                        for i, column in enumerate(row):
                            if i%2 == 1 and column.value is not None:
                                block = {'type':-1, 'price':[], 'exact': not '↓' in column.value}   #exact와 type은 미사용됨
                                cnt = 0
                                for item in column.value.replace('↓', '\n').split('\n'):
                                    if cnt == 0 and item.isdigit():
                                        block['w'] = int(item)
                                        cnt += 1
                                    elif cnt == 1 and item.isdigit():
                                        block['h'] = int(item)
                                        cnt += 1
                                    #elif cnt == 2 and item.strip() == '양개':      엑셀 데이터에서 양개 편개 구분 안되어있으므로 주석 처리
                                    #    block['type'] = 1
                                    #elif cnt == 2 and item.strip() == '편개':
                                    #    block['type'] = 0
                                block['price'].append(row[i+1].value)
                                block['price'].append(doorsheet[getColChar(i+1)+str(j+2)].value)
                                block['price'].append(doorsheet[getColChar(i+1)+str(j+3)].value)
                                if cnt != 0:
                                    self.lammaData[bar].append(block)
                    if stop:
                        break




            #디자인도어 가격 불러오기
            doorsheet=book.worksheets[3]
            for i, row in enumerate(doorsheet.rows):

                if row[0].value is not None and ('DS' in row[0].value or 'A' in row[0].value or '디자인' in row[0].value):     
                    doorNum=row[0].value.strip('A').strip('DS-').strip()
                    self.designData[doorNum]=[]
                    block={'price':[], 'moredata':'false'}       #block에 옵션과 관련된거 넣으면 되겠다.방화핀같은 고정가격 말고 뭐랄까 ...! 규격...?
                    block['price'].append(doorsheet["B"+str(i+1)].value)
                    self.designData[doorNum].append(block)


            #분체도어 가격 불러오기
            doorsheet=book.worksheets[2]
            for j, row in enumerate(doorsheet.rows):
                if row[0].value is not None and ('양개도어' in row[0].value or '편개도어' in row[0].value):  
                    t_size = row[0].value.strip()                                                              
                    self.doorData[t_size] = []                                                                         
                    for i, column in enumerate(row):                                                               
                        if i == 0: continue                                                                         
                        if doorsheet[getColChar(i)+str(j+2)].value is None: continue                                    
                        block = {'price':[], 'exact': not '↓' in doorsheet[getColChar(i)+str(j+2)].value}
                        cnt = 0
                        for item in doorsheet[getColChar(i)+str(j+2)].value.replace('↓', '*').split('*'):        
                            if cnt == 0 and item.isdigit():
                                block['w'] = int(item)
                                cnt += 1
                            elif cnt == 1 and item.isdigit():
                                block['h'] = int(item)
                                cnt += 1                                                                              
                        block['price'].append(doorsheet[getColChar(i)+str(j+3)].value)
                        block['price'].append(doorsheet[getColChar(i)+str(j+4)].value)
                        block['price'].append(doorsheet[getColChar(i)+str(j+5)].value)
                        block['price'].append(doorsheet[getColChar(i)+str(j+6)].value)
                        if cnt != 0:
                            self.doorData[t_size].append(block)
 
        

        except Exception as ex:   
            print("excel")
            sys.exit()



    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.readExcel()
        self.door_type = [self.one_radio, self.two_radio]   #one=편개, two=양개
        self.door_type[0].setChecked(True)
        self.pushButton.clicked.connect(self.onCalcBtnClicked)
       

    def onCalcBtnClicked(self):
        door_type = 0
        for i, r in enumerate(self.door_type):
            if r.isChecked():
                door_type = i       #door type 0=편개, 1양개
        self.ResultTB.clear()
        try:
            #input_wideth=int(self.wideth.text())
            #input_height=int(self.height.text())                 #이게뭔가 여기 선언하고 안쓰면 오류가 생기네? 그냥 선언하지말고 밑에서 바로바로 사용하자
            total_price=0 


            #후렘 계산 
            if not self.lamma.isChecked() and self.frame_check.isChecked():
                hurem = self.frameData[self.bar.currentText()]
                min_price = 1e10
                target = None
                for item in hurem:
                    if item['w'] == int(self.wideth.text()) and item['h'] == int(self.height.text()) and item['exact'] and item['type'] == door_type:
                        target = item
                        break
                    if item['w'] >= int(self.wideth.text()) and item['h'] >= int(self.height.text()) and min_price > item['price'][0] and item['type'] == door_type and not item['exact']:
                        min_price = item['price'][0]
                        target = item
                if target is None:
                    self.ResultTB.addItem("후렘정보 입력 오류입니다.")
                    return
                self.ResultTB.addItem('후렘단가')
                self.ResultTB.addItem('{} {}*{}({}) \t {:8,}원'.format(
                    self.bar.currentText(),
                    target['w'],
                    target['h'],
                    '편개' if target['type'] == 0 else '양개',
                    target['price'][0]
                ))
                total_price += target['price'][0]

                if self.f_option1.isChecked():
                    self.ResultTB.addItem('도장비 \t\t {:8,}원'.format(
                        target['price'][1]
                    ))
                    total_price += target['price'][1]
                if self.f_option2.isChecked():
                    self.ResultTB.addItem('방화핀 1개  \t  5,000원')
                    total_price += 5000

                if self.f_option3.isChecked():
                    self.ResultTB.addItem('후램 그라스울 \t {:8,}원'.format(
                        target['price'][2]
                    ))
                    total_price += target['price'][2]
            self.ResultTB.addItem('----------------------------------------------')    
            #람마용후램 계산
            if self.lamma.isChecked() and self.frame_check.isChecked():
                hurem = self.lammaData[self.bar.currentText()]
                min_price = 1e10
                target = None
                for item in hurem:
                    if item['w'] >= int(self.wideth.text()) and item['h'] >= int(self.height.text()) and min_price > item['price'][0] :
                        min_price = item['price'][0]
                        target = item
                if target is None:
                    self.ResultTB.addItem("람마정보 입력 오류입니다.")
                    return
                self.ResultTB.addItem('람마후렘단가')
                self.ResultTB.addItem('{} {}*{} \t {:8,}원'.format(
                    self.bar.currentText(),
                    target['w'],
                    target['h'],
                    target['price'][0]
                ))
                total_price += target['price'][0]

                if self.f_option1.isChecked():
                    self.ResultTB.addItem('도장비 \t\t {:8,}원'.format(
                        target['price'][1]
                    ))
                    total_price += target['price'][1]
                if self.f_option2.isChecked():
                    self.ResultTB.addItem('방화핀 1개  \t  5,000원')
                    total_price += 5000

                if self.f_option3.isChecked():
                    self.ResultTB.addItem('후램 그라스울 \t {:8,}원'.format(
                        target['price'][2]
                    ))
                    total_price += target['price'][2]
            self.ResultTB.addItem('----------------------------------------------')    





            #디자인도어 계산
            if self.design_check.isChecked():
                if door_type ==0 :
                    target = self.designData[str(self.num_input.text())][0]
                    door_price=target['price'][0]
                    total_price += door_price
                if door_type ==1 :
                    target = self.designData[str(self.num_input.text())][0]
                    door_price=target['price'][0] *2
                    total_price += door_price
                self.ResultTB.addItem('디자인도어 단가')
                
                #디자인 도어 방화핀
                if self.design_fsd.isChecked():
                    if door_type ==0:
                        total_price=total_price + 25000
                        self.ResultTB.addItem('방화핀(편개) 25,000원')
                    if door_type==1:
                        total_price=total_price + 50000
                        self.ResultTB.addItem('방화핀(양개) 50,000원')
              

                #디자인 도어 그라스울 
                if self.design_glass.isChecked():
                    if door_type ==0:
                        total_price=total_price + 50000
                        self.ResultTB.addItem('그라스울(편개) 50,000원')
                    if door_type==1:
                        total_price=total_price + 100000
                        self.ResultTB.addItem('그라스울(양개) 100,000원')
            


                self.ResultTB.addItem('{}({}) \t {:8,}원'.format(
                    self.num_input.text(),
                    '편개' if door_type == 0 else '양개',
                    door_price
                ))
                self.ResultTB.addItem('----------------------------------------------')
                

      
            #분체도어 계산  
            if self.bun_check.isChecked():

                min_price = 1e10
                target = None
                if door_type == 0: # 편개 
                    door = self.doorData[self.depth_combo.currentText() +' 편개도어']
                elif door_type == 1: # 양개
                    door = self.doorData[self.depth_combo.currentText() +' 양개도어']
                for item in door:   #door=[{price,exact,w,h},{price,exact,w,h},,,,] 여기서 각각의 struct가 item
                    if item['w'] == int(self.wideth.text()) and item['h'] == int(self.height.text()) and item['exact']:
                        target = item
                        break
                    if item['w'] >= int(self.wideth.text()) and item['h'] >= int(self.height.text()) and min_price > item['price'][0] and not item['exact']:
                        min_price = item['price'][0]
                        target = item           #단가표의 데이터들이 루프 돌면서 규격에 알맞은 가격들 중 최소한의 가격을 맞춰가는거네
                if target is None:                    
                    return

                self.ResultTB.addItem('분체도어 단가')
                door_price=target['price'][self.fsd_combo.currentIndex()]
                self.ResultTB.addItem('{}\n{}'.format(
                    self.fsd_combo.currentText(),
                    self.depth_combo.currentText()
                ))
                
                self.ResultTB.addItem('{}*{}({}) \t {:8,}원'.format(
                    target['w'], target['h'],
                    '편개' if door_type == 0 else '양개',
                    door_price
                ))
                total_price+=door_price
            
            self.ResultTB.addItem('----------------------------------------------')
            self.ResultTB.addItem('{:8,}'.format(total_price))
      



        except Exception as ex:
            self.ResultTB.clear()
            self.ResultTB.addItem("정보 입력 오류입니다.")
            #sys.exit()
        


if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()