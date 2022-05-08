# pip install pandas
import sys, pulp, random
import pandas as pd
import numpy as np
from PyQt5 import *
from PyQt5.uic import loadUi
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import QCoreApplication


class MainWindow(QDialog):
    def __init__(self):
        super(MainWindow, self).__init__()
        loadUi(r"D:\Pascakampus\2021\[Final] GUI ISOP\GUI_ISOP.ui", self)
        
        #Price Data Box
        self.browse.clicked.connect(self.browsefiles)
        self.show_data.clicked.connect(self.loaddata)
        self.clear_data.clicked.connect(self.delete_data)
        
        #Misc Box
        self.update_misc.clicked.connect(self.update)
        self.clear_misc.clicked.connect(self.delete_misc)
        
        #Result Box
        self.calculate.clicked.connect(self.calculating)
        self.clear_result.clicked.connect(self.delete_result)
        self.exit.clicked.connect(QCoreApplication.instance().quit)
    

    #Fungsi Box Price Data

    def browsefiles(self):
        self.fname=QFileDialog.getOpenFileName(self,'Open file', r'C:','files(*.xls *.xlsx)')	
        self.file.setText(self.fname[0])

    def loaddata(self):
        self.data = pd.read_excel(self.fname[0],sheet_name=self.sheet.text(),nrows = int(self.rows.text()))
        self.m = [i for i in range(len(self.data.columns))]
        self.n = [j for j in range(len(self.data))]
        self.df = self.data.to_dict("records")
        row = 0
        self.preview.setRowCount(len(self.df))
        for price in self.df:
            self.preview.setItem(row,0,QtWidgets.QTableWidgetItem(str(price['x_1'])))
            self.preview.setItem(row,1,QtWidgets.QTableWidgetItem(str(price['x_2'])))
            self.preview.setItem(row,2,QtWidgets.QTableWidgetItem(str(price['x_3'])))
            self.preview.setItem(row,3,QtWidgets.QTableWidgetItem(str(price['x_4'])))
            self.preview.setItem(row,4,QtWidgets.QTableWidgetItem(str(price['x_5'])))
            row = row+1
     
    def delete_data(self):
        self.file.clear()
        self.sheet.clear()
        self.rows.clear()
        self.preview.clear()
    

    #Fungsi Box Misc

    def update(self):
        self.item = []
        self.disc = []
        self.min = []
        self.deliv = []
        
        self.item.append(int(self.q1.text()))
        self.item.append(int(self.q2.text()))
        self.item.append(int(self.q3.text()))
        self.item.append(int(self.q4.text()))
        self.item.append(int(self.q5.text()))
        
        self.disc.append(float(self.d1.text()))
        self.disc.append(float(self.d2.text()))
        self.disc.append(float(self.d3.text()))
        self.disc.append(float(self.d4.text()))
        self.disc.append(float(self.d5.text()))
        
        self.min.append(int(self.min_q1.text()))
        self.min.append(int(self.min_q2.text()))
        self.min.append(int(self.min_q3.text()))
        self.min.append(int(self.min_q4.text()))
        self.min.append(int(self.min_q5.text()))
        
        self.deliv.append(int(self.store1.text()))
        self.deliv.append(int(self.store2.text()))
        self.deliv.append(int(self.store3.text()))
        self.deliv.append(int(self.store4.text()))
        self.deliv.append(int(self.store5.text()))
        self.deliv.append(int(self.store6.text()))
        # print("update success")
  
    def delete_misc(self):
        self.q1.clear()
        self.q2.clear()
        self.q3.clear()
        self.q4.clear()
        self.q5.clear()
        self.d1.clear()
        self.d2.clear()
        self.d3.clear()
        self.d4.clear()
        self.d5.clear()
        self.min_q1.clear()
        self.min_q2.clear()
        self.min_q3.clear()
        self.min_q4.clear()
        self.min_q5.clear()
        self.store1.clear()
        self.store2.clear()
        self.store3.clear()
        self.store4.clear()
        self.store5.clear()
        self.store6.clear()
        

    #Fungsi Box Result
        
    def preproc(self):
        self.data = self.data.values.tolist()
        self.data = np.array(self.data)
        self.data = np.transpose(self.data)
        # print("preproc success")
    
    def disc_func(self):
        if sum(self.item)>200:
            self.discount = self.disc[4]
        elif sum(self.item)>100:
            self.discount = self.disc[3]
        elif sum(self.item)>50:
            self.discount = self.disc[2]
        elif sum(self.item)>25:
            self.discount = self.disc[1]
        elif sum(self.item)>0:
            self.discount = self.disc[0]
        # print(self.discount)
        # print("disc success")
    
    def det(self):
        model_det = pulp.LpProblem('ISOPwD',pulp.LpMinimize)
        x = pulp.LpVariable.dicts("x", ((i, j)for i in (self.m) for j in (self.n)),cat = "Binary")
        y = pulp.LpVariable.dicts("y",((j) for j in (self.n)),cat = "Binary")
        model_det += pulp.lpSum([self.data[i,j]*x[i,j]*self.item[i]*self.discount for i in (self.m) for j in (self.n)]) + pulp.lpSum([self.deliv[j]*y[j] for j in (self.n)])
        for i in (self.m) :
            model_det += pulp.lpSum(x[i,j] for j in (self.n)) == 1
        for i in self.m :
            for j in self.n :
                model_det += 0<=x[i,j]<=y[j]
        model_det.solve()
        print(pulp.LpStatus[model_det.status])
        self.xDet = [[],[],[],[],[]]
        self.yDet = []
        for i in self.m:
            for j in self.n:
                self.xDet[i].append(x[(i,j)].varValue)
        for j in self.n:
            self.yDet.append(y[j].varValue)
        self.zDet = pulp.value(model_det.objective)
        print("det success")
        
    def cetak_det(self):
        self.obj_det.setText(str(self.zDet))
        for i in self.m:
            for j in self.n:
                self.result_det.setItem(j,i,QtWidgets.QTableWidgetItem(str(self.xDet[i][j])))
        print(self.xDet)
        # print("cetak det success")
        
    def arc_func(self):
        d1 = []
        for i in range (0,5):
          a = random.uniform(1, self.deliv[i])
          d1.append(a)
        d1 = np.array(d1)
        print(d1)
        D1 = [[],[]]
        for i in range(0,2):
          for j in range(0,5):
            b = random.uniform(1, self.deliv[j])
            D1[i].append(b)
        D1 = np.array(D1)
        D1 = np.transpose(D1)
        print(D1)
        P = [2.2, 2.4, 2.2, 3, 2.4, 3]
        
        model_arc = pulp.LpProblem('ISOPwD',pulp.LpMinimize)
        x = pulp.LpVariable.dicts("x", ((i, j)for i in (self.m) for j in (self.n)),cat = "Binary")
        y = pulp.LpVariable.dicts("y",((j) for j in (self.n)),cat = "Binary")
        t = pulp.LpVariable("t", lowBound=0, upBound=None)
        Q = pulp.LpVariable.dicts("Q", ((j) for j in (self.n)))
        gamma = pulp.LpVariable.dicts("gamma", ((h) for h in (self.m)), lowBound = 0, upBound = None)
        
        model_arc += t
        model_arc += pulp.lpSum([self.data[i,j]*x[i,j]*self.item[i]*self.discount for i in (self.m) for j in (self.n)]) + pulp.lpSum([self.deliv[j]*y[j] for j in (self.n)]) + pulp.lpSum([d1[h]*gamma[h] for h in (self.m)]) <= t
        for z in range (0,2):
          model_arc += pulp.lpSum([D1[h,z]*gamma[h] for h in (self.m)]) == pulp.lpSum([self.deliv[j]*Q[j]+P[j]*y[j]+P[j]*Q[j] for j in (self.n)])
        for i in (self.m):
          model_arc += pulp.lpSum([x[i,j] for j in (self.n)]) == 1
        for i in self.m :
          for j in self.n :
            model_arc += 0<=x[i,j]<=y[j]
        model_arc.solve()
        print(pulp.LpStatus[model_arc.status])
        
        self.xArc = [[],[],[],[],[]]
        self.yArc = []
        self.QArc = []
        self.gammaArc = []
        self.tArc = t.varValue
        for i in self.m:
          self.gammaArc.append(gamma[i].varValue) 
          for j in self.n:
            self.xArc[i].append(x[(i,j)].varValue)
        for j in self.n:
          self.QArc.append(Q[j].varValue)
          self.yArc.append(y[j].varValue)
        print("arc success")
        
    def cetak_arc(self):
        self.obj_arc.setText(str(self.tArc))
        for i in self.m:
            for j in self.n:
                self.result_arc.setItem(j,i,QtWidgets.QTableWidgetItem(str(self.xArc[i][j])))
        print(self.xArc)
        # print("cetak arc success")
    
    def calculating(self):
        self.preproc()
        self.disc_func()
        self.det()
        self.cetak_det()
        self.arc_func()
        self.cetak_arc()
                
    def delete_result(self):
        self.obj_det.clear()
        self.obj_arc.clear()
        self.result_det.clear()
        self.result_arc.clear()
        

    
app = QApplication(sys.argv)
mainwindow = MainWindow()
widget = QStackedWidget()
widget.addWidget(mainwindow)
widget.setFixedHeight(750)
widget.setFixedWidth(1024)
widget.show()
sys.exit(app.exec_())