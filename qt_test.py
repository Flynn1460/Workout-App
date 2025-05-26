from PyQt5 import QtWidgets, QtCore
from design_temp import Ui_MainWindow
import sys, time
from datetime import datetime
from math import floor


class MyApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.TPS = 20
        
        self.ALLOCATED_REST = 105

        self.current_rest_start_time = 0
        self.WKO_ACTIVE = False

        self.prev_recorded_times = []

        self.UPDATE()

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_BracketLeft:
            self.clock.start()

            if (not self.WKO_ACTIVE):
                self.WKO_ACTIVE = True
                self.START_TIME = time.time()
            else:
                if ((not hasattr(self, 'END_TIME')) or (self.END_TIME < self.current_rest_start_time)): self.END_TIME = time.time()

                self.prev_recorded_times.append(round(self.END_TIME - self.current_rest_start_time, 2))
                self.UPDATE_avr_time()

            self.current_rest_start_time = time.time()
        

        if event.key() == QtCore.Qt.Key_BracketRight:
            self.END_TIME = time.time()
            self.clock.stop()


    def UPDATE(self):
        self.clock = QtCore.QTimer()
        self.raw_clock = QtCore.QTimer()

        self.raw_clock.timeout.connect(self.UPDATE_time)
        self.clock.timeout.connect(self.UPDATE_current_rest_time)
        self.clock.timeout.connect(self.UPDATE_deltatime)
        self.clock.timeout.connect(self.UPDATE_wko_time) 

        self.clock.start(round(1000/self.TPS))
        self.raw_clock.start(round(1000/self.TPS))


    def UPDATE_avr_time(self):
        if (self.WKO_ACTIVE):
            raw_secs = sum(self.prev_recorded_times) / len(self.prev_recorded_times)
        
            time_since = f"{(round(raw_secs) // 60):02}:{(round(raw_secs) % 60):02}.{round((raw_secs - floor(raw_secs))*100):02}"

            diff_from_ideal = abs(self.ALLOCATED_REST - raw_secs)

            if (diff_from_ideal <= 3): self.ui.avr_time.setStyleSheet("border: none; color: #78d95b;")
            if (diff_from_ideal >  3): self.ui.avr_time.setStyleSheet("border: none; color: #d9765b;")

            self.ui.avr_time.setText(str(time_since))

    def UPDATE_current_rest_time(self):
        if (self.WKO_ACTIVE):
            raw_secs = round(time.time() - self.current_rest_start_time)
            time_since = f"{(raw_secs // 60):02}:{(raw_secs % 60):02}"
            self.ui.current_rest_time.setText(str(time_since))

    def UPDATE_deltatime(self):
        if (self.WKO_ACTIVE):
            raw_secs = self.ALLOCATED_REST - round(time.time() - self.current_rest_start_time)
            
            if (raw_secs >  0): 
                self.ui.deltatime.setStyleSheet("border: none; color: #78d95b;")
                time_since = f"+{(raw_secs // 60):02}:{(raw_secs % 60):02}"
                
            if (raw_secs <= 0): 
                self.ui.deltatime.setStyleSheet("border: none; color: #d9765b;")
                time_since = f"-{(abs(raw_secs) // 60):02}:{(abs(raw_secs) % 60):02}"

            self.ui.deltatime.setText(str(time_since))

    def UPDATE_time(self):
        now = datetime.now().strftime("%H:%M:%S.%f")[:-4]
        self.ui.time.setText(now)

    def UPDATE_wko_time(self):
        if (self.WKO_ACTIVE):
            raw_secs = round(time.time() - self.START_TIME)
            time_since = f"{raw_secs // 60}min  {raw_secs % 60}sec"
            self.ui.wko_time.setText(str(time_since))



app = QtWidgets.QApplication(sys.argv)
window = MyApp()
window.show()
sys.exit(app.exec_())
