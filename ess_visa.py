import visa
import time
from datetime import time as tm
import math

class TEST_XM:
    def __init__(self):
        self.rm = visa.ResourceManager()
        self.ports = None
        self.models = None
        self.start_time = None
        self.elapsed_time = None

    def get_rm(self):
        return self.rm

    def open_ports(self):
        self.ports = self.rm.list_resources()
        #mixer.init()
        #mixer.music.load('C:\Users\samscheer\Desktop\TEST\TENNEY\sound.mp3')
        #print(self.ports)
        #self.models = [self.rm.open_resource(port) for port in self.ports]
        try:
            self.models = self.rm.open_resource('USB0::0x05E6::0x2110::1374051::INSTR')
        except:
            print('error: usb?')
            self.models = None
            #self.models = self.rm.open_resource(self.ports[0])
        #return self.models

    def get_ports(self):
        return self.ports
        #return self.models

    def reset(self):
        self.models.write('*RST')

    def set_mode(self):
        self.models.write(':SENSe:VOLTage:DC:RANGe:AUTO %s' % ('ON'))
        self.models.write('FUNCtion2 "TCOuple"')

    def measure(self):
        temp_values = self.models.query_ascii_values(':READ?')
        readings = temp_values[0:]
        #if readings[0] <= 0:
            #mixer.music.play()
            #print('fail')
            #raise SystemExit #######################33
        return readings

    def close_ports(self):
        self.models.close()
        self.rm.close()

    def start_timer(self):
        self.start_time = time.time()

    def get_time(self):
        current_time = time.time()- self.start_time
        formatted_time = self.format_time(current_time)
        #return current_time
        sec = math.floor(current_time)
        time_string = self.string_time(current_time)
        time_list = [formatted_time, sec, time_string]
        return time_list

    def end_time(self):
        self.elapsed_time = time.time() - self.start_time
        formatted_time = self.format_time(math.floor(self.elapsed_time))
        #return self.elapsed_time
        #formatted_time = math.floor(self.elapsed_time)
        return formatted_time

    def format_time(self, input_time):
        formatted_time = tm(math.floor(input_time)//3600, (math.floor(input_time)%3600)//60, math.floor(input_time)%60) # h, m, s
        #formatted_time = time.strftime("%H:%M:%S", time.gmtime(input_time))
        return formatted_time

    def string_time(self, input_time):
        formatted_time = time.strftime("%H:%M:%S", time.gmtime(input_time))
        return formatted_time
