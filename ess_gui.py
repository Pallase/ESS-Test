import tkinter as tk
from tkinter import ttk
import queue
import threading
import sys
import time
from ess_visa import TEST_XM
from ess_xl import TEST_XL

class std_redirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        self.widget.config(state = "normal")
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)
        self.widget.config(state = "disabled")

class GUI(tk.Frame):
    def __init__(self, master, queue, set_port, start_test, end_test):
        super(GUI, self).__init__(master)
        self.queue = queue
        self.grid()
        
        ### TEXT OUTPUT WIDGET ###
        self.text = tk.Text(self, state = "disabled", width = 45, height = 30)
        self.text.grid(row = 1, column = 1, rowspan = 30)
        sys.stdout = std_redirector(self.text)
        ##########################

        ### OPEN PORT BUTTON ###
        self.port_status = tk.StringVar(self, value ="Open Port")
        self.port_button = ttk.Button(self, textvariable = self.port_status, command = lambda : set_port())
        self.port_button.grid(row = 1, column = 3, sticky = "W")
        ########################

        """### SERIAL NUMBER WIDGET ###
        self.serial = tk.StringVar(self, value = "1")

        self.serial_label = ttk.Label(self, text = "Serial Number: ")
        self.serial_label.grid(row = 2, column = 2, padx = 5)

        self.serial_entry = ttk.Entry(self, width = 25, textvariable = self.serial)
        self.serial_entry.grid(row = 2, column = 3)
        #############################"""

        ### ESS FILE WIDGET ###
        self.sheet_name = tk.StringVar(self, value ="ess")

        self.sheet_name_label = ttk.Label(self, text = "File name: ")
        self.sheet_name_label.grid(row = 3, column = 2, padx = 5, sticky = "W")

        self.sheet_name_entry = ttk.Entry(self, width = 25, textvariable = self.sheet_name)
        self.sheet_name_entry.grid(row = 3, column = 3)
        ##############################

        ### UNIT WIDGET ###
        self.unit_title = tk.StringVar(self, value ="PLO")

        self.unit_title_label = ttk.Label(self, text = "Chart Title: ")
        self.unit_title_label.grid(row = 4, column = 2, padx = 5, sticky = "W")

        self.unit_title_entry = ttk.Entry(self, width = 25, textvariable = self.unit_title)
        self.unit_title_entry.grid(row = 4, column = 3)

        ### VOLT WIDGET ###
        self.volt_axis = tk.StringVar(self, value ="LD Voltage")

        self.volt_axis_label = ttk.Label(self, text = "Voltage Detect: ")
        self.volt_axis_label.grid(row = 5, column = 2, padx = 5, sticky = "W")

        self.volt_axis_entry = ttk.Entry(self, width = 25, textvariable = self.volt_axis)
        self.volt_axis_entry.grid(row = 5, column = 3)
        ##############################

        ### START TEST WIDGET ###
        self.start_test = tk.Button(self, text = 'Start Test', command = lambda : start_test())
        self.start_test.grid(row = 7, column = 3, sticky = "E")
        #########################
        
        ### END TEST WIDGET ###
        self.end_test = tk.Button(self, text = 'End Test', command = lambda : end_test())
        self.end_test.grid(row = 8, column = 3, sticky = "E")
        #########################
        

    def processIncoming(self):
        """Handle all messages currently in the queue, if any."""
        while self.queue.qsize(  ):
            try:
                msg = self.queue.get(0)
                # Check contents of message and do whatever is needed. As a
                # simple test, print it (in real life, you would
                # suitably update the GUI's display in a richer fashion).
                print(msg)
            except queue.Empty:
                # just on general principles, although we don't
                # expect this branch to be taken in this case
                pass

class ThreadedClient:
    """
    Launch the main part of the GUI and the worker thread. periodicCall and
    endApplication could reside in the GUI part, but putting them here
    means that you have all the thread controls in a single place.
    """
    def __init__(self, master):
        """
        Start the GUI and the asynchronous threads. We are in the main
        (original) thread of the application, which will later be used by
        the GUI as well. We spawn a new thread for the worker (I/O).
        """
        self.master = master

        # Create the queue
        self.queue = queue.Queue(  )

        # Set up the GUI part
        self.gui = GUI(master, self.queue, self.get_port_status, self.startTest, self.endTest)

        self.ess = TEST_XM()

        self.xl = None

        # Set up the thread to do asynchronous I/O
        # More threads can also be created and used, if necessary
        self.running = 1

        # Start the periodic call in the GUI to check if the queue contains
        # anything
        self.periodicCall(  )

    def periodicCall(self):
        """
        Check every 200 ms if there is something new in the queue.
        """
        self.gui.processIncoming(  )
        """if not self.running:
            # This is the brutal stop of the system. You may want to do
            # some cleanup before actually shutting it down.
            sys.exit(1)"""
        self.master.after(200, self.periodicCall)

    def testing_thread(self):
        """
        This is where we handle the asynchronous I/O. For example, it may be
        a 'select(  )'. One important thing to remember is that the thread has
        to yield control pretty regularly, by select or otherwise.
        """

        if (self.ess.get_port_status() == True):
            self.ess.start_timer()
            self.initialize_xl()

            self.queue.put('Formatting xl sheet...')
            time.sleep(3)
            self.xl.format_print_margins()
            self.xl.page_header()
            self.queue.put('Formatting complete!\n')

            self.queue.put('Initializing ESS Test...')
            time.sleep(1)
            self.ess.set_mode()

            step_1 = 1
            step_2 = 15 * 60
            step_3 = 45 * 60
            step_4 = 25 * 60
            step_5 = 45 * 60
            step_6 = 15 * 60

            tot_cycle = step_1 + ((step_2 + step_3 + step_4 + step_5 + step_6) * 3)

            while True:
                if (self.running is False):
                    break
                else:
                    values = self.ess.measure()
                    voltage = abs(values[0])
                    if (voltage <= 0):
                        self.queue.put('fail')
                    temp = values[1]
                    time_elapsed = self.ess.get_time()
                    if (time_elapsed[1] > tot_cycle):
                        break
                    data = [time_elapsed[0], float(voltage), float(temp)]
                    self.queue.put([time_elapsed[2], float(voltage), float(temp)])
                    self.xl.write_xl(data)
                    time.sleep(1)

            end = self.ess.end_timer()

            self.queue.put('Testing duration = ' + str(end))
            try:
                self.save_graph()
            except:
                self.queue.put('No values?')
                raise SystemError
        else:
            self.queue.put('error: Is the port open?\n')

    def save_graph(self):
            self.queue.put('\nCreating graph...')
            self.xl.create_graph()
            time.sleep(3)
            self.queue.put('Created graph successfully!\n')
            time.sleep(1)
            self.queue.put('Saving file...')
            time.sleep(5)
            self.xl.graph_data()
            self.xl.save_xl()
            self.queue.put('Data saved!\n')

    def get_port_status(self):
        if (self.ess.get_port_status() == True):
            self.queue.put('Closing port...')
            self.ess.close_ports()
            time.sleep(1)
            self.queue.put('Port closed!\n')
            self.gui.port_status.set("Open Port")

        else:
            self.queue.put('Opening port...')
            self.ess.open_ports()
            time.sleep(1)
            self.queue.put('Port Opened!\n')
            self.gui.port_status.set("Close Port")

    def initialize_xl(self):
        file_name = self.gui.sheet_name.get() + '.xlsx'
        self.xl = TEST_XL(file_name, self.gui.unit_title.get(), self.gui.volt_axis.get())

        self.queue.put('Initializing xl sheet...')
        time.sleep(3)
        self.xl.set_wb()

    def init_xl(self):
        self.thread2 = threading.Thread(target=self.initialize_xl)
        self.thread2.start()

    def startTest(self):
        self.running = True
        self.thread1 = threading.Thread(target=self.testing_thread)
        self.thread1.start()


    def endTest(self):
        self.running = False
        
root = tk.Tk()
root.title("ESS TESTING GUI")
root.geometry("650x485")

client = ThreadedClient(root)
root.mainloop()
