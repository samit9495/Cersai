import threading
from os.path import join
import wx
from utility import *
# from send_mails import *
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException


def page_is_loaded(driver):
    return driver.find_element_by_tag_name("body") != None

class MahindraFinance(wx.Frame):
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, "CERSAI", size=(420, 410))
        self.panel = wx.Panel(self)
        try:
            image_file = join("Files", "util", 'sslogo2.jpeg')
            bmp1 = wx.Image(image_file, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            self.bitmap1 = wx.StaticBitmap(self.panel, -1, bmp1, (0, 0))
        except:
            pass
        self.idtext = wx.TextCtrl(self.panel,pos=(120,165),size=(200,20))
        self.passtext = wx.TextCtrl(self.panel,pos=(120,195),size=(200,20))
        self.nametext = wx.TextCtrl(self.panel,pos=(120,225),size=(200,20))
        self.label = wx.StaticText(self.panel, -1, "Press Start to Initiate Processing", (10, 350))
        self.idlabel = wx.StaticText(self.panel, -1, "Maker ID:", (10, 167))
        self.passlabel = wx.StaticText(self.panel, -1, "Password:", (10, 197))
        self.namelabel = wx.StaticText(self.panel, -1, "Machine Name:", (10, 227))
        self.start_button = wx.Button(self.panel, 1, label="START", pos=(80, 270), size=(90, 60))
        self.stop_button = wx.Button(self.panel, label="STOP", pos=(230, 270), size=(90, 60))
        self.Bind(wx.EVT_BUTTON, self.start_thread, self.start_button)
        self.Bind(wx.EVT_BUTTON, self.stop_f, self.stop_button)
        font = wx.SystemSettings.GetFont(wx.SYS_SYSTEM_FONT)
        # self.Bind(wx.EVT_TEXT, self.EvtText2)
        font.SetPointSize(20)
        self.start_button.SetBackgroundColour("#0B0B3B")
        self.start_button.SetFont(font)
        self.start_button.SetForegroundColour("#FAFAFA")
        self.stop_button.SetBackgroundColour("#0B0B3B")
        self.stop_button.SetFont(font)
        self.stop_button.SetForegroundColour("#FAFAFA")

    def stop_f(self, event):
        print("stopping")
        self.label.SetLabelText("Please wait ending process.")
        unproexcel = os.path.join(EXCEL_PATH, "unprocessed_{}.xlsx".format(get_date()))
        try:
            to_excel(unproexcel, self.record, "Stopping",  None, None, self.idtext.GetValue())
            self.driver.quit()
        except:
            pass
        # logger.info("Stopped")
        sys.exit()

    def start_thread(self, event):
        if self.idtext.GetValue() and self.nametext.GetValue() and self.passtext.GetValue():
            self.label.SetLabelText("Starting")
            self.t1 = threading.Thread(target=self.start_f)
            self.t1.setDaemon(True)
            self.t1.start()
        else:
            self.label.SetLabelText("Please Fill all the details to Continue.")

    def start_f(self):
        # if True:
        #     os.environ['http_proxy'] = "http://127.0.0.1:8888"
        #     os.environ['https_proxy'] = "http://127.0.0.1:8888"
        time1 = 0
        rec_no = 1
        self.record = get_record(QUEUE_TBP_PATH)
        self.recordsleft = records_left(QUEUE_TBP_PATH)
        self.driver = webdriver.Chrome(CHRM_DRVR)
        while len(self.record) >= 1:
            self.label.SetLabelText("Processing record {} of {}".format(rec_no, self.recordsleft))
            rec_no += 1
            tt1 = int(time.time())
            print("starting...")
            print(self.record)
            print("#############End")
            try:
                status, rid, code = start_process(self.record, self.driver,self.idtext.GetValue(),self.passtext.GetValue())
                print("rid = ", rid)
            except NoSuchElementException as ex:
                print(ex)
                rid = None
                status = ex
                code = 111
                self.driver.quit()
            except Exception as e:
                print(e)
                rid = None
                status = e
                code = 111
                self.driver.quit()
                if threading.active_count() <= 2 and int(time.time()) - time1 > 5000:
                    time1 = int(time.time())
                    try:
                        self.t2 = threading.Thread(target=mailer, args=(self.nametext.GetValue(),))
                        self.t2.setDaemon(True)
                        self.t2.start()
                    except Exception as e:
                        print(e)
            if str(code) != "200":
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = webdriver.Chrome(CHRM_DRVR)

            if "no such element:" in str(status) and "USERNAME" in str(status):
                print("repeating...")
                rec_no -= 1
                continue

            if rid == "ZERO RECORD":
                al = pyautogui.alert(
                    "No More Records to Process.\n\n\nPress OK to Abort.")
                self.record = []
                if al == "OK":
                    sys.exit()

            # filepath, data, status, mispath = None, rid = None, maker_id = None
            if rid and status == "Success":
                m_path = os.path.join(MIS_FILE, "MIS_{}.xlsx".format(get_date()))
                proexcel = os.path.join(EXCEL_PATH, "processed_{}.xlsx".format(get_date()))
                to_excel(proexcel, self.record, status, m_path, rid, self.idtext.GetValue())
            elif rid:
                unproexcel = os.path.join(EXCEL_PATH, "unprocessed_{}.xlsx".format(get_date()))
                to_excel(unproexcel, self.record, rid, None, None, self.idtext.GetValue())
            else:
                unproexcel = os.path.join(EXCEL_PATH, "unprocessed_{}.xlsx".format(get_date()))
                to_excel(unproexcel, self.record, status, None, None, self.idtext.GetValue())

            if "no such element:" in str(status) and "USERNAME" in str(status):
                self.record = self.record
            else:
                self.record = get_record(QUEUE_TBP_PATH)

            print("time taken in this record = ",int(time.time()) - tt1)

        if len(self.record) < 1:
            self.label.SetLabelText("Finished")
            self.label.SetLabelText("No more Records to Process.")
            pyautogui.alert("No more Records to Process", "Alert")
            self.driver.quit()


if __name__ == "__main__":
    EXCEL_PATH = os.path.join(os.getcwd(), "Files", "Excel",)
    EXCEL_TBP_PATH = os.path.join(os.getcwd(), "Files", "Excel_to_be_processed")
    QUEUE_TBP_PATH = os.path.join(os.getcwd(), "Files", "Queue_to_be_processed")
    CHRM_DRVR = os.path.join(os.getcwd(), "Files", "Chrome_driver", "chromedriver.exe")
    LOG_PATH = os.path.join(os.getcwd(), "Files", "Logs")
    MIS_FILE = os.path.join(os.getcwd(), "Files", "MIS")
    make_dir(EXCEL_PATH, QUEUE_TBP_PATH, EXCEL_TBP_PATH, MIS_FILE, LOG_PATH, CHRM_DRVR)
    # logger_name = "Maker_log_" + get_date() + ".log"
    #
    # logger = create_rotating_logger(join(LOG_PATH, logger_name))
    # logger.info("Starting Up...")

    app = wx.App()
    frame = MahindraFinance(parent=None, id=-1)
    frame.Show()
    app.MainLoop()
    # logger.info("Ending...")



