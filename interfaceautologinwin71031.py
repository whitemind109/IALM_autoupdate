#i1000 interface auto login 개발중
from winreg import QueryReflectionKey
from pywinauto.application import Application
from pywinauto import timings
import psutil, os, win32api, time, pyautogui

PI = "PIAgent.exe" #privacy i

INT = "Ui.Kumc.GR.Interface.exe"
INTF = (r"C:\Program Files\LIS_Interface\차세대")
INT1 = (r"C:\Program Files\LIS_Interface\차세대\Ui.Kumc.GR.Interface.exe")
INTT = "로그인"

BL = "BestLink.exe"
BLF = (r"C:\Beckman\Bestlink")
BL1 = (r"C:\Beckman\Bestlink\Autoupate_BestLink.exe")

QC = "StandardQC.exe"
QCF = (r"C:\Program Files (x86)\LabNDesign\StandardQC2021")
QC1 = (r"C:\Program Files (x86)\LabNDesign\StandardQC2021\LndAutoPatchCloud.exe") 
QC2 = (r"C:\Program Files (x86)\LabNDesign\StandardQC2021\StandardQC.exe")
QCT = "Standard QC v2021"

PHIS = "Nexmed_EHR.exe"
PHISF = (r"C:\Nexmed_EHR\PROD_GR\SYS\BIN")
PHIS1 = (r"C:\Nexmed_EHR\PROD_GR\SYS\BIN\Nexmed_EHR_PROD.exe")
PHIS2 = (r"C:\Nexmed_EHR\PROD_GR\SYS\BIN\Nexmed_EHR.exe")
PHIST = "Samsung Medical Information System, Login."

kumid = "12A27"
kumpw = "asdgqwet15"
qcpw = "1"

def TaskKill(prg):
    PROCNAME = prg
    for proc in psutil.process_iter():
        if proc.name() == PROCNAME:
            proc.kill()
            
def RunTask(run_f, run_prg1): 
    os.chdir(run_f)
    win32api.ShellExecute(0, 'open', run_prg1, '', '', 1)
        
def ConnectTask(con_prg):
    app = Application(backend="uia")
    if con_prg == PHIS:
        time.sleep(5) #업데이트 대비
        timings.wait_until_passes(60, 0.5, lambda: app.connect(path=PHIS2))
        dlg = app.window(title = PHIST)            
        pid_ctrl = dlg.Edit3
        ppw_ctrl = dlg.Edit2
        pid_ctrl.set_focus()
        pid_ctrl.type_keys(kumid)
        ppw_ctrl.set_focus()
        ppw_ctrl.type_keys(kumpw)
        pbtn_ctrl = dlg.Button2 #login
        pbtn_ctrl.click()
        time.sleep(3)
        pyautogui.hotkey('enter')
        
    elif con_prg == QC:        
        timings.wait_until_passes(60, 0.5, lambda: app.connect(path=QC2))
        dlg = app.window(title = QCT)         
        qid_ctrl = dlg.Edit2
        qpw_ctrl = dlg.Edit1
        qid_ctrl.set_focus()
        qid_ctrl.type_keys(kumid)
        qpw_ctrl.set_focus()
        qpw_ctrl.type_keys(qcpw)
        btn_ctrl = dlg.Button3
        btn_ctrl.click()
        time.sleep(3)
        pyautogui.hotkey('enter')
    
    elif con_prg == INT:
        app = Application(backend="win32")
        timings.wait_until_passes(60, 0.5, lambda: app.connect(path=INT1))
        dlg = app.window(title = INTT)         
        iid_ctrl = dlg.Edit2  # ID
        ipw_ctrl = dlg.Edit0  # pw
        ipw_ctrl.click()
        iid_ctrl.set_focus()
        iid_ctrl.type_keys(kumid, set_foreground=False)
        ipw_ctrl.click()
        ipw_ctrl.type_keys(kumpw, set_foreground=False)
        ipw_ctrl.type_keys("{ENTER}", set_foreground=False)        
    
    else :
        print ("error")

def StartTask(con_prg):
    app = Application(backend="uia")
    if con_prg == QC:
        app.start(QC1)    
        timings.wait_until_passes(60, 0.5, lambda: app.connect(title=QCT))
        dlg = app.window(title = QCT)            
        qid_ctrl = dlg.Edit2
        qpw_ctrl = dlg.Edit1
        qid_ctrl.set_focus()
        qid_ctrl.type_keys(kumid)
        qpw_ctrl.set_focus()
        qpw_ctrl.type_keys(qcpw)
        btn_ctrl = dlg.Button3
        btn_ctrl.click()
        time.sleep(3)
        pyautogui.hotkey('enter')
    
    elif con_prg == PHIS:
        app.start(PHIS1)   
        time.sleep(5) #업데이트 대비
        timings.wait_until_passes(60, 0.5, lambda: app.connect(title=PHIST))
        dlg = app.window(title = PHIST)            
        pid_ctrl = dlg.Edit3
        ppw_ctrl = dlg.Edit2
        pid_ctrl.set_focus()
        pid_ctrl.type_keys(kumid)
        ppw_ctrl.set_focus()
        ppw_ctrl.type_keys(kumpw)
        pbtn_ctrl = dlg.Button2 #login
        pbtn_ctrl.click()
        time.sleep(3)
        pyautogui.hotkey('enter')
        
    elif con_prg == INT:
        app = Application(backend="win32")
        app.start(INT1)    
        timings.wait_until_passes(60, 0.5, lambda: app.connect(title=INTT))
        dlg = app.window(title = INTT)       
        dlg.set_focus()     
        iid_ctrl = dlg.Edit2  # ID
        ipw_ctrl = dlg.Edit0  # pw
        ibtn_ctrl = dlg.RadioButton3 #login
        ibtn_ctrl.click()
        iid_ctrl.set_focus()
        iid_ctrl.type_keys(kumid, set_foreground=False)
        ipw_ctrl.click()
        ipw_ctrl.type_keys(kumpw, set_foreground=False)
        ipw_ctrl.type_keys("{ENTER}", set_foreground=False)
        
    elif con_prg == BL:
        app.start(BL1)    
        
    else :
        print ("error")

#프로그램 종료
TaskKill(PI) 
#TaskKill(BL)
#TaskKill(QC)
#TaskKill(PHIS)
TaskKill(INT)

#프로그램 실행
#RunTask(INTF, INT1) # 폴더, 실행프로그램1

#프로그램 연결, 조작
#ConnectTask(INT) # 프로그램

#프로그램 실행, 연결, 조작
#StartTask(BL) # BL은 실행만
#StartTask(QC)
#StartTask(PHIS)
StartTask(INT)



#conda 환경에서 설치 (project1)
#pyqt5 제거(용량 40MB정도 차이남)
#pyinstaller -w -F --uac-admin --exclude pandas, --exclude numpy, --exclude pillow interfaceautologinwin71031.py

